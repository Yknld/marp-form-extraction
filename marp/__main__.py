"""python -m marp --input forms --output results [--no-resume]."""
import argparse
import hashlib
import json
import sys
from collections import Counter
from importlib.metadata import version
from pathlib import Path
from .extract import ROOT, extract
from .export import workbook


def signature():
    paths=sorted((ROOT/"marp").glob("*.py"))+[ROOT/"questions.json"]+sorted((ROOT/"schemas").glob("*.json"))
    h=hashlib.sha256()
    h.update(sys.version.encode())
    for dependency in ["python-docx","pdfplumber","pypdf","openpyxl"]: h.update((dependency+version(dependency)).encode())
    for p in paths:h.update(p.name.encode());h.update(p.read_bytes())
    return h.hexdigest()


def atomic_json(path,data):
    temp=path.with_suffix(".tmp");temp.write_text(json.dumps(data,ensure_ascii=False),encoding="utf-8");temp.replace(path)


def ignored_input(path, base):
    return any(part.startswith((".", "~$")) or part == "__MACOSX" for part in path.relative_to(base).parts) or path.suffix.lower() in {".tmp", ".temp"} or path.name.lower() in {"thumbs.db", "desktop.ini"}


def unsupported_reason(path):
    suffix=path.suffix.lower()
    if suffix in {".doc", ".docm", ".rtf", ".odt"}: return "document_conversion_required"
    if suffix in {".jpg", ".jpeg", ".png", ".tif", ".tiff", ".bmp", ".heic"}: return "image_ocr_required"
    if suffix in {".zip", ".7z", ".rar"}: return "archive_unpack_required"
    return "unsupported_file_type"


def run(input_path, output, resume=True):
    input_path=Path(input_path).resolve();output=Path(output).resolve()
    if not input_path.exists(): raise ValueError("Input path does not exist")
    if input_path==output or output in input_path.parents: raise ValueError("Output must not equal or contain the input directory")
    output.mkdir(parents=True,exist_ok=True);cache=output/"checkpoints";cache.mkdir(exist_ok=True)
    base=input_path.parent if input_path.is_file() else input_path
    candidates=[input_path] if input_path.is_file() else sorted(p for p in input_path.rglob("*") if p.is_file() and output not in p.parents)
    ignored=[str(p.relative_to(base)) for p in candidates if ignored_input(p,base)]
    paths=[p for p in candidates if not ignored_input(p,base)]
    if not paths: raise ValueError("No input files found after excluding temporary and hidden files")
    submission_counts=Counter(p.stem.casefold() for p in paths)
    sig=signature(); results=[];hashes={};resumed=0
    for i,path in enumerate(paths,1):
        relative=str(path.relative_to(base));digest=""
        identity=hashlib.sha256(relative.encode()).hexdigest()[:16]
        saved=cache/(identity+".json")
        used_cache=False
        try:
            digest=hashlib.sha256(path.read_bytes()).hexdigest()
            r=None
            if resume and saved.exists():
                try:
                    old=json.loads(saved.read_text(encoding="utf-8"))
                    if old["source_hash"]==digest and old["signature"]==sig and old.get("source_root")==str(base):
                        r=old["result"];used_cache=True;resumed+=1
                except (ValueError,KeyError,OSError):pass
            if r is None:
                if path.suffix.lower() not in {".pdf",".docx"}:
                    r={"template_id":"unknown","language":"unknown","status":"unsupported_format","review_required":True,"review_reasons":[unsupported_reason(path)],"responses":[],"activities":[],"unmatched":[]}
                else:r=extract(path)
        except Exception as exc:
            r={"template_id":"unknown","language":"unknown","status":"error","review_required":True,"review_reasons":["file_read_failed"],"error":f"{type(exc).__name__}: {exc}","responses":[],"activities":[],"unmatched":[]}
        r.update(file_id=identity,submission_id=path.stem,relative_path=relative,source_hash=digest)
        # Recompute duplicate status for the current input set, including resumed files.
        r.pop("duplicate_of",None)
        r["review_reasons"]=[x for x in r.get("review_reasons",[]) if x not in {"duplicate_file_content","duplicate_submission_id"}]
        if submission_counts[path.stem.casefold()]>1:r["review_reasons"].append("duplicate_submission_id")
        r["review_required"]=bool(r["review_reasons"] or r["status"]!="extracted" or any(a["review_required"] for a in r.get("responses",[])+r.get("activities",[])))
        if digest and digest in hashes:
            r["duplicate_of"]=hashes[digest];r["review_reasons"].append("duplicate_file_content");r["review_required"]=True
        elif digest:hashes[digest]=identity
        atomic_json(saved,{"source_hash":digest,"source_root":str(base),"signature":sig,"result":r})
        results.append(r)
        print(f"[{i}/{len(paths)}] {'cached ' if used_cache else ''}{relative}: {r['status']}",flush=True)
    target=output/"results.jsonl";temp=output/"results.jsonl.tmp"
    temp.write_text("\n".join(json.dumps(r,ensure_ascii=False) for r in results)+"\n",encoding="utf-8");temp.replace(target)
    workbook(results,output/"extracted_responses.xlsx")
    summary={"files":len(results),"resumed":resumed,"errors":sum(r["status"]=="error" for r in results),"responses":sum(len(r["responses"]) for r in results),"activities":sum(len(r["activities"]) for r in results),"output":str(output)}
    summary.update(status_counts=dict(Counter(r["status"] for r in results)),ignored_files=ignored)
    atomic_json(output/"summary.json",summary)
    print(json.dumps(summary),flush=True)
    return results


def main():
    p=argparse.ArgumentParser(description="Extract local MARP Word/PDF forms into reviewable Excel and JSON.")
    p.add_argument("--input",required=True,type=Path);p.add_argument("--output",required=True,type=Path)
    p.add_argument("--no-resume",action="store_true")
    a=p.parse_args()
    try:run(a.input,a.output,not a.no_resume)
    except (ValueError,OSError) as exc:p.exit(2,f"Error: {exc}\n")


if __name__=="__main__":main()

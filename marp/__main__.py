"""python -m marp --input forms --output results ."""
import argparse
from collections import Counter
from pathlib import Path
from .extract import extract
from .export import workbook


def ignored_input(path, base):
    return any(part.startswith((".", "~$")) or part == "__MACOSX" for part in path.relative_to(base).parts) or path.suffix.lower() in {".tmp", ".temp"} or path.name.lower() in {"thumbs.db", "desktop.ini"}


def unsupported_reason(path):
    suffix=path.suffix.lower()
    if suffix in {".doc", ".docm", ".rtf", ".odt"}: return "document_conversion_required"
    if suffix in {".jpg", ".jpeg", ".png", ".tif", ".tiff", ".bmp", ".heic"}: return "image_ocr_required"
    if suffix in {".zip", ".7z", ".rar"}: return "archive_unpack_required"
    return "unsupported_file_type"


def run(input_path, output, reader=None):
    input_path=Path(input_path).resolve();output=Path(output).resolve()
    if not input_path.exists(): raise ValueError("Input path does not exist")
    if input_path==output or output in input_path.parents: raise ValueError("Output must not equal or contain the input directory")
    output.mkdir(parents=True,exist_ok=True)
    base=input_path.parent if input_path.is_file() else input_path
    candidates=[input_path] if input_path.is_file() else sorted(p for p in input_path.rglob("*") if p.is_file() and output not in p.parents)
    ignored=[str(p.relative_to(base)) for p in candidates if ignored_input(p,base)]
    paths=[p for p in candidates if not ignored_input(p,base)]
    if not paths: raise ValueError("No input files found after excluding temporary and hidden files")
    submission_counts=Counter(p.stem.casefold() for p in paths)
    results=[]
    for i,path in enumerate(paths,1):
        relative=str(path.relative_to(base))
        print(f"[{i}/{len(paths)}] Reading {relative}...",flush=True)
        try:
            if path.suffix.lower() not in {".pdf", ".docx"}:
                reason=unsupported_reason(path)
                status="skipped_ocr_required" if reason=="image_ocr_required" else "unsupported_format"
                r={"template_id":"unknown","language":"unknown","status":status,"review_required":True,"review_reasons":[reason],"responses":[],"activities":[],"unmatched":[]}
            else:
                r=extract(path, reader=reader)
        except Exception as exc:
            r={"template_id":"unknown","language":"unknown","status":"error","review_required":True,"review_reasons":["file_read_failed"],"error":type(exc).__name__,"responses":[],"activities":[],"unmatched":[]}
        r.update(submission_id=path.stem,relative_path=relative)
        if submission_counts[path.stem.casefold()]>1:r["review_reasons"].append("duplicate_submission_id")
        r["review_required"]=bool(r["review_reasons"] or r["status"]!="extracted" or any(a["review_required"] for a in r.get("responses",[])+r.get("activities",[])))
        results.append(r)
        workbook(results, output/"extracted_responses.xlsx")
        print(f"[{i}/{len(paths)}] {relative}: {r['status'].replace('_',' ')}",flush=True)
    for relative in ignored:
        results.append(dict(submission_id=Path(relative).stem,relative_path=relative,template_id="unknown",language="unknown",status="skipped",review_required=False,review_reasons=[],responses=[],activities=[],unmatched=[]))
    target=output/"extracted_responses.xlsx"
    workbook(results,target)
    print(f"Saved {target}. Processed {len(paths)} files; skipped {len(ignored)} temporary/hidden files.",flush=True)
    return results


def main():
    p=argparse.ArgumentParser(description="Extract local MARP Word/PDF forms into reviewable Excel.")
    p.add_argument("--input", required=True, type=Path)
    p.add_argument("--output", required=True, type=Path)
    a=p.parse_args()
    try:
        run(a.input,a.output)
    except ImportError as exc:
        p.exit(2,f"Required installed package unavailable: {exc.name}. Check the Python environment with IT; nothing was installed.\n")
    except PermissionError:
        p.exit(2,"Access denied. Close the results workbook and check folder permissions. The last completed save is retained.\n")
    except (ValueError,OSError) as exc:
        p.exit(2,f"Error: {exc}\n")
    except KeyboardInterrupt:
        p.exit(130,"Stopped. Completed files are saved in Excel; rerunning starts over.\n")


if __name__=="__main__":main()

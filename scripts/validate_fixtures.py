"""Validate fixture completeness and answer-key consistency, not extraction accuracy."""
import argparse
import hashlib
import json
from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET
from pypdf import PdfReader


def main():
    parser=argparse.ArgumentParser()
    parser.add_argument("directory",type=Path,nargs="?",default=Path(__file__).resolve().parents[1]/"fixtures/synthetic")
    root=parser.parse_args().directory
    records=[json.loads(line) for line in (root/"expected.jsonl").read_text().splitlines()]
    assert len({r["file"] for r in records})==len(records)
    actual={str(p.relative_to(root)) for p in root.rglob("*") if p.suffix in {".pdf",".docx"}}
    assert actual=={r["file"] for r in records}
    total_pages=0
    for r in records:
        p=root/r["file"]
        assert hashlib.sha256(p.read_bytes()).hexdigest()==r["sha256"],p
        assert len(r["expected"])==(33 if r["scope"]=="full_form" else 3),p
        for fid,e in r["expected"].items():
            assert e["review_required"]==bool(e["review_reasons"]),(p,fid)
            if e.get("presence")=="missing": assert e["normalized_answer"] is None
        if p.suffix==".docx":
            with ZipFile(p) as z:
                xml=ET.fromstring(z.read("word/document.xml"))
                if r["variant"]=="native_controls":
                    checks=xml.findall(".//{http://schemas.microsoft.com/office/word/2010/wordml}checkbox")
                    assert len(checks)==6
        else:
            pdf=PdfReader(p); assert len(pdf.pages)==r["pages"]>0; total_pages+=len(pdf.pages)
            text="".join(page.extract_text() or "" for page in pdf.pages)
            if p.name.startswith("scan_") and r["variant"]!="mixed_text_scan": assert not text.strip()
            else: assert "SYNTHETIC" in text.upper(),p
            if r["variant"]=="interactive":
                fields=pdf.get_fields(); assert len(fields)==6
                widgets=[a.get_object() for page in pdf.pages for a in page.get("/Annots",[]) if a.get_object().get("/Subtype")=="/Widget"]
                assert len(widgets)==6
                for w in widgets:
                    assert w["/V"]==fields[w["/T"]]["/V"]
                    assert w["/AP"]["/N"]
            if r["variant"]=="radio_buttons":
                fields=pdf.get_fields(); assert len(fields)==3
                for fid,e in r["expected"].items(): assert fields[fid]["/V"]=="/"+e["normalized_answer"]
    print(f"Validated {len(records)} files, {total_pages} PDF pages, hashes, fields, scan text layers, and interactive checkbox values.")


if __name__=="__main__": main()

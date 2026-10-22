"""Validate the Anti-Racism corpus and activity row ground truth."""
import argparse
import hashlib
import json
from pathlib import Path
from docx import Document
from pypdf import PdfReader

ROOT=Path(__file__).resolve().parents[1]


def main():
    parser=argparse.ArgumentParser()
    parser.add_argument("directory",nargs="?",type=Path,default=ROOT/"fixtures/anti_racism_call")
    out=parser.parse_args().directory
    schema=json.loads((ROOT/"schemas/anti_racism_call.json").read_text())
    ids={f["field_id"] for f in schema["questions"]}
    records=[json.loads(line) for line in (out/"expected.jsonl").read_text().splitlines()]
    assert len({r["file"] for r in records})==len(records)
    assert {str(p.relative_to(out)) for p in out.rglob("*") if p.suffix in {".pdf",".docx"}}=={r["file"] for r in records}
    pages=0
    for r in records:
        path=out/r["file"]; truth=r["expected"]
        assert hashlib.sha256(path.read_bytes()).hexdigest()==r["sha256"]
        assert r["template_id"]==schema["template_id"]
        assert set(truth["responses"])==ids
        for e in truth["responses"].values(): assert e["review_required"]==bool(e["review_reasons"])
        rows=truth["activities"]
        assert [a["activity_index"] for a in rows]==list(range(1,len(rows)+1))
        for a in rows:
            assert a["review_required"]==bool(a["review_reasons"])
            assert a["source_segments"]
            for field in ["description","completed","participant_count"]:
                e=a[field]; assert e["review_required"]==bool(e["review_reasons"])
        if path.suffix==".pdf":
            pdf=PdfReader(path); assert len(pdf.pages)==r["pages"]; pages+=len(pdf.pages)
            text=" ".join(" ".join((p.extract_text() or "").split()) for p in pdf.pages)
            assert "SYNTHETIC TEST FORM" in text
            for a in rows: assert f"Activity {a['activity_index']}:" in text
        else:
            doc=Document(path)
            order=["description","completed","participant_count"]
            if r["variant"]=="reordered_columns": order=["participant_count","description","completed"]
            for a in rows:
                seen={key:[] for key in order}
                for segment in a["source_segments"]:
                    cells=doc.tables[segment["table_index"]-1].rows[segment["row"]-1].cells
                    for c,key in enumerate(order):
                        if cells[c].text: seen[key].append(cells[c].text)
                for key in order:
                    assert " ".join(" ".join(seen[key]).split())==" ".join(a[key]["answer_text"].split()),(path,a["activity_index"],key)
            for row in truth["physical_activity_rows"]:
                if row["kind"]=="empty_spacer":
                    assert all(not c.text for c in doc.tables[row["table_index"]-1].rows[row["row"]-1].cells)
            for e in truth["responses"].values():
                loc=e["source_location"]
                observed=doc.tables[loc["table_index"]-1].rows[loc["row"]-1].cells[0].text
                value=e["answer_text"]
                assert observed==value or (value=="Yes" and observed=="Yes [X] / No [ ]") or (value=="No" and observed=="Yes [ ] / No [X]")
        if r["variant"]=="missing_counts":
            assert rows[0]["participant_count"]["normalized_answer"]==0
            assert rows[1]["participant_count"]["normalized_answer"] is None
        if r["variant"]=="split_activity_row": assert len(rows[1]["source_segments"])==2
    print(f"Validated {len(records)} files, {pages} PDF pages, 32 ordinary fields per form, and activity row content/continuations.")


if __name__=="__main__": main()

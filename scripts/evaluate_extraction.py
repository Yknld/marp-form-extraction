"""Compare a completed run against synthetic keys. Never imported by the extractor."""
import argparse
import json
from pathlib import Path
import sys
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from marp.extract import norm


def evaluate(fixtures, results_path):
    results={r["relative_path"]:r for r in map(json.loads,results_path.read_text().splitlines())}
    summary={"files":0,"supported_full_forms":0,"ocr_files":0,"unsupported_probes":0,"missing_files":[],
             "answers_correct":0,"answers_total":0,"selection_correct":0,"selection_total":0,
             "activity_cells_correct":0,"activity_cells_total":0,"activity_row_count_errors":0,
             "required_reviews_detected":0,"required_reviews_total":0,"additional_reviews":0,"clear_fields_total":0,"failures":[]}
    def review(expected,actual):
        if expected:
            summary["required_reviews_total"]+=1;summary["required_reviews_detected"]+=bool(actual)
        else:
            summary["clear_fields_total"]+=1;summary["additional_reviews"]+=bool(actual)
    for key in sorted(fixtures.rglob("expected.jsonl")):
        for record in map(json.loads,key.read_text().splitlines()):
            relative=str(key.parent.relative_to(fixtures)/record["file"]);summary["files"]+=1
            r=results.get(relative)
            if r is None:summary["missing_files"].append(relative);continue
            if record.get("scope")=="yes_no_probe":
                summary["unsupported_probes"]+=int(r["status"]=="unsupported_template");continue
            if record["file"].split("/")[-1].startswith("scan_"):
                summary["ocr_files"]+=1
                review(True,bool(r.get("ocr_pages")));continue
            summary["supported_full_forms"]+=1
            expected=record["expected"].get("responses",record["expected"])
            actual={a["field_id"]:a for a in r["responses"]}
            for fid,e in expected.items():
                a=actual.get(fid,{})
                is_selection="selected_options" in e or fid in {"ARC-C01","ARC-C02","ARC-D09","ARC-F01"}
                if is_selection:
                    correct=a.get("normalized_answer")==e["normalized_answer"] and bool(a)
                    summary["selection_total"]+=1;summary["selection_correct"]+=correct
                else:
                    correct=bool(a) and norm(a.get("answer_text") or "")==norm(e.get("answer_text") or "")
                    if e.get("presence")=="missing":correct=correct and a.get("presence")=="missing"
                summary["answers_total"]+=1;summary["answers_correct"]+=correct
                review(e["review_required"],a.get("review_required",False))
                if not correct:summary["failures"].append({"file":relative,"field":fid,"expected":e.get("normalized_answer") if is_selection else e.get("answer_text"),"actual":a.get("normalized_answer") if is_selection else a.get("answer_text"),"flagged":a.get("review_required",False)})
            if "activities" in record["expected"]:
                expected_rows=record["expected"]["activities"];actual_rows=r["activities"]
                summary["activity_row_count_errors"]+=len(expected_rows)!=len(actual_rows)
                for i,e in enumerate(expected_rows):
                    a=actual_rows[i] if i<len(actual_rows) else {}
                    for target,source in [("description","description"),("completed","completed"),("participant_count","participant_count")]:
                        ev=e[source]["answer_text"] if target=="description" else e[source]["normalized_answer"]
                        av=a.get(target)
                        correct=bool(a) and (norm(av or "")==norm(ev or "") if target=="description" else av==ev)
                        summary["activity_cells_total"]+=1;summary["activity_cells_correct"]+=correct
                        if not correct:summary["failures"].append({"file":relative,"activity":i+1,"field":target,"expected":ev,"actual":av,"flagged":a.get("review_required",False)})
                    review(e["review_required"],a.get("review_required",False))
    return summary


def main():
    p=argparse.ArgumentParser();p.add_argument("--fixtures",type=Path,default=Path(__file__).resolve().parents[1]/"fixtures")
    p.add_argument("--results",type=Path,required=True);p.add_argument("--report",type=Path,required=True)
    a=p.parse_args();report=evaluate(a.fixtures,a.results);a.report.parent.mkdir(parents=True,exist_ok=True)
    a.report.write_text(json.dumps(report,ensure_ascii=False,indent=2)+"\n")
    print(json.dumps({k:v for k,v in report.items() if k!="failures"},indent=2));print(f"Answer mismatches: {len(report['failures'])}")


if __name__=="__main__":main()

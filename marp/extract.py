"""Read layout-aware source blocks, match questions, and retain review evidence."""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from pathlib import Path
import json
import re
import unicodedata

ROOT = Path(__file__).resolve().parents[1]


def norm(text):
    return " ".join(re.findall(r"[^\W_]+", unicodedata.normalize("NFKC", text).casefold(), re.UNICODE))


def compact(text):
    return " ".join(text.split())


@dataclass
class Cell:
    text: str
    active: str = ""
    corrected: bool = False


@dataclass
class Block:
    cells: list[Cell]
    location: dict
    is_table: bool = False


def load_schemas():
    events = [{"field_id": r[0], "section": r[1], "question": r[2],
               "response_type": r[3], "condition": r[4]} for r in json.loads((ROOT / "questions.json").read_text())]
    arc = json.loads((ROOT / "schemas/anti_racism_call.json").read_text())["questions"]
    return {"marp_events": events, "marp_anti_racism_call": arc}


def read_docx(path):
    from docx import Document
    from docx.oxml.ns import qn
    doc = Document(path)
    blocks = []; table_index = 0; paragraph_index = 0

    def read_cell(el):
        texts=[]; active=[]; corrected=False
        for p in el.iter(qn("w:p")):
            full=[]; live=[]
            for r in p.iter(qn("w:r")):
                t="".join((n.text or "") if n.tag==qn("w:t") else "\n" if n.tag in {qn("w:br"),qn("w:cr")} else "\t" if n.tag==qn("w:tab") else "" for n in r)
                for symbol in r.iter(qn("w:sym")):
                    # Unmapped legacy symbol fonts must not be silently interpreted.
                    t += " [UNMAPPED_SYMBOL] "
                struck=any(s.get(qn("w:val"),"true").lower() not in {"0","false","off"}
                           for tag in ["strike","dstrike"] for s in r.findall("./"+qn("w:rPr")+"/"+qn("w:"+tag)))
                full.append(t)
                if struck: corrected=True
                else: live.append(t)
            texts.append("".join(full)); active.append("".join(live))
        return Cell("\n".join(texts).strip(), "\n".join(active).strip(), corrected)

    for child in doc.element.body:
        if child.tag == qn("w:tbl"):
            table_index+=1
            for row_index,row in enumerate(child.findall(qn("w:tr")),1):
                cells=[read_cell(c) for c in row.findall(qn("w:tc"))]
                blocks.append(Block(cells,{"table":table_index,"row":row_index},True))
        elif child.tag==qn("w:p"):
            paragraph_index+=1
            c=read_cell(child)
            if c.text: blocks.append(Block([c],{"paragraph":paragraph_index}))
    # Drawing text boxes and nested tables need dedicated reading-order handling.
    warnings=[]
    if list(doc.element.iter(qn("w:drawing"))) or list(doc.element.iter(qn("w:pict"))): warnings.append("embedded_images_require_review")
    if list(doc.element.iter(qn("w:txbxContent"))): warnings.append("text_boxes_require_review")
    if any(list(doc.element.iter(qn("w:"+tag))) for tag in ["ins","del"]): warnings.append("tracked_changes_require_review")
    if any(t.find(".//"+qn("w:tbl")) is not None for t in doc.element.body.findall(qn("w:tbl"))):
        warnings.append("nested_tables_require_review")
    return blocks, [], warnings


def read_pdf(path):
    """Summarize repeated font-metadata warnings as one file review reason."""
    import logging

    class FontWarningFilter(logging.Filter):
        seen = False

        def filter(self, record):
            if record.getMessage().startswith("Could not get FontBBox from font descriptor because"):
                self.seen = True
                return False
            return True

    logger = logging.getLogger("pdfminer.pdffont")
    capture = FontWarningFilter()
    logger.addFilter(capture)
    try:
        blocks, pages, warnings = _read_pdf(path)
    finally:
        logger.removeFilter(capture)
    if capture.seen:
        warnings.append("PDF font metadata missing or invalid; check extracted text against the original")
    return blocks, pages, warnings


def _read_pdf(path):
    import pdfplumber
    blocks=[]; ocr_pages=[]; warnings=[]
    with pdfplumber.open(path) as pdf:
        if not pdf.pages: raise ValueError("PDF has no pages")
        for number,page in enumerate(pdf.pages,1):
            text=page.extract_text() or ""
            image_area=sum(max(0,im["x1"]-im["x0"])*max(0,im["bottom"]-im["top"]) for im in page.images)
            if image_area/(page.width*page.height)>.5 and len(text.strip())<150:
                ocr_pages.append(number)
                continue
            for im in page.images:
                area=max(0,im["x1"]-im["x0"])*max(0,im["bottom"]-im["top"])
                if area/(page.width*page.height)>.12:
                    chars=sum(im["x0"]<=c["x0"]<=im["x1"] and im["top"]<=c["top"]<=im["bottom"] for c in page.chars)
                    if chars<20 and number not in ocr_pages: ocr_pages.append(number)
            found=page.find_tables()
            regions=[]
            for ti,table in enumerate(found,1):
                values=table.extract() or []
                for ri,(row,vals) in enumerate(zip(table.rows,values),1):
                    cells=[Cell(v or "",v or "") for v in vals]
                    if any(c.text for c in cells):
                        regions.append((row.bbox[1],row.bbox[0],Block(cells,{"page":number,"table_on_page":ti,"row":ri,"bbox":list(row.bbox)},True)))
            def outside(obj):
                if obj.get("object_type")!="char": return True
                x=(obj["x0"]+obj["x1"])/2; y=(obj["top"]+obj["bottom"])/2
                return not any(t.bbox[0]<=x<=t.bbox[2] and t.bbox[1]<=y<=t.bbox[3] for t in found)
            for line in page.filter(outside).extract_text_lines():
                # Ignore only margin page numbers; never scrub this phrase from an answer.
                if re.fullmatch(r"Page\s+\d+\s+(?:of|sur)\s+\d+",line["text"],re.I) and line["top"]>page.height*.85: continue
                regions.append((line["top"],line["x0"],Block([Cell(line["text"],line["text"])],{"page":number,"bbox":[line["x0"],line["top"],line["x1"],line["bottom"]]})))
            blocks.extend(r[2] for r in sorted(regions,key=lambda r:(r[0],r[1])))
        # Widget-only answers are not inferred from printed option labels.
        if any(page.annots for page in pdf.pages): warnings.append("pdf_annotations_or_fields_require_review")
    return blocks,ocr_pages,warnings


def choose_template(blocks, schemas):
    text=" ".join(c.text for b in blocks for c in b.cells)
    t=norm(text)
    if "anti racism call" in t: return "marp_anti_racism_call","en",[]
    french=any(s in t for s in ["nom de votre organisme","activités de l événement","résultats de l événement"])
    if french: return "unknown","fr",["unsupported_french_template"]
    if any(s in t for s in ["expected total number of participants","increase awareness of democratic institutions"]):
        return "unknown","en",["unsupported_legacy_events_variant"]
    unique=["has your event changed from what was originally planned","what helped make your event a success","total number of participants for the event"]
    if sum(s in t for s in unique)>=2 or "events component" in t: return "marp_events","en",[]
    return "unknown","unknown",["unknown_template"]


def yes_no(text, corrected=False):
    reasons=["corrected_selection"] if corrected else []
    s=compact(text).casefold()
    if s in {"yes","oui"}: return "Yes",reasons
    if s in {"no","non"}: return "No",reasons
    circled=re.findall(r"\(\s*(yes|no|oui|non)\s*\)",s)
    if len(circled)==1: return ("Yes" if circled[0] in {"yes","oui"} else "No"),reasons+["nonstandard_selection_mark"]
    token_re=r"\[\s*[xX✓✔]\s*\]|\[\s*\]|[☒☑✓✔☐]|\b(?:yes|no|oui|non|x)\b"
    tokens=re.findall(token_re,s,re.I)
    labels={"yes":"Yes","oui":"Yes","no":"No","non":"No"}
    checked=lambda token: token in {"☒","☑","✓","✔","x"} or bool(re.fullmatch(r"\[\s*[x✓✔]\s*\]",token,re.I))
    positions=[i for i,t in enumerate(tokens) if t in labels]
    if not positions: return None,reasons+["selection_unreadable"]
    selected=[]; observed={}
    prefix=positions[0]>0
    for k,i in enumerate(positions):
        between=tokens[(positions[k-1]+1 if k else 0):i] if prefix else tokens[i+1:(positions[k+1] if k+1<len(positions) else len(tokens))]
        label=labels[tokens[i]]
        observed.setdefault(label,[]).extend(between)
        if any(checked(v) for v in between): selected.append(label)
    if len(set(selected))>1: return None,reasons+["both_options_selected"]
    if len(positions)>2: return None,reasons+["multiple_selection_groups"]
    if len(selected)==1: return selected[0],reasons
    return None,reasons+["no_option_selected" if all(observed.values()) else "selection_symbol_missing"]


def normalize_value(text, typ, active=None, corrected=False):
    s=compact(text); reasons=[]
    if typ=="Yes / No": return yes_no(active if active is not None else text,corrected)
    if not s: return None,["blank_answer"]
    if s.casefold() in {"n/a","na","not applicable","sans objet"}: return None,[]
    if typ in {"Number","Currency"}:
        stripped=s.replace(",","").replace("$","").strip()
        if re.fullmatch(r"\d+(?:\.\d+)?",stripped):
            value=float(stripped) if "." in stripped else int(stripped)
            if typ=="Number" and value!=int(value): return value,["non_integer_count"]
            return value,[]
        match=re.fullmatch(r"(?:about|approximately|approx\.?)\s+(\d+)",s,re.I)
        if match: return int(match[1]),["approximate_count"]
        return None,["invalid_number"]
    if typ=="Date":
        try:
            if not re.fullmatch(r"\d{4}-\d{2}-\d{2}",s): return None,["ambiguous_date"]
            return date.fromisoformat(s).isoformat(),[]
        except ValueError: return None,["invalid_date"]
    return s,reasons


def match_question(text, fields):
    s=norm(text)
    if not s: return None
    found=[]
    for f in fields:
        q=norm(f["question"])
        # A header can span a PDF page; short exact labels still require a boundary.
        if s==q or s.startswith(q+" ") or (len(s.split())>=9 and q.startswith(s+" ")):
            found.append((min(len(s),len(q)),f))
    return max(found,key=lambda x:x[0])[1] if found else None


def remove_prompt(text, prompt):
    # Use normalized word matching for the prefix, while returning original answer text.
    tokens=list(re.finditer(r"[^\W_]+",text,re.UNICODE))
    q=norm(prompt).split(); words=[norm(t.group()) for t in tokens]
    matched=0
    for x,y in zip(words,q):
        if x!=y: break
        matched+=1
    if matched==len(q) or (matched>=9 and matched==len(words)):
        return text[tokens[matched-1].end():].lstrip(" \n\t:?).")
    return text


def extract(path):
    schemas=load_schemas()
    blocks,ocr_pages,warnings=read_docx(path) if path.suffix.lower()==".docx" else read_pdf(path)
    template,language,detected_warnings=choose_template(blocks,schemas)
    warnings+=detected_warnings
    if ocr_pages: warnings.append("ocr_required")
    result={"template_id":template,"language":language,"ocr_pages":ocr_pages,"responses":[],"activities":[],"unmatched":[],"review_reasons":warnings}
    if template=="unknown":
        result["status"]="ocr_required" if ocr_pages else "unsupported_template"
        return result
    fields=schemas[template]; pending={}; current=None; activity_mode=False; columns=None
    section_names={norm(f["section"]) for f in fields}
    if template=="marp_anti_racism_call":
        section_names.update(norm(s) for s in json.loads((ROOT/"schemas/anti_racism_call.json").read_text())["sections"].values())
        warnings.append("template_inventory_has_unresolved_source_labels")
    def append_response(fid,c,location):
        pending.setdefault(fid,{"chunks":[],"active":[],"locations":[],"corrected":False,"matches":0})
        p=pending[fid]; p["chunks"].append(c.text);p["active"].append(c.active);p["locations"].append(location);p["corrected"]|=c.corrected

    for b in blocks:
        texts=[c.text for c in b.cells]; joined=" ".join(texts); n=norm(joined)
        # Parse activity columns by their labels, independent of order.
        if template=="marp_anti_racism_call" and len(texts)==3:
            mapped={}
            for i,t in enumerate(texts):
                nt=norm(t)
                if "number of participants" in nt: mapped["participant_count"]=i
                elif nt=="completed": mapped["completed"]=i
                elif nt=="activity" or nt.startswith("activity list and describe"): mapped["description"]=i
            if len(mapped)==3:
                columns=mapped;activity_mode=True;current=None;continue
            if activity_mode and columns:
                d=b.cells[columns["description"]]; c=b.cells[columns["completed"]]; p=b.cells[columns["participant_count"]]
                if not any([d.text.strip(),c.text.strip(),p.text.strip()]): continue
                if d.text.strip() and not c.text.strip() and not p.text.strip() and result["activities"]:
                    previous=result["activities"][-1]
                    previous["description"] += "\n"+d.text
                    previous["source_locations"].append(b.location)
                    previous["review_reasons"].append("activity_row_continuation")
                    previous["review_required"]=True
                    continue
                completed,cr=normalize_value(c.text,"Yes / No",c.active,c.corrected)
                count,pr=normalize_value(p.text,"Number")
                reasons=cr+pr+([] if d.text.strip() else ["activity_description_blank"])
                result["activities"].append({"activity_index":len(result["activities"])+1,"description":d.text,
                    "completed_raw":c.text,"completed":completed,"participant_count_raw":p.text,"participant_count":count,
                    "source_locations":[b.location],"review_required":bool(reasons),"review_reasons":reasons})
                continue
        # Section boundaries stop accumulation, so boilerplate is not treated as an answer.
        partial_section=(not b.is_table and len(n.split())>=4 and any(s.startswith(n+" ") for s in section_names))
        if n in section_names or partial_section or re.fullmatch(r"[a-g] (?:recipient|event|project|events).*",n) and len(n.split())<8:
            current=None
            if "activities" not in n: activity_mode=False
            continue
        if not b.is_table:
            if n in {"activities continued","fictional data only"} or n.startswith("synthetic test") or n.startswith("fictional test data only"):
                continue
        matches=[match_question(t,fields) for t in texts]
        if any(matches):
            activity_mode=False
            for c,f in zip(b.cells,matches):
                if f:
                    current=f["field_id"]
                    append_response(current,c,b.location);pending[current]["matches"]+=1
                elif current and c.text: append_response(current,c,b.location)
        elif current:
            if any(c.text for c in b.cells):
                append_response(current,Cell(" ".join(c.text for c in b.cells if c.text)," ".join(c.active for c in b.cells if c.active),any(c.corrected for c in b.cells)),b.location)
        elif any(c.text for c in b.cells) and b.is_table:
            result["unmatched"].append({"text":joined,"source_location":b.location,"review_reason":"unmatched_table_content"})
    for f in fields:
        fid=f["field_id"]; p=pending.get(fid)
        if p:
            raw=remove_prompt("\n".join(p["chunks"]),f["question"])
            active=remove_prompt("\n".join(p["active"]),f["question"])
            value,reasons=normalize_value(raw,f["response_type"],active,p["corrected"])
            if p["matches"]>1: reasons.append("duplicate_question")
            locations=p["locations"]
            presence="present" if raw.strip() else "blank"
            if presence=="blank" and "optional" in (f.get("condition","")+f["question"]).casefold():
                reasons=[reason for reason in reasons if reason!="blank_answer"]
        else:
            raw=None;value=None;reasons=["question_missing"];locations=[];presence="missing"
        if f.get("wording_status")=="incomplete_in_source": reasons.append("source_question_incomplete")
        if ocr_pages: reasons.append("document_contains_unread_pages")
        result["responses"].append({"field_id":fid,"question":f["question"],"answer_text":raw,"normalized_answer":value,
             "presence":presence,"source_locations":locations,"review_required":bool(reasons),"review_reasons":reasons})
    indexed={r["field_id"]:r for r in result["responses"]}
    # A follow-up asserting a change contradicts an explicit No. Do not infer Yes from prose.
    if template=="marp_events" and indexed["C02"]["normalized_answer"]=="No":
        follow=norm(indexed["C03"]["answer_text"] or "")
        if follow and follow not in {"n a","na","not applicable","no changes","none"}:
            for fid in ["C02","C03"]:
                indexed[fid]["review_reasons"].append("contradictory_followup");indexed[fid]["review_required"]=True
    if result["unmatched"]: warnings.append("unmatched_table_content")
    result["review_required"]=bool(warnings or any(r["review_required"] for r in result["responses"]+result["activities"]))
    result["status"]="partial_ocr_required" if ocr_pages else "extracted"
    return result

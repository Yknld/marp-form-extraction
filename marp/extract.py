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
    ocb = json.loads((ROOT / "schemas/organizational_capacity.json").read_text())["questions"]
    return {"marp_events": events, "marp_anti_racism_call": arc, "marp_organizational_capacity": ocb}


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
                        regions.append((row.bbox[1],row.bbox[0],Block(cells,{"page":number,"table_on_page":ti,"row":ri,"bbox":list(row.bbox),"cell_bboxes":list(row.cells)},True)))
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
    if "organizational capacity building component" in t: return "marp_organizational_capacity","en",[]
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
    if typ=="Rating":
        if s in {"To a large extent","To some extent","To a little extent","To no extent","N/A"}: return s,[]
        return None,["rating_layout_unreadable"]
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



# Only these known printed instructions are discarded, never arbitrary "Note:" prose.
PRINTED_NOTES = [
    "Note: If applicable, you may submit copies of your bilingual resources or material you have developed for this event along with your final report.",
    "Note: The authorized signatory must be the same person, or hold the same authorization to sign official documents, as the person who signed the contribution agreement.",
    "Note: The authorized signatory must be the same person, or hold the same authorization to sign official documents, as the person who signed the application form.",
]


def wording_pattern(wording):
    return r"[\W_]*".join(re.escape(word) for word in norm(wording).split())


def clean_blocks(blocks, extra_notes=()):
    # Match full known notes across cell/page boundaries; preserve all other text
    # and each cell's original source location and direct-strikethrough state.
    for attr in ("text", "active"):
        cells = [cell for block in blocks for cell in block.cells]
        joined = "\n".join(getattr(c, attr) for c in cells)
        spans = []
        for note in [*PRINTED_NOTES, *extra_notes]:
            pattern = r"(?<!\w)\*?\s*" + wording_pattern(note) + r"[:.)]*"
            spans.extend((m.start(), m.end()) for m in re.finditer(pattern, joined, re.I))
        offset = 0
        for cell in cells:
            value = getattr(cell, attr)
            fragments = []
            cursor = 0
            for start, end in sorted(spans):
                lo, hi = max(0, start-offset), min(len(value), end-offset)
                if hi > lo:
                    fragments.append(value[cursor:lo]); cursor = max(cursor, hi)
            fragments.append(value[cursor:])
            setattr(cell, attr, "".join(fragments).strip())
            offset += len(value)+1
    return blocks


def question_segments(cell, fields):
    """Split only complete labels at line boundaries or in a label-led block."""
    hits = []
    starts_with_question = match_question(cell.text, fields) is not None
    for f in fields:
        for m in re.finditer(wording_pattern(f["question"]), cell.text, re.I):
            if m.start() and cell.text[m.start()-1].isalnum(): continue
            if m.end()<len(cell.text) and cell.text[m.end()].isalnum(): continue
            prefix = cell.text[:m.start()]
            line_prefix = prefix.rsplit("\n", 1)[-1]
            at_line_start = not line_prefix.strip(" \t*•-0123456789.)")
            # Inline labels are accepted only when this block already begins
            # with a recognized question and the label isn't quoted.
            quoted = prefix.rstrip().endswith(('"', "'", '“', '‘'))
            if not quoted and (at_line_start or starts_with_question):
                start=m.start()-len(line_prefix) if at_line_start else m.start()
                decoration=re.search(r"[ \t]*\*[ \t]*$", cell.text[:start])
                if decoration: start=decoration.start()
                hits.append((start, m.end(), f))
    accepted = []
    for hit in sorted(hits, key=lambda h: (h[0], -(h[1]-h[0]))):
        if not accepted or hit[0] >= accepted[-1][1]: accepted.append(hit)
    if not accepted: return [(match_question(cell.text, fields), cell)]
    segments = []
    if cell.text[:accepted[0][0]].strip():
        segments.append((None, Cell(cell.text[:accepted[0][0]], cell.active[:accepted[0][0]], cell.corrected)))
    # Match active text separately: strikethroughs can change character offsets.
    active_hits = []
    active_cursor = 0
    for _, _, f in accepted:
        match = re.search(wording_pattern(f["question"]), cell.active[active_cursor:], re.I)
        start = active_cursor+match.start() if match else active_cursor
        active_hits.append(start)
        if match: active_cursor += match.end()
    for i,(start,end,f) in enumerate(accepted):
        stop = accepted[i+1][0] if i+1<len(accepted) else len(cell.text)
        active_stop = active_hits[i+1] if i+1<len(accepted) else len(cell.active)
        segments.append((f, Cell(cell.text[start:stop].strip(), cell.active[active_hits[i]:active_stop].strip(), cell.corrected)))
    return segments


def assemble_headings(blocks, fields):
    """Join wrapped non-table headings only while they remain a known prefix."""
    out=[]; i=0
    while i<len(blocks):
        b=blocks[i]
        if not b.is_table and len(b.cells)==1:
            start=norm(b.cells[0].text)
            candidates=[f for f in fields if len(start.split())>=2 and norm(f["question"]).startswith(start+" ")]
            if candidates:
                end=i+1; collected=[b]
                while end<len(blocks) and end<i+15 and not blocks[end].is_table and len(blocks[end].cells)==1:
                    collected.append(blocks[end])
                    combined="\n".join(x.cells[0].text for x in collected)
                    n=norm(combined)
                    if any(n==norm(f["question"]) or n.startswith(norm(f["question"])+" ") for f in candidates):
                        location=dict(b.location);location["continued_locations"]=[x.location for x in collected[1:]]
                        b=Block([Cell(combined,"\n".join(x.cells[0].active for x in collected),any(x.cells[0].corrected for x in collected))],location)
                        i=end; break
                    if not any(norm(f["question"]).startswith(n+" ") for f in candidates): break
                    end+=1
        out.append(b);i+=1
    return out

def prepare_ocb_blocks(blocks, fields):
    # PDF table grids can include phantom cells where a row spans columns.
    compacted=[]
    for b in blocks:
        boxes=b.location.get("cell_bboxes")
        if boxes:
            keep=[i for i,box in enumerate(boxes) if box is not None]
            b.cells=[b.cells[i] for i in keep]
            b.location["cell_bboxes"]=[boxes[i] for i in keep]
        if (compacted and len(b.cells)==6 and b.cells[0].text.strip()
                and not any(c.text.strip() for c in b.cells[1:])
                and len(compacted[-1].cells)==6
                and any(match_question(compacted[-1].cells[0].text,[f]) for f in fields if f["response_type"]=="Rating")):
            previous=compacted[-1]
            for attr in ("text","active"):
                setattr(previous.cells[0],attr,getattr(previous.cells[0],attr)+"\n"+getattr(b.cells[0],attr))
            previous.location.setdefault("continued_locations",[]).append(b.location)
            continue
        compacted.append(b)
    # Restore known labels broken within words by narrow printed columns. Only
    # labels at cell/line starts qualify; respondent prose is not spell-corrected.
    for b in compacted:
        for c in b.cells:
            for attr in ("text","active"):
                value=getattr(c,attr)
                for f in sorted(fields,key=lambda f:-len(f["question"])):
                    words=norm(f["question"]).split()
                    pattern=r"[\W_]*".join(r"\n?".join(re.escape(ch) for ch in word) for word in words)
                    value=re.sub(r"(?im)^[ \t*]*"+pattern+r"(?!\w)",lambda m:f["question"],value)
                setattr(c,attr,value)
    return compacted


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
    boundary_fields=fields+[{"field_id":"__signoff", "question":"Name and signature (required)"}]
    ocb_schema=json.loads((ROOT/"schemas/organizational_capacity.json").read_text()) if template=="marp_organizational_capacity" else None
    if ocb_schema: blocks=prepare_ocb_blocks(blocks,fields)
    blocks=assemble_headings(clean_blocks(blocks,ocb_schema["printed_instructions"] if ocb_schema else ()),fields)
    rating_values={};rating_options=ocb_schema["rating_options"] if ocb_schema else []
    lanes={}; lane_table=None; lane_width=0; ambiguous_pending=False
    section_names={norm(f["section"]) for f in fields}
    section_names.update({"name and signature required", "name and signature"})
    if template=="marp_anti_racism_call":
        section_names.update(norm(s) for s in json.loads((ROOT/"schemas/anti_racism_call.json").read_text())["sections"].values())
        warnings.append("template_inventory_has_unresolved_source_labels")
    if ocb_schema: section_names.update(norm(value) for value in ocb_schema["sections"].values())
    def append_response(fid,c,location):
        pending.setdefault(fid,{"chunks":[],"active":[],"locations":[],"corrected":False,"matches":0})
        p=pending[fid]; p["chunks"].append(c.text);p["active"].append(c.active);p["locations"].append(location);p["corrected"]|=c.corrected

    for b in blocks:
        table_key=(b.location.get("page"),b.location.get("table_on_page",b.location.get("table"))) if b.is_table else None
        if table_key!=lane_table: lanes={};lane_table=table_key
        texts=[c.text for c in b.cells]; joined=" ".join(texts); n=norm(joined)
        # Parse activity columns by their labels, independent of order.
        if template in {"marp_anti_racism_call","marp_organizational_capacity"} and len(texts)==(4 if ocb_schema else 3):
            mapped={}
            for i,t in enumerate(texts):
                nt=norm(t)
                if "number of participants" in nt: mapped["participant_count"]=i
                elif nt=="completed": mapped["completed"]=i
                elif ocb_schema and nt=="results achieved": mapped["results_achieved"]=i
                elif nt=="activity" or nt.startswith("activity list and describe"): mapped["description"]=i
            if len(mapped)==(4 if ocb_schema else 3):
                columns=mapped;activity_mode=True;current=None;continue
            if activity_mode and columns:
                d=b.cells[columns["description"]]; c=b.cells[columns["completed"]]; p=b.cells[columns["participant_count"]]
                achieved=b.cells[columns["results_achieved"]].text if ocb_schema else ""
                if not any([d.text.strip(),c.text.strip(),p.text.strip(),achieved.strip()]): continue
                if ocb_schema and not any([d.text.strip(),p.text.strip(),achieved.strip()]) and yes_no(c.text)[0] is None and "no_option_selected" in yes_no(c.text)[1]: continue
                if d.text.strip() and not c.text.strip() and not p.text.strip() and result["activities"]:
                    previous=result["activities"][-1]
                    previous["description"] += "\n"+d.text
                    if ocb_schema: previous["results_achieved"] += ("\n" if previous["results_achieved"] else "")+achieved
                    previous["source_locations"].append(b.location)
                    previous["review_reasons"].append("activity_row_continuation")
                    previous["review_required"]=True
                    continue
                completed,cr=normalize_value(c.text,"Yes / No",c.active,c.corrected)
                count,pr=normalize_value(p.text,"Number")
                reasons=cr+pr+([] if d.text.strip() else ["activity_description_blank"])
                result["activities"].append({"activity_index":len(result["activities"])+1,"description":d.text,
                    "completed_raw":c.text,"completed":completed,"participant_count_raw":p.text,"participant_count":count,"results_achieved":achieved,
                    "source_locations":[b.location],"review_required":bool(reasons),"review_reasons":reasons})
                continue
        if ocb_schema:
            # Rating grids need column labels, not a concatenation of tick marks.
            if n.startswith(norm(ocb_schema["rating_heading"])):
                if len(texts)==6 and set(norm(t) for t in texts[1:])==set(norm(t) for t in ocb_schema["rating_options"]):
                    rating_options=[next(v for v in ocb_schema["rating_options"] if norm(v)==norm(t)) for t in texts[1:]]
                current=None;lanes={};activity_mode=False;continue
            parts=question_segments(b.cells[0],fields) if b.cells else []
            first=next((f for f,c in parts if f),None)
            if first and first["response_type"]=="Rating" and len(texts)==6:
                selected=[];invalid=False
                for label,c in zip(rating_options,b.cells[1:]):
                    value=c.active.strip()
                    if re.fullmatch(r"(?:[☒☑✓✔xX]|\[\s*[xX✓✔]\s*\])",value): selected.append(label)
                    elif value and not re.fullmatch(r"(?:☐|\[\s*\])",value): invalid=True
                reasons=[]
                if first["field_id"] in rating_values: reasons.append("duplicate_rating_question")
                if invalid: reasons.append("rating_mark_unreadable")
                if len(selected)!=1: reasons.append("multiple_rating_options_selected" if selected else "no_rating_option_selected")
                if any(c.corrected for c in b.cells[1:]):reasons.append("corrected_selection")
                answer=selected[0] if len(selected)==1 and not invalid else "Unresolved: "+"; ".join(f"{label}: {c.text}" for label,c in zip(rating_options,b.cells[1:]))
                rating_values[first["field_id"]]=(answer,selected[0] if len(selected)==1 and not invalid else None,reasons,b.location)
                current=None;lanes={};activity_mode=False
                for f,c in parts:
                    if f and f["response_type"]!="Rating":
                        current=f["field_id"];append_response(current,c,b.location);pending[current]["matches"]+=1
                continue
            # The Yes/No cell belongs to the first question, not its explanatory follow-up.
            if first and first["response_type"]=="Yes / No" and len(texts)==2:
                current=None;lanes={};activity_mode=False
                for f,c in parts:
                    if f:
                        current=f["field_id"];append_response(current,c,b.location);pending[current]["matches"]+=1
                append_response(first["field_id"],b.cells[1],b.location)
                continue
        # Section boundaries stop accumulation, so boilerplate is not treated as an answer.
        partial_section=(not b.is_table and len(n.split())>=4 and any(s.startswith(n+" ") for s in section_names))
        if n in section_names or partial_section or re.fullmatch(r"[a-g] (?:recipient|event|project|events).*",n) and len(n.split())<8:
            current=None;lanes={}
            if "activities" not in n: activity_mode=False
            continue
        if not b.is_table:
            if n in {"activities continued","fictional data only"} or n.startswith("synthetic test") or n.startswith("fictional test data only"):
                continue
        # Each label/answer segment is assigned independently. Column lanes
        # retain side-by-side header-to-answer relationships on following rows.
        segmented=[question_segments(c,boundary_fields) for c in b.cells]
        has_labels=any(f for parts in segmented for f,_ in parts)
        if has_labels: ambiguous_pending=False
        combined_headers=(len(b.cells)==1 and len(segmented[0])>1 and
                          all(f and not remove_prompt(c.text,f["question"]).strip() for f,c in segmented[0]))
        header_lanes={i:parts[-1][0]["field_id"] for i,parts in enumerate(segmented)
                      if parts and parts[-1][0] and len(parts)==1
                      and not remove_prompt(parts[-1][1].text,parts[-1][0]["question"]).strip()}
        if len(header_lanes)>=2:
            lanes=header_lanes;lane_width=len(b.cells)
        elif any(f for parts in segmented for f,_ in parts):
            lanes={}
        if lanes and not has_labels and len(b.cells)!=lane_width:
            lanes={};current=None;ambiguous_pending=True
        for ci,parts in enumerate(segmented):
            location=dict(b.location)
            if b.is_table:
                location["column"]=ci+1
                boxes=location.pop("cell_bboxes",None)
                if boxes and ci<len(boxes) and boxes[ci]: location["bbox"]=list(boxes[ci])
            for f,c in parts:
                if not c.text.strip(): continue
                # Section headings inside a cell are boundaries too.
                if norm(c.text) in section_names:
                    current=None;lanes={};continue
                if f and f["field_id"]=="__signoff":
                    current=None;lanes={}
                    trailing=remove_prompt(c.text,f["question"])
                    if trailing.strip():
                        result["unmatched"].append({"text":trailing,"source_location":location,"review_reason":"ambiguous_signoff_answer"})
                    continue
                if f:
                    activity_mode=False
                    current=f["field_id"]
                    append_response(current,c,location);pending[current]["matches"]+=1
                elif ambiguous_pending:
                    result["unmatched"].append({"text":c.text,"source_location":location,"review_reason":"ambiguous_question_boundary"})
                elif lanes and ci in lanes:
                    append_response(lanes[ci],c,location)
                elif lanes:
                    result["unmatched"].append({"text":c.text,"source_location":location,"review_reason":"ambiguous_question_boundary"})
                elif current:
                    append_response(current,c,location)
                elif b.is_table:
                    result["unmatched"].append({"text":c.text,"source_location":location,"review_reason":"unmatched_table_content"})
        if combined_headers:
            current=None;ambiguous_pending=True
        if lanes: current=None
    for f in fields:
        if f["response_type"]=="Signature": continue
        fid=f["field_id"]; p=pending.get(fid)
        if fid in rating_values:
            raw,value,reasons,location=rating_values[fid];locations=[location];presence="present"
        elif p:
            raw=remove_prompt("\n".join(p["chunks"]),f["question"])
            active=remove_prompt("\n".join(p["active"]),f["question"])
            if ocb_schema and f["response_type"]=="Currency":
                raw=re.sub(r"^[\s:$]*","",raw);active=re.sub(r"^[\s:$]*","",active)
            if ocb_schema and re.fullmatch(r"[\s()$]*",raw): raw="";active=""
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

"""Excel-only output with submission IDs and readable review flags."""
from pathlib import Path
import math
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE


def render_value(value):
    if value is None: return ""
    if isinstance(value,dict): value=", ".join(f"{k.replace(chr(95), chr(32))}: {v}" for k,v in value.items())
    if isinstance(value,list): value="; ".join(str(render_value(v)) for v in value)
    if isinstance(value,str): return ILLEGAL_CHARACTERS_RE.sub(lambda m: "\\u%04x" % ord(m.group()),value)
    return value


def workbook(results, target):
    wb=Workbook(); wb.remove(wb.active)
    files=[];responses=[];activities=[];unmatched=[]
    for r in results:
        common=[r["submission_id"],r["relative_path"]]
        reasons=list(r.get("review_reasons",[])); any_review=r.get("review_required",True)
        for a in r.get("responses",[]):
            ar=list(a["review_reasons"])
            if len(a.get("answer_text") or "")>32767:
                ar.append("long_answer_continued_on_Long_text_sheet"); reasons.append("long_answer_continued_on_Long_text_sheet");any_review=True
            responses.append(common+[a["field_id"],a["question"],a["answer_text"],a["normalized_answer"],a["presence"],a["source_locations"],"Yes" if ar else "No","; ".join(x.replace("_"," ") for x in ar)])
        for a in r.get("activities",[]):
            ar=list(a["review_reasons"])
            if any(len(a.get(k) or "")>32767 for k in ["description","completed_raw","participant_count_raw"]):
                ar.append("long_answer_continued_on_Long_text_sheet"); reasons.append("long_answer_continued_on_Long_text_sheet");any_review=True
            activities.append(common+[a["activity_index"],a["description"],a["completed_raw"],a["completed"],a["participant_count_raw"],a["participant_count"],a["source_locations"],"Yes" if ar else "No","; ".join(x.replace("_"," ") for x in ar)])
        for a in r.get("unmatched",[]): unmatched.append(common+[a["text"],a["source_location"],"Yes",a["review_reason"]])
        files.append(common+[r["status"].replace("_"," ").capitalize(),len(r.get("responses",[])),len(r.get("activities",[])),"Yes" if any_review else "No","; ".join(x.replace("_"," ") for x in dict.fromkeys(reasons)),r.get("error","")])

    long_text=[]
    def add(name,headers,rows,widths):
        for row in rows:
            for i,value in enumerate(row):
                value=render_value(value)
                if isinstance(value,str) and len(value)>30000:
                    for part,start in enumerate(range(0,len(value),30000),1):
                        long_text.append([row[0],row[1],name,headers[i],len(long_text)+1,part,value[start:start+30000],"Yes","Long text continued from another sheet"])
                    row[i]=value[:29000]+"\n[Full text in Long text sheet]"
                    row[headers.index("Flagged") ]="Yes"
                    row[headers.index("Reason")]+="; Full text in Long text sheet"
        ws=wb.create_sheet(name);ws.sheet_view.showGridLines=False;ws.freeze_panes="D2"
        ws.append(headers)
        for row in rows:
            ws.append([render_value(v) for v in row])
            for cell in ws[ws.max_row]:
                if isinstance(cell.value,str):
                    # Untrusted responses and filenames must never become Excel formulas.
                    cell.data_type="s"
        for cell in ws[1]: cell.font=Font(name="Arial",size=11,bold=True,color="FFFFFF");cell.fill=PatternFill("solid",fgColor="243C59");cell.alignment=Alignment(wrap_text=True,vertical="center")
        ws.row_dimensions[1].height=32
        for i,w in enumerate(widths,1):ws.column_dimensions[get_column_letter(i)].width=w
        for row in ws.iter_rows(min_row=2):
            for c in row:c.font=Font(name="Arial",size=11);c.alignment=Alignment(vertical="top",wrap_text=True)
            lines=max(sum(max(1,math.ceil(len(line)/(width*.9))) for line in str(c.value or "").split("\n")) for c,width in zip(row,widths))
            ws.row_dimensions[row[0].row].height=min(409,max(30,lines*15+10))
        if rows:
            tab=Table(displayName=name.replace(" ","")+"Table",ref=f"A1:{get_column_letter(len(headers))}{len(rows)+1}")
            tab.tableStyleInfo=TableStyleInfo(name="TableStyleMedium2",showRowStripes=True);ws.add_table(tab)
            rc=get_column_letter(headers.index("Flagged")+1)
            dv=DataValidation(type="list",formula1='"Yes,No"');dv.errorTitle="Select Yes or No";dv.error="Use the dropdown.";dv.showErrorMessage=True;ws.add_data_validation(dv);dv.add(f"{rc}2:{rc}{len(rows)+1}")
            ws.conditional_formatting.add(f"{rc}2:{rc}{len(rows)+1}",FormulaRule(formula=[f'{rc}2="Yes"'],fill=PatternFill("solid",fgColor="FFF0C2")))
        ws.auto_filter.ref=ws.dimensions
    common=["Submission ID","Source file"]
    add("Files",common+["Status","Response count","Activity count","Flagged","Reason","Error"],files,[20,44,26,17,17,14,65,55])
    add("Responses",common+["Question number","Question","Answer","Interpreted answer","Answer status","Location","Flagged","Reason"],responses,[20,44,20,70,80,40,15,50,14,55])
    add("Activities",common+["Activity number","Description","Completed original","Completed","Participant count original","Participant count","Location","Flagged","Reason"],activities,[20,44,16,80,30,15,28,22,50,14,55])
    add("Unmatched",common+["Original text","Location","Flagged","Reason"],unmatched,[20,44,90,50,14,55])
    if long_text:
        add("Long text",common+["Sheet","Field","Sequence","Part","Text","Flagged","Reason"],long_text,[20,44,20,30,12,12,100,14,50])
    target=Path(target);tmp=target.with_suffix(".tmp.xlsx");wb.save(tmp);tmp.replace(target)

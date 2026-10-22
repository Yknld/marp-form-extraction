"""One worksheet: one source file per row, one question per column."""
from pathlib import Path
from collections import Counter
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from .extract import load_schemas


def text(value):
    if value is None:
        return ""
    return ILLEGAL_CHARACTERS_RE.sub(lambda m: "\\u%04x" % ord(m.group()), str(value))


def readable(reason):
    return reason.replace("_", " ")


def workbook(results, target):
    schemas = load_schemas()
    families = [name for name in schemas if any(r["template_id"] == name for r in results)]
    if not families:
        families = list(schemas)
    fields = [(family, q) for family in families for q in schemas[family]]
    counts = Counter(q["question"].casefold() for _, q in fields)
    labels = {"marp_events": "Events", "marp_anti_racism_call": "Anti-Racism"}
    columns = []
    for family, q in fields:
        label = q["question"]
        if counts[label.casefold()] > 1:
            label += f" [{labels[family]} — {q['field_id']}]"
        columns.append(((family, q["field_id"]), label))
    # Preserve each activity's correspondence across the three cells.
    activity_columns = [("description", "Activities — description"),
                        ("completed_raw", "Activities — completed?"),
                        ("participant_count_raw", "Activities — participant count")]
    has_activities = any(r.get("activities") for r in results)
    headers = ["Submission ID"] + [label for _, label in columns]
    if has_activities:
        headers += [label for _, label in activity_columns]
    headers += ["Flagged", "Reason", "Source file", "Status", "Unmatched text"]
    rows = []
    for result in results:
        answers = {a["field_id"]: a for a in result.get("responses", [])}
        reasons = [readable(x) for x in result.get("review_reasons", [])]
        row = [text(result["submission_id"])]
        for (family, fid), label in columns:
            a = answers.get(fid) if result["template_id"] == family else None
            row.append(text(a["answer_text"]) if a else "")
            if a:
                reasons.extend(f"{label}: {readable(x)}" for x in a.get("review_reasons", []))
        if has_activities:
            for key, _ in activity_columns:
                row.append("\n\n".join(f"Activity {a['activity_index']}: {text(a.get(key))}" for a in result.get("activities", [])))
        for a in result.get("activities", []):
            reasons.extend(f"Activity {a['activity_index']}: {readable(x)}" for x in a.get("review_reasons", []))
        for a in result.get("unmatched", []):
            reasons.append("Unmatched text: " + readable(a["review_reason"]))
        if result.get("error"):
            reasons.append(result["error"])
        flagged = result.get("review_required", False) or bool(reasons)
        row += ["Yes" if flagged else "No", "\n".join(dict.fromkeys(reasons)),
                text(result["relative_path"]), readable(result["status"]).capitalize(),
                "\n\n".join(a["text"] for a in result.get("unmatched", []))]
        rows.append([text(v) for v in row])
    # Excel cells hold at most 32,767 characters. Continue across columns,
    # keeping exactly one sheet and one row per file, without losing text.
    expanded_headers = []
    expanded_rows = [[] for _ in rows]
    for col, header in enumerate(headers):
        parts = max([1] + [(len(row[col]) + 29999) // 30000 for row in rows])
        for part in range(parts):
            expanded_headers.append(header if part == 0 else f"{header} (continued {part + 1})")
            for source, dest in zip(rows, expanded_rows):
                dest.append(source[col][part * 30000:(part + 1) * 30000])
    wb = Workbook()
    ws = wb.active
    ws.title = "Responses"
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "B2"
    ws.append(expanded_headers)
    for row in expanded_rows:
        ws.append(row)
    for row in ws:
        for cell in row:
            cell.data_type = "s"  # Preserve IDs and prevent response formulas.
            cell.font = Font(name="Arial", size=11)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
    for cell in ws[1]:
        cell.font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="243C59")
    ws.row_dimensions[1].height = 120
    for i, header in enumerate(expanded_headers, 1):
        width = 48
        if header == "Submission ID": width = 20
        elif header == "Flagged": width = 12
        elif header == "Status": width = 24
        ws.column_dimensions[get_column_letter(i)].width = width
    for i in range(2, ws.max_row + 1):
        ws.row_dimensions[i].height = 90
        if i % 2 == 0:
            for cell in ws[i]: cell.fill = PatternFill("solid", fgColor="EEF3F8")
    ws.auto_filter.ref = ws.dimensions
    if rows:
        flag = get_column_letter(expanded_headers.index("Flagged") + 1)
        validation = DataValidation(type="list", formula1='"Yes,No"')
        ws.add_data_validation(validation)
        validation.add(f"{flag}2:{flag}{ws.max_row}")
        ws.conditional_formatting.add(f"{flag}2:{flag}{ws.max_row}", FormulaRule(
            formula=[f'{flag}2="Yes"'], fill=PatternFill("solid", fgColor="FFF0C2")))
    target = Path(target)
    tmp = target.with_suffix(".tmp.xlsx")
    wb.save(tmp)
    tmp.replace(target)

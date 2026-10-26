# Extract MARP forms into Excel

**Run `extract_forms.py`. No new packages or model downloads are required if the six existing dependencies are installed.** Keep the code folder together. This release uses `python-docx`, `pdfplumber`, `pypdf`, `openpyxl`, `pypdfium2`, and Pillow. `requirements.txt` records those dependencies; do not install or upgrade anything while your IT restriction is in effect.

Docling and AI/OCR model dependencies are not part of this release. The extraction script does not access the network, launch Office/Acrobat, run VBA, or move source files.

## Run

From the code folder, with Forms and Results beside it:

```powershell
python extract_forms.py --input "..\Forms" --output "..\Results"
```

From any other location, use the full script and folder paths:

```powershell
python "C:\Tools\marp-form-extraction\extract_forms.py" --input "C:\MARP\Forms" --output "C:\MARP\Results"
```

Use your actual paths. If your computer uses `py` instead of `python`, substitute `py`. Keep the code folder outside Forms. Forms can contain mixed file types and subfolders. Filenames such as `001234.pdf` become submission IDs; leading zeros remain text in Excel.

If necessary, this read-only check verifies that the existing packages can be imported by the selected Python. It installs nothing:

```powershell
python -c "import docx, pdfplumber, pypdf, openpyxl, pypdfium2, PIL; print('Existing packages available')"
```

If a package is missing, check which Python environment is selected with IT. Do not install additional packages as a workaround.

## Results

Open **Results → extracted_responses.xlsx**. Keep the workbook closed while running.

- Five sheets: **English Type 1** (Events), **English Type 2** (older Events), **French Type 1** (Events), **French Type 2** (older Events), and **Miscellaneous**.
- One row per source file, placed on exactly one sheet. Each survey sheet has only its own question columns; empty types still have a tab.
- French Type 2 is recognized but not yet supported for extraction: its rows remain flagged, with readable text in Unmatched text.
- Miscellaneous holds unknown forms, OCR skips, errors, unsupported files, and other recognized families (Anti-Racism and Capacity Building). Any extracted answers from those families are preserved there.
- **Submission ID** first, **Flagged** second.
- Each question is a column. ID and flag stay visible while scrolling.
- Recognized Yes/No selections display plain **Yes** or **No**. Ambiguous choices retain their original text and a review reason.
- Names and email addresses are separate fields. Authorized Signature is excluded; printed signatory name and date remain.
- Activities retain their correspondence through numbered entries within cells.
- Reason, Source file, Status, and Unmatched text appear at the end.

The workbook is saved after **every completed file**, and again at completion. Ctrl+C leaves the last completed save available. Rows accumulate on the matching sheets during a run. Each run starts over and replaces the workbook; this is not resumable processing. Save manual corrections in a separate copy. Duplicate IDs remain separate flagged rows.

No hashes, JSON result files, or separate diagnostic dumps are created. The JSON files in the code folder are internal question definitions. Very long answers continue in additional columns to avoid Excel's cell-length limit. Values are written as text, preserving IDs and preventing answers from becoming formulas.

## Files needing OCR

OCR is deliberately disabled in extraction.

- Scanned PDFs detected as unreadable: **Skipped OCR required**, Flagged = Yes, blank answers.
- PDFs containing both readable and unread scanned pages: the **whole file is skipped**, with unread page numbers in Reason. This prevents incomplete answers appearing complete.
- Image-only Word documents and standalone images: skipped for OCR, with the ID retained.
- Word files with both text and embedded images: readable text may be extracted, but embedded images require review.
- Old `.doc`, `.rtf`, `.odt`, and other unsupported formats: retained as flagged rows needing conversion.
- Corrupt/password-protected files: retained as error rows; other files continue.
- Temporary/hidden files and Word lock files: listed as skipped.

Nothing is deleted or moved. Filter Status to collect the OCR backlog later. The scan detector is a heuristic: an existing bad OCR text layer or an unusual image layout can still need manual review. Blank/unrecognized documents are never treated as verified answers.

## Questionnaire rules

The Python parser uses the supplied template definitions and selected wording/cleanup rules recovered from the legacy VBA scripts. It does not run the macros or use their hard-coded workbook columns.

| Questionnaire | Coverage |
| --- | --- |
| English Events | Existing supplied template, plus alternate contact labels, date labels, participant-count/reach wording and known instructions |
| Anti-Racism Call | Existing supplied template and repeated activity tables; incomplete source labels remain flagged |
| Organizational Capacity Building | Existing supplied template, activity tables, five ratings and separate follow-ups |
| Older English Events (form type 2 in the VBA scripts) | Separate provisional definition, including four Yes/No outcomes and four separate “Explain how” answers |
| French Events (form type 1 in the VBA scripts) | Provisional definition recovered from the French extraction and cleaning scripts |

Script-derived definitions are flagged for review. Some labels were reconstructed from script fragments and require confirmation against blank forms. The older French questionnaire remains unsupported. Unrecognized questionnaire text is retained in Unmatched text rather than assigned to guessed question columns. French official-language wording not established by the supplied scripts is also left for review.

Known printed instructions are removed by specific wording, including wrapped text. Genuine respondent notes are preserved; the parser does not delete every sentence starting with “Note.” Repeated questions, missing fields, malformed email addresses, ambiguous dates/numbers, multiple selections, and unclear table relationships generate review reasons.

The parser preserves table cells and recognizes question boundaries. Ambiguous text stays under Unmatched text where possible. It does not use the VBA fallback that grabbed the next 500 characters when an ending phrase was missing.

## Review limitations

Flagged = No means no automated issue was detected, not that a human verified the answer. Start with a small batch and compare it against originals before processing all submissions. Font metadata warnings are summarized in Reason. Drawing-based checkmarks, crossed-out PDF choices, interactive fields, unusual layouts, and previously unseen question wording can still need manual review.

Python reads Word directly: there is no need to convert readable `.docx` files to PDF. Converting a scan to PDF does not make its text readable.

## Optional questionnaire inventory

`discover_versions.py` groups candidate questionnaire versions in a separate report. OCR is off by default:

```powershell
python discover_versions.py --input "..\Forms" --output "..\VersionReview" --ocr off
```

This produces `questionnaire_versions.xlsx`, with IDs, candidate groups, and representative filenames. It does not redact documents or export response snippets. Grouping is provisional. Its earlier Tesseract option remains available only for environments where that software is already approved and installed; extraction never invokes it.

## Files to keep

Keep `extract_forms.py`, `marp/`, `questions.json`, and `schemas/` together. `discover_versions.py` is optional. `requirements.txt` is the dependency inventory. `tests/` contains small fictional regression tests and is not needed to run extraction.

Run the tests with `python -m unittest discover -s tests -v`. Development validation uses fictional forms; it does not establish accuracy on classified submissions. Keep real forms and results in your approved work environment, outside this public repository.

## Validation of this release

On the Mac, using the existing dependency stack: 12 regression tests passed; 1,921 of 1,924 ordinary-answer comparisons matched across the existing synthetic corpus, fictional capacity-building forms, and new script-derived template probes. The three mismatches were existing crossed-out PDF selections, retained as unresolved and flagged. All 70 reference activity rows matched. Six scan variants, including a mixed text/scan PDF, were skipped with no answers. A mixed nine-file batch retained corrupt and unsupported files as review rows. The new-template probes use generated labels, so they test parser mechanics rather than independently confirming the reconstructed wording. Windows runtime validation and a local pilot on actual forms are still needed.

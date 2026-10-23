# Extract forms into Excel

**Run `extract_forms.py`.** Keep the downloaded folder together: the script needs the supporting code and question definitions beside it.

**To find different questionnaire versions first, run `discover_versions.py` instead.** This is a separate inventory tool; it does not extract answers into the report or change the existing extraction workflow.

## Discover questionnaire versions

From inside the downloaded code folder, with Forms beside it:

```powershell
python discover_versions.py --input "..\Forms" --output "..\VersionReview"
```

Open **VersionReview → questionnaire_versions.xlsx**. The single sheet lists every file, its submission ID, candidate version group, and **Representative file to open**. Filter **Use this file as representative?** to **Yes** to find one suggested example per group. Files with insufficient evidence remain ungrouped and need manual inspection.

Open the representative files locally, make copies, and remove responses from those copies before sharing templates. The script does not redact or copy originals. Its report contains filenames and IDs, which may themselves be sensitive, but no extracted answer text or question snippets.

Discovery uses known question labels plus question-like wording from unfamiliar forms; it does not require the form to pass the existing template detector. It ignores font differences and uses a conservative similarity threshold. Grouping is provisional: missing questions or OCR errors may split one version into several groups, and very small revisions can still be grouped together. Similarity is a question-cue comparison, not an accuracy/confidence score. Candidate group numbers can change when the input set changes.

The report is saved every 25 inspected files and at completion. Ctrl+C saves the files completed so far. Keep the workbook closed while running. Each rerun starts over and replaces the report; these are progress snapshots, not resumable checkpoints. No JSON files are generated.

### Scans in version discovery

Discovery can use **local Tesseract OCR** when installed. It looks on PATH and in the usual Windows Tesseract installation folders. No cloud service or download is used by the script. Python dependencies alone do not install Tesseract or its language data. Install these through your approved work-machine process if needed.

If Tesseract is elsewhere, specify its executable:

```powershell
python discover_versions.py --input "..\Forms" --output "..\VersionReview" --tesseract "C:\Program Files\Tesseract-OCR\tesseract.exe"
```

The default OCR language is English. If English and French language data are installed, add `--ocr-language eng+fra`. To disable OCR, add `--ocr off`. Scanned PDFs and common image files are rendered/read locally; temporary images are removed afterward. OCR runs on PDFs with unread pages or too few question cues. Without OCR, unreadable scans appear as **Needs OCR**, not as invented questionnaire versions.

Old `.doc`, archives and other unsupported formats are listed for conversion/manual review. Image-only Word documents need manual conversion to PDF for OCR. Corrupt or encrypted files are reported without stopping the batch. Discovery OCR is for grouping only; it does not yet add scan answer extraction to `extract_forms.py`.

Optional tuning: `--similarity 0.96` is the default; a higher value separates more groups. At least eight question cues are required for grouping. Review representatives rather than treating candidate groups as confirmed templates.

Discovery prints a **Reason** for failed/unresolved files as well as writing it to Excel. `Read failed` means document reading or question analysis failed; `OCR failed` means the OCR attempt failed; `OCR unavailable` means the startup check found an OCR installation/language issue. Any native question cues already found are retained when OCR fails, but incomplete files remain ungrouped. Zero cues after a failure does not prove that the form contains no questions. Check the Reason before changing templates. Files stored in OneDrive must be available locally to the Python process.

## 1. Install once

Use Python 3.10 or newer in your approved work environment. In PowerShell, replace this example location with your downloaded folder:

```powershell
python -m pip install -r "C:\Tools\marp-form-extraction\requirements.txt"
```

If Windows recognizes `py` instead of `python`, use `py` in both commands.

## 2. Run

Put the forms together in a folder. PDFs and Word documents can be mixed, including subfolders. Keep filenames as submission IDs, such as `001234.pdf`.

```powershell
python "C:\Tools\marp-form-extraction\extract_forms.py" --input "C:\MARP\Forms" --output "C:\MARP\Results"
```

**This full-path command works from any current directory.** Use your actual paths. Keep the code folder outside your Forms folder.

## 3. Open your Excel results

Open `C:\MARP\Results\extracted_responses.xlsx`:

The workbook has **one sheet**, with **one row per input file**:

- **Submission ID** is the first column (from the filename, preserving leading zeros).
- **Flagged** is the second column, immediately after Submission ID.
- Each **question is a column heading**, with the extracted answer underneath.
- Repeated activities are numbered inside three cells: description, completed, participant count.
- **Flagged** and **Reason** identify files needing review and the affected questions.
- Source file, processing status and unmatched text appear at the end.

The first two columns and header row stay visible when scrolling. Filters let you show only flagged files. Headers and answers wrap; expand row heights or use the formula bar to read long text. Matching question labels from different form versions are qualified so their answers stay separate. Questions from other form versions remain blank. Duplicate IDs remain separate rows and are flagged.

No in Flagged means no automated issue was detected, not that someone verified the answer. Compare a pilot batch against the originals before processing all submissions.

The only generated result is **extracted_responses.xlsx**. There are no hash columns, JSON outputs or checkpoint files. Every run processes all inputs again and replaces the workbook: save manual review edits in a separate copy. Existing JSON/checkpoint files from older versions are not used or deleted. Duplicate filenames (submission IDs) are still flagged; identical content under different IDs is no longer detected.

The bundled JSON question definitions are internal settings, not output files; keep them with the script. You do not need to open or map them.

## What stays in the download

| File or folder | Purpose |
| --- | --- |
| `extract_forms.py` | The script you run |
| `discover_versions.py` | Find candidate questionnaire versions before preparing blank templates |
| `requirements.txt` | Dependencies to install once |
| `marp/` | Required supporting code |
| `questions.json` | Required Events question definitions |
| `schemas/` | Required Anti-Racism question definitions |
| `README.md` | These instructions |

Development fixtures, generators, tests and inventory workbooks were removed from the current download. They are recoverable in Git history.

## Limits

Supports the supplied English MARP Events Component, Anti-Racism Call, and Organizational Capacity Building forms in DOCX and selectable-text PDF. Runs locally without Office, Acrobat or a network connection after installation. Source files are read only.

Scans need OCR, which this script does not perform. Other formats appear in the sheet with a review reason: older Word/RTF/ODT files need conversion, images need OCR, and archives need unpacking. Unreadable/password-protected files are reported without stopping other files. Hidden/temporary files and Word lock files are listed as skipped rows. The selected output folder is excluded.

French and unknown templates, ambiguous choices and unsupported visual content require review. PDF drawing marks, crossed-out choices and interactive checkbox values are not reliably interpreted. Word images/text boxes may require manual review. Review flags cannot detect every mistake. Very long answers continue in additional columns on the same sheet, preserving the full text and one row per file. Rows may be taller than Excel can display; inspect the cell contents when necessary.

The Excel-only output has been checked for mixed formats, leading-zero IDs, duplicate-ID flags, long-text preservation and formula safety. The earlier synthetic benchmark matched 1,829 of 1,832 ordinary answers and all 70 activity rows; three unresolved selections were flagged. These development results do not establish accuracy on real submissions or throughput for 1,000+ files.

Keep real forms and extracted responses in your approved work environment, outside this public repository.

### PDF font warnings and unrecognized forms

Missing or invalid PDF font measurements (`FontBBox`) do not necessarily stop extraction. Repeated instances of this warning are summarized once in that file's Excel Reason column and set Flagged to Yes. Compare its extracted text against the original; this change does not repair the PDF. Other PDF diagnostics remain visible.

`Unsupported template` is separate: the text did not match a supported form version. The file keeps its row and submission ID, with a review reason, but its question cells are not filled. Consult Reason to distinguish an unknown template from a recognized unsupported French or legacy version.

### Answers without printed instructions

Authorized Signature is excluded from the output and does not generate missing-signature flags. Printed signatory name and sign-off date remain separate columns. The parser still recognizes signature headings to keep that text out of preceding answers.

Known printed bilingual-material and authorized-signatory notes are removed, including wrapped versions. Respondent notes and quoted question labels are preserved. The cleaner uses specific wording; it does not discard every line starting with Note. Other unrecognized instructions may still require review.

Multiple question labels in a cell and side-by-side table headings are handled separately. Where the following answers cannot be assigned confidently, their text is kept under Unmatched text and flagged instead of being guessed. The Excel-only command and output path are unchanged.

### Organizational Capacity Building version

The Organizational Capacity Building component is detected separately. Its 38 non-signature fields include project costs, activity-change/location follow-ups, five result ratings with separate explanations, and printed signatory name/date. Activity cells include description, completion, results achieved, and participant count.

Rating selections are mapped to their column labels (large/some/little/no extent or N/A). Multiple selections, unreadable marks, and unselected ratings are flagged. Empty activity rows with unselected boxes are ignored; checked rows without descriptions remain flagged. The default financial-report request and other printed instructions are not exported as answers.

Validation covers a fictional Word form and a bordered-table PDF conversion, including narrow/wrapped rating columns and page continuations. Borderless PDFs and materially changed layouts can still need review; OCR answer extraction remains outside the extraction script. Only template wording is published, not the supplied document or its remaining selections.

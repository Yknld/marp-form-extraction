# Extract forms into Excel

**Run `extract_forms.py`.** Keep the downloaded folder together: the script needs the supporting code and question definitions beside it.

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
| `requirements.txt` | Dependencies to install once |
| `marp/` | Required supporting code |
| `questions.json` | Required Events question definitions |
| `schemas/` | Required Anti-Racism question definitions |
| `README.md` | These instructions |

Development fixtures, generators, tests and inventory workbooks were removed from the current download. They are recoverable in Git history.

## Limits

Supports the supplied English MARP Events Component and Anti-Racism Call forms in DOCX and selectable-text PDF. Runs locally without Office, Acrobat or a network connection after installation. Source files are read only.

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

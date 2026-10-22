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

- **Files:** status and review flags for each source file.
- **Responses:** questions, original answers, interpreted values and source locations.
- **Activities:** one row per activity.
- **Unmatched:** table text that could not be assigned to a question.

Every sheet includes **Submission ID**, taken from the filename without its extension, stored as text to preserve leading zeros. Source filenames remain attached to answers. Repeated IDs are flagged and are not merged.

Start with **Review required** and **Review reason** in Files. No means no automated issue was detected, not that someone verified the answer. Compare a small pilot batch against the originals before processing all submissions.

Rerun the same command to resume; unchanged files reuse checkpoints. Add `--no-resume` to force reprocessing. The workbook is replaced each run: save manual review edits in a separate copy. Keep `checkpoints`, `results.jsonl` (complete extracted text) and `summary.json` in your results folder.

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

Scans need OCR, which this script does not perform. Other formats appear in Files with a review reason: older Word/RTF/ODT files need conversion, images need OCR, and archives need unpacking. Unreadable/password-protected files are reported without stopping other files. Hidden/temporary files and Word lock files are skipped and listed in `summary.json`. The selected output folder is excluded.

French and unknown templates, ambiguous choices and unsupported visual content require review. PDF drawing marks, crossed-out choices and interactive checkbox values are not reliably interpreted. Word images/text boxes may require manual review. Review flags cannot detect every mistake. Very long answers remain complete in JSON even when Excel cell/display limits are reached.

Before cleanup, 13 tests passed, including mixed formats, unreadable files, resume and ID preservation. The earlier synthetic benchmark matched 1,829 of 1,832 ordinary answers and all 70 activity rows; three unresolved selections were flagged. These development results do not establish accuracy on real submissions or throughput for 1,000+ files.

Keep real forms and extracted responses in your approved work environment, outside this public repository.

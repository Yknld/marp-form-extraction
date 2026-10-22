# MARP form extraction

Python extraction of Multiculturalism and Anti-Racism Program (MARP) Word documents and selectable-text PDFs, with source locations and review flags. The first version supports the supplied English Events Component and Anti-Racism Call inventories.

## Run

Use Python 3.10 or newer. Run these commands from this repository's root, installing dependencies in your approved environment:

```bash
python -m pip install -r requirements.txt
python -m marp --input /path/to/forms --output /path/to/results
```

Windows example:

```powershell
python -m marp --input "C:\Forms" --output "C:\Results"
```

The extractor runs locally. It does not require Microsoft Word, Excel, Adobe Acrobat, LibreOffice, an API key, or a network connection once dependencies are installed. LibreOffice/Poppler are only used by the synthetic-file generators.

The output contains `extracted_responses.xlsx`, full `results.jsonl`, a `summary.json`, and per-file `checkpoints`. The workbook has **Files**, **Responses**, **Activities**, and **Unmatched** sheets. Review flags and reasons are separate columns. Each activity has its own row; original text is retained alongside interpreted values.

Rerun the same command to resume. Completed unchanged files are loaded from checkpoints; input changes, code/schema changes, and dependency-version changes invalidate the relevant cache. Use `--no-resume` to force extraction. Individual file failures appear in the Files sheet and do not stop other files. Source files are read only. The output directory's generated workbook is replaced on each run, so save manual review edits in a separate copy.

See [scope and review guidance](docs/EXTRACTION.md) and the [synthetic benchmark](docs/BENCHMARK.md).

### Mixed input folders and submission IDs

Keep PDF and DOCX files together, including subfolders and uppercase extensions. The **Submission ID** column on every output sheet is the filename without its final extension, stored as text to preserve leading zeros (`001234.PDF` → `001234`). The original relative path and internal File ID remain available. Files with the same submission ID (case-insensitive) are all flagged `duplicate_submission_id`; their responses remain separate pending review.

All other file types appear in the Files sheet as `unsupported_format` with a review reason: older Word/RTF/ODT documents need conversion, images need OCR, and archives need unpacking. No conversion or archive extraction happens automatically. Corrupt and password-protected files that cannot be read appear as errors; other files continue processing. A folder containing only unsupported files still produces a report.

Hidden files/folders, Word lock files (`~$…`), `.tmp`/`.temp` files, `__MACOSX`, `Thumbs.db`, and `desktop.ini` are ignored and listed in `summary.json`. The selected output folder is excluded from input discovery. Keep unrelated files outside the forms folder, since other visible files are included in the report. `summary.json` also contains counts by extraction status.

## Files

- **MARP_question_inventory.xlsx**: 33 questions and individual fields, organized by section with original wording, assigned field IDs, response types, conditional follow-ups, and review columns.
- **questions.json**: the five question-definition columns as an array of rows. Each row contains field ID, section, question or field wording, response type, and condition or format, in that order.

The inventory was prepared from a Word sample named marp (1).docx. Field IDs are assigned for extraction and are not official question numbers. Existing question numbers remain in the wording.

## Project scope

Question labels and table structure guide extraction. Unknown forms, French versions, the legacy Events variant, scans, and unsupported visual selections require further work or review. This release is a pilot implementation, not a claim of accuracy on the 1,000+ real submissions.

The repository contains question definitions and fictional test forms. Real completed forms and respondent answers belong in the approved work environment.

## Synthetic extraction test suite

[Browse the test suite and coverage guide](fixtures/README.md). It includes 52 fictional files: 20 DOCX forms and 32 PDFs, with known answers and per-response review flags. PDF files include actual Word conversions, scan simulations, interactive checkboxes/radio buttons, and graphical selection marks.

Run `python scripts/validate_fixtures.py` to validate the checked-in suite. The generator supports a 1,052-file throughput batch; see the coverage guide for prerequisites and commands. Large batches repeat scenario families and do not establish real-world accuracy.

### Anti-Racism Call form family

The [Anti-Racism Call guide](fixtures/ANTI_RACISM_CALL.md) describes a separate inventory and 16 additional synthetic files focused on activity tables spanning pages. See `MARP_anti_racism_question_inventory.xlsx` for 32 ordinary fields, three activity fields, and source review issues, or `schemas/anti_racism_call.json` for the machine-readable definitions.

The combined corpora contain **68 synthetic files**. Run `python scripts/validate_anti_racism_fixtures.py` to check the second corpus, including split activity rows and blank spacers. Both inventories retain review flags. The original uploaded PDF is not included.

## Review flags

Question inventories include **Review required** (Not assessed / Yes / No) and **Review reason**. The extractor's output uses Yes / No flags for each response, activity row, and file. The Anti-Racism inventory also records unresolved source questions.

Missing questions, ambiguous checkboxes, unread pages, invalid field formats, and other detected issues are flagged. Blank answers remain distinct from No and N/A. A No review flag means no automated issue was detected, not that a human verified the response.

# MARP form extraction

Question inventory for the Multiculturalism and Anti-Racism Program (MARP) Final Activity Report – Events Component.

## Files

- **MARP_question_inventory.xlsx**: 33 questions and individual fields, organized by section with original wording, assigned field IDs, response types, conditional follow-ups, and review columns.
- **questions.json**: the five question-definition columns as an array of rows. Each row contains field ID, section, question or field wording, response type, and condition or format, in that order.

The inventory was prepared from a Word sample named marp (1).docx. Field IDs are assigned for extraction and are not official question numbers. Existing question numbers remain in the wording.

## Project scope

This inventory is the starting point for future code that extracts responses from Word documents and PDFs whose layout may vary. Extraction code has not yet been implemented.

The repository contains question definitions and fictional test forms. Real completed forms and respondent answers belong in the approved work environment.

## Synthetic extraction test suite

[Browse the test suite and coverage guide](fixtures/README.md). It includes 52 fictional files: 20 DOCX forms and 32 PDFs, with known answers and per-response review flags. PDF files include actual Word conversions, scan simulations, interactive checkboxes/radio buttons, and graphical selection marks.

Run `python scripts/validate_fixtures.py` to validate the checked-in suite. The generator supports a 1,052-file throughput batch; see the coverage guide for prerequisites and commands. Large batches repeat scenario families and do not establish real-world accuracy.

## Review flags

The workbook includes **Review required** (Not assessed / Yes / No) and **Review reason**. All fields start as Not assessed. These are placeholders in the question inventory; the future extractor will create separate review flags for each response in each submitted form.

Flag missing questions, unclear answer boundaries, ambiguous checkboxes, unreadable text, and invalid field formats for review. Keep blank answers distinct from No and N/A. A No review flag will mean no automated issue was detected, not that a human verified the response.

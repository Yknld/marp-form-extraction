# Anti-Racism Call synthetic cases

This is a separate form family from Events Component. The inventory is in `schemas/anti_racism_call.json` and `MARP_anti_racism_question_inventory.xlsx`. Its 32 ordinary fields use ARC-prefixed IDs. Activity tables have three columns per logical activity: description, completion, and participant count. Use the form family and field ID together when mapping responses.

The source was inspected for its question labels and layout only. The uploaded PDF, its identifying details, and its answers are not published. Synthetic forms use entirely fictional responses and do not reproduce the source's organisation, signatory, or date.

## Sixteen files

Each of these eight scenarios has a DOCX and a PDF converted from it:

1. Four activity rows continuing on the next page without column headers.
2. Activity table continuation with repeated headers.
3. One activity split across two physical tables/pages; join into one logical activity and flag for review.
4. Seven activity rows with multiline lists inside descriptions.
5. Both selected, neither selected, a missing selection symbol, and a clear No.
6. A completely blank spacer row that must not become an activity.
7. Participant counts of zero, blank, N/A, and an approximate value.
8. Reordered columns: participant count, description, then completion.

Each form also contains all identifiable ordinary fields. The generated layouts are representative test cases, not pixel-identical copies of the original 13-page form.

## Source limitations

Questions 1, 2, and 4 visibly end mid-sentence. The inventory retains those visible fragments and flags them. The numbering gap between 6 and 8 is recorded as a document review issue; no question 7 has been invented. Unlabelled boxes in the source are another document review issue: their contents, if any, must remain unmatched until reviewed. These limitations are carried into the expected review policy even though the synthetic forms do not recreate every unlabelled source box.

## Answer keys and output design

`anti_racism_call/expected.jsonl` contains one record per file. `expected.responses` is keyed by field ID. `expected.activities` contains logical activity records with raw text, normalized values, per-field review flags, row-level flags, and reasons. `physical_activity_rows` distinguishes activities, continuation rows, and empty spacers.

Table and row indices are one-based locations in the paired Word document. They are not PDF coordinates or PDF page references. PDF page counts are independently recorded; the future extractor must supply actual source pages/coordinates in its output.

Future Excel output should contain a **Responses** table and an **Activities** table, joined by file ID and template ID. Keep raw checkbox/count text alongside normalized values, preserve zero versus blank versus N/A, and retain multiple source locations for continued activities. Unknown form versions must be flagged rather than forced into an inventory. The JSON schema lists the proposed output columns.

## Commands

Use the dependencies in `requirements-fixtures.txt` and LibreOffice on PATH:

```bash
python scripts/generate_anti_racism_fixtures.py --output /path/to/new-empty-directory
python scripts/validate_anti_racism_fixtures.py /path/to/new-empty-directory
```

Without an argument, the validator checks the committed corpus. Validation checks fixture structure, answer-key consistency, activity text and continuation joins, not the accuracy of an extractor. Extraction code is still to be implemented.

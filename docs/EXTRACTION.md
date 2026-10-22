# Extraction scope and review

## What version 0.1 does

- Reads `.docx` XML paragraphs, tables, and checkbox display text, retaining table and row locations. Text marked with direct strikethrough is excluded from interpretation but retained in the original answer and flagged.
- Reads PDF table cells and positioned text with pdfplumber, retaining page numbers and bounding boxes. It handles wrapped headings and questions that cross pages in the tested layouts.
- Detects the two supplied English form families from their title/question wording. Matching tolerates whitespace, punctuation, and case differences; it does not yet implement general fuzzy matching or all legacy aliases.
- Separates question responses from repeated Anti-Racism activity rows, identifies columns by labels, skips entirely empty spacer rows, and joins possible continuation rows with a review flag.
- Interprets printed Yes/No checkbox text, distinguishes both/neither selected, and keeps numeric/date validation separate from original answers.
- Routes image-only pages and sufficiently large image regions without text to OCR review. OCR is not run. Other readable pages may still be extracted when the form family can be identified.
- Logs unsupported `.doc` files and failed PDFs without stopping the batch, marks byte-identical duplicates without deleting them, and saves per-file resumable checkpoints.

## Deliberate limitations

French forms and the legacy Events form variant are identified when recognisable, but their question inventories have not been implemented. They are flagged as unsupported rather than mapped to a different form. Language detection is based on known labels, not a general language classifier.

Interactive PDF widget values, freehand marks, vector-drawn ticks/circles, and PDF strikethrough geometry are not interpreted in this release. PDF annotations/widgets trigger file review. Printed checkbox symbols can be interpreted when they are present in the text layer. Crossed-out PDF choices may therefore appear as both selected and remain unresolved. Standalone checkbox probes have no supported full-form identity and are explicitly unsupported.

Word image answers, text boxes, nested tables, and tracked changes trigger file review; their reading order is not fully resolved. Not all small embedded images or corrupt text encodings can be detected automatically. Review flags are not an exhaustive correctness guarantee.

Form variants can introduce new questions or rearrange information beyond the tested layouts. Unmatched table content is retained for review. Unknown/unlabelled content is not assigned invented question wording. The Anti-Racism schema contains incomplete question labels inherited from the supplied sample, so it triggers source-inventory review.

## Outputs

`Files` contains one row for every processed source, including unsupported and failed files. `Responses` contains one record for each expected field of a recognised form, including explicitly missing questions. `Activities` contains logical activity rows. `Unmatched` contains table content that was not assigned a question. File IDs are derived from relative paths; source hashes are retained in JSON.

Raw text, normalized values, source locations, and review reasons are retained in `results.jsonl` and checkpoints. Excel strings are written as literal text, so an answer or filename starting with `=` is not executed as a formula. XML-invalid control characters are shown as escaped character codes in Excel. Excel limits a cell to 32,767 characters: longer answers are flagged and their complete text remains in JSON. Very long rows can exceed Excel's maximum displayed row height; read the full cell or JSON rather than assuming the visible portion is complete.

Counts of zero remain numeric zero. Blank and N/A both have no normalized numeric value but retain distinct raw text and presence. Ambiguous dates are not guessed. Approximate counts retain their qualifier in raw text and are flagged. A review flag on an activity continuation means the proposed merge needs checking.

The Files flag is the aggregate starting point for review. A file-level annotation/OCR/source-inventory issue may require checking responses even if an individual response has no separate issue. Known source-inventory concerns should be resolved before trusting bulk results.

## Batch operation

The reader processes files sequentially with bounded per-file processing state and writes checkpoints after each file. Final aggregation/Excel export loads the current batch results in memory; very large answers across thousands of files may need future streaming exports. There is no per-file execution timeout yet. Checkpoint JSON is saved atomically. A corrupt checkpoint is recomputed. Unchanged cached failures can be retried with `--no-resume`.

Use a separate output directory and save manual workbook review edits elsewhere before rerunning. Checkpoints are caches, not the authoritative record of human review. At this stage the 68-file suite has been exercised; the 1,052-file generator option has not been used to establish throughput or runtime limits.

## Tests

```bash
python -m unittest discover -s tests -v
python -m marp --input fixtures --output /path/to/synthetic-results
python scripts/evaluate_extraction.py --results /path/to/synthetic-results/results.jsonl --report /path/to/evaluation.json
```

The extractor never reads `expected.jsonl`. Answer keys are loaded only by the separate evaluator. Keep Word/PDF pairs together when constructing future held-out datasets.

# Synthetic test forms

This page describes the Events Component corpus. A separate [Anti-Racism Call corpus](ANTI_RACISM_CALL.md) adds 16 files with multi-page activity tables and its own question inventory.

All names, organisations, contact details, signatures, and answers in these files are fictional. Email addresses and websites use the reserved example.invalid domain. None of the submitted sample's respondent answers are included.

The starter suite contains **52 files**: 20 Word documents, their 20 LibreOffice-converted PDFs, six scan variants, and six focused PDF selection probes. The first 46 files are full forms. The six probes contain only the three Yes/No fields.

## Coverage

| Case | Variation |
| --- | --- |
| 01 | ASCII [X] / [ ] selections |
| 02 | Unicode checked / unchecked boxes |
| 03 | Native Word checkbox content controls |
| 04 | Both Yes and No selected; requires review |
| 05 | Neither selected; requires review |
| 06 | No printed before Yes |
| 07 | Typed Yes or No without boxes |
| 08 | Parentheses around the chosen option; conservative review flag |
| 09 | Crossed-out Yes followed by selected No; requires review |
| 10 | Added table rows for each listed answer |
| 11 | Bulleted lists inside a single answer cell |
| 12 | Answer continues on another page in a separate table |
| 13 | Entire question and answer omitted |
| 14 | Sections appear in a different order |
| 15 | Zero, blank, N/A, comma-separated count, ambiguous date, impossible date |
| 16 | No selected but follow-up describes a change |
| 17 | Answer quotes text that resembles another question header |
| 18 | Merged question cells with answers in multiple columns |
| 19 | Long multiline answer |
| 20 | Accents, French answer text, and varied header fonts/sizes |

The scan PDFs cover clean image-only pages, two-degree skew, lower resolution, blur, 90-degree rotation, and alternating selectable-text and scanned pages. Scan source forms also include checkbox variations. Scan variants are simulations, not real scanner or handwriting samples. Focused PDF probes include clear selection and both selected with actual AcroForm checkbox fields, grouped radio buttons, vector-drawn circles, offset ticks, and faint marks.

## Expected results

`synthetic/expected.jsonl` has one record per file:

- `file`, `case_id`, `variant`, `scope`, `sha256`, and PDF `pages`.
- `expected`: keyed by the 33 inventory field IDs (three IDs for selection probes).
- `answer_text`: intended text, with newlines joining list items. It is not a byte-exact PDF text dump. Checkbox ambiguity cases retain the visible marks.
- `normalized_answer`: intended interpretation, or null where no unique answer exists.
- `selected_options`: checkbox choices when relevant.
- `presence`: present, blank, or missing for full forms.
- `review_required` and `review_reasons`: expected minimum review policy.

All scan-derived responses are conservatively flagged for verification, including the mixed-page form. This can be refined once OCR is evaluated. A false review flag means the fixture has no intentionally planted uncertainty; it does not certify a future extractor's output. A wrong or uncertain extraction must still be flagged.

Do not expose the answer keys to extraction code. Use them only in a separate evaluator. Word/PDF pairs share answers deliberately: keep both members of each pair together when splitting development and held-out evaluation sets. Larger generated batches repeat the 20 scenario families and test throughput; they are not hundreds of independent layout designs.

## Generate and validate

Install Python dependencies from `requirements-fixtures.txt`, plus LibreOffice and Poppler. Both must be available on PATH. Explicit `--soffice` and `--pdftoppm` executable paths are supported.

```bash
python -m pip install -r requirements-fixtures.txt
python scripts/generate_fixtures.py --output /path/to/new-empty-directory
python scripts/validate_fixtures.py /path/to/new-empty-directory
```

The default generates 52 files. For a **1,052-file throughput batch**, use 520 Word cases plus their 520 PDF conversions and 12 special PDFs:

```bash
python scripts/generate_fixtures.py --count 520 --output /path/to/new-empty-large-batch
```

An existing nonempty output directory is rejected. Source content is repeatable; PDF/DOCX metadata can produce different file hashes on regeneration. Large batches take longer and are intended to be generated locally.

## Remaining coverage

This is a starter corpus, not exhaustive coverage. Real handwriting, real pen strokes, severe photocopy damage, stamps over text, annotations, password-protected PDFs, legacy .doc files, French question translations, and cropped/missing pages still need separate cases. Batch unit tests cover damaged files and byte-identical duplicate submissions. See [the benchmark](../docs/BENCHMARK.md) for the current extractor's measured results and unsupported cases.

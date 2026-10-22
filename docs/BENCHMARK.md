# Synthetic extraction benchmark

Run on 2026-10-22 using the committed synthetic fixtures. These are development fixtures used while implementing the extractor, not an independent held-out test or a measure of accuracy on real submissions.

| Check | Result |
| --- | --- |
| Input files processed | 68, no file errors |
| Supported full text forms | 56 |
| Ordinary answers matching expected values | 1,829 / 1,832 |
| Yes/no answers (subset of ordinary answers) | 181 / 184 |
| Activity cells matching expected values | 210 / 210 across 70 activity rows |
| Activity row count errors | 0 |
| Expected review conditions detected | 100 / 100 |
| Additional flags among 1,808 expected-clear checks | 0 |
| Scan fixtures routed to OCR required | 6 / 6 |
| Standalone selection probes routed to unsupported template | 6 / 6 |
| Unit tests | 11 passed |

All three answer mismatches are crossed-out PDF selections. The text contains both selected choices, and the extractor returns an unresolved value with a review flag. It does not infer which mark was cancelled from PDF drawing geometry.

Text comparisons normalize punctuation, case and whitespace; selection comparisons use normalized values. The review metric combines response-level, activity-level and scan-file checks, so it is not a uniform field-level accuracy estimate. File-level warnings are separate from the additional-review count. Scans and standalone probes are excluded from answer accuracy: routing them to review is not successful answer extraction.

The batch supports checkpoints and duplicate detection. Resume, changed inputs, duplicate files and per-file failure handling have unit coverage. A second unchanged 68-file batch reuses all checkpoints. A 1,000+ file throughput benchmark and validation on real work-machine submissions have not been performed.

Reproduce the extraction and evaluation using the commands in [EXTRACTION.md](EXTRACTION.md). Keep real submissions and their extracted outputs on the authorized work machine.

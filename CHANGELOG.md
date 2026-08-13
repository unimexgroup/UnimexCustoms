# Changelog

All notable changes to UnimexCustoms. Versions are tagged `customs-vX.Y.Z`;
the number must match `CUSTOMS_VERSION` in `_version.py` or CI refuses to build.

## v1.0.0 — 2026-08-13

First release. Replaces the hand-run `customs-summary-prompt.txt` workflow with
a self-updating executable.

### Added

- **Customs summary output** — one row per part number aggregated across every
  invoice, sorted by value descending, with a bold TOTAL row using live `=SUM()`
  formulas. `Charges` = 3.00 and `Zone` = P as confirmed by the customs team.
- **HTS from the parts database** — the third column is the US HTS code looked
  up from `database\Kohler Parts for Upload ....xlsx`, replacing the Mexican
  fracción that used to come from the reception report. The most recently
  modified workbook in `database\` is used, so refreshing the database needs no
  rebuild.
- **Four-step HTS resolution** — exact Product Key, then punctuation-insensitive,
  then same key minus the finish suffix, then same base part number. The last two
  are accepted only when every candidate agrees; 113 of 24,596 base numbers in
  the current database carry conflicting codes across their suffixes. An
  unresolvable part ships a **blank** cell, lands on a `Review` sheet, and counts
  as a problem — never a guess.
- **Total cartons on the first row** — shipping paperwork states a shipment total
  only, so the whole total goes on the first (highest-value) row and the rest stay
  blank. The column ties out to the shipment without inventing per-part counts.
- **Weight allocation per HS group** in proportion to piece count, using
  largest-remainder rounding in integer hundredths so columns sum exactly and no
  float artefacts reach the file. Single-part groups are exact; the log names
  which groups are exact and which are estimates.
- **Glyph-run column reconstruction** for the reception report, so the tariff
  column is read correctly where it physically overlaps the purchase-order
  column and flat text extraction interleaves the two.
- **Typo-tolerant part matching** — punctuation-blind, then edit distance 1,
  accepted only when exactly one candidate fits. Every rewrite is reported.
- **Reconciliation suite** — line arithmetic, per-page totals, all Invoice
  Summary sets combined, packing-list quantities and weights, reception-report
  parts and quantities, HTS plausibility, and country of origin. A missing value
  is a failure, not a pass.
- **Content-based document identification** — filenames are never trusted.
- **Multi-shipment runs** — one subfolder per shipment under `input\`.
- **Self-update from GitHub Releases**, checked only after a run that had
  problems, with the detached swap-and-relaunch helper.
- **Run logs** in `logs\`, mirroring the console output.

# Changelog

All notable changes to UnimexCustoms. Versions are tagged `customs-vX.Y.Z`;
the number must match `CUSTOMS_VERSION` in `_version.py` or CI refuses to build.

## v1.1.0 — 2026-08-14

### Added

- **Packing lists in PDF form.** Previously a PDF packing list was reported as
  unrecognized and the run continued without weights or cartons. PDFs carry no
  column structure, only ink at coordinates, and every supplier's template
  differs — so the table is rebuilt geometrically: find the header band, take
  each header cell's horizontal span as a column, and assign every word below it
  to the column it sits under. Header cells stacked over several lines
  (`N.W/CTN` over `(KGS)`) are one column, and a line item's description is
  gathered from the printed lines above and below its figures, since suppliers
  wrap it either way.
- **A validation gate on PDF packing lists.** A parse is only used if it
  reproduces every total the document states, taken from both the TOTAL row and
  any prose restatement (`TOTAL SAYS 262CTNS ,G.W 2648.7KGS`). One that does not
  tie out is refused outright and the run continues without a packing list —
  blank weights and a loud warning — because wrong weights in a filing are worse
  than absent ones. When a document states no totals at all, the parse is used
  but reported as uncorroborated.
- **Per-part carton counts** when the packing list actually states them per row,
  replacing the total-on-first-row rule for those documents. Used only when
  every matched row has a count and they add to the shipment total, so a
  partly-filled column can never look exact. Documents that state only a
  shipment total are unchanged.
- **Combined invoice + packing list (CIPL) PDFs**: when no separate packing list
  arrives, the invoice PDF is re-read for a packing table. Same validation gate,
  so a false positive cannot pass.
- Per-carton columns (`Pcs/ctn`, `N.W/CTN`, `Vol./ctn`) are recognized and
  excluded, so a per-unit figure is never mistaken for a row total.

### Fixed

- The packing-list totals check treated a missing total as a pass. When a column
  was mis-identified, that column's total cell is empty — so the check was
  skipped on exactly the parses that had gone wrong. It now validates against
  every stated total from every source, and a mismatch anywhere refuses the
  parse.

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

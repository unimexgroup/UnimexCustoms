# Changelog

All notable changes to UnimexCustoms. Versions are tagged `customs-vX.Y.Z`;
the number must match `CUSTOMS_VERSION` in `_version.py` or CI refuses to build.

## v1.5.1 — 2026-08-21

Says so plainly when a document is a **photocopy rather than an export**. The
customs team sent `220509656.pdf` as a file the exe "had issues with": 29 pages
covering the whole of shipment `KBU07122026491` / invoice `IV00075089` — the
reception report, the reconocimiento previo, four printings of the Runner
(Xiamen) packing list, two commercial invoices, the waybills and the FCR. Every
page is a 1-bit image from a RICOH MP 3055 copier and carries **zero
characters**. It is the shipment straight after `KBU07052026468`, whose native
workbook v1.5.0 was written to read.

### Fixed

- **A scan is now named as a scan.** The run said `[SKIP] shipment: no
  commercial invoice found (found 1 unknown)`, which reads as a parsing gap the
  tool should grow out of — the same wording it uses for a template nobody has
  taught it yet. Nothing was wrong with the parsers: there was no text on the
  page to parse. A PDF whose every page holds an image and not one character is
  now identified as its own kind of document, and the run says what it is and
  what to ask for instead — the file the sender's own system produced. The
  `Troubleshooting` table has promised this diagnosis since v1.0.0; the tool now
  actually gives it.
- **A part-scanned bundle no longer loses pages in silence.** Where some pages
  carry text and others are images, the readable ones parsed and the rest simply
  were not there — no line said so, and a shipment can lose a whole invoice that
  way. The run now names how many pages of a document could not be read.

### Not done, deliberately: OCR

Reading these pages means OCR, and OCR was measured against this file before
being ruled out, not assumed to be unusable. On the *cleanest* page — the
packing list, large type, no annotations over most of it — a current engine
dropped the CARTONS cell on **10 of 22 rows**, every one of them a row a
warehouse pencil-tick crosses. That is not a garbled character anyone would
notice; it shifts the row left, so `306 | 34 | 265.586` is read as
`306 | 265.586 | 2.612`. Weights here are allocated **by carton count**, so a
silently missing carton figure corrupts the weight of every part sharing that
pallet.

The commercial invoice is worse, and it is the document the tool cannot do
without. Across its 52 lines the engine returned `1013858883BU` as
`101385888380`, `1413306-BL` as `1413306-8L`, `1547190-SN` as `1547190-5N`,
`MRKU4670621` as `MRXU4670621`, the HS code `8481909000` as `84 81909000`, and
the quantity `500.000` as `$00.000`. Those first two are the keys everything in
this tool joins on — the purchase-order number and the part number — corrupted
on roughly half the lines, and a part number that misses the database gets no
HTS at all.

So OCR here buys nothing. Either a misread digit reaches a customs filing as a
real figure, or the reconciliation catches it and the run is a wall of `[CHECK]`
lines with no usable output. The documents all exist as exports upstream — this
supplier's previous shipment arrived as the workbook v1.5.0 reads — so the fix
that produces a filing is asking for that file, and the fix in this version is
saying so clearly enough that nobody spends an afternoon on the parsers instead.

## v1.5.0 — 2026-08-19

Reads Runner (Xiamen)'s combined invoice + packing-list workbook, the first
template that writes a **token weight instead of leaving the cell blank**. The
customs team reported that weight "wasn't being picked up" for shipment
`KBU07052026468` / invoices `IV00075087`-`88`, and it was not: 11 of 25 parts
reached the file carrying between 0.01 and 0.10 kg, including 26,000 stem
adapters at 0.10 kg and 1,000 hose assemblies at 0.01 kg.

### Fixed

- **A token weight is now read as the blank cell it stands in for.** Every
  weight rule keyed off "a non-blank gross weight starts a new group", which
  assumed a supplier leaves the cell empty on a row that shares the weight
  stated above it. Runner types `0.01` there. So all 53 rows opened a group of
  their own, each holding a hundredth of a kilo, and the 22 real pallet weights
  were split off from the rows they cover. A figure that is a kilo or less on
  the whole line **and** under a gram per piece is a placeholder, not a
  measurement; it carries down like a blank, and is added to the group it rides
  on so the column still ties to the packing list's printed total exactly.
  Across the three older suppliers the lightest real figure is 1.97 kg at
  0.123 kg/piece and the lightest real per-piece figure is 0.020 kg, 20× clear
  of the threshold — none of them trips it, and their output is byte-identical.
- **Nothing was flagged.** Not one `[CHECK]` fired on a file declaring ten
  grams for 26,000 pieces, because the only figure the weights were checked
  against was the packing list's own total row — a `=SUM()` over the column the
  tokens sit in. A document is not evidence about itself. The run now reports
  how many rows stated a token weight and that those parts' weights are
  allocated rather than measured.

### Changed

- **A group's weight is split by cartons, not pieces, where the packing list
  states cartons per row.** Piece count assumes every piece in a group weighs
  the same, which breaks as soon as a group mixes sizes: one of Runner's
  pallets holds 100 tub/shower assemblies (100 cartons) beside 1,000 hose
  assemblies (20 cartons), and by piece count the assemblies take a tenth of
  the weight of the hoses — 0.31 kg each against a real 2.8 kg. A carton is
  packed to a working weight whatever is inside it. Piece count stays the
  fallback, so the Kohler Nanchang template — which states cartons for the
  shipment only — is unaffected; both older shipments produce byte-identical
  files, verified.

### Known limitation

- Runner states each pallet's weight on **one** row of that pallet and `0.01`
  on the rest, and the file does not say reliably which row. Read as
  carry-down — the convention every other supplier here follows, and the only
  reading under which no row is left unweighed — some groups still imply a
  per-piece weight that disagrees with the same part elsewhere in the file by
  around 2×. Every part now lands in a plausible range and the column total is
  exact, but the per-part split inside a shared pallet is an estimate, as the
  log says. A packing list stating weight per line would remove the guess.

## v1.4.0 — 2026-08-19

Corrects what the log says about a folder holding several invoices. The merging
itself is right and stays: verified against the customs team's hand-built file
for shipment `um260545`, which puts invoices 202604441 and 202604923 on **one
sheet** with a subtotal per invoice and a grand total across both — the same
figures this tool already produced ($36,559.93 + $23,158.80 = $59,718.73,
cartons 413 + 7 = 420, gross 1,975.20 + 1,363.67 = 3,338.87 kg).

### Fixed

- **The multi-document warning described something the tool never did.**
  `2 invoices documents found; using 202604441.0713A.xlsx` named one file while
  the run read and merged all of them. Anyone reconciling the output against
  that one document found figures that could not come from it, and the natural
  reading — that half the paperwork had been dropped — was wrong. It now says
  plainly that every one of them is read and merged into the shipment's file.
- A shipment or guide number containing a slash (`202604441/00199`, exactly how
  invoices print it) would have sent the output file to a folder that does not
  exist and failed the shipment. The file **name** is sanitized; nothing inside
  the file changes.

### Changed

- The reception-report case is a separate `[WARN]`, since only the first one is
  cross-checked — that one really is "using X".

## v1.3.0 — 2026-08-14

Reads the remaining supplier format from the August batch: a PDF packing list
whose measures are split into PER and TTL sub-columns. All eight shipments in
that batch now produce a customs file that ties to its own documents.

### Added

- **PER/TTL packing lists (PDF).** One supplier prints every measure twice --
  per carton and total -- with a line item spread over three printed lines:
  description, carton range, then the figures. Values are read from contiguous
  glyph runs, because the two sub-columns do not merely abut: they physically
  overlap, so a flat read interleaves them into a number that is wrong but
  still plausible (`1.9200` + `38.40` extracts as `1.920038.40`, and worse,
  `204.55` + `1,227.30` extracts as `204.5510,0227.30`).
- **Reconstruction from PER x CARTONS** where a total is missing from the page
  entirely. This is the relationship the document is built on, and the result
  still has to agree with the shipment totals or the whole packing list is
  refused — on the batch it reproduced 828.90 and 1,227.30 exactly.
- **Several packing documents per shipment** are merged, as invoices already
  were. One supplier sends one packing list per invoice.
- Figures that wrap onto a line of their own are attached to the row they sit
  nearest, and never overwrite a value that row already has.

### Fixed

- Characters were re-bucketed into lines independently of the words, which
  could split one visual line and break the glyph runs across it. Each line's
  runs are now built from the characters of that line's own words.
- A value narrow enough to reach neither the PER nor the TTL label (a two-digit
  quantity under a right-aligned heading) was dropped rather than assigned.

## v1.2.0 — 2026-08-14

Adds the supplier formats in the August batch: three families that send the
commercial invoice as a **spreadsheet**, which the tool previously could not
read at all. Verified against all eight shipments in that batch; every one now
ties to the totals printed on its own documents.

### Added

- **Invoices in Excel.** Found by header keywords like the packing list, so a
  new supplier's column order costs nothing. Handles headers stacked over two
  or three rows, several invoice sheets in one workbook (`CI 1` + `CI 2` is one
  shipment), and country of origin stated once in the header rather than per
  line.
- **Invoices as a PDF table.** Suppliers who print one table instead of
  one-invoice-per-page are read with the same geometric parser as the PDF
  packing list. Currency marks are stripped (`US$12,150.00`).
- **Several invoice documents per shipment** are merged. One supplier sends one
  file per invoice, and the customs file covers the whole shipment.
- **Several packing sheets per workbook** are merged, with their totals summed.
  A shipment split across `PK40GP1` and `PK40GP2` previously lost half its
  weights silently.
- **Zip input.** Drop the carrier's zip in `input\` and it is extracted and
  processed as one shipment; already-extracted folders keep working unchanged.
  Archive paths are flattened and sanitized so no member can escape the folder.
- **Duplicate detection.** Byte-identical copies (`X.pdf` and `X (1).pdf`, which
  is how these arrive by email) are counted once instead of doubling every
  figure.
- **Transport paperwork is recognized and skipped** — waybills, bills of lading,
  FCRs, arrival notices, container checklists, images. It used to be reported as
  unrecognized, which counted as a problem and triggered pointless update checks.
- **Blank net weight** where a supplier states gross only, with a plain warning.
  Nothing is copied from gross to fill the column.
- **Value and weight true-up.** Rounding each part to the cent could leave a
  column a cent or two off the figure printed on the invoice or packing list;
  the residual is now pushed onto the largest rows so the file ties exactly.
  Bounded, so a genuine discrepancy is still reported rather than smeared away.

### Changed

- **Invoice and packing list are now reconciled per (PO, part), not per line.**
  The two documents agree on what shipped, not on how it was written down: one
  supplier splits a part across four invoice lines and two packing rows, another
  prints one packing row per pallet. Weights and cartons are read from the
  packing rows directly, which is the document that states what physically
  shipped. This removed the per-line join and the pallet-merging pass it needed.
- Packing lists that never repeat the purchase order are matched on part number
  alone, and say so.

### Fixed

- A column was mapped to the **leftmost** header containing its keyword. One
  supplier leaves a stray `PO#` label above an unrelated column, so the PO was
  read from an empty column and every join failed. The closest-matching header
  now wins.
- An invoice sheet with PO, part and quantity was read as a **packing list with
  no weights**, which silently suppressed the "no packing list" warning. A
  packing sheet must now also carry weights or cartons.
- An invoice sheet carrying a cartons column was read as a **second packing
  sheet**, double-counting every carton and weight. A sheet with a unit price
  is an invoice.
- A totals row announcing itself as `TOTAL` in a label column was read as a line
  item for a part called "TOTAL".
- `PACKING/WEIGHT LIST` was not recognized as a packing list.

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

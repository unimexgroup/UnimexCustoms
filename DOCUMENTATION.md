# UnimexCustoms — Technical Documentation

Version 1.0.0 · US Customs Brokerage

---

## 1. What this replaces

The customs summary used to be produced by pasting `customs-summary-prompt.txt` into a chat model and attaching the shipment's documents. That worked, but every shipment needed a human to drive it, re-confirm the constants, and re-read the traps.

This tool is that prompt, executed. Same output, same reconciliation checks, same refusal to guess — as a double-clickable exe that self-updates.

**The one deliberate change from the old workflow:** the third column is now the **US HTS code from the Kohler parts database**, not the Mexican *fracción* from the reception report. The reception report is now only a cross-check.

---

## 2. Inputs

Documents are identified by **content**, never by filename (`classify()`), because filenames vary by client and by whoever forwarded the email.

| Document | Required | Supplies | Recognized by |
|---|---|---|---|
| Commercial invoices (PDF **or** Excel) | **Yes** | part, qty, unit price, extended value, PO, country of origin, invoice HS | PDF: "Invoice" + "P.O. No.", or a table with qty/amount columns. Excel: a sheet with PO, part, qty, amount **and a unit price** |
| Packing list (`.xls`/`.xlsx`/PDF) | Strongly wanted | gross/net weight, cartons, shipment number | Excel: PO + part + qty **and** weights or cartons, but no unit price. PDF: "PACKING LIST" (or "PACKING/WEIGHT LIST") plus a readable table |
| Warehouse reception report (PDF) | Optional | independent part/qty check, Mexican fracción | "REPORTE DE RECEPCION" / "RECONOCIMIENTO PREVIO" |
| Transport paperwork | Ignored | nothing | bills of lading, waybills, FCRs, arrival notices, and any `.docx`/image/email file |
| Parts database (`.xlsx`) | **Yes** | `Product Key` → `HTS` | lives in `database\`, not `input\` |

A single workbook is frequently **both** invoice and packing list, one sheet
each. The discriminator is the unit price: only an invoice has one. Getting this
wrong in either direction is costly — an invoice sheet read as a packing list
silently suppresses the "no packing list" warning, and an invoice sheet with a
cartons column read as a *second* packing sheet double-counts every weight.

### Input forms

Shipment paperwork arrives from the carrier as one zip. Dropping that zip
straight into `input\` works: it is extracted to its own folder and processed as
one shipment, with archive paths flattened and sanitized so no member can write
outside it. Already-extracted files and hand-made subfolders behave exactly as
before.

Byte-identical duplicates are dropped before anything is read. These documents
arrive as email attachments and the same file routinely appears twice, once as
`X.pdf` and once as `X (1).pdf`; without this every figure would double.

Without a packing list the file still writes, with weights and cartons blank and a loud warning — the tool does not block on it.

### The parts database

Loaded from `database\` — the **most recently modified** workbook there wins. Refreshing it is a copy-paste, with no rebuild and no release. If the folder is empty the run stops immediately with an explanation, since an HTS-less customs file is worthless.

Columns are found by name: `Product Key` (or `Part`/`SKU`) and any column starting with `HTS`. The current export is 66,350 parts.

---

## 3. HTS resolution

`lookup_hts()` tries four steps, most specific first, and **stops at the first unambiguous answer**:

| # | Match | Example |
|---|---|---|
| 1 | Exact Product Key | `1522807-U-CP` → `8481901000` |
| 2 | Same key ignoring punctuation | `1332370-U-2-MB` → `1332370-U-2MB` |
| 3 | Same key except the finish suffix | `1522807-U-BV` → `1522807-U-*` |
| 4 | Same leading part number | `830915-XX` → `830915-*` |

Steps 3 and 4 are accepted **only when every candidate agrees on the code**. This matters: in the current database **113 of 24,596 base part numbers carry more than one HTS across their suffixes** — `1044620-A-*` is `7419805010` while `1044620-*` is `8481901000`. Falling back to "same base number" blindly would file the wrong code on those.

When the candidates disagree, or the part is absent entirely, **the cell ships blank**, the part is listed on a `Review` sheet in the workbook, and the run is marked as having problems. Nothing is ever guessed into a filing.

Every non-exact match is printed as a `[NOTE]` line so the resolution path is auditable.

---

## 4. Processing pipeline

1. **Parse invoices.** One page is normally one invoice; multiple line items per page are handled. Pages titled "Invoice Summary" are roll-ups, kept aside for reconciliation.
2. **Parse the packing list.** A non-blank gross weight starts a new HS weight group; the weight is stated once per group and carried down. Cartons and volume are shipment-wide, not per line.
3. **Reconcile** the invoice against the packing list **per (PO, part)**, not per line.
4. **Allocate weight inside each HS group** across only the parts in that group, in proportion to piece count, using largest-remainder rounding in integer hundredths.
5. **Resolve the HTS** for every part.
6. **Reconcile** (section 6). Nothing is suppressed.
7. **Write** the workbook.

### Why the join is per (PO, part)

The two documents agree on **what shipped**, not on how it was written down. Real examples from one batch: one supplier splits a single part across four invoice lines but two packing rows; another prints one packing row per pallet, so one invoice line appears fifteen times. A per-line join fails on both, and the failure is expensive — the unmatched rows' weights are simply dropped.

So both sides are totalled per `(PO, part)` and compared there. Weights and cartons are then read from the **packing rows directly**, because the packing list is the document that states what physically shipped and it is free to break a part across any number of rows. The invoice supplies value, country and its own quantity. Where a packing list never repeats the purchase order, matching falls back to the part number alone and says so.

This replaced a per-line join and the pallet-merging pass that existed only to prop it up.

### Weight allocation

`allocate_largest_remainder()` works in integer hundredths, so each column sums to its group total **exactly** and the file total matches the packing list exactly, with no `530.3199999999999` float artefacts. Leftover cents go to the largest fractional parts, ties broken by the larger quantity, so the result is identical run to run.

A group holding one part is exact. Any other group assumes equal weight per piece — an estimate, and the run log says so and names which groups are which. Packing lists that state a weight per line produce one group per line, so every figure is exact.

**True-up.** Each group's split is exact within that group, but the group totals are themselves rounded to the cent, so a shipment with fifty single-row groups can land a few cents from the figure printed on the packing list. The residual is pushed onto the heaviest rows — the same largest-remainder principle applied once at the file level. Value gets the same treatment against the invoice total. Both are bounded to a few cents, so a genuine discrepancy is still reported as a `[CHECK]` rather than smeared across the rows.

**Net weight** is left blank where the document states none — several suppliers give gross only. It is never filled from gross, and the run says plainly that the column is empty.

### Header matching

Both the invoice and packing readers find their table by keyword, allowing the header to span up to three rows (`QTY` over `(PCS)`, or an `HS CODE` printed two rows below its neighbours). Two rules matter more than they look:

- The winning column is the one whose header is **closest** to the keyword, not the leftmost containing it. One supplier leaves a stray `PO#` label above an unrelated column; taking that in preference to the real `PO #` header mapped the purchase order onto an empty column and made every join fail.
- Columns naming a **per-carton** figure (`Pcs/ctn`, `N.W/CTN`, `Vol./ctn`) are excluded before matching, so a per-unit weight is never read as a row total.

### Packing lists in PDF form

A PDF has no columns — only ink at coordinates — and every supplier's template differs, so there is no regex to write. `parse_packing_list_pdf()` rebuilds the table geometrically:

1. Cluster words into visual lines by vertical position.
2. Find the header: the run of up to 5 consecutive lines that together name the most canonical columns. A header cell stacked over several lines (`N.W/CTN` over `N` over `(KGS)`) is **one** column, so header words are clustered by horizontal overlap across the whole band, then merged across gaps narrower than a column gutter — otherwise `Kohler PO#` becomes two columns and the values beneath land in a column with no name.
3. Assign each word below the header to the column it sits under, by centre position, falling back to the nearest column since values often overhang their header.
4. Gather each line item's description from the printed lines above **and** below its figures — suppliers wrap it either way — assigning each text-only line to the row it sits nearest, which is how a person reads it off the page.

Columns naming a per-carton figure (`Pcs/ctn`, `N.W/CTN`, `Vol./ctn`) are excluded before the mapping runs, so a per-unit weight is never mistaken for a row total.

### PER/TTL packing lists

One supplier prints every measure twice — per carton and total — under a two-level header (`QUANTITY` / `Net` / `Gross`, each split into `PER` and `TTL`), with a single line item spread over three printed lines: description, carton range, then the figures.

Three things make this layout hostile, and each has a specific answer:

- **The sub-columns overlap.** They do not merely abut: `1.9200` and `38.40` extract as `1.920038.40`, and `204.55` with `1,227.30` extracts as `204.5510,0227.30` — characters from the two values interleaved. This is TRAP 1 again, so the same answer applies: values are read from contiguous glyph runs, which separate the two columns cleanly where no split of the flat string can.
- **Labels do not sit over their own values.** A group label is centred over its PER/TTL *pair*, and a super-group (`Weight` over `Net` and `Gross`) can sit closer to a leaf than its real parent. So the PER and TTL columns are paired first, and the pair's centre is matched to the nearest label. Matching each leaf on its own picks the wrong owner.
- **A figure can be missing from the page**, having wrapped onto a line of its own or been swallowed by the overlap. Where a total is absent but its per-carton figure is not, the total is reconstructed as **PER × CARTONS** — the relationship the document is built on. This is not a guess that ships unchecked: the shipment's stated totals still have to agree. On the batch it reproduced 828.90 and 1,227.30 exactly, both confirmed by the document's own totals.

**The validation gate.** A geometric parse can go wrong in ways that still look plausible, so the result is only used if it reproduces every total the document states — from the TOTAL row *and* from any prose restatement (`TOTAL SAYS 262CTNS ,G.W 2648.7KGS`). Anything that does not tie out is **refused**, and the run continues as though no packing list arrived: blank weights, blank cartons, loud warning. Wrong weights in a filing are worse than absent ones.

Both sources are needed, and that is not belt-and-braces. When a column is mis-identified, that column's cell in the TOTAL row is *empty* — so a check that skips missing values would skip the check precisely on the parses that had gone wrong. This bug shipped in v1.0.0's gate and was fixed in v1.1.0; the sabotage test that catches it is described below. When a document states no totals at all, the parse is used but reported as uncorroborated rather than being presented as verified.

### Cartons

Most paperwork gives a shipment total only. The full total is then written to the **first output row** (highest value, since rows sort by value descending) and every other row is left blank. The column still ties out to the shipment; it is explicitly not a per-part count and not an allocation.

Some templates do state a carton count per line. Those real counts are used per part instead — but only when **every** matched row has one **and** they add to the shipment total, so a partly-filled column can never masquerade as exact. The run log states which rule applied.

---

## 5. Document traps this tool defends against

All of these were hit on real shipments.

**Overlapping columns break flat text extraction.** On the reception report the tariff column runs to x=698 while the purchase-order column begins at x=692. Every text extractor sorts glyphs left to right and interleaves them:

```
7419809999 + 1013878363BU  ->  74198099919013878363BU
```

whose first ten characters are `7419809991` — wrong, but still shaped like a valid code, so it passes unnoticed. `_glyph_chains()` sorts a row's characters by x, chains each onto the one that **ends where it starts** (tolerance 0.6), then merges runs that continue an earlier chain. That reunites `848190050` with its final `3` across the intruding character and separates the two columns cleanly. Validated against a hand-checked shipment: 36 of 36 comparable codes exact.

**Documents do not share a row order.** A reception report once matched its packing list for the first ~32 rows then diverged, mismatching 87 of 123 rows when joined by position. Everything here joins on real keys.

**Hand-keyed part numbers contain typos.** Real examples from this shipment: `129929-U-BL` (a digit short of `1299295-U-BL`) and `1332370-U-2-MB` (an extra hyphen). Unknown parts are matched punctuation-blind, then at edit distance 1, and accepted **only when exactly one candidate fits**. Every rewrite is printed.

**Weight must be allocated per HS group, never globally.** A global split once gave a part 1,148.60 kg against an actual 66.00 kg — 17× over — because it was a tiny copper nut shipped alongside heavy assemblies.

**PDF backends are not interchangeable.** pdfplumber is tried first, pypdf second, and whichever actually yields text wins. A coordinate-capable backend is required for the glyph-chain work; without one the reception report's overlapping rows are **dropped rather than recorded with a mangled code**.

**There can be more than one Invoice Summary set.** This shipment has a 21-item summary mid-file and a 102-item summary at the end. Reading only the last one invents a discrepancy that does not exist. Every set is found and summed before anything is compared.

**Transcribe verbatim.** Truncated fields, unclosed parentheses and short invoice HS codes are left exactly as printed. Nothing is tidied, repaired, expanded, or filled in from a different page.

**Excel locks files.** If the output workbook is open, the save falls back to a timestamped filename and says so, rather than failing silently.

---

## 6. Reconciliation checks

Every check treats a **missing value as a failure, not a pass** — a check written as `if value and value != expected` silently passes when the value is empty and the column ships blank.

- qty × unit price == extended total, on every line
- each page's line sum vs. its printed items total
- every invoice and amount vs. **all** Invoice Summary sets combined, plus the printed grand totals
- invoices listed in a summary but not parsed, and parsed invoices absent from every summary
- packing-list qty vs. invoice qty, and packing rows vs. its own total row
- HS-group weights vs. the packing-list totals, and the written output vs. both
- reception-report qty and part list vs. the invoices
- every part has an HTS, each 8–12 digits, and a country of origin

Failures print as `[CHECK]` lines and are returned as the run's problem count, which is what triggers the update check.

---

## 7. Output workbook

Sheet `Summary`:

| Part # | Qty | HTS | Value | Cartons | Gross Weight (kg) | Net Weight (kg) | Country Origin | Charges | Zone |
|---|---|---|---|---|---|---|---|---|---|

- Sorted by value descending
- HTS stored as **text** (`@` format) so leading zeros and long codes are not mangled
- Weights rounded to 2 decimals, only at the end so the totals still land exactly
- Country as a 2-letter ISO code (`CN`, not `China`); an unmapped country is written as printed and flagged
- `Charges` = 3.00, `Zone` = P — constants at the top of `customs_processor.py`
- Bold TOTAL row summing Qty, Value, Cartons and both weights with live `=SUM()`. **HTS is never summed.**

Sheet `Review` appears **only** when a part could not be resolved, listing the part and why.

---

## 8. Versioning and release

`_version.py` holds `CUSTOMS_VERSION` and is the single source of truth.

To ship a new version:

1. Bump `CUSTOMS_VERSION`.
2. Add a `CHANGELOG.md` entry.
3. Commit, then push a tag `customs-vX.Y.Z`.

GitHub Actions verifies the tag matches `_version.py` (and fails fast if not), builds with the same flags as `build.bat`, and publishes `UnimexCustoms.exe` to a Release.

### How the self-update works

`updater.py` runs **only** from the frozen exe and **only** after a run with problems. It queries the repo's releases, filters to the `customs-v` prefix, picks the highest semver, and compares against the running version. A newer build is streamed to `UnimexCustoms.exe.new` beside the running exe and byte-count-verified against the size the API reported.

A running Windows exe is locked and cannot overwrite itself, so a throwaway `.cmd` is written to `%TEMP%` and spawned fully detached. It retries `move` — which fails while the exe is alive and succeeds the instant it exits — then relaunches and deletes itself. It sleeps with a full-path `ping` rather than `timeout`, which aborts immediately in a detached process with no console.

The updater never raises. Offline, DNS failure, GitHub 5xx, rate-limiting, a changed JSON shape — everything falls through to a one-line note and `return False`, so the user's real work is never blocked.

**Install location matters:** the exe must live somewhere user-writable. In Program Files it cannot replace itself.

**The repo must stay public.** A distributed exe carries no credentials, so the release check is an anonymous API call. Making this repo private returns 404 to every installed copy, and self-update silently stops working — the team would be back to hand-copying builds. Nothing sensitive is tracked here: the parts database, invoices, packing lists and reception reports are all gitignored, and `input\`, `output\`, `database\` and `logs\` hold only a `.gitkeep`.

---

## 9. Troubleshooting

| Symptom | Cause / fix |
|---|---|
| `no parts database found` | `database\` is empty. Drop the latest `Kohler Parts for Upload ....xlsx` there. |
| `refusing this packing list` | The PDF parsed, but the figures did not match the totals printed on it. The run continued without weights. Send the PDF to Andy — the template needs a look. |
| PDF packing list read but weights look odd | Check the log for "could not be corroborated": that document states no totals, so nothing could be verified against it. |
| `[SKIP] ...: no commercial invoice PDF found` | The invoice PDF is missing, or it is a scan with no text layer. |
| Part on the `Review` sheet, blank HTS | Either the part is not in the database, or its base number carries conflicting codes. Resolve in the database, then re-run. |
| `output gross != packing list` | The join dropped lines. Look for the `no packing-list row` / `no invoice line` `[CHECK]` lines above it. |
| `reception report qty != invoice qty` | Genuine mismatch, or the report has rows the parser dropped. Check the row count printed for the report. |
| Output file has a timestamp in its name | The normal file was open in Excel. |
| Nothing happens on double-click | The exe was quarantined by antivirus, or it is in a non-writable folder. |

Every run writes `logs\run_<timestamp>.log` containing exactly what the console showed. That file is what to send when something looks wrong.

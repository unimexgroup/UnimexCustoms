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
| Commercial invoices (PDF) | **Yes** | part, qty, unit price, extended value, PO, country of origin, invoice HS | pages containing "Invoice" and "P.O. No." |
| Packing list (`.xls`/`.xlsx`) | Strongly wanted | gross/net weight per HS group, cartons, shipment number | a header row with Purchase Order + Material + Qty |
| Warehouse reception report (PDF) | Optional | independent part/qty check, Mexican fracción | "REPORTE DE RECEPCION" / "RECONOCIMIENTO PREVIO" |
| Parts database (`.xlsx`) | **Yes** | `Product Key` → `HTS` | lives in `database\`, not `input\` |

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
3. **Join** invoice lines to packing rows on `(PO, part, qty)`, one to one, consuming each row at most once.
4. **Allocate weight inside each HS group** across only the parts in that group, in proportion to piece count, using largest-remainder rounding in integer hundredths.
5. **Resolve the HTS** for every part.
6. **Reconcile** (section 6). Nothing is suppressed.
7. **Write** the workbook.

### Weight allocation

`allocate_largest_remainder()` works in integer hundredths, so each column sums to its group total **exactly** and the file total matches the packing list exactly, with no `530.3199999999999` float artefacts. Leftover cents go to the largest fractional parts, ties broken by the larger quantity, so the result is identical run to run.

A group holding one part is exact. Any other group assumes equal weight per piece — an estimate, and the run log says so and names which groups are which.

### Cartons

The paperwork gives a shipment total only. The full total is written to the **first output row** (highest value, since rows sort by value descending) and every other row is left blank. The column still ties out to the shipment; it is explicitly not a per-part count and not an allocation.

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

---

## 9. Troubleshooting

| Symptom | Cause / fix |
|---|---|
| `no parts database found` | `database\` is empty. Drop the latest `Kohler Parts for Upload ....xlsx` there. |
| `[SKIP] ...: no commercial invoice PDF found` | The invoice PDF is missing, or it is a scan with no text layer. |
| Part on the `Review` sheet, blank HTS | Either the part is not in the database, or its base number carries conflicting codes. Resolve in the database, then re-run. |
| `output gross != packing list` | The join dropped lines. Look for the `no packing-list row` / `no invoice line` `[CHECK]` lines above it. |
| `reception report qty != invoice qty` | Genuine mismatch, or the report has rows the parser dropped. Check the row count printed for the report. |
| Output file has a timestamp in its name | The normal file was open in Excel. |
| Nothing happens on double-click | The exe was quarantined by antivirus, or it is in a non-writable folder. |

Every run writes `logs\run_<timestamp>.log` containing exactly what the console showed. That file is what to send when something looks wrong.

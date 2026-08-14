# UnimexCustoms — Shipment Paperwork to Customs Summary

A standalone Windows tool that turns one inbound shipment's paperwork into the customs team's Excel file: **one row per part number**, aggregated across every invoice in the shipment, with the **US HTS code** pulled from the Kohler parts database.

**Department:** US Customs Brokerage
**Current version:** v1.2.0

Replaces the hand-run `customs-summary-prompt.txt` workflow. Same output, same checks, no prompting.

## Automatic Updates

The exe **updates itself** from GitHub Releases — no hand-copying a new build to every machine. To avoid wasting time when the tool is working fine, an update check only runs after a run that had a problem (a failed reconciliation check, an unrecognized file, or a crash). When a newer build exists it downloads, installs, and restarts automatically, then reprocesses the input. Offline or GitHub unreachable → it prints a short note and keeps running the current version.

Install v1.0.0 **once**, manually, into a **user-writable folder** (e.g. a subfolder of Documents or `%LOCALAPPDATA%` — *not* Program Files, or the exe can't replace itself). Every later version arrives on its own. The first launch may show a one-time Windows "unknown publisher" SmartScreen prompt (the exe is unsigned); choose *More info → Run anyway*.

Publishing a new version is automated: push a `customs-vX.Y.Z` tag and GitHub Actions builds the exe and creates the Release (see [DOCUMENTATION.md](DOCUMENTATION.md) and `.github/workflows/release.yml`).

## Quick Start (End Users)

1. Copy `UnimexCustoms.exe` to any user-writable folder.
2. Double-click once — creates `input\`, `output\`, `database\`, `logs\` next to it.
3. Put the latest **`Kohler Parts for Upload ....xlsx`** in `database\`. The tool always uses the most recently modified file there, so refreshing the database is just dropping in the new file.
4. Put the shipment's paperwork into `input\` — either the carrier's **.zip exactly as received**, or the loose documents:
   - commercial invoices (PDF **or** Excel) — **required**
   - packing list (`.xls`, `.xlsx` or PDF) — supplies weights and cartons; often the same workbook as the invoice
   - warehouse reception report (PDF) — optional cross-check
   - waybills, FCRs, container checklists — harmless, they're recognized and skipped
5. Double-click again. `output\customs_summary_<shipment>.xlsx` appears.

Processing several shipments at once: drop in one zip per shipment, or give each one its own subfolder under `input\`. Duplicate copies of the same document (`X.pdf` and `X (1).pdf`) are detected and counted once.

Documents are identified by their **contents**, not their filenames — the team can drop them in exactly as received.

## Output

| Part # | Qty | HTS | Value | Cartons | Gross Weight (kg) | Net Weight (kg) | Country Origin | Charges | Zone |
|---|---|---|---|---|---|---|---|---|---|

Sorted by value descending, with a bold TOTAL row using live `=SUM()` formulas. HTS is stored as **text** so leading zeros survive, and is never summed.

Two things worth knowing before filing:

- **Cartons.** Most shipping paperwork states a shipment total only, never a per-part count. In that case the whole total sits on the **first row** so the column still ties out; it is not an allocation and the other rows are blank by design. When a packing list *does* state a count per line, those real per-part counts are used instead. The run log says which rule applied.
- **Per-row weights are allocated**, not measured, *when* the packing list states weight once per HS group: each group is split across its parts in proportion to piece count. Column totals are exact; a group holding a single part is exact; everything else assumes equal weight per piece. The run log names which groups came out which way. Packing lists that state a weight per line give exact figures throughout.
- **Net weight can be blank.** Some suppliers state gross only. Nothing is invented in its place; the run log says so plainly.

## Reconciliation

The tool refuses to stay quiet about a file that does not tie out. Every run checks qty × unit price against each line's extended total, each page against its printed items total, every invoice against **all** Invoice Summary sets combined, packing-list qty against invoice qty, HS-group weights against the packing-list totals, the reception report's parts and quantities against the invoices, and that every part has a plausible HTS and a country of origin. Anything that fails prints as a `[CHECK]` line and is counted; a missing value is treated as a failure, never as a pass.

If something looks wrong, send the relevant `logs\run_...log` file to Andy.

## Building the Executable

Requires Python 3.10+ and the packages in `requirements.txt`.

```bat
build.bat
```

Produces `dist\UnimexCustoms.exe`.

## Tech Stack

- **Language:** Python 3.10+
- **Libraries:** pandas, openpyxl (Excel), pdfplumber + pypdf (PDF), xlrd (legacy `.xls`)
- **Packaging:** PyInstaller (single-file `.exe`, no Python required on end-user machines)

## Project Files

| File | Purpose |
|---|---|
| `customs_processor.py` | The processor |
| `_version.py` | Single source of truth for the version |
| `updater.py` | Self-updater (checks GitHub Releases, swaps the exe in place) |
| `build.bat` | One-step build script |
| `requirements.txt` | Python dependencies |
| `UnimexCustoms.spec` | PyInstaller spec |
| `.github/workflows/release.yml` | CI: builds the exe and publishes a Release on a version tag |
| `DOCUMENTATION.md` | Full technical documentation and user manual |
| `CHANGELOG.md` | Version history |

## Full Documentation

See [DOCUMENTATION.md](DOCUMENTATION.md) for the complete technical reference, the document-parsing traps this tool defends against, and the user manual.

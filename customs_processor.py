"""
Unimex Customs - Shipment Paperwork to Customs Summary
=====================================================
Reads one inbound shipment's paperwork and writes the customs team's Excel
file: one row per part number, aggregated across every invoice in the shipment.

Inputs (identified by CONTENT, not filename -- filenames vary by client):
  * Commercial invoices (PDF)        -- REQUIRED. Part, qty, price, PO, origin.
  * Packing list (.xls/.xlsx)        -- gross/net weight, cartons, HS groups.
  * Warehouse reception report (PDF) -- optional; independent qty cross-check.
  * Parts database (database\\*.xlsx) -- Product Key -> HTS. Not in input\\.

Output columns:
  Part # | Qty | HTS | Value | Cartons | Gross Weight (kg) | Net Weight (kg) |
  Country Origin | Charges | Zone

Run:
    python customs_processor.py                # ./input -> ./output
    python customs_processor.py /path/in /path/out

Drop one shipment's documents straight into input\\, or give each shipment its
own subfolder under input\\ to process several in one go.
"""

from __future__ import annotations

import math
import os
import re
import sys
import traceback
import warnings
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

warnings.filterwarnings("ignore", message="Workbook contains no default style")

# Single source of truth for the version (see _version.py). Guarded so a stray
# copy without _version.py alongside it still runs.
try:
    from _version import CUSTOMS_VERSION as __version__
except Exception:
    __version__ = "0.0.0"

# ---------------------------------------------------------------------------
# Constants the customs team confirmed. Change here, not in the code below.
# ---------------------------------------------------------------------------
CHARGES = 3.00
ZONE = "P"

OUT_COLUMNS = [
    "Part #", "Qty", "HTS", "Value", "Cartons",
    "Gross Weight (kg)", "Net Weight (kg)", "Country Origin", "Charges", "Zone",
]

# Invoices spell the country out; the filing wants the 2-letter ISO code.
COUNTRY_ISO = {
    "CHINA": "CN", "MEXICO": "MX", "UNITED STATES": "US", "USA": "US",
    "INDIA": "IN", "THAILAND": "TH", "VIETNAM": "VN", "TAIWAN": "TW",
    "CANADA": "CA", "GERMANY": "DE", "ITALY": "IT", "SPAIN": "ES",
    "FRANCE": "FR", "JAPAN": "JP", "KOREA": "KR", "SOUTH KOREA": "KR",
    "MALAYSIA": "MY", "INDONESIA": "ID", "BRAZIL": "BR", "POLAND": "PL",
    "PORTUGAL": "PT", "TURKEY": "TR", "UNITED KINGDOM": "GB",
}

# Packing-list header keywords -> canonical column. First header whose
# normalized text contains any of these tokens wins.
PL_COLUMN_TOKENS = {
    "po":          ["purchasedorder", "purchaseorder", "po#", "pono", "order#"],
    "part":        ["materialno", "material", "partno", "part#", "sku", "item#"],
    "description": ["description", "descripcion"],
    "qty":         ["qty", "quantity", "pcs", "cantidad"],
    "hs":          ["hsncode", "hscode", "hsn", "hs", "tariff", "fraccion"],
    "gross":       ["grossweight", "gross", "pesobruto"],
    "net":         ["netweight", "net", "pesoneto"],
    "volume":      ["volume", "cbm", "volumen"],
    "cartons":     ["carton", "ctn", "cajas", "bultos", "package"],
}

# ---------------------------------------------------------------------------
# Regexes. Kept whitespace-tolerant on purpose (TRAP 6): PDF backends pad
# columns differently, so a literal space would work on one and fail on another.
# ---------------------------------------------------------------------------
# "10   24  1308865-U-BV   29.87   716.88"
RE_ITEM_LINE = re.compile(
    r"^(\d{1,4})\s+([\d,]+)\s+([A-Za-z0-9][A-Za-z0-9\-./#]*)\s+"
    r"([\d,]*\.?\d+)\s+([\d,]*\.?\d+)\s*$"
)
# The HS code the invoice prints, on its own line under the description.
RE_HS_ONLY = re.compile(r"^(\d{6,12})$")
RE_ORIGIN = re.compile(r"Country\s+of\s+Origin\s*[-:]\s*(.+?)\s*$", re.I)
RE_ITEMS_TOTAL = re.compile(r"Items?\s+total\s+([\d,]+\.\d{2})", re.I)
# Header/value pairs are not always on adjacent lines (TRAP 6): scan forward.
RE_NUM_DATE = re.compile(r"^(\S+)\s+(\d{1,2}/\d{1,2}/\d{4})\b")
# Summary page row: "19  9079663585  1030012622  1013749677BU  18,487.98"
RE_SUMMARY_ROW = re.compile(
    r"^(\d{1,4})\s+(\d{6,})\s+(\d{6,})\s+(\S+)\s+([\d,]+\.\d{2})\s*$"
)
# Reception report line, e.g.
# "CHN3609436 (L/E) 1 830915 ARANDELA DE ACERO 900.00 PIEZA 900.00 PIEZA 7318229199 S/N"
RE_RECEPTION_ROW = re.compile(
    r"\(L/E\)\s+(\d{1,4})\s+(\S+)\s+.*?([\d,]+\.\d{2})\s+([A-Za-zÁÉÍÓÚÑ]+)\s+"
    r"([\d,]+\.\d{2})\s+([A-Za-zÁÉÍÓÚÑ]+)\s+(\d{6,12})\b"
)
RE_SHIPMENT_NO = re.compile(r"shipment\s*(?:no\.?|number)", re.I)


# ---------------------------------------------------------------------------
# Small helpers
# ---------------------------------------------------------------------------
def base_dir() -> Path:
    """
    Folder the script (or .exe) lives in. When packaged with PyInstaller and
    double-clicked, the current working directory can be C:\\Windows\\System32,
    so we always anchor input/output/database to the .exe's actual location.
    """
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).parent


class TeeLogger:
    """Write everything printed to console to a log file as well."""
    def __init__(self, log_path: Path):
        self.terminal = sys.stdout
        self.log = open(log_path, "w", encoding="utf-8")
    def write(self, msg: str) -> None:
        self.terminal.write(msg)
        self.log.write(msg)
        self.log.flush()
    def flush(self) -> None:
        self.terminal.flush()
        self.log.flush()
    def close(self) -> None:
        try:
            self.log.close()
        except Exception:
            pass


def pause_for_user() -> None:
    """Wait for a keypress so the console doesn't flash shut on double-click."""
    if not sys.stdout.isatty():
        return
    try:
        input("\nPress ENTER to close...")
    except (EOFError, KeyboardInterrupt):
        pass


def _num(text: object) -> float | None:
    """'1,234.56' -> 1234.56. Returns None for blanks and non-numbers."""
    if text is None:
        return None
    if isinstance(text, (int, float)) and not isinstance(text, bool):
        return None if (isinstance(text, float) and math.isnan(text)) else float(text)
    s = str(text).strip().replace(",", "")
    if not s:
        return None
    try:
        return float(s)
    except ValueError:
        return None


def _norm_key(text: object) -> str:
    """Uppercase, whitespace-stripped. The join key for part numbers and POs."""
    return "".join(str(text).split()).upper() if text is not None else ""


def _depunct(text: object) -> str:
    """Uppercase alphanumerics only -- for punctuation-blind part matching."""
    return re.sub(r"[^A-Z0-9]", "", str(text).upper()) if text is not None else ""


def _edit_distance_1(a: str, b: str) -> bool:
    """True when a and b differ by exactly one insert, delete or substitution."""
    la, lb = len(a), len(b)
    if abs(la - lb) > 1:
        return False
    if la == lb:
        diff = sum(1 for x, y in zip(a, b) if x != y)
        return diff == 1
    if la > lb:
        a, b, la, lb = b, a, lb, la
    i = j = 0
    skipped = False
    while i < la and j < lb:
        if a[i] != b[j]:
            if skipped:
                return False
            skipped = True
            j += 1
            continue
        i += 1
        j += 1
    return True


def _iso_country(name: str) -> tuple[str, bool]:
    """('China') -> ('CN', True). Unknown names come back as-is with False so
    the caller can flag them instead of silently inventing a code."""
    raw = str(name).strip().rstrip(".").upper()
    if len(raw) == 2 and raw.isalpha():
        return raw, True
    if raw in COUNTRY_ISO:
        return COUNTRY_ISO[raw], True
    return str(name).strip(), False


def allocate_largest_remainder(total: float, weights: list[float]) -> list[float]:
    """
    Split `total` across `weights` in proportion, rounded to 2 decimals, so the
    parts sum to `total` EXACTLY. Works in integer hundredths -- floats would
    reintroduce the 530.3199999999999 artefacts the filing must not contain.
    """
    if not weights:
        return []
    grand = sum(weights)
    if grand <= 0:
        return [0.0] * len(weights)
    cents_total = int(round(total * 100))
    raw = [cents_total * w / grand for w in weights]
    floors = [int(math.floor(r)) for r in raw]
    short = cents_total - sum(floors)
    # Hand the leftover cents to the largest fractional parts, breaking ties by
    # the bigger weight so the result is deterministic run to run.
    order = sorted(range(len(raw)), key=lambda i: (-(raw[i] - floors[i]), -weights[i], i))
    for i in order[:max(short, 0)]:
        floors[i] += 1
    return [f / 100.0 for f in floors]


# ---------------------------------------------------------------------------
# PDF text extraction. Backends are NOT interchangeable (TRAP 5) -- try each
# and keep whichever actually yields text.
# ---------------------------------------------------------------------------
def extract_pdf_pages(path: Path) -> list[str]:
    """Return the text of every page, best available backend first."""
    attempts: list[str] = []
    try:
        import pdfplumber
        with pdfplumber.open(str(path)) as pdf:
            pages = [(p.extract_text() or "") for p in pdf.pages]
        if any(t.strip() for t in pages):
            return pages
        attempts.append("pdfplumber returned no text")
    except Exception as e:
        attempts.append(f"pdfplumber failed ({type(e).__name__})")

    try:
        from pypdf import PdfReader
        reader = PdfReader(str(path))
        pages = [(pg.extract_text() or "") for pg in reader.pages]
        if any(t.strip() for t in pages):
            print(f"  [INFO] {path.name}: read with pypdf ({attempts[0]}).")
            return pages
        attempts.append("pypdf returned no text")
    except Exception as e:
        attempts.append(f"pypdf failed ({type(e).__name__})")

    raise RuntimeError(f"could not read any text from {path.name}: {'; '.join(attempts)}")


# ---------------------------------------------------------------------------
# Invoices
# ---------------------------------------------------------------------------
@dataclass
class InvoiceLine:
    invoice: str
    po: str
    item: str
    part: str
    qty: float
    unit_price: float
    ext_total: float
    invoice_hs: str
    country_raw: str
    page: int


@dataclass
class SummaryEntry:
    summary_invoice: str
    item: str
    invoice: str
    sales_order: str
    po: str
    total: float


@dataclass
class InvoiceDoc:
    path: Path
    lines: list[InvoiceLine] = field(default_factory=list)
    summaries: list[SummaryEntry] = field(default_factory=list)
    summary_totals: dict[str, float] = field(default_factory=dict)
    page_total_mismatches: list[str] = field(default_factory=list)
    invoice_pages: int = 0
    summary_pages: int = 0


def _scan_forward(lines: list[str], start: int, pattern: re.Pattern, span: int = 8):
    """Find the first match of `pattern` within `span` lines after `start`.
    A header and its value are not always adjacent (TRAP 6)."""
    for line in lines[start + 1: start + 1 + span]:
        m = pattern.search(line)
        if m:
            return m
    return None


def parse_invoice_page(text: str, page_no: int) -> list[InvoiceLine]:
    """Pull every line item off one invoice page."""
    lines = [ln.rstrip() for ln in text.splitlines()]

    invoice_no = ""
    for i, ln in enumerate(lines):
        if re.search(r"\bNumber\b.*\bDate\b", ln):
            m = _scan_forward(lines, i, RE_NUM_DATE)
            if m and m.group(1).isdigit():
                invoice_no = m.group(1)
                break

    po = ""
    for i, ln in enumerate(lines):
        if re.search(r"P\.?\s*O\.?\s*No", ln, re.I):
            m = _scan_forward(lines, i, RE_NUM_DATE)
            if m:
                po = m.group(1)
                break

    # Locate the item lines first, then read each one's trailing block (the
    # description, the HS code and the country line) up to the next item.
    item_at: list[tuple[int, re.Match]] = []
    for i, ln in enumerate(lines):
        m = RE_ITEM_LINE.match(ln.strip())
        if m:
            item_at.append((i, m))

    out: list[InvoiceLine] = []
    for n, (idx, m) in enumerate(item_at):
        end = item_at[n + 1][0] if n + 1 < len(item_at) else len(lines)
        block = lines[idx + 1: end]
        hs = ""
        country = ""
        for bl in block:
            s = bl.strip()
            if not hs:
                hm = RE_HS_ONLY.match(s)
                if hm:
                    hs = hm.group(1)
            om = RE_ORIGIN.search(s)
            if om and not country:
                country = om.group(1).strip()
        out.append(InvoiceLine(
            invoice=invoice_no,
            po=po,
            item=m.group(1),
            part=_norm_key(m.group(3)),
            qty=float(m.group(2).replace(",", "")),
            unit_price=float(m.group(4).replace(",", "")),
            ext_total=float(m.group(5).replace(",", "")),
            invoice_hs=hs,
            country_raw=country,
            page=page_no,
        ))
    return out


def parse_invoices(path: Path) -> InvoiceDoc:
    """
    Read the invoice PDF. Pages titled 'Invoice Summary' are roll-ups, not
    invoices: they are collected separately for reconciliation. There may be
    MORE THAN ONE summary set in different places in the file -- every set is
    kept, keyed by its own summary invoice number.
    """
    doc = InvoiceDoc(path=path)
    pages = extract_pdf_pages(path)

    for pno, text in enumerate(pages, start=1):
        if re.search(r"Invoice\s+Summary", text, re.I):
            doc.summary_pages += 1
            lines = [ln.rstrip() for ln in text.splitlines()]
            sum_no = ""
            for i, ln in enumerate(lines):
                if re.search(r"\bNumber\b.*\bDate\b", ln):
                    m = _scan_forward(lines, i, RE_NUM_DATE)
                    if m and m.group(1).isdigit():
                        sum_no = m.group(1)
                        break
            for ln in lines:
                m = RE_SUMMARY_ROW.match(ln.strip())
                if m:
                    doc.summaries.append(SummaryEntry(
                        summary_invoice=sum_no,
                        item=m.group(1),
                        invoice=m.group(2),
                        sales_order=m.group(3),
                        po=m.group(4),
                        total=float(m.group(5).replace(",", "")),
                    ))
                tm = RE_ITEMS_TOTAL.search(ln)
                if tm and sum_no:
                    doc.summary_totals[sum_no] = float(tm.group(1).replace(",", ""))
            continue

        page_lines = parse_invoice_page(text, pno)
        if not page_lines:
            continue
        doc.invoice_pages += 1
        doc.lines.extend(page_lines)

        # Each page prints its own line-items total; check ours against it.
        printed = None
        for ln in text.splitlines():
            tm = RE_ITEMS_TOTAL.search(ln)
            if tm:
                printed = float(tm.group(1).replace(",", ""))
        if printed is not None:
            ours = round(sum(l.ext_total for l in page_lines), 2)
            if abs(ours - printed) > 0.01:
                doc.page_total_mismatches.append(
                    f"page {pno} (invoice {page_lines[0].invoice}): "
                    f"lines sum to {ours:,.2f} but page prints {printed:,.2f}")
    return doc


# ---------------------------------------------------------------------------
# Packing list
# ---------------------------------------------------------------------------
@dataclass
class PackingRow:
    po: str
    part: str
    qty: float
    hs: str
    group: int
    consumed: bool = False


@dataclass
class PackingList:
    path: Path
    rows: list[PackingRow] = field(default_factory=list)
    group_gross: dict[int, float] = field(default_factory=dict)
    group_net: dict[int, float] = field(default_factory=dict)
    total_qty: float | None = None
    total_gross: float | None = None
    total_net: float | None = None
    cartons: float | None = None
    carton_source: str = ""
    shipment_no: str = ""
    notes: list[str] = field(default_factory=list)


def _read_sheet(path: Path, sheet) -> pd.DataFrame:
    engine = "xlrd" if path.suffix.lower() == ".xls" else "openpyxl"
    return pd.read_excel(path, sheet_name=sheet, header=None, dtype=object, engine=engine)


def _map_pl_columns(header_row: list[object]) -> dict[str, int]:
    """Map canonical names to column indexes by keyword. Longer tokens are
    tried first so 'grossweight' can't be claimed by the bare 'gross' rule."""
    normed = ["".join(str(h).lower().split()) if h is not None else "" for h in header_row]
    found: dict[str, int] = {}
    for canon, tokens in PL_COLUMN_TOKENS.items():
        for tok in tokens:
            for idx, h in enumerate(normed):
                if idx in found.values():
                    continue
                if tok in h:
                    found[canon] = idx
                    break
            if canon in found:
                break
    return found


def parse_packing_list(path: Path) -> PackingList:
    """
    Read the packing list. Two layout habits matter and both are common:
      * gross/net weight is stated ONCE per HS group, on the group's first row
        and carried down implicitly -- so a non-blank gross starts a new group;
      * cartons/volume are stated once for the WHOLE shipment, not per line.
    """
    pl = PackingList(path=path)
    xl = pd.ExcelFile(path, engine="xlrd" if path.suffix.lower() == ".xls" else "openpyxl")

    best = None
    for sheet in xl.sheet_names:
        df = _read_sheet(path, sheet)
        for r in range(min(40, len(df))):
            cols = _map_pl_columns(list(df.iloc[r]))
            if {"po", "part", "qty"} <= set(cols):
                best = (sheet, df, r, cols)
                break
        if best:
            break
    if not best:
        raise RuntimeError(f"{path.name}: could not find a packing-list header row "
                           f"(need Purchase Order / Material / Qty columns)")

    sheet, df, hrow, cols = best

    # Shipment number, printed above the table.
    for r in range(hrow + 1):
        for c in range(df.shape[1]):
            if RE_SHIPMENT_NO.search(str(df.iat[r, c])):
                for rr in range(r + 1, min(r + 4, len(df))):
                    v = df.iat[rr, c]
                    if v is not None and str(v).strip() and str(v).strip().lower() != "nan":
                        pl.shipment_no = str(v).strip().split(".")[0]
                        break
            if pl.shipment_no:
                break
        if pl.shipment_no:
            break

    def cell(r: int, name: str):
        idx = cols.get(name)
        return df.iat[r, idx] if idx is not None and idx < df.shape[1] else None

    group = -1
    carton_values: list[float] = []
    for r in range(hrow + 1, len(df)):
        po = _norm_key(cell(r, "po")) if cell(r, "po") is not None else ""
        part = _norm_key(cell(r, "part")) if cell(r, "part") is not None else ""
        qty = _num(cell(r, "qty"))
        gross = _num(cell(r, "gross"))
        net = _num(cell(r, "net"))
        ctn = _num(cell(r, "cartons"))
        po = "" if po.lower() in {"nan", "none"} else po
        part = "" if part.lower() in {"nan", "none"} else part

        if not po and not part:
            # Trailing totals row: no PO, no part, but the columns still add up.
            if qty is not None:
                pl.total_qty = qty
                pl.total_gross = gross
                pl.total_net = net
                if ctn is not None:
                    pl.cartons = ctn
                    pl.carton_source = "packing-list total row"
            continue
        if qty is None:
            continue

        if gross is not None:
            group += 1
            pl.group_gross[group] = gross
            pl.group_net[group] = net if net is not None else 0.0
        if group < 0:  # rows before any stated weight
            group = 0
            pl.group_gross.setdefault(group, 0.0)
            pl.group_net.setdefault(group, 0.0)
        if ctn is not None:
            carton_values.append(ctn)

        pl.rows.append(PackingRow(
            po=po, part=part, qty=qty,
            hs=str(cell(r, "hs") or "").strip(), group=group,
        ))

    if pl.cartons is None and carton_values:
        distinct = {round(v, 4) for v in carton_values}
        if len(distinct) == 1:
            # One number for the whole shipment, printed on the first row --
            # summing the column here would double-count it.
            pl.cartons = carton_values[0]
            pl.carton_source = "single shipment-wide carton figure"
        else:
            pl.cartons = sum(carton_values)
            pl.carton_source = "sum of the per-row carton column"

    # An HS group whose code is not constant means the group boundaries (drawn
    # from where weights are stated) don't line up with the tariff grouping.
    for g in sorted(pl.group_gross):
        codes = {r.hs for r in pl.rows if r.group == g and r.hs}
        if len(codes) > 1:
            pl.notes.append(f"weight group {g + 1} spans more than one HS code: "
                            f"{', '.join(sorted(codes))}")
    return pl


# ---------------------------------------------------------------------------
# Reception report (optional cross-check)
# ---------------------------------------------------------------------------
@dataclass
class ReceptionRow:
    partida: str
    part: str
    qty_doc: float
    qty_physical: float
    fraccion: str


@dataclass
class ReceptionReport:
    path: Path
    rows: list[ReceptionRow] = field(default_factory=list)
    gross_weight: float | None = None
    packages: str = ""
    guide: str = ""


def _glyph_chains(chars: list[dict], tol: float = 0.6) -> list[tuple[str, float]]:
    """
    Rebuild one printed row's cells by following CONTIGUOUS GLYPH RUNS.

    Two columns can physically overlap in the page layout -- on this reception
    report the tariff column runs to x=698 while the purchase-order column
    begins at x=692. Every text extractor sorts glyphs left to right and
    interleaves them:
        7419809999 + 1013878363BU  ->  74198099919013878363BU
    whose first ten characters are 7419809991: wrong, but still shaped like a
    valid code, so it passes unnoticed (TRAP 1).

    Sorting by x and chaining each character onto the one that ENDS where it
    STARTS separates the two columns cleanly, because a character of the tariff
    column continues the tariff column's chain, not the PO's.

    Returns [(text, x0)] in left-to-right order of each chain's start.
    """
    if not chars:
        return []
    ordered = sorted(chars, key=lambda c: c["x0"])
    runs: list[list[dict]] = []
    cur = [ordered[0]]
    for prev, nxt in zip(ordered, ordered[1:]):
        if abs(nxt["x0"] - prev["x1"]) <= tol:
            cur.append(nxt)
        else:
            runs.append(cur)
            cur = [nxt]
    runs.append(cur)

    # A run that starts exactly where an earlier run ended continues that same
    # chain -- this is the step that reunites '848190050' with its final '3'
    # across the interleaved character from the neighbouring column.
    chains: list[list[dict]] = []
    used = [False] * len(runs)
    for i, run in enumerate(runs):
        if used[i]:
            continue
        used[i] = True
        chain = list(run)
        extended = True
        while extended:
            extended = False
            for j, other in enumerate(runs):
                if used[j]:
                    continue
                if abs(other[0]["x0"] - chain[-1]["x1"]) <= tol:
                    chain.extend(other)
                    used[j] = True
                    extended = True
        chains.append(chain)
    chains.sort(key=lambda ch: ch[0]["x0"])
    return [("".join(c["text"] for c in ch).strip(), ch[0]["x0"]) for ch in chains]


def _reception_rows_by_glyph(path: Path) -> list[ReceptionRow]:
    """Read the report's line items cell by cell using glyph runs."""
    import pdfplumber

    rows: list[ReceptionRow] = []
    with pdfplumber.open(str(path)) as pdf:
        for page in pdf.pages:
            by_line: dict[float, list[dict]] = {}
            for ch in page.chars:
                by_line.setdefault(round(ch["top"], 0), []).append(ch)
            for top in sorted(by_line):
                cells = _glyph_chains(by_line[top])
                texts = [t for t, _ in cells]
                marker = next((i for i, t in enumerate(texts) if "(L/E)" in t), None)
                if marker is None:
                    continue
                rest = texts[marker + 1:]
                if len(rest) < 3 or not rest[0].isdigit():
                    continue
                qtys = [t for t in rest if re.fullmatch(r"[\d,]+\.\d{1,3}", t)]
                if not qtys:
                    continue
                # The tariff column is printed before the purchase-order column,
                # so the FIRST bare digit run after the quantities is the code.
                after_qty = rest[rest.index(qtys[-1]) + 1:]
                frac = next((t for t in after_qty if re.fullmatch(r"\d{6,12}", t)), "")
                rows.append(ReceptionRow(
                    partida=rest[0],
                    part=_norm_key(rest[1]),
                    qty_doc=float(qtys[0].replace(",", "")),
                    qty_physical=float(qtys[-1].replace(",", "")),
                    fraccion=frac,
                ))
    return rows


def parse_reception_report(path: Path) -> ReceptionReport:
    """
    Parse the warehouse intake report. Since the HTS now comes from the parts
    database, this document is only a cross-check: part numbers and quantities.
    The fraccion is still read so a Mexican-vs-US code difference can be
    reported, never to fill a cell.
    """
    rep = ReceptionReport(path=path)
    try:
        rep.rows = _reception_rows_by_glyph(path)
    except Exception as e:
        print(f"  [INFO] {path.name}: glyph reconstruction unavailable "
              f"({type(e).__name__}); falling back to flat text.")

    pages = extract_pdf_pages(path)
    if not rep.rows:
        # Flat-text fallback. Rows whose tariff column overlaps the PO column
        # simply will not match here -- that is deliberate. A row is dropped
        # rather than recorded with a mangled code.
        for text in pages:
            for ln in text.splitlines():
                m = RE_RECEPTION_ROW.search(ln)
                if m:
                    rep.rows.append(ReceptionRow(
                        partida=m.group(1),
                        part=_norm_key(m.group(2)),
                        qty_doc=float(m.group(3).replace(",", "")),
                        qty_physical=float(m.group(5).replace(",", "")),
                        fraccion=m.group(7),
                    ))

    for text in pages:
        for ln in text.splitlines():
            gm = re.search(r"Peso\s+bruto\s+([\d,]+\.?\d*)", ln, re.I)
            if gm and rep.gross_weight is None:
                rep.gross_weight = float(gm.group(1).replace(",", ""))
            bm = re.search(r"Bultos\s+(.+)$", ln, re.I)
            if bm and not rep.packages:
                rep.packages = bm.group(1).strip()
            gu = re.search(r"Gu[ií]?[óo]?a\(s\)\s+(\S+)", ln, re.I)
            if gu and not rep.guide:
                rep.guide = gu.group(1).strip()

    # Every code should come out the same length; a short one means a chain was
    # cut and the value cannot be trusted.
    odd = {r.fraccion for r in rep.rows if r.fraccion and not (8 <= len(r.fraccion) <= 12)}
    if odd:
        print(f"  [WARN] {path.name}: implausible tariff code(s) read from the report: "
              f"{', '.join(sorted(odd))}")
    return rep


# ---------------------------------------------------------------------------
# Parts database  ->  HTS
# ---------------------------------------------------------------------------
@dataclass
class PartsDB:
    path: Path
    by_key: dict[str, str] = field(default_factory=dict)          # exact key -> HTS
    by_depunct: dict[str, set] = field(default_factory=dict)      # A1B2 -> {HTS}
    by_stem: dict[str, set] = field(default_factory=dict)         # key minus finish -> {HTS}
    by_base: dict[str, set] = field(default_factory=dict)         # leading segment -> {HTS}


def _segments(key: str) -> list[str]:
    return [s for s in re.split(r"[-\s]+", key.strip().upper()) if s]


def load_parts_db(db_dir: Path) -> PartsDB:
    """
    Load the newest Product Key -> HTS workbook from database\\.

    Newest = most recently modified, so dropping in a fresh export is all the
    team has to do; nothing needs rebuilding or renaming.
    """
    if not db_dir.exists():
        raise RuntimeError(f"no database folder at {db_dir}")
    candidates = [p for p in db_dir.iterdir()
                  if p.suffix.lower() in {".xlsx", ".xlsm", ".xls"}
                  and not p.name.startswith("~$")]
    if not candidates:
        raise RuntimeError(f"no parts database found in {db_dir} "
                           f"(expected an .xlsx with Product Key and HTS columns)")
    path = max(candidates, key=lambda p: p.stat().st_mtime)

    df = pd.read_excel(path, dtype=str,
                       engine="xlrd" if path.suffix.lower() == ".xls" else "openpyxl")
    normed = {"".join(str(c).lower().split()): c for c in df.columns}
    key_col = next((normed[k] for k in normed if "productkey" in k or k in {"part", "partno", "sku"}), None)
    hts_col = next((normed[k] for k in normed if k.startswith("hts")), None)
    if key_col is None or hts_col is None:
        raise RuntimeError(f"{path.name}: expected 'Product Key' and 'HTS' columns, "
                           f"found {list(df.columns)[:6]}")

    db = PartsDB(path=path)
    for key, hts in zip(df[key_col], df[hts_col]):
        if key is None or hts is None:
            continue
        k = _norm_key(key)
        h = str(hts).strip()
        if not k or not h or h.lower() == "nan":
            continue
        db.by_key[k] = h
        db.by_depunct.setdefault(_depunct(k), set()).add(h)
        segs = _segments(k)
        if segs:
            db.by_base.setdefault(segs[0], set()).add(h)
            if len(segs) > 1:
                db.by_stem.setdefault("-".join(segs[:-1]), set()).add(h)
    return db


def lookup_hts(part: str, db: PartsDB) -> tuple[str, str, str]:
    """
    Resolve a part number to its HTS code.

    Returns (hts, how, note). `hts` is "" when nothing safe could be resolved,
    in which case `note` explains why -- the cell ships BLANK and the part is
    listed for review rather than being filled with a guess.

    Match order, most specific first:
      1. exact Product Key
      2. same key ignoring punctuation
      3. same key except the finish suffix   (1522807-U-BV -> 1522807-U-*)
      4. same leading part number            (1522807-U-BV -> 1522807-*)
    Steps 3 and 4 are only accepted when every candidate agrees on the code.
    """
    key = _norm_key(part)
    if key in db.by_key:
        return db.by_key[key], "exact", ""

    dp = _depunct(key)
    if dp in db.by_depunct and len(db.by_depunct[dp]) == 1:
        return next(iter(db.by_depunct[dp])), "punctuation-insensitive", ""

    segs = _segments(key)
    if segs:
        stem = "-".join(segs[:-1]) if len(segs) > 1 else ""
        if stem and stem in db.by_stem:
            codes = db.by_stem[stem]
            if len(codes) == 1:
                return next(iter(codes)), f"matched {stem}-* (finish suffix ignored)", ""
        base = segs[0]
        if base in db.by_base:
            codes = db.by_base[base]
            if len(codes) == 1:
                return next(iter(codes)), f"matched {base}-* (base part number)", ""
            return "", "conflict", (f"base {base} carries {len(codes)} different HTS codes "
                                    f"in the database ({', '.join(sorted(codes))}) and no "
                                    f"exact key exists -- resolve before filing")
    return "", "missing", "part number is not in the parts database"


# ---------------------------------------------------------------------------
# Join, allocate, aggregate
# ---------------------------------------------------------------------------
@dataclass
class ShipmentResult:
    rows: list[dict] = field(default_factory=list)
    problems: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    review: list[dict] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)
    exact_groups: list[str] = field(default_factory=list)
    estimated_groups: list[str] = field(default_factory=list)


def join_invoice_to_packing(lines: list[InvoiceLine], pl: PackingList,
                            res: ShipmentResult) -> dict[int, int]:
    """
    Match every invoice line to its packing-list row on (PO, part, qty) and
    return {index into lines -> weight group}.

    Never joins by row position: documents do not share a row order (TRAP 2).
    Hand-keyed part numbers carry typos (TRAP 3), so an unmatched part is
    retried punctuation-blind and then at edit distance 1 -- accepted only when
    exactly one candidate fits, and every rewrite is reported.
    """
    by_triple: dict[tuple, list[PackingRow]] = {}
    for row in pl.rows:
        by_triple.setdefault((row.po, row.part, row.qty), []).append(row)

    group_of: dict[int, int] = {}
    for i, ln in enumerate(lines):
        cand = by_triple.get((ln.po, ln.part, ln.qty), [])
        row = next((r for r in cand if not r.consumed), None)

        if row is None:  # same PO+part, different qty
            same = [r for r in pl.rows
                    if not r.consumed and r.po == ln.po and r.part == ln.part]
            if len(same) == 1:
                row = same[0]
                res.warnings.append(
                    f"invoice {ln.invoice} PO {ln.po} part {ln.part}: invoice qty "
                    f"{ln.qty:,.0f} vs packing list {row.qty:,.0f} -- matched anyway")

        if row is None:  # typo'd part number
            pool = [r for r in pl.rows if not r.consumed and r.po == ln.po]
            hits = [r for r in pool if _depunct(r.part) == _depunct(ln.part)]
            if not hits:
                hits = [r for r in pool if _edit_distance_1(_depunct(r.part), _depunct(ln.part))]
            if len(hits) == 1:
                row = hits[0]
                res.warnings.append(
                    f"part number rewritten for matching: invoice '{ln.part}' -> "
                    f"packing list '{row.part}' (PO {ln.po}, invoice {ln.invoice})")
            elif len(hits) > 1:
                res.problems.append(
                    f"invoice {ln.invoice} part {ln.part}: {len(hits)} possible packing-list "
                    f"rows, none unique -- left unmatched")

        if row is None:
            res.problems.append(
                f"invoice {ln.invoice} PO {ln.po} part {ln.part} qty {ln.qty:,.0f}: "
                f"no packing-list row")
            continue

        row.consumed = True
        group_of[i] = row.group

    for row in pl.rows:
        if not row.consumed:
            res.problems.append(
                f"packing-list row PO {row.po} part {row.part} qty {row.qty:,.0f}: "
                f"no invoice line")
    return group_of


def build_rows(lines: list[InvoiceLine], pl: PackingList | None,
               group_of: dict[int, int], db: PartsDB,
               res: ShipmentResult) -> list[dict]:
    """Aggregate to one row per part, allocating weight inside each HS group."""
    # Per (group, part): qty and value.
    per_group: dict[int, dict[str, dict]] = {}
    per_part: dict[str, dict] = {}
    for i, ln in enumerate(lines):
        p = per_part.setdefault(ln.part, {"qty": 0.0, "value": 0.0,
                                          "countries": set(), "inv_hs": set()})
        p["qty"] += ln.qty
        p["value"] += ln.ext_total
        if ln.country_raw:
            p["countries"].add(ln.country_raw)
        if ln.invoice_hs:
            p["inv_hs"].add(ln.invoice_hs)

        g = group_of.get(i)
        if g is None:
            continue
        gp = per_group.setdefault(g, {}).setdefault(ln.part, {"qty": 0.0})
        gp["qty"] += ln.qty

    gross_by_part: dict[str, float] = {}
    net_by_part: dict[str, float] = {}
    if pl is not None:
        for g, parts in sorted(per_group.items()):
            names = list(parts)
            qtys = [parts[n]["qty"] for n in names]
            gross = allocate_largest_remainder(pl.group_gross.get(g, 0.0), qtys)
            net = allocate_largest_remainder(pl.group_net.get(g, 0.0), qtys)
            for n, gv, nv in zip(names, gross, net):
                gross_by_part[n] = round(gross_by_part.get(n, 0.0) + gv, 2)
                net_by_part[n] = round(net_by_part.get(n, 0.0) + nv, 2)
            label = (f"group {g + 1} ({pl.group_gross.get(g, 0.0):,.2f} kg gross, "
                     f"{len(names)} part(s))")
            (res.exact_groups if len(names) == 1 else res.estimated_groups).append(label)

    rows: list[dict] = []
    for part, agg in per_part.items():
        hts, how, note = lookup_hts(part, db)
        if not hts:
            res.review.append({"Part #": part, "Issue": how, "Detail": note})
            res.problems.append(f"{part}: no HTS ({note})")
        elif how != "exact":
            res.notes.append(f"{part}: HTS {hts} {how}")

        iso, ok = ("", True)
        if agg["countries"]:
            if len(agg["countries"]) > 1:
                res.problems.append(
                    f"{part}: invoices disagree on country of origin "
                    f"({', '.join(sorted(agg['countries']))})")
            iso, ok = _iso_country(sorted(agg["countries"])[0])
            if not ok:
                res.warnings.append(f"{part}: country '{iso}' has no ISO code mapping "
                                    f"-- written as printed")
        else:
            res.problems.append(f"{part}: no country of origin on the invoices")

        rows.append({
            "Part #": part,
            "Qty": int(agg["qty"]) if float(agg["qty"]).is_integer() else agg["qty"],
            "HTS": hts,
            "Value": round(agg["value"], 2),
            "Cartons": None,
            "Gross Weight (kg)": gross_by_part.get(part),
            "Net Weight (kg)": net_by_part.get(part),
            "Country Origin": iso,
            "Charges": CHARGES,
            "Zone": ZONE,
        })

    rows.sort(key=lambda r: r["Value"], reverse=True)

    # Cartons: shipping paperwork gives a shipment TOTAL only, never a per-part
    # count. Rather than invent an allocation, the whole total sits on the first
    # row so the column still ties out to the shipment.
    if pl is not None and pl.cartons is not None and rows:
        rows[0]["Cartons"] = int(pl.cartons) if float(pl.cartons).is_integer() else pl.cartons
        res.notes.append(
            f"cartons: shipment total {pl.cartons:,.0f} placed on the first row "
            f"({rows[0]['Part #']}); source was the {pl.carton_source}. The paperwork "
            f"does not state cartons per part.")
    return rows


# ---------------------------------------------------------------------------
# Reconciliation
# ---------------------------------------------------------------------------
def reconcile(doc: InvoiceDoc, pl: PackingList | None, rep: ReceptionReport | None,
              rows: list[dict], res: ShipmentResult) -> None:
    """Every check below treats a MISSING value as a failure, not a pass."""
    # qty x unit price == extended total, on every line
    for ln in doc.lines:
        if abs(round(ln.qty * ln.unit_price, 2) - ln.ext_total) > 0.02:
            res.problems.append(
                f"invoice {ln.invoice} part {ln.part}: {ln.qty:,.0f} x {ln.unit_price} "
                f"= {ln.qty * ln.unit_price:,.2f}, but the line reads {ln.ext_total:,.2f}")
    for msg in doc.page_total_mismatches:
        res.problems.append(msg)

    inv_value = round(sum(l.ext_total for l in doc.lines), 2)
    inv_qty = sum(l.qty for l in doc.lines)

    # ALL summary sets combined -- reading only the last one invents a
    # discrepancy that does not exist.
    if doc.summaries:
        by_invoice: dict[str, float] = {}
        for s in doc.summaries:
            by_invoice[s.invoice] = round(by_invoice.get(s.invoice, 0.0) + s.total, 2)
        summary_value = round(sum(by_invoice.values()), 2)
        if abs(summary_value - inv_value) > 0.01:
            res.problems.append(
                f"invoice pages total {inv_value:,.2f} but the {len(doc.summary_totals) or 1} "
                f"summary set(s) total {summary_value:,.2f}")
        if doc.summary_totals:
            printed = round(sum(doc.summary_totals.values()), 2)
            if abs(printed - summary_value) > 0.01:
                res.problems.append(
                    f"summary rows add to {summary_value:,.2f} but the printed "
                    f"Items Total(s) say {printed:,.2f}")
        parsed_invoices = {l.invoice for l in doc.lines}
        missing = set(by_invoice) - parsed_invoices
        extra = parsed_invoices - set(by_invoice)
        if missing:
            res.problems.append(f"{len(missing)} invoice(s) listed in a summary but not "
                                f"parsed from any page: {', '.join(sorted(missing)[:8])}")
        if extra:
            res.problems.append(f"{len(extra)} parsed invoice(s) absent from every summary: "
                                f"{', '.join(sorted(extra)[:8])}")
    else:
        res.warnings.append("no Invoice Summary pages found -- the invoice total could not "
                            "be checked against a roll-up")

    if pl is not None:
        pl_qty = sum(r.qty for r in pl.rows)
        if abs(pl_qty - inv_qty) > 0.001:
            res.problems.append(f"packing list qty {pl_qty:,.0f} != invoice qty {inv_qty:,.0f}")
        if pl.total_qty is not None and abs(pl.total_qty - pl_qty) > 0.001:
            res.problems.append(f"packing-list rows sum to {pl_qty:,.0f} but its total row "
                                f"says {pl.total_qty:,.0f}")
        group_gross = round(sum(pl.group_gross.values()), 2)
        group_net = round(sum(pl.group_net.values()), 2)
        if pl.total_gross is None:
            res.problems.append("packing list states no total gross weight")
        elif abs(group_gross - pl.total_gross) > 0.01:
            res.problems.append(f"HS-group gross weights sum to {group_gross:,.2f} but the "
                                f"packing-list total is {pl.total_gross:,.2f}")
        if pl.total_net is None:
            res.problems.append("packing list states no total net weight")
        elif abs(group_net - pl.total_net) > 0.01:
            res.problems.append(f"HS-group net weights sum to {group_net:,.2f} but the "
                                f"packing-list total is {pl.total_net:,.2f}")
        if pl.cartons is None:
            res.warnings.append("packing list states no carton count -- the Cartons column "
                                "is blank")
        for n in pl.notes:
            res.warnings.append(n)

        out_gross = round(sum(r["Gross Weight (kg)"] or 0 for r in rows), 2)
        out_net = round(sum(r["Net Weight (kg)"] or 0 for r in rows), 2)
        if pl.total_gross is not None and abs(out_gross - pl.total_gross) > 0.01:
            res.problems.append(f"output gross {out_gross:,.2f} != packing list "
                                f"{pl.total_gross:,.2f}")
        if pl.total_net is not None and abs(out_net - pl.total_net) > 0.01:
            res.problems.append(f"output net {out_net:,.2f} != packing list "
                                f"{pl.total_net:,.2f}")
    else:
        res.warnings.append("no packing list -- weights and cartons are blank, and the "
                            "invoice quantities could not be cross-checked")

    if rep is not None and rep.rows:
        rep_qty = sum(r.qty_physical for r in rep.rows)
        if abs(rep_qty - inv_qty) > 0.001:
            res.problems.append(f"reception report qty {rep_qty:,.0f} != invoice qty "
                                f"{inv_qty:,.0f}")
        inv_parts = {l.part for l in doc.lines}
        rep_parts = {r.part for r in rep.rows}
        for p in sorted(rep_parts - inv_parts):
            close = [q for q in inv_parts if _depunct(q) == _depunct(p)
                     or _edit_distance_1(_depunct(q), _depunct(p))]
            if len(close) == 1:
                res.warnings.append(f"reception report part '{p}' reads as '{close[0]}' on "
                                    f"the invoices (hand-keying typo)")
            else:
                res.problems.append(f"reception report lists part '{p}', absent from every "
                                    f"invoice")
        for p in sorted(inv_parts - rep_parts):
            close = [q for q in rep_parts if _depunct(q) == _depunct(p)
                     or _edit_distance_1(_depunct(q), _depunct(p))]
            if not close:
                res.problems.append(f"part '{p}' is on the invoices but not on the "
                                    f"reception report")
        if rep.gross_weight is not None and pl is not None and pl.total_gross is not None:
            if abs(rep.gross_weight - pl.total_gross) > 0.5:
                res.problems.append(f"reception report gross {rep.gross_weight:,.2f} != "
                                    f"packing list {pl.total_gross:,.2f}")
        # The Mexican fraccion and the US HTS are different code systems, so a
        # difference is expected -- surfaced once, as information, not an error.
        frac = {r.part: r.fraccion for r in rep.rows if r.fraccion}
        conflicts = sum(1 for r in rows if r["HTS"] and frac.get(r["Part #"])
                        and frac[r["Part #"]] != r["HTS"])
        if conflicts:
            res.notes.append(f"{conflicts} part(s) carry a Mexican fraccion on the reception "
                             f"report that differs from the US HTS -- expected; the HTS "
                             f"column is the database's US code.")
    elif rep is not None:
        res.warnings.append("reception report found but no line items could be read from it")

    for r in rows:
        if not r["HTS"]:
            continue
        if not (8 <= len(str(r["HTS"])) <= 12) or not str(r["HTS"]).isdigit():
            res.problems.append(f"{r['Part #']}: HTS '{r['HTS']}' is not a plausible "
                                f"8-12 digit code")
    for r in rows:
        if not r["Country Origin"]:
            res.problems.append(f"{r['Part #']}: country of origin is blank")


# ---------------------------------------------------------------------------
# Output workbook
# ---------------------------------------------------------------------------
def write_workbook(rows: list[dict], review: list[dict], out_path: Path) -> Path:
    """Write the customs file. If Excel has the workbook open the save fails
    (TRAP 8), so fall back to a timestamped name and say so rather than dying."""
    df = pd.DataFrame(rows, columns=OUT_COLUMNS)
    try:
        _save(df, review, out_path)
        return out_path
    except PermissionError:
        alt = out_path.with_name(
            f"{out_path.stem}_{datetime.now().strftime('%H%M%S')}{out_path.suffix}")
        print(f"  [WARN] {out_path.name} is open in Excel; writing {alt.name} instead.")
        _save(df, review, alt)
        return alt


def _save(df: pd.DataFrame, review: list[dict], path: Path) -> None:
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Summary", index=False)
        if review:
            pd.DataFrame(review).to_excel(writer, sheet_name="Review", index=False)
    format_workbook(path, len(df))


def format_workbook(path: Path, n_rows: int) -> None:
    wb = load_workbook(path)
    ws = wb["Summary"]
    header_font = Font(name="Calibri", bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", start_color="1F4E78")
    widths = {"Part #": 18, "Qty": 10, "HTS": 14, "Value": 14, "Cartons": 10,
              "Gross Weight (kg)": 15, "Net Weight (kg)": 15,
              "Country Origin": 14, "Charges": 10, "Zone": 7}
    fmt = {"Qty": "#,##0", "Value": "#,##0.00", "Cartons": "#,##0",
           "Gross Weight (kg)": "#,##0.00", "Net Weight (kg)": "#,##0.00",
           "Charges": "0.00"}

    headers = [c.value for c in ws[1]]
    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for idx, name in enumerate(headers, start=1):
        letter = get_column_letter(idx)
        ws.column_dimensions[letter].width = widths.get(name, 12)
        if name in fmt:
            for row in ws.iter_rows(min_row=2, max_row=n_rows + 1, min_col=idx, max_col=idx):
                for cell in row:
                    cell.number_format = fmt[name]
        # HTS must be TEXT or Excel eats leading zeros and reformats long codes.
        if name == "HTS":
            for row in ws.iter_rows(min_row=2, max_row=n_rows + 1, min_col=idx, max_col=idx):
                for cell in row:
                    cell.number_format = "@"

    total_row = n_rows + 2
    ws.cell(row=total_row, column=1, value="TOTAL")
    for idx, name in enumerate(headers, start=1):
        letter = get_column_letter(idx)
        cell = ws.cell(row=total_row, column=idx)
        cell.font = Font(bold=True)
        # Qty, Value, Cartons and the weights sum. HTS is a code, not a number.
        if name in {"Qty", "Value", "Cartons", "Gross Weight (kg)", "Net Weight (kg)"}:
            cell.value = f"=SUM({letter}2:{letter}{n_rows + 1})"
            cell.number_format = fmt[name]
    ws.freeze_panes = "A2"

    if "Review" in wb.sheetnames:
        rv = wb["Review"]
        for cell in rv[1]:
            cell.font = header_font
            cell.fill = header_fill
        for idx in range(1, rv.max_column + 1):
            letter = get_column_letter(idx)
            longest = max((len(str(c.value)) for c in rv[letter] if c.value is not None),
                          default=12)
            rv.column_dimensions[letter].width = min(max(longest + 2, 14), 70)
    wb.save(path)


# ---------------------------------------------------------------------------
# Document classification and per-shipment driver
# ---------------------------------------------------------------------------
def classify(path: Path) -> str:
    """Identify a document by its CONTENT: filenames vary by client."""
    suffix = path.suffix.lower()
    if suffix == ".pdf":
        try:
            pages = extract_pdf_pages(path)
        except Exception:
            return "unknown"
        head = "\n".join(pages[:3])
        if re.search(r"REPORTE\s+DE\s+RECEPCION|RECONOCIMIENTO\s+PREVIO", head, re.I):
            return "reception"
        if re.search(r"\bInvoice\b", head, re.I) and re.search(r"P\.?\s*O\.?\s*No", head, re.I):
            return "invoices"
        if re.search(r"Invoice\s+Summary", head, re.I):
            return "invoices"
        return "unknown"
    if suffix in {".xls", ".xlsx", ".xlsm"}:
        try:
            xl = pd.ExcelFile(path, engine="xlrd" if suffix == ".xls" else "openpyxl")
            for sheet in xl.sheet_names[:3]:
                df = _read_sheet(path, sheet).head(40)
                flat = " ".join(str(v).lower() for v in df.values.ravel())
                if "product key" in flat and "hts" in flat:
                    return "partsdb"
                for r in range(len(df)):
                    cols = _map_pl_columns(list(df.iloc[r]))
                    if {"po", "part", "qty"} <= set(cols):
                        return "packing"
        except Exception:
            return "unknown"
    return "unknown"


def process_shipment(docs: dict[str, list[Path]], db: PartsDB,
                     out_dir: Path, label: str) -> tuple[int, str]:
    """Run one shipment end to end. Returns (problem count, message)."""
    res = ShipmentResult()

    doc = parse_invoices(docs["invoices"][0])
    print(f"  invoices ....... {doc.invoice_pages} invoice page(s), {len(doc.lines)} line "
          f"item(s), {doc.summary_pages} summary page(s), "
          f"{len(doc.summary_totals)} summary set(s)")
    if not doc.lines:
        raise RuntimeError("no invoice line items could be parsed")

    pl = None
    if docs["packing"]:
        pl = parse_packing_list(docs["packing"][0])
        print(f"  packing list ... {len(pl.rows)} row(s), {len(pl.group_gross)} weight "
              f"group(s), cartons {pl.cartons if pl.cartons is not None else 'n/a'}")

    rep = None
    if docs["reception"]:
        rep = parse_reception_report(docs["reception"][0])
        print(f"  reception ...... {len(rep.rows)} row(s)"
              f"{', guide ' + rep.guide if rep.guide else ''}")

    group_of = join_invoice_to_packing(doc.lines, pl, res) if pl else {}
    rows = build_rows(doc.lines, pl, group_of, db, res)
    reconcile(doc, pl, rep, rows, res)

    shipment = (pl.shipment_no if pl and pl.shipment_no
                else (rep.guide if rep and rep.guide else label))
    out_path = out_dir / f"customs_summary_{shipment}.xlsx"
    written = write_workbook(rows, res.review, out_path)

    print()
    for n in res.notes:
        print(f"  [NOTE] {n}")
    for w in res.warnings:
        print(f"  [WARN] {w}")
    for p in res.problems:
        print(f"  [CHECK] {p}")
    if res.estimated_groups or res.exact_groups:
        print()
        print("  Per-row weights are ALLOCATED within each HS group in proportion to")
        print("  piece count; column totals are exact, individual rows assume equal")
        print("  weight per piece.")
        for g in res.exact_groups:
            print(f"    exact (single part)  : {g}")
        for g in res.estimated_groups:
            print(f"    estimated (multi part): {g}")

    total_value = sum(r["Value"] for r in rows)
    total_qty = sum(r["Qty"] for r in rows)
    msg = (f"{len(doc.lines)} line items -> {len(rows)} parts, "
           f"{total_qty:,.0f} pcs, ${total_value:,.2f} -> {written.name}")
    return len(res.problems), msg


def collect(folder: Path) -> dict[str, list[Path]]:
    docs: dict[str, list[Path]] = {"invoices": [], "packing": [], "reception": [],
                                   "partsdb": [], "unknown": []}
    for p in sorted(folder.iterdir()):
        if p.is_dir() or p.name.startswith("~$"):
            continue
        docs[classify(p)].append(p)
    return docs


def run(in_dir: Path, out_dir: Path, db_dir: Path) -> int:
    if not in_dir.exists():
        print(f"Creating input folder: {in_dir}")
        in_dir.mkdir(parents=True, exist_ok=True)
    if not db_dir.exists():
        db_dir.mkdir(parents=True, exist_ok=True)
    out_dir.mkdir(parents=True, exist_ok=True)

    try:
        db = load_parts_db(db_dir)
    except Exception as e:
        print(f"[ERROR] {e}")
        print(f"\nPut the latest 'Kohler Parts for Upload ....xlsx' in:")
        print(f"  {db_dir}")
        print("The tool always uses the most recently modified file there.")
        return 1
    print(f"Parts database: {db.path.name}  ({len(db.by_key):,} part numbers)")

    # A shipment is either the input folder itself or one subfolder per shipment.
    subfolders = [p for p in sorted(in_dir.iterdir()) if p.is_dir()]
    shipments = subfolders or [in_dir]

    processed = problems = 0
    for folder in shipments:
        label = folder.name if folder != in_dir else "shipment"
        docs = collect(folder)
        if not docs["invoices"]:
            if folder == in_dir and not any(docs[k] for k in docs):
                print(f"\nNo documents in {in_dir}. Drop one shipment's invoices, packing "
                      f"list and reception report there, then run again.")
                return 0
            print(f"\n[SKIP] {label}: no commercial invoice PDF found")
            problems += 1
            continue

        print(f"\n{'=' * 62}\n{label}\n{'=' * 62}")
        for kind in ("invoices", "packing", "reception"):
            for p in docs[kind]:
                print(f"  [{kind.upper():9}] {p.name}")
        for p in docs["partsdb"]:
            print(f"  [PARTSDB  ] {p.name}  (ignored here -- the database\\ copy is used)")
        for p in docs["unknown"]:
            print(f"  [?????????] {p.name}  (not recognized -- check the file)")
            problems += 1
        for kind in ("invoices", "packing", "reception"):
            if len(docs[kind]) > 1:
                print(f"  [WARN] {len(docs[kind])} {kind} documents found; using "
                      f"{docs[kind][0].name}. Give each shipment its own subfolder "
                      f"under input\\ if these belong to different shipments.")
        print()

        try:
            n_problems, msg = process_shipment(docs, db, out_dir, label)
            print(f"\n  [OK  ] {msg}")
            if n_problems:
                print(f"  [!!!!] {n_problems} check(s) did not tie out -- read the "
                      f"[CHECK] lines above before filing.")
            problems += n_problems
            processed += 1
        except Exception as e:
            print(f"  [FAIL] {label}: {e}")
            traceback.print_exc()
            problems += 1

    print(f"\n{'=' * 62}")
    print(f"Done. {processed} shipment(s) processed.")
    print(f"Output folder: {out_dir}")
    print("=" * 62)
    return problems


def main() -> int:
    here = base_dir()
    in_dir = Path(sys.argv[1]) if len(sys.argv) > 1 else here / "input"
    out_dir = Path(sys.argv[2]) if len(sys.argv) > 2 else here / "output"
    db_dir = Path(sys.argv[3]) if len(sys.argv) > 3 else here / "database"

    log_dir = here / "logs"
    log_dir.mkdir(exist_ok=True)
    log_path = log_dir / f"run_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
    tee = TeeLogger(log_path)
    sys.stdout = tee
    sys.stderr = tee

    print(f"Unimex Customs Summary  v{__version__}")
    print(f"Run started: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"Log file: {log_path}")
    print()

    problems = 0
    try:
        problems = run(in_dir, out_dir, db_dir)
    except Exception as e:
        print(f"\n[ERROR] Something went wrong: {e}")
        print("\nFull details (please send this log file to Andy):")
        traceback.print_exc()
        problems = 1  # a crash is also a reason to look for a newer build

    # Only check for updates when this run had trouble. A clean run never
    # touches the network. If a newer build installs it relaunches and
    # reprocesses the input, self-healing what this version couldn't handle.
    updated = False
    if getattr(sys, "frozen", False) and problems > 0:
        try:
            import updater  # lazy: only the frozen exe ever needs it
            updated = updater.check_and_update("customs", "UnimexCustoms.exe", __version__)
        except Exception as e:
            print(f"[update] skipped ({type(e).__name__}).")

    sys.stdout = tee.terminal
    sys.stderr = tee.terminal
    tee.close()
    if not updated:
        pause_for_user()  # on the update path the relaunched instance owns the UX
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

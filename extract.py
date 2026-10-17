"""
Extractor: customer account-statement PDFs -> normalized ledger rows.

TUNED 2026-10-16 to Noman's REAL documents (customer account statements
from packaging companies; Noman = "NOMAN NADEEM", the salesperson).

Ledger format:
  - Page 1 header block: "Account Statement", Code, Starting/Ending Date,
    Company name, NTN, STRN, Sales Person, Previous Balance.
  - Every page: one 12-column table
      Date | InvoiceNo | Description | Size | Length | Qty | Rate |
      Amount | SalesTax | Debit | Credit | Balance
    (header row repeats on every page; last page may end with an empty row)
  - Dates are DD-MM-YYYY, parsed STRICTLY day-first (never M/D).
  - Balance carries a Dr/Cr suffix -> stored as signed number
      Dr (or no suffix, positive) = +,  Cr / leading "-" = -.
  - JV rows (payments / tax deductions) have only Credit filled;
    Size/Length/Qty/Rate/Amount/SalesTax/Debit stay blank.

Provenance on every row: Company, CustomerCode, SourceFile (visible),
plus internal _source_page (goes to the "Source" sheet).

Pages that do not contain the ledger table fall back to a per-page text
dump so nothing is silently lost.

====================================================================
ADAPT THIS WHEN NEW LAYOUTS ARRIVE  (the ONLY section you should
need to touch):
  1. LEDGER_HEADERS - the expected 12 column headers, in order.
  2. COLUMN_MAP     - extra header spellings -> canonical name.
  3. CANONICAL_COLUMNS - final Excel column order (visible).
  4. HEADER_PATTERNS - regexes for the page-1 statement header block.
====================================================================
"""

import os
import re
from datetime import datetime

import pdfplumber

# ------------------------------------------------------------------
# 1. LEDGER_HEADERS: the 12 expected columns, in order.
# ------------------------------------------------------------------
LEDGER_HEADERS = [
    "Date", "InvoiceNo", "Description", "Size", "Length", "Qty",
    "Rate", "Amount", "SalesTax", "Debit", "Credit", "Balance",
]

# ------------------------------------------------------------------
# 2. COLUMN_MAP: extra header spellings (normalized) -> canonical name.
# ------------------------------------------------------------------
COLUMN_MAP = {
    "invoice no": "InvoiceNo", "invoice #": "InvoiceNo", "inv no": "InvoiceNo",
    "inv #": "InvoiceNo", "voucher": "InvoiceNo", "doc no": "InvoiceNo",
    "sales tax": "SalesTax", "gst": "SalesTax", "tax": "SalesTax",
    "wht": "SalesTax",
    "particulars": "Description", "narration": "Description",
    "desc": "Description",
    "quantity": "Qty", "qnty": "Qty",
    "unit price": "Rate", "price": "Rate",
    "dr": "Debit", "cr": "Credit",
    "bal": "Balance", "running balance": "Balance",
}

# ------------------------------------------------------------------
# 3. CANONICAL_COLUMNS: final visible Excel column order.
# ------------------------------------------------------------------
CANONICAL_COLUMNS = LEDGER_HEADERS + ["Company", "CustomerCode", "SourceFile"]

# Internal provenance (goes to the "Source" sheet, not the main sheet).
META_COLUMNS = ["_source_page"]

MONEY_COLUMNS = {"Size", "Length", "Qty", "Rate", "Amount",
                 "SalesTax", "Debit", "Credit"}

# ------------------------------------------------------------------
# 4. HEADER_PATTERNS: page-1 statement header block.
# ------------------------------------------------------------------
HEADER_PATTERNS = {
    "customer_code": re.compile(r"Code\s*:\s*(\d+)"),
    "stmt_start": re.compile(r"Starting Date\s*:\s*([0-9]{1,2}-[A-Za-z]{3}-[0-9]{4})"),
    "stmt_end": re.compile(r"Ending Date\s*:\s*([0-9]{1,2}-[A-Za-z]{3}-[0-9]{4})"),
    "ntn": re.compile(r"NTN\s*:([^\n]*)"),
    "strn": re.compile(r"STRN\s*:([^\n]*)"),
    "sales_person": re.compile(r"Sales Person\s*:([^\n]*)"),
    "prev_balance": re.compile(r"Previous Balance\s*:\s*([\d,\.\-]+)"),
}

LEDGER_MARKER = "Account Statement"


def _clean_header(h):
    if h is None:
        return ""
    h = str(h).strip().lower()
    h = re.sub(r"\s+", " ", h).strip(" .:-")
    return h


def _canonical_header(h):
    key = _clean_header(h)
    if not key:
        return ""
    return COLUMN_MAP.get(key, h.strip())


def _is_ledger_header(row):
    """True if this row is the 12-col ledger header (any page)."""
    if len(row) != len(LEDGER_HEADERS):
        return False
    canon = [_canonical_header(h) for h in row]
    return canon == LEDGER_HEADERS


def _parse_money(s):
    """'1,596,732.50' -> 1596732.5 ; ''/'-' -> None."""
    s = str(s or "").replace(",", "").strip()
    if s in ("", "-"):
        return None
    try:
        return float(s)
    except ValueError:
        return None


def _parse_balance(s):
    """'2,635,210.28Dr' -> 2635210.28 ; '-272,712.49Cr' -> -272712.49.

    Sign rule: an explicit leading '-' always wins; otherwise Dr = +,
    Cr = -. (Seen in the wild: '-272,712.49Cr' where sign and suffix agree.)
    """
    raw = str(s or "").replace(",", "").replace(" ", "").strip()
    if raw in ("", "-"):
        return None
    m = re.match(r"^(-)?([\d.]+)(Dr|Cr)?$", raw, re.IGNORECASE)
    if not m:
        return None
    neg, num, suffix = m.groups()
    val = float(num)
    if neg or (suffix and suffix.upper() == "CR"):
        val = -val
    return val


def _parse_date(s):
    """'28-10-2025' -> date(2025,10,28). STRICTLY day-first; None if bad."""
    s = str(s or "").strip()
    try:
        return datetime.strptime(s, "%d-%m-%Y").date()
    except ValueError:
        return None


def _parse_statement_header(page1_text):
    """Extract the statement header block from page-1 text."""
    meta = {"company": "", "customer_code": "", "stmt_start": None,
            "stmt_end": None, "ntn": "", "strn": "",
            "sales_person": "", "previous_balance": None}
    if not page1_text or LEDGER_MARKER not in page1_text:
        return meta
    lines = [ln.strip() for ln in page1_text.splitlines()]

    def grab(name):
        m = HEADER_PATTERNS[name].search(page1_text)
        return (m.group(1).strip() if m else "")

    meta["customer_code"] = grab("customer_code")
    meta["ntn"] = grab("ntn")
    meta["strn"] = grab("strn")
    meta["sales_person"] = grab("sales_person")

    for key, pat in (("stmt_start", "stmt_start"), ("stmt_end", "stmt_end")):
        raw = grab(pat)
        try:
            meta[key] = datetime.strptime(raw, "%d-%b-%Y").date() if raw else None
        except ValueError:
            meta[key] = None

    pb = grab("prev_balance")
    meta["previous_balance"] = _parse_money(pb)

    # Company = first non-empty line right after the "Ending Date" line.
    for i, ln in enumerate(lines):
        if "Ending Date" in ln:
            for nxt in lines[i + 1:]:
                if nxt:
                    meta["company"] = nxt
                    break
            break
    return meta


def _extract_ledger_tables(page):
    """Yield (data_rows) for ledger tables on a page (header skipped)."""
    for table in page.extract_tables() or []:
        rows = [[(c or "").strip() for c in r] for r in table]
        rows = [r for r in rows if any(c for c in r)]  # drop empty rows
        if not rows or not _is_ledger_header(rows[0]):
            continue
        yield rows[1:]


def extract_pdf(path):
    """Extract normalized ledger rows from one PDF. Returns (rows, notes, meta)."""
    rows, notes = [], []
    fname = os.path.basename(path)
    try:
        pdf = pdfplumber.open(path)
    except Exception as e:
        return [], [f"{fname}: could not open ({e})"], {}

    meta = {}
    with pdf:
        pages = list(pdf.pages)
        if pages:
            meta = _parse_statement_header(pages[0].extract_text() or "")
        for i, page in enumerate(pages, start=1):
            page_rows = 0
            for data in _extract_ledger_tables(page):
                for dr in data:
                    dr = dr + [""] * (len(LEDGER_HEADERS) - len(dr))
                    cells = dict(zip(LEDGER_HEADERS, dr))
                    if not any(cells.values()):
                        continue  # trailing empty row
                    row = {}
                    bad = []
                    for col in LEDGER_HEADERS:
                        v = cells[col]
                        if col == "Date":
                            d = _parse_date(v)
                            if d is None and v:
                                bad.append(f"date {v!r}")
                            row[col] = d if d else ""
                        elif col == "Balance":
                            b = _parse_balance(v)
                            if b is None and v:
                                bad.append(f"balance {v!r}")
                            row[col] = b if b is not None else ""
                        elif col in MONEY_COLUMNS:
                            m = _parse_money(v)
                            if m is None and v not in ("", "-"):
                                bad.append(f"{col} {v!r}")
                            row[col] = m if m is not None else ""
                        elif col == "Description":
                            row[col] = re.sub(r"\s*\n\s*", " ", v).strip()
                        else:
                            row[col] = v.strip()
                    if bad:
                        notes.append(f"{fname} p{i}: unparsed {', '.join(bad)} "
                                     f"(row kept as text)")
                    row["Company"] = meta.get("company", "")
                    row["CustomerCode"] = meta.get("customer_code", "")
                    row["SourceFile"] = fname
                    row["_source_page"] = i
                    rows.append(row)
                    page_rows += 1
            if page_rows == 0:
                text = (page.extract_text() or "").strip()
                if text:
                    notes.append(f"{fname} p{i}: no ledger table found")
    if not rows:
        notes.append(f"{fname}: no extractable ledger rows")
    return rows, notes, meta


def _dedupe(rows):
    """Drop exact-duplicate transactions (same customer+file+full row).

    InvoiceNo is unique per real transaction, so this only removes
    extraction artefacts or the same file uploaded twice.
    """
    seen, out = set(), []
    for r in rows:
        key = (r.get("CustomerCode"), r.get("SourceFile"),
               str(r.get("Date")), r.get("InvoiceNo"), r.get("Description"),
               str(r.get("Debit")), str(r.get("Credit")), str(r.get("Balance")))
        if key in seen:
            continue
        seen.add(key)
        out.append(r)
    return out


def normalize_rows(all_rows):
    cols = CANONICAL_COLUMNS + META_COLUMNS
    rows = _dedupe(all_rows)
    return cols, [{c: r.get(c, "") for c in cols} for r in rows]


def process_files(paths):
    """Full pipeline: list of PDF paths -> (columns, rows, notes, statements).

    statements: one dict per file for the "Statements" sheet.
    """
    all_rows, notes, statements = [], [], []
    for p in sorted(paths):
        rows, n, meta = extract_pdf(p)
        all_rows.extend(rows)
        notes.extend(n)
        statements.append({
            "Company": meta.get("company", ""),
            "CustomerCode": meta.get("customer_code", ""),
            "SourceFile": os.path.basename(p),
            "StmtStart": meta.get("stmt_start") or "",
            "StmtEnd": meta.get("stmt_end") or "",
            "SalesPerson": meta.get("sales_person", ""),
            "PreviousBalance": meta.get("previous_balance")
            if meta.get("previous_balance") is not None else "",
            "NTN": meta.get("ntn", ""),
            "STRN": meta.get("strn", ""),
            "Rows": len(rows),
        })
    cols, rows = normalize_rows(all_rows)
    return cols, rows, notes, statements


if __name__ == "__main__":
    import sys
    cols, rows, notes, statements = process_files(sys.argv[1:])
    print(f"rows={len(rows)} cols={len(cols)}")
    for s in statements:
        print("statement:", {k: v for k, v in s.items()})
    for n in notes:
        print("note:", n)
    for r in rows[:3]:
        print({k: v for k, v in r.items() if v != "" and not k.startswith("_")})

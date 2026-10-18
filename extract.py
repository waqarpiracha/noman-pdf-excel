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
  5. AGING_*        - the "Customer's Balances" aging report (2026-10-18):
                     a second document type with its own parser + sheet.
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
            "sales_person": "", "previous_balance": None,
            "report": "Account Statement"}
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


# ------------------------------------------------------------------
# 5. AGING REPORT: "Customer's Balances" (Crystal Reports).
#    Added 2026-10-18: Noman's real aging PDF produced 0 ledger rows
#    because it is a different document type, not a broken file.
#
#    Text layout (NOT a table):
#      Customer line: 1010023  BIN HASHIM ... LTD<2+sp>ISLAMABAD<2+sp>0
#      Invoice line:  STINV-489  02-Aug-2026  3,070,961  2,622,440
#                     448,521  448,521  76
#      Summary line:  19-Oct-2026  -2,013  0  -2,013  -2,013  -2
#                     (customer balance, no invoice number -> kept)
#    Subtotal lines (bare amounts, no date/invoice) and the Grand Total
#    line carry no new information -> skipped, never guessed.
# ------------------------------------------------------------------
AGING_MARKER = "Customer's Balances"

AGING_COLUMNS = ["CustomerCode", "CustomerName", "City", "InvoiceNo",
                 "InvoiceDate", "Amount", "Received", "InvoiceBalance",
                 "Balance", "Days", "SourceFile"]

_MONEY = r"-?[\d,]+(?:\.\d+)?"
_RE_AGING_CUST = re.compile(rf"^(\d{{7}})\s+(.*?)\s+({_MONEY})\s*$")
_RE_AGING_INV = re.compile(
    rf"^\s*([A-Z]{{2,10}}-\d+)\s+(\d{{2}}-[A-Za-z]{{3}}-\d{{4}})\s+"
    rf"({_MONEY})\s+({_MONEY})\s+({_MONEY})\s+({_MONEY})\s+(-?\d+)\s*$")
_RE_AGING_SUM = re.compile(
    rf"^\s*(\d{{2}}-[A-Za-z]{{3}}-\d{{4}})\s+"
    rf"({_MONEY})\s+({_MONEY})\s+({_MONEY})\s+({_MONEY})\s+(-?\d+)\s*$")
_RE_AGING_SKIP = re.compile(
    r"Customer's Balances|To Date|CustomerCode|CustomerName|"
    r"invoice No\.|Page \d+ of \d+|Grand Total")


# Pakistani cities, used to split "NAME ... CITY" customer lines.
# pdfplumber collapses the name/city column gap to single spaces, so the
# city is found by lookup (1-3 trailing tokens). A trailing token that is
# not a known city is NEVER guessed -> city stays blank.
_PAK_CITIES = frozenset([
    "ISLAMABAD", "KARACHI", "LAHORE", "MULTAN", "FAISALABAD", "RAWALPINDI",
    "GUJRAT", "RAIWIND", "PESHAWAR", "QUETTA", "HYDERABAD", "SIALKOT",
    "GUJRANWALA", "BAHAWALPUR", "SARGODHA", "SUKKUR", "LARKANA",
    "SHEIKHUPURA", "JHELUM", "SAHIWAL", "OKARA", "KASUR", "MARDAN",
    "ABBOTTABAD", "MANSEHRA", "NOWSHERA", "ATTOCK", "CHAKWAL", "TAXILA",
    "KHAIRPUR", "SHIKARPUR", "JACOBABAD", "NAWABSHAH", "DADU", "BADIN",
    "THATTA", "UMERKOT", "MIRPURKHAS", "TANDO ADAM", "MUZAFFARABAD",
    "GILGIT", "SKARDU", "TURBAT", "GAWADAR", "KHUZDAR", "CHITRAL", "SWAT",
    "MINGORA", "KOHAT", "BANNU", "KAMALIA", "CHISHTIAN", "VEHARI",
    "KHANEWAL", "LODHRAN", "PAKPATTAN", "ARIFWALA", "JARANWALA", "HAFIZABAD",
    "NAROWAL", "PASRUR", "WAZIRABAD", "MANDI BAHAUDDIN", "PHALIA", "BHALWAL",
    "KOT MOMIN", "CHINIOT", "JHANG", "SHORKOT", "GOJRA", "SAMUNDRI",
    "DUNYAPUR", "MAILSI", "BUREWALA", "CHICHAWATNI", "SHUJABAD", "JALALPUR",
    "PIR MAHAL", "TOBA TEK SINGH",
    # multi-word towns (checked as trailing 2-3 tokens)
    "DERA ISMAIL KHAN", "DERA GHAZI KHAN", "RAHIM YAR KHAN", "WAH CANTT",
    "KOT ADDU", "KOT ABDUL MALIK",
])


def _split_aging_customer(middle):
    """'NAME ... CITY' (balance already stripped) -> (name, city).

    The city is the trailing 1-3 tokens when they form a known Pakistani
    city; otherwise the whole text is the name and city stays blank.
    """
    toks = middle.strip().split()
    for n in (3, 2, 1):
        if len(toks) > n:
            cand = " ".join(toks[-n:])
            if cand in _PAK_CITIES:
                return " ".join(toks[:-n]), cand
    return " ".join(toks), ""


def _parse_date_mmm(s):
    """'02-Aug-2026' -> date(2026,8,2). None if bad."""
    s = str(s or "").strip()
    try:
        return datetime.strptime(s, "%d-%b-%Y").date()
    except ValueError:
        return None


def _aging_row(ctx, fname, inv_no, dt_raw, amt, rec, ibal, bal, days,
               pageno, notes):
    d = _parse_date_mmm(dt_raw)
    if d is None:
        notes.append(f"{fname} p{pageno}: bad date {dt_raw!r} (kept as text)")
    def money(v):
        m = _parse_money(v)
        return m if m is not None else ""
    row = {
        "CustomerCode": ctx.get("code", ""),
        "CustomerName": ctx.get("name", ""),
        "City": ctx.get("city", ""),
        "InvoiceNo": inv_no,
        "InvoiceDate": d if d else (dt_raw or ""),
        "Amount": money(amt),
        "Received": money(rec),
        "InvoiceBalance": money(ibal),
        "Balance": money(bal),
        "Days": int(days),
        "SourceFile": fname,
    }
    return row


def _extract_aging_page(text, fname, pageno, ctx):
    """Parse one page of a Customer's Balances report.

    ctx carries the current customer across lines (and pages).
    Returns (rows, notes)."""
    rows, notes = [], []
    for raw in (text or "").splitlines():
        ln = raw.rstrip()
        if not ln.strip() or _RE_AGING_SKIP.search(ln):
            continue
        m = _RE_AGING_CUST.match(ln)
        if m:
            name, city = _split_aging_customer(m.group(2))
            ctx.update(code=m.group(1), name=name, city=city)
            continue
        m = _RE_AGING_INV.match(ln)
        if m:
            rows.append(_aging_row(
                ctx, fname, m.group(1), m.group(2), m.group(3), m.group(4),
                m.group(5), m.group(6), m.group(7), pageno, notes))
            continue
        m = _RE_AGING_SUM.match(ln)
        if m:
            # customer-level balance line (no invoice number) -> kept
            rows.append(_aging_row(
                ctx, fname, "", m.group(1), m.group(2), m.group(3),
                m.group(4), m.group(5), m.group(6), pageno, notes))
            continue
        # subtotal lines and anything unrecognized: skipped, never guessed
    return rows, notes


def _dedupe_aging(rows):
    """Drop exact-duplicate aging rows (same customer+invoice+amount+file)."""
    seen, out = set(), []
    for r in rows:
        key = (r.get("CustomerCode"), r.get("InvoiceNo"),
               str(r.get("InvoiceDate")), str(r.get("Amount")),
               r.get("SourceFile"))
        if key in seen:
            continue
        seen.add(key)
        out.append(r)
    return out


def normalize_aging(all_rows):
    rows = _dedupe_aging(all_rows)
    return AGING_COLUMNS, [{c: r.get(c, "") for c in AGING_COLUMNS}
                           for r in rows]


# ------------------------------------------------------------------
# 6. GENERIC FALLBACK (2026-10-18, Waqar): "if it does not recognize,
#    pull everything with headings in a new sheet." Any PDF that yields
#    zero ledger AND zero aging rows gets every table and every text
#    line pulled into the "Raw" sheet instead of an empty workbook.
# ------------------------------------------------------------------

def _extract_generic_tables(page):
    """Yield (headings, data_rows) for every table on the page."""
    for table in page.extract_tables() or []:
        rows = [[(c or "").strip() for c in r] for r in table]
        rows = [r for r in rows if any(r)]
        if not rows:
            continue
        heads, seen, fixed = rows[0], {}, []
        for i, h in enumerate(heads):
            h = h or f"Column {i + 1}"
            if h in seen:
                seen[h] += 1
                h = f"{h} ({seen[h]})"
            else:
                seen[h] = 1
            fixed.append(h)
        width = len(fixed)
        data = []
        for r in rows[1:]:
            r = (r + [""] * width)[:width]
            if any(r):
                data.append(r)
        yield fixed, data


def extract_generic_pdf(path):
    """Last-resort pull for unrecognized PDFs.

    Returns {"SourceFile", "tables": [(pageno, tableno, headings, rows)],
             "texts": [(pageno, lineno, text)], "notes": [...]}.
    Tables win over text on any page (no duplication).
    """
    fname = os.path.basename(path)
    doc = {"SourceFile": fname, "tables": [], "texts": [], "notes": []}
    try:
        pdf = pdfplumber.open(path)
    except Exception as e:
        doc["notes"].append(f"{fname}: could not open ({e})")
        return doc
    with pdf:
        for i, page in enumerate(pdf.pages, start=1):
            tables = list(_extract_generic_tables(page))
            if tables:
                for tno, (heads, data) in enumerate(tables, start=1):
                    if data:
                        doc["tables"].append((i, tno, heads, data))
                    else:
                        doc["notes"].append(
                            f"{fname} p{i} table {tno}: headings only, no data")
                continue
            text = (page.extract_text() or "").strip()
            if text:
                for lno, ln in enumerate(text.splitlines(), start=1):
                    ln = ln.strip()
                    if ln:
                        doc["texts"].append((i, lno, ln))
    if not doc["tables"] and not doc["texts"]:
        doc["notes"].append(f"{fname}: no extractable content at all")
    return doc


def extract_pdf(path):
    """Extract normalized rows from one PDF.

    Returns (rows, notes, meta, aging_rows): ledger rows go to the
    Transactions sheet, aging-report rows to the Aging sheet.
    """
    rows, notes, aging_rows = [], [], []
    fname = os.path.basename(path)
    try:
        pdf = pdfplumber.open(path)
    except Exception as e:
        return [], [f"{fname}: could not open ({e})"], {}, []

    meta = {}
    with pdf:
        pages = list(pdf.pages)
        page1_text = pages[0].extract_text() or "" if pages else ""
        if AGING_MARKER in page1_text:
            # ---- second document type: Customer's Balances aging ----
            m = re.search(r"To Date\s*:\s*([0-9]{1,2}-[A-Za-z]{3}-[0-9]{4})",
                          page1_text)
            to_date = ""
            if m:
                try:
                    to_date = datetime.strptime(m.group(1), "%d-%b-%Y").date()
                except ValueError:
                    pass
            meta = {"report": "Customer's Balances", "to_date": to_date}
            ctx = {}
            for i, page in enumerate(pages, start=1):
                pr, pn = _extract_aging_page(page.extract_text() or "",
                                             fname, i, ctx)
                aging_rows.extend(pr)
                notes.extend(pn)
            if not aging_rows:
                notes.append(f"{fname}: aging report, no invoice rows parsed")
            return rows, notes, meta, aging_rows
        # ---- default: account-statement ledger ----
        meta = _parse_statement_header(page1_text)
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
    return rows, notes, meta, aging_rows


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
    """Full pipeline: list of PDF paths ->
    (columns, rows, notes, statements, aging_columns, aging_rows, raw_docs).

    statements: one dict per file for the "Statements" sheet.
    raw_docs: generic pulls for files no parser recognized (Raw sheet).
    """
    all_rows, notes, statements = [], [], []
    aging_all, raw_docs = [], []
    for p in sorted(paths):
        rows, n, meta, aging_rows = extract_pdf(p)
        all_rows.extend(rows)
        aging_all.extend(aging_rows)
        notes.extend(n)
        if not rows and not aging_rows:
            # unrecognized (or unparseable) -> pull everything, with headings
            g = extract_generic_pdf(p)
            notes.extend(g["notes"])
            if g["tables"] or g["texts"]:
                raw_docs.append(g)
        statements.append({
            "Company": meta.get("company", ""),
            "CustomerCode": meta.get("customer_code", ""),
            "SourceFile": os.path.basename(p),
            "Report": meta.get("report", "Account Statement"),
            "StmtStart": meta.get("stmt_start") or "",
            "StmtEnd": meta.get("stmt_end") or "",
            "SalesPerson": meta.get("sales_person", ""),
            "PreviousBalance": meta.get("previous_balance")
            if meta.get("previous_balance") is not None else "",
            "NTN": meta.get("ntn", ""),
            "STRN": meta.get("strn", ""),
            "Rows": len(rows) + len(aging_rows),
        })
    cols, rows = normalize_rows(all_rows)
    aging_cols, aging_rows = normalize_aging(aging_all)
    return cols, rows, notes, statements, aging_cols, aging_rows, raw_docs


if __name__ == "__main__":
    import sys
    cols, rows, notes, statements, aging_cols, aging_rows, raw_docs = (
        process_files(sys.argv[1:]))
    print(f"rows={len(rows)} cols={len(cols)} aging_rows={len(aging_rows)} "
          f"raw_docs={len(raw_docs)}")
    for s in statements:
        print("statement:", {k: v for k, v in s.items()})
    for n in notes:
        print("note:", n)
    for r in rows[:3]:
        print({k: v for k, v in r.items() if v != "" and not k.startswith("_")})
    for r in aging_rows[:5]:
        print({k: v for k, v in r.items() if v != ""})

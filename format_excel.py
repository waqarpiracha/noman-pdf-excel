"""
Formatter: rows -> print/share-ready .xlsx.

Kills Noman's #1 pain (manual formatting). Every output file gets:
  - bold white-on-dark header row, autofilter, frozen top row
  - sensible column widths (auto, capped), wrapped text
  - date columns as DD-MMM-YYYY, numeric columns as #,##0.00
  - thin borders, light zebra striping, print titles + landscape fit
Internal `_source_*` provenance columns are written on a second
"Source" sheet instead of cluttering the main sheet.
"""

from copy import copy
from datetime import datetime, date
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo  # noqa: F401 (kept for future use)

HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
TITLE_FONT = Font(bold=True, size=14, color="1F4E79")
THIN = Side(style="thin", color="B0B0B0")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
ZEBRA = PatternFill("solid", fgColor="F2F7FB")
WRAP = Alignment(vertical="center", wrap_text=True)
CENTER = Alignment(vertical="center", wrap_text=True, horizontal="center")

NUMERIC_HINTS = {"qty", "rate", "amount", "disc", "net", "s.no", "sno", "no",
                 "size", "length", "salestax", "tax", "debit", "credit", "balance",
                 "received", "days"}
DATE_HINTS = {"date"}


def _looks_numeric(values):
    vals = [v for v in values if str(v).strip() not in ("", "-")]
    if not vals:
        return False
    ok = 0
    for v in vals:
        s = str(v).replace(",", "").replace("%", "").strip()
        try:
            float(s)
            ok += 1
        except ValueError:
            pass
    return ok / len(vals) >= 0.7


def _looks_date(values):
    vals = [v for v in values if str(v).strip() not in ("", "-")]
    if not vals:
        return False
    fmts = ("%d-%b-%Y", "%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d",
            "%d.%m.%Y", "%d %b %Y", "%d-%b-%y")
    ok = 0
    for v in vals:
        s = str(v).strip()
        hit = False
        for f in fmts:
            try:
                datetime.strptime(s, f)
                hit = True
                break
            except ValueError:
                continue
        if hit:
            ok += 1
    return ok / len(vals) >= 0.7


def _to_number(v):
    s = str(v).replace(",", "").strip()
    pct = s.endswith("%")
    s = s.rstrip("%")
    try:
        num = float(s)
        return num
    except ValueError:
        return v


def _to_date(v):
    if isinstance(v, (date, datetime)):
        return v.date() if isinstance(v, datetime) else v
    s = str(v).strip()
    for f in ("%d-%b-%Y", "%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d",
              "%d.%m.%Y", "%d %b %Y", "%d-%b-%y"):
        try:
            return datetime.strptime(s, f).date()
        except ValueError:
            continue
    return v


def build_workbook(columns, rows, title="Customer Account Statements",
                   statements=None, dispatches=None, aging=None, raw=None):
    wb = Workbook()
    ws = wb.active
    ws.title = "Transactions"
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    # visible columns = everything except internal provenance
    vis_cols = [c for c in columns if not c.startswith("_")]
    meta_cols = [c for c in columns if c.startswith("_")]

    if not vis_cols:
        # no ledger transactions: keep a placeholder sheet
        if aging:
            ws.cell(row=1, column=1,
                    value="No ledger statements in this batch — see the "
                          "Aging sheet for customer balances.")
        elif raw:
            ws.cell(row=1, column=1,
                    value="No recognized statements in this batch — see the "
                          "Raw sheet for everything pulled from the file(s).")
        else:
            ws.cell(row=1, column=1,
                    value="No PDF statements in this batch — see the "
                          "Dispatches sheet for photo OCR results.")
        ws.column_dimensions["A"].width = 70
        if statements:
            _add_statements_sheet(wb, statements)
        if dispatches:
            _add_dispatches_sheet(wb, dispatches)
        if aging:
            _add_aging_sheet(wb, *aging)
        if raw:
            _add_raw_sheet(wb, raw)
        return wb

    # ---- title block ----
    ws.merge_cells(start_row=1, start_column=1,
                   end_row=1, end_column=len(vis_cols))
    tc = ws.cell(row=1, column=1, value=title)
    tc.font = TITLE_FONT
    ws.cell(row=2, column=1,
            value=f"Generated {datetime.now():%d-%b-%Y %I:%M %p}  |  {len(rows)} line items")
    ws.row_dimensions[1].height = 24

    hdr_row = 4
    # ---- header ----
    for j, col in enumerate(vis_cols, start=1):
        c = ws.cell(row=hdr_row, column=j, value=col)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = CENTER
        c.border = BORDER
    ws.row_dimensions[hdr_row].height = 20

    # ---- detect column types from data ----
    col_vals = {c: [r.get(c, "") for r in rows] for c in vis_cols}
    col_kind = {}
    for c in vis_cols:
        lc = c.lower()
        if any(h in lc for h in DATE_HINTS) and _looks_date(col_vals[c]):
            col_kind[c] = "date"
        elif any(h in lc for h in NUMERIC_HINTS) and _looks_numeric(col_vals[c]):
            col_kind[c] = "number"
        else:
            col_kind[c] = "text"

    # ---- data ----
    for i, r in enumerate(rows):
        rr = hdr_row + 1 + i
        if i % 2 == 1:
            for j in range(1, len(vis_cols) + 1):
                ws.cell(row=rr, column=j).fill = ZEBRA
        for j, col in enumerate(vis_cols, start=1):
            v = r.get(col, "")
            kind = col_kind[col]
            if kind == "date":
                v = _to_date(v)
            elif kind == "number":
                v = _to_number(v)
            c = ws.cell(row=rr, column=j, value=v)
            c.border = BORDER
            c.alignment = CENTER if kind in ("date", "number") else WRAP
            if kind == "date":
                c.number_format = "DD-MMM-YYYY"
            elif kind == "number":
                c.number_format = '#,##0.00'

    last_row = hdr_row + len(rows)
    last_col = get_column_letter(len(vis_cols))

    # ---- widths (auto, capped) ----
    for j, col in enumerate(vis_cols, start=1):
        vals = [str(col)] + [str(r.get(col, "")) for r in rows]
        w = min(max(len(v) for v in vals) + 4, 42)
        w = max(w, 12)
        if col_kind[col] == "number":
            w = min(w, 16)
        ws.column_dimensions[get_column_letter(j)].width = w

    # ---- filter + freeze + print ----
    if rows:
        ws.auto_filter.ref = f"A{hdr_row}:{last_col}{last_row}"
    ws.freeze_panes = f"A{hdr_row + 1}"
    ws.print_title_rows = f"{hdr_row}:{hdr_row}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0

    # ---- Source sheet (provenance) ----
    if meta_cols and rows:
        s2 = wb.create_sheet("Source")
        s2.append(["#"] + meta_cols)
        for c in s2[1]:
            c.font = HEADER_FONT
            c.fill = HEADER_FILL
        for i, r in enumerate(rows, start=1):
            s2.append([i] + [r.get(c, "") for c in meta_cols])
        s2.freeze_panes = "A2"
        s2.auto_filter.ref = s2.dimensions

    # ---- Statements sheet (one row per source PDF header block) ----
    if statements:
        _add_statements_sheet(wb, statements)

    # ---- Dispatches sheet (one row per shipment, from photo OCR) ----
    if dispatches:
        _add_dispatches_sheet(wb, dispatches)

    # ---- Aging sheet (Customer's Balances report) ----
    if aging:
        _add_aging_sheet(wb, *aging)

    # ---- Raw sheet (unrecognized PDFs: everything pulled) ----
    if raw:
        _add_raw_sheet(wb, raw)

    return wb


STATEMENT_COLUMNS = [
    ("Company", "text"), ("CustomerCode", "text"), ("SourceFile", "text"),
    ("Report", "text"),
    ("StmtStart", "date"), ("StmtEnd", "date"), ("SalesPerson", "text"),
    ("PreviousBalance", "number"), ("NTN", "text"), ("STRN", "text"),
    ("Rows", "int"),
]


def _add_statements_sheet(wb, statements):
    ws = wb.create_sheet("Statements")
    cols = [c for c, _ in STATEMENT_COLUMNS]
    kinds = dict(STATEMENT_COLUMNS)
    for j, col in enumerate(cols, start=1):
        c = ws.cell(row=1, column=j, value=col)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = CENTER
        c.border = BORDER
    for i, s in enumerate(statements, start=2):
        if i % 2 == 0:
            for j in range(1, len(cols) + 1):
                ws.cell(row=i, column=j).fill = ZEBRA
        for j, col in enumerate(cols, start=1):
            v = s.get(col, "")
            kind = kinds[col]
            if kind == "date":
                v = _to_date(v)
            elif kind in ("number", "int"):
                v = _to_number(v) if str(v).strip() not in ("", "-") else ""
            c = ws.cell(row=i, column=j, value=v)
            c.border = BORDER
            c.alignment = CENTER if kind in ("date", "number", "int") else WRAP
            if kind == "date":
                c.number_format = "DD-MMM-YYYY"
            elif kind == "number":
                c.number_format = "#,##0.00"
            elif kind == "int":
                c.number_format = "#,##0"
    last_col = get_column_letter(len(cols))
    ws.auto_filter.ref = f"A1:{last_col}{len(statements) + 1}"
    ws.freeze_panes = "A2"
    for j, col in enumerate(cols, start=1):
        vals = [col] + [str(s.get(col, "")) for s in statements]
        w = min(max(max(len(v) for v in vals) + 4, 14), 48)
        ws.column_dimensions[get_column_letter(j)].width = w


def write_excel(columns, rows, path, title="Customer Account Statements",
                statements=None, dispatches=None, aging=None, raw=None):
    wb = build_workbook(columns, rows, title, statements=statements,
                        dispatches=dispatches, aging=aging, raw=raw)
    wb.save(path)
    return path


# ------------------------------------------------------------------
# Aging sheet ("Customer's Balances" report). One row per invoice
# (InvoiceNo blank = customer-level balance line, no open invoice).
# ------------------------------------------------------------------

AGING_DATE_COLS = {"InvoiceDate"}
AGING_NUM_COLS = {"Amount", "Received", "InvoiceBalance", "Balance", "Days"}


def _add_aging_sheet(wb, acols, arows):
    ws = wb.create_sheet("Aging")
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    n = len(acols)

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n)
    tc = ws.cell(row=1, column=1, value="Customer Balances — aging")
    tc.font = TITLE_FONT
    ws.cell(row=2, column=1,
            value=f"Generated {datetime.now():%d-%b-%Y %I:%M %p}  |  "
                  f"{len(arows)} invoice/balance rows")
    ws.row_dimensions[1].height = 24

    hdr_row = 4
    for j, col in enumerate(acols, start=1):
        c = ws.cell(row=hdr_row, column=j, value=col)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = CENTER
        c.border = BORDER
    ws.row_dimensions[hdr_row].height = 20

    for i, r in enumerate(arows):
        rr = hdr_row + 1 + i
        if i % 2 == 1:
            for j in range(1, n + 1):
                ws.cell(row=rr, column=j).fill = ZEBRA
        for j, col in enumerate(acols, start=1):
            v = r.get(col, "")
            if col in AGING_DATE_COLS and v:
                v = _to_date(v)
            elif col in AGING_NUM_COLS and str(v).strip() not in ("", "-"):
                v = _to_number(v)
            c = ws.cell(row=rr, column=j, value=v if v != "" else None)
            c.border = BORDER
            c.alignment = (CENTER if col in AGING_DATE_COLS | AGING_NUM_COLS
                           else WRAP)
            if col in AGING_DATE_COLS:
                c.number_format = "DD-MMM-YYYY"
            elif col in AGING_NUM_COLS:
                c.number_format = "#,##0.00" if col != "Days" else "#,##0"

    last_row = hdr_row + len(arows)
    last_col = get_column_letter(n)
    if arows:
        ws.auto_filter.ref = f"A{hdr_row}:{last_col}{last_row}"
    ws.freeze_panes = f"A{hdr_row + 1}"
    ws.print_title_rows = f"{hdr_row}:{hdr_row}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    for j, col in enumerate(acols, start=1):
        vals = [str(col)] + [str(r.get(col, "")) for r in arows]
        w = min(max(max(len(v) for v in vals) + 4, 12), 42)
        if col in AGING_NUM_COLS:
            w = min(w, 16)
        ws.column_dimensions[get_column_letter(j)].width = w


# ------------------------------------------------------------------
# Raw sheet: unrecognized PDFs. Every table and every text line is
# pulled with the document's own headings, each block labelled with
# its page. The yellow banner says it plainly: format not recognized,
# please verify. (Waqar, 2026-10-18: "pull everything".)
# ------------------------------------------------------------------

RAW_LABEL_FONT = Font(bold=True, size=11, color="1F4E79")


def _add_raw_sheet(wb, raw_docs):
    ws = wb.create_sheet("Raw")
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    maxw = 2
    for doc in raw_docs:
        for (_, heads, _) in doc["tables"]:
            maxw = max(maxw, len(heads))
    ncols = maxw + 1  # col A = line no / blank, rest = content

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
    tc = ws.cell(row=1, column=1, value="Raw extract — format not recognized")
    tc.font = TITLE_FONT
    ntables = sum(len(d["tables"]) for d in raw_docs)
    ntexts = sum(len(d["texts"]) for d in raw_docs)
    ws.cell(row=2, column=1,
            value=f"Generated {datetime.now():%d-%b-%Y %I:%M %p}  |  "
                  f"{ntables} table(s), {ntexts} text line(s) from "
                  f"{len(raw_docs)} file(s). Headings are the document's "
                  f"own — please verify before using.")
    ws.row_dimensions[1].height = 24

    def banner(text):
        ws.merge_cells(start_row=r[0], start_column=1,
                       end_row=r[0], end_column=ncols)
        c = ws.cell(row=r[0], column=1, value=text)
        c.font = RAW_LABEL_FONT
        c.fill = LOWCONF_FILL
        for j in range(1, ncols + 1):
            ws.cell(row=r[0], column=j).fill = LOWCONF_FILL
            ws.cell(row=r[0], column=j).border = BORDER

    r = [4]
    for doc in raw_docs:
        banner(f"\U0001F4C4 {doc['SourceFile']} — layout not recognized; "
               f"everything pulled below, please verify.")
        r[0] += 1
        for (pageno, tno, heads, data) in doc["tables"]:
            banner(f"Page {pageno} · Table {tno} ({len(heads)} columns)")
            r[0] += 1
            for j, h in enumerate(heads, start=2):
                c = ws.cell(row=r[0], column=j, value=h)
                c.font = HEADER_FONT
                c.fill = HEADER_FILL
                c.alignment = CENTER
                c.border = BORDER
            ws.cell(row=r[0], column=1, value="#").font = HEADER_FONT
            ws.cell(row=r[0], column=1).fill = HEADER_FILL
            ws.cell(row=r[0], column=1).alignment = CENTER
            ws.cell(row=r[0], column=1).border = BORDER
            r[0] += 1
            for i, row in enumerate(data):
                if i % 2 == 1:
                    for j in range(1, ncols + 1):
                        ws.cell(row=r[0], column=j).fill = ZEBRA
                ws.cell(row=r[0], column=1, value=i + 1).border = BORDER
                ws.cell(row=r[0], column=1).alignment = CENTER
                for j, v in enumerate(row, start=2):
                    c = ws.cell(row=r[0], column=j, value=v or None)
                    c.border = BORDER
                    c.alignment = WRAP
                r[0] += 1
        if doc["texts"]:
            banner(f"Text lines ({len(doc['texts'])} total, no tables found)")
            r[0] += 1
            for (pageno, lno, ln) in doc["texts"]:
                ws.merge_cells(start_row=r[0], start_column=2,
                               end_row=r[0], end_column=ncols)
                ws.cell(row=r[0], column=1, value=lno).border = BORDER
                ws.cell(row=r[0], column=1).alignment = CENTER
                c = ws.cell(row=r[0], column=2, value=f"[p{pageno}] {ln}")
                c.border = BORDER
                c.alignment = WRAP
                r[0] += 1
        r[0] += 1  # blank line between documents

    ws.column_dimensions["A"].width = 8
    for j in range(2, ncols + 1):
        ws.column_dimensions[get_column_letter(j)].width = 30
    ws.freeze_panes = "A5"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0


# ------------------------------------------------------------------
# Dispatches sheet (photo OCR). One row per shipment; low-confidence
# cells are highlighted yellow so Noman's review is a 2-minute glance.
# ------------------------------------------------------------------

DISPATCH_COLUMNS = [
    ("Date", "date"), ("Customer", "text"), ("CustomerCity", "text"),
    ("VehicleNo", "text"), ("DriverName", "text"), ("DriverCell", "text"),
    ("GatePassNo", "text"), ("BiltyNo", "text"), ("DispatchNo", "text"),
    ("DispatchRef", "text"), ("TotalReels", "int"), ("TotalPallets", "number"),
    ("NetWeight", "number"), ("FreightTotal", "number"),
    ("AdvancePaid", "number"), ("BalanceDue", "number"), ("Adda", "text"),
    ("SourceFiles", "text"), ("Confidence", "text"), ("Notes", "text"),
]

LOWCONF_FILL = PatternFill("solid", fgColor="FFF2CC")  # review-me yellow

# internal field names -> Dispatches column
_DISPATCH_FIELD_MAP = {
    "date": "Date", "customer": "Customer", "customer_city": "CustomerCity",
    "vehicle_no": "VehicleNo", "driver_name": "DriverName",
    "driver_cell": "DriverCell", "gate_pass_no": "GatePassNo",
    "bilty_no": "BiltyNo", "dispatch_no": "DispatchNo",
    "dispatch_ref": "DispatchRef", "total_reels": "TotalReels",
    "total_pallets": "TotalPallets", "net_weight": "NetWeight",
    "freight_total": "FreightTotal", "advance_paid": "AdvancePaid",
    "balance_due": "BalanceDue", "adda": "Adda", "contact_no": None,
}


def _add_dispatches_sheet(wb, shipments):
    """shipments: list of {"fields": {name: (value, conf)},
    "notes": [...], "source_files": [...]} from ocr_dispatch."""
    ws = wb.create_sheet("Dispatches")
    cols = [c for c, _ in DISPATCH_COLUMNS]
    kinds = dict(DISPATCH_COLUMNS)

    # title block
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(cols))
    tc = ws.cell(row=1, column=1, value="Dispatches (photo OCR)")
    tc.font = TITLE_FONT
    ws.cell(row=2, column=1,
            value=f"Generated {datetime.now():%d-%b-%Y %I:%M %p}  |  "
                  f"{len(shipments)} shipment(s). Yellow cells = low OCR "
                  f"confidence — please glance and correct.")
    ws.row_dimensions[1].height = 24

    hdr_row = 4
    for j, col in enumerate(cols, start=1):
        c = ws.cell(row=hdr_row, column=j, value=col)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = CENTER
        c.border = BORDER
    ws.row_dimensions[hdr_row].height = 20

    for i, ship in enumerate(shipments):
        rr = hdr_row + 1 + i
        fields = ship.get("fields", {})
        if i % 2 == 1:
            for j in range(1, len(cols) + 1):
                ws.cell(row=rr, column=j).fill = ZEBRA
        for j, col in enumerate(cols, start=1):
            kind = kinds[col]
            if col == "SourceFiles":
                v, conf = (", ".join(ship.get("source_files", [])), "")
            elif col == "Confidence":
                v, conf = (ship.get("fields", {}).get("_overall_conf",
                                                      "low"), "")
                v = str(v).upper()
            elif col == "Notes":
                extra = []
                cc, cc_conf = fields.get("contact_no", ("", ""))
                dc, _ = fields.get("driver_cell", ("", ""))
                if cc and cc != dc:
                    extra.append(f"Customer contact: {cc}")
                v = "; ".join(extra + ship.get("notes", []))
                conf = ""
            else:
                fname = next((k for k, c2 in _DISPATCH_FIELD_MAP.items()
                              if c2 == col), None)
                v, conf = fields.get(fname, ("", "")) if fname else ("", "")
            if kind == "date" and v:
                v = _to_date(v)
            elif kind in ("number", "int") and str(v).strip() not in ("", "-"):
                v = _to_number(v)
            c = ws.cell(row=rr, column=j, value=v if v != "" else None)
            c.border = BORDER
            c.alignment = CENTER if kind in ("date", "number", "int") else WRAP
            if kind == "date":
                c.number_format = "DD-MMM-YYYY"
            elif kind == "number":
                c.number_format = "#,##0.00"
            elif kind == "int":
                c.number_format = "#,##0"
            # low-confidence -> yellow (draws the review eye); zebra
            # striping underneath is replaced, not blended.
            if conf == "low" and v != "":
                c.fill = LOWCONF_FILL

    last_row = hdr_row + len(shipments)
    last_col = get_column_letter(len(cols))
    ws.auto_filter.ref = f"A{hdr_row}:{last_col}{last_row}"
    ws.freeze_panes = f"A{hdr_row + 1}"
    ws.print_title_rows = f"{hdr_row}:{hdr_row}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    for j, col in enumerate(cols, start=1):
        ws.column_dimensions[get_column_letter(j)].width = (
            32 if col == "Notes" else 22 if col == "SourceFiles"
            else 14 if kinds[col] in ("date", "int", "number") else 18)


if __name__ == "__main__":
    # tiny smoke test
    cols = ["S.No", "Customer", "Description", "Qty", "Rate", "Amount",
            "Delivery Date", "_source_file"]
    rows = [
        {"S.No": "1", "Customer": "ABC Traders", "Description": "Sugar 50kg",
         "Qty": "100", "Rate": "7,250", "Amount": "725000",
         "Delivery Date": "20-Oct-2026", "_source_file": "a.pdf"},
        {"S.No": "2", "Customer": "ABC Traders", "Description": "Rice 25kg",
         "Qty": "50", "Rate": "4,100", "Amount": "205000",
         "Delivery Date": "20-Oct-2026", "_source_file": "a.pdf"},
    ]
    write_excel(cols, rows, "/tmp/format_smoke.xlsx")
    print("smoke OK -> /tmp/format_smoke.xlsx")

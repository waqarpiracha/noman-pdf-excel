# Noman — PDF Requirements → Excel (pilot)

Mobile web app. Noman uploads the day's customer-requirement PDFs from his phone;
one tap builds a single **formatted** Excel file he downloads. No login, no
real-time plumbing — end-of-day batch, per his own requirement.

## Run locally

```bash
cd ~/workspace/noman-pdf-excel
python3 -m venv .venv && .venv/bin/pip install -r requirements.txt
.venv/bin/python server.py
# open http://localhost:5000  (on his phone: http://<machine-ip>:5000)
```

## Daily routine (Noman)

1. PDFs arrive on WhatsApp during the day (nothing changes).
2. Evening: open the link → **Choose PDFs or photos** → pick the day's files → upload.
3. Tap **⚙️ Make my Excel**.
4. Tap **⬇️ Download Excel** — formatted, filtered, ready to share/print.

## How it works

```
app/            mobile PWA page (vanilla JS, installable to home screen)
server.py       Flask: /upload /files /process /download /status
extract.py      ledger-format detector: statement header block (company,
                code, period, balances) + 12-col transaction tables ->
                normalized rows; day-first dates; Dr/Cr -> signed Balance;
                per-file/per-row dedupe; text fallback for odd pages
format_excel.py openpyxl: bold headers, autofilter, freeze panes,
                auto widths, date/number formats, Source sheet,
                Dispatches sheet (photo OCR, yellow = low confidence)
ocr_dispatch.py photo OCR: classify gate pass / bilty / dispatch note,
                per-doc field extraction with confidence, merge to one
                row per shipment (date+vehicle+customer; no triple-count);
                blank + REVIEW, never guessed
uploads/<date>/ today's PDFs + photos (flat files, kept per day)
output/         requirements_<date>.xlsx
samples/        synthetic test PDFs (A=clean table, B=messy, C=text-only)
```

**Adapting to his real PDFs:** only `extract.py` needs touching — the
`COLUMN_MAP` / `CANONICAL_COLUMNS` section at the top is clearly marked.
Add his real header spellings there; everything downstream stays the same.

## Hosting

- **Frontend** (`app/`): any static host. Netlify: drag-drop the `app/`
  folder or `netlify deploy --dir app`. It must reach the backend, so set
  the API base if they end up on different origins (currently same-origin).
- **Backend** (`server.py`): needs to be always-on with a persistent disk
  for `uploads/` and `output/`. Use his RDP (Windows: run under NSSM or
  Task Scheduler; Linux: systemd unit) behind a reverse proxy
  (nginx/Caddy) with HTTPS, e.g. `https://pdf2excel.<his-domain>/`.
  A cheap VPS works too. **Do not put the backend on Netlify** — functions
  are ephemeral and uploads would vanish.
- **PWA install:** served over HTTPS, the page is installable to the home
  screen (manifest + icons + service worker included).

## Status (2026-10-16)

**Tuned to Noman's REAL samples and verified end-to-end.** His PDFs are
customer account statements (text-based, clean tables — no OCR needed):
12-column ledger (Date | InvoiceNo | Description | Size | Length | Qty |
Rate | Amount | SalesTax | Debit | Credit | Balance), dates DD-MM-YYYY
(day-first), Balance as Dr/Cr → signed number.

Test on both real PDFs: **148/148 rows extracted, zero parse failures**;
running-balance continuity (PrevBalance + Debit − Credit = Balance) holds
for every row — the strongest possible correctness check. Output workbook
has 3 sheets: **Transactions** (15 cols incl. Company/CustomerCode/SourceFile),
**Statements** (one row per source file: company, code, period, salesperson,
previous balance, NTN/STRN, row count), **Source** (file/page provenance).

**Open question — RESOLVED 2026-10-17:** the 3 sample photos ARE in scope
(Waqar confirmed). They are dispatch documents (gate pass + bilty +
dispatch note, one Shine→Adnan Traders shipment) and get a **Dispatches**
sheet in the same workbook, one row per shipment.

## Dispatch photos (added 2026-10-17)

1. Evening: **Choose PDFs or photos** → pick the day's PDFs *and* dispatch
   photos (JPG/PNG) → upload.
2. Tap **⚙️ Make my Excel** (photo OCR adds ~30–40s per photo).
3. Tap **⬇️ Download Excel** — the **Dispatches** sheet has one row per
   shipment: Date | Customer | … | VehicleNo | DriverName | BiltyNo |
   NetWeight | AdvancePaid | … | SourceFiles | Confidence | Notes.
4. **2-minute review:** glance at the **yellow cells** (low OCR confidence)
   and the Notes column (REVIEW flags) — type any missing values straight
   off the photos. That's it; don't recheck the white cells.

**Accuracy limits (by design, not bugs):** the printed dispatch note OCRs
near-perfectly; handwritten gate-pass entries and the Urdu bilty are
partial — unreadable fields stay **blank with a REVIEW note** rather than
guessed. Photos of the same shipment (gate pass + bilty + dispatch note)
are merged into one row via date+vehicle+customer so nothing is
triple-counted; cross-document agreement raises confidence.

## Photo OCR setup (tesseract)

```bash
sudo apt-get install tesseract-ocr tesseract-ocr-eng tesseract-ocr-urd
.venv/bin/pip install -r requirements.txt   # adds pytesseract, Pillow
```

**IMPORTANT:** tesseract is a *system* package — if the VM is ever
replaced, `/usr` is wiped and photo OCR breaks. Re-run the apt line above
(the server fails loudly with install instructions if the binary is
missing; PDFs keep working regardless).

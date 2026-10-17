"""
Noman PDF -> Excel : tiny backend.

Endpoints:
  GET  /            -> the mobile web app (app/index.html)
  POST /upload      -> multipart PDFs + dispatch photos -> uploads/<YYYY-MM-DD>/
  GET  /files       -> today's uploaded files (JSON)
  POST /process     -> extract PDFs + OCR photos -> output/*.xlsx
                       (adds a "Dispatches" sheet, one row per shipment)
  GET  /download    -> today's finished .xlsx
  GET  /status      -> today, upload count, output readiness (JSON)

No auth, no database. Flat files only (pilot scale).
Run:  .venv/bin/python server.py   ->  http://localhost:5000

Photo OCR needs the tesseract system binary (apt: tesseract-ocr
tesseract-ocr-eng tesseract-ocr-urd). Without it, PDFs still process;
photos are skipped with a note.
"""

import os
import glob
import json
from datetime import date
from flask import Flask, request, jsonify, send_from_directory, send_file

from extract import process_files
from format_excel import write_excel

BASE = os.path.dirname(os.path.abspath(__file__))
APP_DIR = os.path.join(BASE, "app")
UPLOAD_ROOT = os.path.join(BASE, "uploads")
OUTPUT_ROOT = os.path.join(BASE, "output")
os.makedirs(UPLOAD_ROOT, exist_ok=True)
os.makedirs(OUTPUT_ROOT, exist_ok=True)

app = Flask(__name__, static_folder=APP_DIR, static_url_path="")
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024  # 100 MB total

ALLOWED_EXTS = {".pdf", ".jpg", ".jpeg", ".png"}


def today_str():
    return date.today().isoformat()


def today_dir():
    d = os.path.join(UPLOAD_ROOT, today_str())
    os.makedirs(d, exist_ok=True)
    return d


def output_path(day=None):
    return os.path.join(OUTPUT_ROOT, f"requirements_{day or today_str()}.xlsx")


def safe_name(directory, filename):
    """Collision-proof filename inside directory."""
    base = os.path.basename(filename) or "upload.pdf"
    name, ext = os.path.splitext(base)
    cand, i = base, 1
    while os.path.exists(os.path.join(directory, cand)):
        i += 1
        cand = f"{name}_{i}{ext}"
    return cand


def _ext(path):
    return os.path.splitext(path)[1].lower()


@app.get("/")
def index():
    return send_from_directory(APP_DIR, "index.html")


@app.post("/upload")
def upload():
    files = request.files.getlist("files") or request.files.getlist("file")
    if not files:
        return jsonify({"ok": False, "error": "no files received"}), 400
    saved, skipped = [], []
    dest = today_dir()
    for f in files:
        if not f.filename:
            continue
        if _ext(f.filename) not in ALLOWED_EXTS:
            skipped.append(f.filename)
            continue
        name = safe_name(dest, f.filename)
        f.save(os.path.join(dest, name))
        saved.append(name)
    return jsonify({"ok": True, "saved": saved, "skipped": skipped,
                    "day": today_str()})


@app.get("/files")
def files():
    day = request.args.get("day", today_str())
    d = os.path.join(UPLOAD_ROOT, day)
    names = []
    if os.path.isdir(d):
        names = sorted(os.path.basename(p) for p in glob.glob(d + "/*")
                       if _ext(p) in ALLOWED_EXTS and os.path.isfile(p))
    sizes = {n: os.path.getsize(os.path.join(d, n)) for n in names}
    return jsonify({"ok": True, "day": day, "files": names, "sizes": sizes})


@app.post("/process")
def process():
    day = (request.get_json(silent=True) or {}).get("day", today_str())
    d = os.path.join(UPLOAD_ROOT, day)
    all_files = sorted(glob.glob(os.path.join(d, "*")))
    pdfs = [p for p in all_files if _ext(p) == ".pdf"]
    photos = [p for p in all_files if _ext(p) in {".jpg", ".jpeg", ".png"}]
    if not pdfs and not photos:
        return jsonify({"ok": False,
                        "error": "no PDFs or photos uploaded for " + day}), 400

    cols, rows, notes, statements = process_files(pdfs) if pdfs else (
        [], [], [], [])

    dispatches, ocr_notes = [], []
    if photos:
        try:
            from ocr_dispatch import process_dispatch_images
            dispatches, ocr_notes = process_dispatch_images(photos)
        except RuntimeError as e:
            # tesseract binary missing: PDFs still process, photos noted.
            ocr_notes = [f"photo OCR skipped: {e}"]
        except Exception as e:
            ocr_notes = [f"photo OCR failed: {e}"]

    out = output_path(day)
    write_excel(cols, rows, out,
                title=f"Customer Account Statements — {day}",
                statements=statements, dispatches=dispatches)
    return jsonify({
        "ok": True, "day": day,
        "files_processed": len(pdfs),
        "photos_processed": len(photos),
        "shipments": len(dispatches),
        "rows": len(rows),
        "columns": [c for c in cols if not c.startswith("_")],
        "notes": notes + ocr_notes,
        "download": "/download?day=" + day,
    })


@app.get("/download")
def download():
    day = request.args.get("day", today_str())
    out = output_path(day)
    if not os.path.exists(out):
        return jsonify({"ok": False, "error": "not processed yet"}), 404
    return send_file(out, as_attachment=True,
                     download_name=f"requirements_{day}.xlsx")


@app.get("/status")
def status():
    day = today_str()
    d = os.path.join(UPLOAD_ROOT, day)
    names = []
    if os.path.isdir(d):
        names = [p for p in glob.glob(os.path.join(d, "*"))
                 if _ext(p) in ALLOWED_EXTS and os.path.isfile(p)]
    out = output_path(day)
    return jsonify({
        "ok": True, "day": day, "uploads": len(names),
        "ready": os.path.exists(out),
        "download": "/download?day=" + day if os.path.exists(out) else None,
    })


if __name__ == "__main__":
    print("Noman PDF->Excel  |  http://localhost:5000")
    app.run(host="0.0.0.0", port=5000)

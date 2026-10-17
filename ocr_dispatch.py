"""
Photo OCR for dispatch documents (gate pass / bilty / dispatch note).

HONESTY CONTRACT (read before touching this file):
  - The PRINTED dispatch note OCRs well: most of its fields come out HIGH.
  - HANDWRITTEN gate-pass entries and the URDU bilty OCR poorly: most of
    their fields come out LOW or blank.
  - NEVER GUESS. A field we cannot read is left blank and flagged REVIEW.
    A wrong value in Noman's Excel is worse than an empty cell.
    (Observed while building: fuzzy plate matching manufactured "OGL-1350",
    "TEP-4555" from pure noise; a phone fragment "0300" posed as a bilty
    number. All such cleverness has been REMOVED.)
  - Confidence travels with every field: "high" | "low". The Excel writer
    highlights low-confidence cells yellow so Noman's review is a 2-minute
    glance, not a re-read of every document.
  - The 3 sample photos describe ONE shipment. merge_shipments() links them
    so nothing is triple-counted: strict key match first, scored fuzzy match
    second, single-shipment-in-batch fallback last (always noted).

Requires: tesseract binary + eng + urd tessdata
  (apt: tesseract-ocr tesseract-ocr-eng tesseract-ocr-urd),
  pytesseract, Pillow.
NOTE: /usr is wiped on VM replacement — if OCR suddenly returns empty,
reinstall tesseract (README documents this). process_dispatch_images()
fails LOUDLY when the binary is missing; it never silently returns "".
"""

import os
import re
import shutil
import difflib
from collections import defaultdict

from PIL import Image, ImageOps
import pytesseract

# ------------------------------------------------------------------
# OCR plumbing
# ------------------------------------------------------------------

def preprocess(path):
    """Grayscale -> 2x upscale -> autocontrast. (No deskew: keep it simple
    and honest. A badly skewed photo yields LOW/blank, not guesses.)"""
    img = Image.open(path).convert("L")
    w, h = img.size
    img = img.resize((w * 2, h * 2), Image.LANCZOS)
    return ImageOps.autocontrast(img, cutoff=2)


def crop_frac(img, box):
    w, h = img.size
    x0, y0, x1, y1 = box
    return img.crop((int(w * x0), int(h * y0), int(w * x1), int(h * y1)))


def ocr(img, lang="eng", psm=4):
    # Deliberately loud: a missing binary must never silently empty results.
    return pytesseract.image_to_string(img, lang=lang,
                                       config=f"--psm {psm}") or ""


# Fractional regions, tuned to Noman's 3 sample photos. Layout-specific;
# new document layouts need their boxes re-tuned.
REGIONS = {
    "dispatch_note": [
        ("full", None, "eng", 4),
        ("full3", None, "eng", 3),
        ("full6", None, "eng", 6),
        ("rightmid", (0.52, 0.13, 1.0, 0.38), "eng", 6),
        ("table", (0.02, 0.36, 1.00, 0.62), "eng", 6),
    ],
    "gate_pass": [
        ("full", None, "eng", 6),
        ("header", (0.00, 0.25, 1.00, 0.50), "eng", 6),
    ],
    "bilty": [
        ("full", None, "eng+urd", 6),
        ("topstrip", (0.00, 0.18, 1.00, 0.34), "eng", 6),
    ],
}


def ocr_all(path, doc_type):
    base = preprocess(path)
    out = {}
    for name, box, lang, psm in REGIONS.get(doc_type, REGIONS["gate_pass"]):
        img = crop_frac(base, box) if box else base
        out[name] = ocr(img, lang=lang, psm=psm)
    return out


def classify_image(path):
    """eng PSM 3+4 for Latin titles (urd mangles them; no single PSM reads
    every title), plus a urd pass purely to count Urdu script for bilties."""
    base = preprocess(path)
    txt_en = ocr(base, lang="eng", psm=3) + "\n" + ocr(base, lang="eng", psm=4)
    txt_ur = ocr(base, lang="urd", psm=6)
    urdu_chars = sum(1 for ch in txt_ur if "\u0600" <= ch <= "\u06FF")
    t = txt_en.upper()
    if "DISPATCH NOTE" in t:
        return "dispatch_note", "high"
    if "GATE PASS" in t:
        return "gate_pass", "high"
    if urdu_chars > 30:
        return "bilty", "high"
    return "unknown", "low"


# ------------------------------------------------------------------
# Field helpers. Every field -> (value, confidence in {"high","low",""}).
# ------------------------------------------------------------------

def _clean(s):
    return re.sub(r"\s+", " ", (s or "")).strip(" .,:;|[]()")


def _denoise(text):
    """'|' separators are pure OCR noise in these docs; flatten them."""
    return re.sub(r"[ \t]+", " ", text.replace("|", " "))


def _first(pattern, text, flags=re.IGNORECASE):
    m = re.search(pattern, text, flags)
    return _clean(m.group(1)) if m else ""


def _parse_doc_date(s):
    """'16/Oct/26', '16-10-26', '16-10-2026' -> date. Day-first. None if bad."""
    from datetime import datetime as _dt
    s = _clean(s)
    for fmt in ("%d/%b/%y", "%d/%b/%Y", "%d-%m-%y", "%d-%m-%Y",
                "%d/%m/%y", "%d/%m/%Y"):
        try:
            return _dt.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def _parse_money(s):
    s = _clean(s).replace(",", "")
    if not s or s == "-":
        return None
    try:
        return float(s)
    except ValueError:
        return None


def _norm_phone(s):
    """'3476551926' -> '0347-6551926'. Anything else: cleaned raw or ''."""
    d = re.sub(r"\D", "", s or "")
    if len(d) == 10 and d.startswith("3"):
        d = "0" + d
    if len(d) == 11 and d.startswith("03"):
        return d[:4] + "-" + d[4:]
    return ""


_VEHICLE_STRICT = re.compile(r"\b([A-Z]{2,4})-(\d{3,4})\b")


def _find_vehicle(text):
    """Pakistani plate like GLT-3509. STRICT pattern only -> high.
    No fuzzy fallback (it manufactured plates from noise). Blank + REVIEW
    when unreadable; the shipment still merges on (date, customer)."""
    m = _VEHICLE_STRICT.search(text.upper())
    if m:
        return f"{m.group(1)}-{m.group(2)}", "high"
    return None, None


_MONEY = re.compile(r"(\d{1,3}(?:,\d{3})*\.\d{2})")


def _item_weights(table_text):
    """Per-row weights from the item table — weight ONLY (no reels/pallets).

    Row = a line mentioning a film code like SPMZ-25-810-16000.
    weight = the largest money-shaped number in the row (weights dwarf the
    pallet counts, so max() is sound). Reels/pallets are NOT parsed: the
    OCR mangles small ints ("18" -> "is", "825.00" -> reels=825) and a
    wrong count is worse than a blank cell.
    The same row appears in several OCR regions -> dedupe by film code.
    Returns (weights_by_code, n_codes).
    """
    by_code = {}
    for line in table_text.splitlines():
        m = re.search(r"SPM[A-Z0-9]*[- ]25[- ]\d+[- ]16000", line,
                      re.IGNORECASE)
        if not m:
            continue
        code = re.sub(r"[^A-Z0-9]", "", m.group(0).upper())
        moneys = [_parse_money(x.group(1)) for x in _MONEY.finditer(line)]
        moneys = [x for x in moneys if x is not None]
        if moneys and code not in by_code:
            by_code[code] = max(moneys)
    return by_code


# ------------------------------------------------------------------
# Per-document extractors -> {field: (value, confidence)}
# ------------------------------------------------------------------

def extract_dispatch_note(regions):
    flat = _denoise("\n".join(regions.values()))
    F = {"doc_type": ("dispatch_note", "high")}

    F["dispatch_ref"] = _hv(_first(r"REF\s*#?\s*([A-Z]{2,}\d+)", flat))
    F["dispatch_no"] = _hv(_first(r"D\.?\s*NO\.?\s*(\d{3,5})", flat))
    F["bilty_no"] = _hv(_first(r"[Bb]il\w*\s*No\.?\s*(\d{3,5})", flat))

    cust = _first(
        r"DISPATCH TO:?\s*(?:\d{1,2}/[A-Za-z]{3}/\d{2,4}\s*)?"
        r"([A-Z][A-Z &]{2,40}?)(?:\s{2,}|$)", flat)
    F["customer"] = ((cust.title(), "high") if cust else ("", ""))

    city = ""
    if re.search(r"ISLAMABAD", flat, re.IGNORECASE):
        city = "Islamabad"
    elif re.search(r"HUMAK", flat, re.IGNORECASE):
        city = "Humak (Islamabad)"
    F["customer_city"] = (city, "high" if city else "")

    v, vc = _find_vehicle(flat)
    F["vehicle_no"] = (v or "", vc or "")

    drv = _first(r"DRIVER NAME\s+([A-Z][A-Z ]{2,30}?)(?:\s{2,}|$)", flat)
    F["driver_name"] = ((drv.title(), "high") if drv else ("", ""))

    dcell = _norm_phone(_first(r"DRIVER CELL[^\d]*(\d[\d\s\-]{6,14})", flat))
    F["driver_cell"] = ((dcell, "high") if dcell else ("", ""))

    ccell = _norm_phone(_first(r"CONTACT\s*NO\.?\s*([\d\-\s]{7,16})", flat))
    F["contact_no"] = ((ccell, "high") if ccell else ("", ""))

    adda = _first(r"ADDA\s+([A-Z][A-Z ]{2,30}?)(?:\s{2,}|$)", flat)
    F["adda"] = ((adda.title(), "high") if adda else ("", ""))

    d = None
    m = re.search(r"(\d{1,2}/[A-Za-z]{3}/\d{2,4})", flat)
    if m:
        d = _parse_doc_date(m.group(1))
    if d is None:
        m = re.search(r"\b(\d{1,2}-\d{1,2}-\d{2,4})\b", flat)
        if m:
            d = _parse_doc_date(m.group(1))
    F["date"] = ((d.isoformat(), "high") if d else ("", ""))

    # Net weight: direct label (HIGH); else sum of item-row weights (LOW).
    # The sum only counts when EVERY item row yielded a weight — a partial
    # sum would be a wrong total.
    nw, nw_conf, nw_note = "", "", ""
    m = re.search(r"NET WEIGHT\s*([\d,\.]+)", flat, re.IGNORECASE)
    if m:
        val = _parse_money(m.group(1))
        if val is not None:
            nw, nw_conf = val, "high"
    weights = _item_weights(regions.get("table", "") + "\n" + flat)
    if nw == "" and weights:
        # every distinct film code needs a weight, else the sum is partial
        codes_in_text = {re.sub(r"[^A-Z0-9]", "", m.group(0).upper())
                         for m in re.finditer(
                             r"SPM[A-Z0-9]*[- ]25[- ]\d+[- ]16000",
                             regions.get("table", "") + "\n" + flat,
                             re.IGNORECASE)}
        if codes_in_text and codes_in_text <= set(weights):
            nw, nw_conf = round(sum(weights.values()), 2), "low"
            nw_note = "summed from item rows"
    F["net_weight"] = (nw, nw_conf)
    if nw_note:
        F["_nw_note"] = (nw_note, "")

    # Reels/pallets totals: not parsed from OCR (small ints mangle too
    # easily: "18"->"is", "825.00"->reels=825). Blank + REVIEW; Noman
    # reads two numbers off the photo.
    F["total_reels"] = ("", "")
    F["total_pallets"] = ("", "")
    return F


def _hv(val):
    """(value, 'high' if non-blank else '')."""
    return (val, "high") if val else ("", "")


def extract_gate_pass(regions):
    flat = _denoise("\n".join(regions.values()))
    F = {"doc_type": ("gate_pass", "high")}

    # Handwriting OCR is unreliable: everything here caps at LOW, and most
    # fields will be blank + REVIEW (the dispatch note fills the gaps).
    sr = _first(r"Sr\.?\s*No:?\s*(\d+)", flat)
    F["gate_pass_no"] = ((sr, "low") if sr else ("", ""))

    cust = _first(r"Delivered to M/S:?\s*([A-Za-z][A-Za-z ]{2,40}?)(?:\s{2,}|$)",
                  flat)
    F["customer"] = ((cust, "low")
                     if cust and re.search(r"[A-Za-z]{3,}", cust) else ("", ""))

    city = ""
    m = re.search(r"\b(Islamabad|Multan|Lahore|Karachi|Rawalpindi|Peshawar|"
                  r"Faisalabad)\b", flat, re.IGNORECASE)
    if m:
        city = m.group(1).title()
    F["customer_city"] = ((city, "low") if city else ("", ""))

    cell = _norm_phone(_first(r"Cell No:?\s*([\d\-\s]{7,18})", flat))
    F["contact_no"] = ((cell, "low") if cell else ("", ""))

    v, _ = _find_vehicle(flat)
    F["vehicle_no"] = ((v, "low") if v else ("", ""))

    d = None
    m = re.search(r"Date:?\s*([0-9\-/\.]{6,12})", flat)
    if m:
        d = _parse_doc_date(m.group(1))
    F["date"] = ((d.isoformat(), "low") if d else ("", "")

    )
    F["net_weight"] = ("", "")
    F["total_reels"] = ("", "")
    F["total_pallets"] = ("", "")
    return F


def extract_bilty(regions):
    top = _denoise(regions.get("topstrip", ""))
    full = _denoise(regions.get("full", ""))
    F = {"doc_type": ("bilty", "high")}

    # Bilty number: 4-digit in the top strip. Printed phones are 7 digits;
    # skip 03-prefixed fragments (Pakistani mobile prefix).
    bilty = ""
    for cand in re.findall(r"(?<!\d)(\d{4})(?!\d)", top):
        if not cand.startswith("03"):
            bilty = cand
            break
    F["bilty_no"] = ((bilty, "low") if bilty else ("", ""))

    d = None
    m = re.search(r"\b(\d{1,2}-\d{1,2}-\d{2,4})\b", top)
    if m:
        d = _parse_doc_date(m.group(1))
    F["date"] = ((d.isoformat(), "low") if d else ("", ""))

    v, _ = _find_vehicle(top + " " + full)
    F["vehicle_no"] = ((v, "low") if v else ("", ""))

    drv = ""
    m = re.search(r"\b(Rizwan|Usman|Muhammad|Ahmed|Ali|Hassan|Imran|Razaq)\b",
                  top + " " + full)
    if m:
        drv = m.group(1)
    F["driver_name"] = ((drv, "low") if drv else ("", ""))

    cust = ""
    if re.search(r"Adnan\s*Traders", full, re.IGNORECASE):
        cust = "Adnan Traders"
    elif re.search(r"Shine\s*Packages", full, re.IGNORECASE):
        cust = "Shine Packages (sender)"
    F["customer"] = ((cust, "low") if cust else ("", ""))

    city = "Islamabad" if re.search(r"اسلام آباد", regions.get("full", "")) \
        else ""
    F["customer_city"] = ((city, "low") if city else ("", ""))

    adv = ""
    m = re.search(r"[Pp]aid\s*(\d[\d,\.]*)", full)
    if m:
        val = _parse_money(m.group(1))
        if val is not None:
            adv = val
    F["advance_paid"] = ((adv, "low") if adv != "" else ("", ""))

    # Freight total / balance columns: handwriting too poor -> REVIEW.
    F["freight_total"] = ("", "")
    F["balance_due"] = ("", "")
    F["net_weight"] = ("", "")
    F["total_reels"] = ("", "")
    F["total_pallets"] = ("", "")
    return F


EXTRACTORS = {
    "dispatch_note": extract_dispatch_note,
    "gate_pass": extract_gate_pass,
    "bilty": extract_bilty,
}


def process_image(path):
    """OCR one dispatch photo -> (doc_type, fields, notes)."""
    fname = os.path.basename(path)
    doc_type, class_conf = classify_image(path)
    notes = []
    if doc_type == "unknown":
        notes.append(f"{fname}: could not classify document type (REVIEW)")
        return doc_type, {}, notes
    if class_conf == "low":
        notes.append(f"{fname}: document type uncertain (REVIEW)")
    regions = ocr_all(path, doc_type)
    fields = EXTRACTORS[doc_type](regions)
    return doc_type, fields, notes


# ------------------------------------------------------------------
# Merge: one row per shipment.
# ------------------------------------------------------------------

def _norm_key(v):
    return re.sub(r"[^A-Z0-9]", "", str(v or "").upper())


def _fuzzy(a, b):
    a, b = _norm_key(a), _norm_key(b)
    if not a or not b:
        return 0.0
    return difflib.SequenceMatcher(None, a, b).ratio()


def _key_tuple(fields):
    d = (fields.get("date") or ("", ""))[0]
    v = (fields.get("vehicle_no") or ("", ""))[0]
    c = (fields.get("customer") or ("", ""))[0]
    return (d or "NODATE", _norm_key(v) or "NOVEH",
            _norm_key(c).rstrip("S")[:14] or "UNKNOWN")


def _link_score(f_new, f_grp):
    """How strongly does a doc belong to a shipment?"""
    s = 0.0
    dn = (f_new.get("date") or ("", ""))[0]
    dg = (f_grp.get("date") or ("", ""))[0]
    if dn and dg and dn == dg:
        s += 3
    vn = (f_new.get("vehicle_no") or ("", ""))[0]
    vg = (f_grp.get("vehicle_no") or ("", ""))[0]
    if vn and vg and _norm_key(vn) == _norm_key(vg):
        s += 3
    cn = (f_new.get("customer") or ("", ""))[0]
    cg = (f_grp.get("customer") or ("", ""))[0]
    r = _fuzzy(cn, cg)
    if r >= 0.8:
        s += 3
    elif r >= 0.55:
        s += 2
    nn = (f_new.get("net_weight") or ("", ""))[0]
    ng = (f_grp.get("net_weight") or ("", ""))[0]
    try:
        if nn != "" and ng != "" and abs(float(nn) - float(ng)) \
                < max(1.0, float(ng) * 0.01):
            s += 2
    except (ValueError, TypeError):
        pass
    return s


_DOC_RANK = {"dispatch_note": 0, "gate_pass": 1, "bilty": 2, "unknown": 3}
_CONF_RANK = {"high": 0, "low": 1, "": 2}

SHIPMENT_FIELDS = [
    "date", "customer", "customer_city", "vehicle_no", "driver_name",
    "driver_cell", "gate_pass_no", "bilty_no", "dispatch_no",
    "dispatch_ref", "total_reels", "total_pallets", "net_weight",
    "freight_total", "advance_paid", "balance_due", "adda", "contact_no",
]


def _merge_group(members):
    """members: list of (fname, doc_type, fields) -> shipment dict."""
    merged, notes = {}, []
    files = [f for f, _, _ in members]
    for fname, doc_type, fields in members:
        # internal annotations (underscore fields) become notes, not columns
        for k, (v, _) in fields.items():
            if k.startswith("_") and v:
                notes.append(f"{fname}: {v}")
    for col in SHIPMENT_FIELDS:
        cands = []
        for fname, doc_type, fields in members:
            val, conf = fields.get(col, ("", ""))
            if val != "":
                cands.append((val, conf, doc_type, fname))
        if not cands:
            merged[col] = ("", "")
            continue
        cands.sort(key=lambda c: (_CONF_RANK.get(c[1], 2),
                                  _DOC_RANK.get(c[2], 3)))
        best_val, best_conf, best_doc, _ = cands[0]
        # Conflict note only when a rival is at least as confident —
        # low-confidence OCR noise must not drag down a high reading.
        rivals = {(str(v), d) for v, c, d, _ in cands
                  if str(v) != str(best_val)
                  and _CONF_RANK.get(c, 2) <= _CONF_RANK.get(best_conf, 2)}
        if rivals:
            detail = "; ".join(f"{d}={v}" for v, d in sorted(rivals))
            notes.append(f"CONFLICT {col}: kept '{best_val}' ({best_doc}); "
                         f"also saw {detail} — REVIEW")
            best_conf = "low"
        merged[col] = (best_val, best_conf)

    ft, _ = merged.get("freight_total", ("", ""))
    ap, _ = merged.get("advance_paid", ("", ""))
    if ft != "" and ap != "":
        try:
            merged["balance_due"] = (round(float(ft) - float(ap), 2), "low")
        except (ValueError, TypeError):
            pass

    key_fields = ["date", "customer", "vehicle_no", "net_weight"]
    blanks = [c for c in key_fields if merged[c][0] == ""]
    confs = {merged[c][1] for c in key_fields if merged[c][0] != ""}
    if blanks:
        notes.append("REVIEW: " + ", ".join(blanks) + " unreadable")
        overall = "low"
    elif confs == {"high"}:
        overall = "high"
    else:
        overall = "mixed"
    merged["_overall_conf"] = overall
    return {"fields": merged, "notes": notes, "source_files": files}


def merge_shipments(docs):
    """docs: list of (fname, doc_type, fields) -> list of shipment dicts.

    1. Strict pass: docs with >=2 of (date, vehicle, customer) group by key.
    2. Scored pass: leftovers join the best group on link score >= 4.
    3. Batch fallback: if exactly ONE group exists in the batch, leftovers
       join it ("only shipment in batch — REVIEW"); else they get their
       own rows + REVIEW. Nothing is silently dropped or triple-counted.
    """
    groups = defaultdict(list)   # key -> members
    keyed = set()
    for i, (fname, doc_type, fields) in enumerate(docs):
        filled = sum(1 for c in ("date", "vehicle_no", "customer")
                     if (fields.get(c) or ("", ""))[0] != "")
        if filled >= 2:
            groups[_key_tuple(fields)].append((fname, doc_type, fields))
            keyed.add(i)

    # anchor order: dispatch notes first (printed = most reliable keys)
    leftovers = [(fname, doc_type, fields)
                 for i, (fname, doc_type, fields) in enumerate(docs)
                 if i not in keyed]
    fallback_linked = set()
    for fname, doc_type, fields in leftovers:
        best_key, best_score = None, 0.0
        for key, members in groups.items():
            rep_fields = _merge_group(members)["fields"]
            sc = _link_score(fields, rep_fields)
            if sc > best_score:
                best_key, best_score = key, sc
        if best_key is not None and best_score >= 4:
            groups[best_key].append((fname, doc_type, fields))
        elif len(groups) == 1:
            only = next(iter(groups))
            groups[only].append((fname, doc_type, fields))
            fallback_linked.add(fname)
        else:
            groups[("UNLINKED", fname)].append((fname, doc_type, fields))

    shipments = []
    for key, members in groups.items():
        ship = _merge_group(members)
        if key[0] == "UNLINKED":
            ship["notes"].append(
                "REVIEW: could not link this photo to a shipment")
        for fname, _, _ in members:
            if fname in fallback_linked:
                ship["notes"].append(
                    f"{fname}: linked as the only shipment in this batch — "
                    f"REVIEW the link")
                fallback_linked.discard(fname)
        shipments.append(ship)
    return shipments


def process_dispatch_images(paths):
    """paths: list of image files -> (shipments, notes)."""
    if not shutil.which("tesseract"):
        raise RuntimeError(
            "tesseract binary not found. Install it first: "
            "sudo apt-get install tesseract-ocr tesseract-ocr-eng tesseract-ocr-urd "
            "(see README.md). Photo OCR cannot run without it.")
    docs, notes = [], []
    for p in sorted(paths):
        try:
            doc_type, fields, n = process_image(p)
        except Exception as e:  # never let one bad photo kill the batch
            notes.append(f"{os.path.basename(p)}: OCR failed ({e})")
            continue
        notes.extend(n)
        docs.append((os.path.basename(p), doc_type, fields))
    shipments = merge_shipments(docs)
    return shipments, notes


if __name__ == "__main__":
    import sys
    import json
    shipments, notes = process_dispatch_images(sys.argv[1:])
    for s in shipments:
        print("=" * 60)
        print("overall:", s["fields"].get("_overall_conf"),
              "| files:", len(s["source_files"]))
        for f in s["source_files"]:
            print("   -", f)
        for col in SHIPMENT_FIELDS:
            v, c = s["fields"][col]
            if v != "":
                print(f"  {col:15s} {str(v):28s} [{c}]")
        for n in s["notes"]:
            print("  NOTE:", n)
    for n in notes:
        print("NOTE:", n)
    print(json.dumps({"shipments": len(shipments)}))

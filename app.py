import importlib.util
import sys

_NEED = {"flask": "flask", "PIL": "pillow", "openpyxl": "openpyxl", "numpy": "numpy"}
_missing = [pip for mod, pip in _NEED.items() if importlib.util.find_spec(mod) is None]
if _missing:
    print("\n  Some packages are missing. Run this command, then start the app again:\n")
    print("      pip install " + " ".join(_missing) + " rapidocr-onnxruntime\n")
    input("  Press Enter to close...")
    sys.exit(1)

import csv
import hashlib
import io
import json
import re
import secrets
import socket
import statistics
import threading
from datetime import datetime
from pathlib import Path

import numpy as np
from flask import Flask, Response, jsonify, redirect, request, send_file
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from PIL import Image, ImageOps

BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR / "data"
PHOTO_DIR = DATA_DIR / "photos"
PHOTO_DIR.mkdir(parents=True, exist_ok=True)
FILE_DIR = DATA_DIR / "files"
FILE_DIR.mkdir(parents=True, exist_ok=True)
XLSX_PATH = DATA_DIR / "entries.xlsx"
INDEX = "Index"
INDEX_COLS = ["Entry ID", "Entry Date", "Photo", "Sheet", "Rows", "Columns", "File", "Hash", "Type"]
FONT = "Arial"

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024
lock = threading.Lock()

PIN = None
LOCAL = ("127.0.0.1", "::1")
LOCK_HTML = """<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Smart Data Entry</title></head><body style="margin:0;min-height:100vh;display:flex;align-items:center;justify-content:center;background:#000;color:#e6ebf3;font-family:system-ui,sans-serif">
<form method="get" action="/" style="text-align:center;padding:24px"><h2>Smart Data Entry</h2><p>Enter the PIN shown on the computer</p>
<input name="pin" inputmode="numeric" autofocus style="font-size:24px;padding:10px;width:9em;text-align:center;border-radius:8px;border:1px solid #2fb8ff;background:#0a0e1c;color:#fff">
<p><button style="font-size:18px;padding:10px 24px;border-radius:8px;border:0;background:#2fb8ff;font-weight:700">Open</button></p></form></body></html>"""


@app.before_request
def guard():
    if not PIN or request.remote_addr in LOCAL:
        return None
    if request.cookies.get("sde_pin") == PIN:
        return None
    if request.args.get("pin") == PIN:
        resp = redirect(request.path)
        resp.set_cookie("sde_pin", PIN, max_age=60 * 60 * 24 * 30, httponly=True, samesite="Lax")
        return resp
    if request.path.startswith("/api/"):
        return jsonify(error="Locked - open the link with the PIN first."), 401
    return Response(LOCK_HTML, 401, mimetype="text/html")


@app.after_request
def cors(resp):
    resp.headers["Access-Control-Allow-Origin"] = "*"
    resp.headers["Cache-Control"] = "no-store"
    resp.headers["Access-Control-Allow-Methods"] = "GET, POST, DELETE, OPTIONS"
    resp.headers["Access-Control-Allow-Headers"] = "Content-Type"
    return resp


_rapid = None


def words_rapidocr(img):
    global _rapid
    from rapidocr_onnxruntime import RapidOCR
    if _rapid is None:
        _rapid = RapidOCR()
    result, _ = _rapid(np.array(img))
    words = []
    for box, text, score in (result or []):
        xs = [p[0] for p in box]
        ys = [p[1] for p in box]
        text = re.sub(r"(?<=[a-z])(?=[A-Z])", " ", str(text).strip())
        if text:
            words.append(dict(text=text, x=min(xs), w=max(xs) - min(xs),
                              y=min(ys), h=max(ys) - min(ys)))
    return words


def words_tesseract(img):
    import pytesseract
    from pytesseract import Output
    g = ImageOps.autocontrast(ImageOps.grayscale(img))
    if g.width < 1600:
        g = g.resize((1600, int(g.height * 1600 / g.width)), Image.LANCZOS)
    d = pytesseract.image_to_data(g, config="--psm 6", output_type=Output.DICT)
    words = []
    for i, t in enumerate(d["text"]):
        try:
            ok = float(d["conf"][i]) >= 25
        except ValueError:
            ok = False
        if t.strip() and ok:
            words.append(dict(text=t.strip(), x=d["left"][i], w=d["width"][i],
                              y=d["top"][i], h=d["height"][i]))
    return words


def read_words(img):
    errors = []
    for engine in (words_rapidocr, words_tesseract):
        try:
            return engine(img)
        except ImportError as e:
            errors.append(str(e))
        except Exception as e:
            errors.append(f"{engine.__name__}: {e}")
    raise RuntimeError("No OCR engine available. Run: pip install rapidocr-onnxruntime  (" + "; ".join(errors) + ")")


def build_grid(words):
    if not words:
        return []
    med_h = statistics.median(w["h"] for w in words) or 20

    words = sorted(words, key=lambda w: w["y"] + w["h"] / 2)
    lines, cur, cur_y = [], [], None
    for w in words:
        yc = w["y"] + w["h"] / 2
        if cur and abs(yc - cur_y) > med_h * 0.6:
            lines.append(cur)
            cur = []
        cur.append(w)
        cur_y = statistics.mean(x["y"] + x["h"] / 2 for x in cur)
    lines.append(cur)

    rows = []
    for ws in lines:
        ws.sort(key=lambda w: w["x"])
        cells, c = [], [ws[0]]
        for a, b in zip(ws, ws[1:]):
            if b["x"] - (a["x"] + a["w"]) > med_h * 1.2:
                cells.append(c)
                c = [b]
            else:
                c.append(b)
        cells.append(c)
        rows.append([dict(text=" ".join(x["text"] for x in c), x=c[0]["x"]) for c in cells])

    xs = sorted(c["x"] for r in rows for c in r)
    tol = med_h * 2.5
    centers, grp = [], [xs[0]]
    for a, b in zip(xs, xs[1:]):
        if b - a > tol:
            centers.append(statistics.mean(grp))
            grp = [b]
        else:
            grp.append(b)
    centers.append(statistics.mean(grp))

    grid = []
    for r in rows:
        row = [""] * len(centers)
        for c in r:
            k = min(range(len(centers)), key=lambda i: abs(centers[i] - c["x"]))
            row[k] = (row[k] + " " + c["text"]).strip()
        grid.append(row)
    keep = [j for j in range(len(centers)) if any(r[j] for r in grid)]
    return [[r[j] for j in keep] for r in grid]


DATE_FORMATS = ["%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y", "%m/%d/%Y", "%d %b %Y", "%d %B %Y", "%b %d, %Y"]


def smart_cast(v):
    s = str(v).strip()
    if not s:
        return None
    if re.fullmatch(r"[-+]?\d{1,3}(,\d{3})+(\.\d+)?|[-+]?\d+(\.\d+)?", s) and not re.fullmatch(r"0\d+", s):
        n = float(s.replace(",", ""))
        return int(n) if n.is_integer() and "." not in s else n
    for f in DATE_FORMATS:
        try:
            return datetime.strptime(s, f)
        except ValueError:
            pass
    return s


def header_guess(grid):
    if len(grid) < 3 or not all(grid[0]):
        return False
    if any(isinstance(smart_cast(c), (int, float, datetime)) for c in grid[0]):
        return False
    for j in range(len(grid[0])):
        col = [smart_cast(r[j]) for r in grid[1:] if j < len(r) and str(r[j]).strip()]
        if len(col) >= 2 and all(isinstance(v, (int, float, datetime)) for v in col):
            return True
    return False


def safe_sheet_name(name, existing):
    base = re.sub(r"[\[\]:*?/\\]", "", Path(name).stem).strip()[:24] or "photo"
    n, cand = 1, base
    while cand in existing or cand == INDEX:
        n += 1
        cand = f"{base[:24]}_{n}"
    return cand


def style_header(ws):
    for c in ws[1]:
        c.font = Font(name=FONT, bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", start_color="1F3A5F")
        c.alignment = Alignment(horizontal="center")
    ws.freeze_panes = "A2"


def autosize(ws):
    for col in ws.columns:
        longest = max((len(str(c.value)) for c in col if c.value is not None), default=8)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max(longest + 3, 10), 50)


def open_book():
    if XLSX_PATH.exists():
        wb = load_workbook(XLSX_PATH)
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = INDEX
        ws.append(INDEX_COLS)
        style_header(ws)
    return wb


def fill_sheet(ws, headers, body):
    ws.append(headers)
    for r in body:
        ws.append([smart_cast(c) for c in r])
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.font = Font(name=FONT)
            if isinstance(c.value, datetime):
                c.number_format = "yyyy-mm-dd"
    style_header(ws)
    autosize(ws)


def split_header(grid, force_header=False):
    width = max((len(r) for r in grid), default=1)
    grid = [list(r) + [""] * (width - len(r)) for r in grid]
    if grid and (force_header or header_guess(grid)):
        return grid[0], grid[1:], width
    return [f"Column {i + 1}" for i in range(width)], grid, width


def store(photo_name, grid, file_name="", digest="", kind="image", force_header=False):
    with lock:
        wb = open_book()
        idx = wb[INDEX]
        entry_id = max((r[0] for r in idx.iter_rows(min_row=2, values_only=True) if isinstance(r[0], int)), default=0) + 1
        sheet = safe_sheet_name(photo_name, wb.sheetnames)
        ws = wb.create_sheet(sheet)
        headers, body, width = split_header(grid, force_header)
        fill_sheet(ws, headers, body)

        now = datetime.now().replace(microsecond=0)
        idx.append([entry_id, now, photo_name, sheet, len(body), width, file_name, digest, kind])
        idx.cell(row=idx.max_row, column=2).number_format = "yyyy-mm-dd hh:mm"
        for c in idx[idx.max_row]:
            c.font = Font(name=FONT)
        autosize(idx)
        wb.save(XLSX_PATH)
        return entry_id, sheet, headers, body


IMAGE_EXT = {".jpg", ".jpeg", ".png", ".webp", ".bmp", ".tif", ".tiff", ".gif"}


def pad(rows):
    width = max((len(r) for r in rows), default=1)
    return [list(r) + [""] * (width - len(r)) for r in rows]


def sv(v):
    if v is None:
        return ""
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, (dict, list)):
        return json.dumps(v, ensure_ascii=False)
    return str(v)


def decode_text(raw):
    if raw[:2] in (b"\xff\xfe", b"\xfe\xff"):
        return raw.decode("utf-16", "replace")
    for enc in ("utf-8-sig", "cp1252"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            pass
    return raw.decode("latin-1", "replace")


def text_to_grid(text):
    words = []
    for i, line in enumerate(text.splitlines()):
        line = line.replace("\t", "    ")
        for m in re.finditer(r"\S+(?: \S+)*", line):
            words.append(dict(text=m.group(), x=m.start() * 13, w=len(m.group()) * 13, y=i * 30, h=20))
    return build_grid(words)


def pdf_to_grid(raw):
    from pypdf import PdfReader
    reader = PdfReader(io.BytesIO(raw))
    grids = []
    for page in reader.pages[:300]:
        try:
            txt = page.extract_text(extraction_mode="layout")
        except TypeError:
            txt = page.extract_text()
        if txt and txt.strip():
            grids += text_to_grid(txt)
    if grids:
        return pad(grids)
    import pypdfium2 as pdfium
    pdf = pdfium.PdfDocument(raw)
    for i in range(min(len(pdf), 15)):
        img = pdf[i].render(scale=2).to_pil().convert("RGB")
        grids += build_grid(read_words(img))
    return pad(grids)


def docx_to_grid(raw):
    import zipfile
    import xml.etree.ElementTree as ET
    W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    with zipfile.ZipFile(io.BytesIO(raw)) as z:
        root = ET.fromstring(z.read("word/document.xml"))

    def text_of(el):
        parts = []
        for n in el.iter():
            if n.tag == W + "t" and n.text:
                parts.append(n.text)
            elif n.tag in (W + "tab", W + "br"):
                parts.append(" ")
        return "".join(parts).strip()

    rows = []
    for child in root.find(W + "body"):
        if child.tag == W + "p":
            t = text_of(child)
            if t:
                rows.append([t])
        elif child.tag == W + "tbl":
            for tr in child.iter(W + "tr"):
                rows.append([" ".join(filter(None, (text_of(p) for p in tc.findall(W + "p"))))
                             for tc in tr.findall(W + "tc")])
    return pad(rows)


def xlsx_to_grid(raw):
    wb = load_workbook(io.BytesIO(raw), data_only=True, read_only=True)
    rows = []
    for ws in wb.worksheets:
        block = [[sv(v) for v in r] for r in ws.iter_rows(values_only=True, max_row=5000)]
        block = [r for r in block if any(c.strip() for c in r)]
        if not block:
            continue
        if len(wb.worksheets) > 1:
            rows.append([f"[Sheet: {ws.title}]"])
        rows += block
    return pad(rows)


def csv_to_grid(raw):
    text = decode_text(raw)
    try:
        dialect = csv.Sniffer().sniff(text[:4096], delimiters=",;\t|")
    except csv.Error:
        dialect = csv.excel
    return pad([r for r in csv.reader(io.StringIO(text), dialect) if any(c.strip() for c in r)][:20000])


def json_to_grid(raw):
    data = json.loads(decode_text(raw))
    if isinstance(data, list) and data and all(isinstance(x, dict) for x in data):
        keys = list(dict.fromkeys(k for x in data for k in x))
        return [keys] + [[sv(x.get(k)) for k in keys] for x in data], True
    if isinstance(data, dict):
        return [["Key", "Value"]] + [[str(k), sv(v)] for k, v in data.items()], True
    return text_to_grid(json.dumps(data, indent=2, ensure_ascii=False)), False


def info_grid(name, raw, note):
    ext = Path(name).suffix.lower().lstrip(".") or "unknown"
    return [["Property", "Value"], ["File name", name], ["File type", ext],
            ["Size (KB)", round(len(raw) / 1024, 1)], ["Note", note]]


NEEDS = {"pdf": "pypdf pypdfium2"}


def extract_any(raw, name):
    ext = Path(name).suffix.lower()
    try:
        if ext in IMAGE_EXT:
            img = ImageOps.exif_transpose(Image.open(io.BytesIO(raw))).convert("RGB")
            return "image", build_grid(read_words(img)), False, ""
        if ext == ".pdf":
            return "pdf", pdf_to_grid(raw), False, ""
        if ext == ".docx":
            return "word", docx_to_grid(raw), False, ""
        if ext in (".xlsx", ".xlsm"):
            return "excel", xlsx_to_grid(raw), True, ""
        if ext in (".csv", ".tsv"):
            return "csv", csv_to_grid(raw), True, ""
        if ext == ".json":
            grid, hdr = json_to_grid(raw)
            return "json", grid, hdr, ""
        if ext in (".txt", ".md", ".log", ".text"):
            return "text", text_to_grid(decode_text(raw)), False, ""
        try:
            img = ImageOps.exif_transpose(Image.open(io.BytesIO(raw))).convert("RGB")
            return "image", build_grid(read_words(img)), False, ""
        except Exception:
            pass
    except ImportError as e:
        pkg = NEEDS.get({".pdf": "pdf", ".docx": "word"}.get(ext, ""), "")
        note = f"Reader not installed ({e.name}). Run: pip install {pkg or e.name}  - the file itself is stored safely."
        return "file", info_grid(name, raw, note), True, note
    except RuntimeError:
        raise
    except Exception as e:
        note = f"Could not read the contents ({type(e).__name__}). The file itself is stored safely."
        return "file", info_grid(name, raw, note), True, note
    return "file", info_grid(name, raw, "Stored as-is. No readable table data in this file type."), True, ""


def index_row(wb, sheet):
    for r in wb[INDEX].iter_rows(min_row=2, values_only=True):
        if r[3] == sheet:
            return r
    return None


def detect_fields(headers, rows):
    cells = [str(c) for r in ([headers] + rows) for c in r if c not in ("", None)]
    blob = "\n".join(cells)
    emails = sorted(set(re.findall(r"[\w.+-]+@[\w-]+\.[\w.-]+", blob)))
    phones = sorted({p.strip() for p in re.findall(r"(?<!\d)\+?\d[\d\s-]{8,14}\d(?!\d)", blob)
                     if 10 <= len(re.sub(r"\D", "", p)) <= 13})
    dates = sorted(set(re.findall(r"\b\d{1,2}[/.-]\d{1,2}[/.-]\d{2,4}\b|\b\d{4}-\d{2}-\d{2}\b", blob)))
    nums = []
    for r in rows:
        for c in r:
            v = c if isinstance(c, (int, float)) else smart_cast(c)
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                nums.append(v)
    low = blob.lower()
    if any(k in low for k in ("invoice", "bill", "receipt", "gst", "tax", "amount due", "total")):
        kind = "Invoice / receipt"
    elif len(headers) >= 3 and nums:
        kind = "Data table"
    elif len(headers) <= 2:
        kind = "Form / key-value / text"
    else:
        kind = "Document"
    out = dict(type=kind, emails=emails[:8], phones=phones[:8], dates=dates[:8])
    if nums:
        out["numbers"] = dict(count=len(nums), sum=round(sum(nums), 2), min=min(nums), max=max(nums))
    return out


def sheet_rows(ws):
    out = []
    for r in ws.iter_rows(values_only=True):
        out.append([(v.strftime("%Y-%m-%d") if isinstance(v, datetime) else ("" if v is None else v)) for v in r])
    return out


@app.route("/")
def home():
    return send_file(BASE_DIR / "index.html")


@app.post("/api/process")
def process():
    f = request.files.get("photo") or request.files.get("file")
    if not f:
        return jsonify(error="No file received."), 400
    raw = f.read()
    if not raw:
        return jsonify(error="The file is empty."), 400
    digest = hashlib.sha1(raw).hexdigest()
    if request.form.get("force") != "1" and XLSX_PATH.exists():
        with lock:
            wb = load_workbook(XLSX_PATH, data_only=True)
            for r in wb[INDEX].iter_rows(min_row=2, values_only=True):
                if len(r) > 7 and r[7] == digest:
                    return jsonify(duplicate=True, sheet=r[3], photo=r[2])
    name = f.filename or "upload"
    try:
        kind, grid, force_header, note = extract_any(raw, name)
    except Exception as e:
        return jsonify(error=str(e)), 500
    ext = re.sub(r"[^A-Za-z0-9.]", "", Path(name).suffix)[:10]
    file_name = f"{datetime.now().strftime('%Y%m%d_%H%M%S_%f')}_{re.sub(r'[^A-Za-z0-9._-]', '_', Path(name).stem)[:60]}{ext}"
    (FILE_DIR / file_name).write_bytes(raw)
    if not grid:
        grid = [["(no readable text found in this file)"]]
        force_header = False
    entry_id, sheet, headers, body = store(name, grid, file_name, digest, kind, force_header)
    return jsonify(entry_id=entry_id, sheet=sheet, headers=headers, rows=body, photo=name,
                   kind=kind, note=note, fields=detect_fields(headers, body))


@app.get("/api/entries")
def entries():
    if not XLSX_PATH.exists():
        return jsonify(entries=[])
    with lock:
        ws = load_workbook(XLSX_PATH, data_only=True)[INDEX]
        rows = sheet_rows(ws)
    return jsonify(entries=[dict(zip(["id", "date", "photo", "sheet", "rows", "cols", "file", "hash", "type"], r)) for r in rows[1:]])


@app.get("/api/entries/<sheet>")
def entry_data(sheet):
    with lock:
        wb = load_workbook(XLSX_PATH, data_only=True)
        if sheet not in wb.sheetnames:
            return jsonify(error="Not found"), 404
        rows = sheet_rows(wb[sheet])
    return jsonify(headers=rows[0], rows=rows[1:], fields=detect_fields(rows[0], rows[1:]))


@app.put("/api/entries/<sheet>")
def entry_update(sheet):
    b = request.get_json(force=True)
    headers = [str(h).strip() or f"Column {i + 1}" for i, h in enumerate(b.get("headers", []))]
    rows = b.get("rows", [])
    with lock:
        wb = load_workbook(XLSX_PATH)
        if sheet not in wb.sheetnames or sheet == INDEX:
            return jsonify(error="Not found"), 404
        pos = wb.sheetnames.index(sheet)
        del wb[sheet]
        fill_sheet(wb.create_sheet(sheet, pos), headers, rows)
        idx = wb[INDEX]
        for r in range(2, idx.max_row + 1):
            if idx.cell(row=r, column=4).value == sheet:
                idx.cell(row=r, column=5, value=len(rows))
                idx.cell(row=r, column=6, value=len(headers))
        wb.save(XLSX_PATH)
    return jsonify(saved=len(rows))


@app.get("/api/entries/<sheet>/export")
def entry_export(sheet):
    fmt = request.args.get("fmt", "csv")
    with lock:
        wb = load_workbook(XLSX_PATH, data_only=True)
        if sheet not in wb.sheetnames:
            return jsonify(error="Not found"), 404
        rows = sheet_rows(wb[sheet])
    if fmt == "json":
        data = [dict(zip(rows[0], r)) for r in rows[1:]]
        return Response(json.dumps(data, indent=2, ensure_ascii=False), mimetype="application/json",
                        headers={"Content-Disposition": f'attachment; filename="{sheet}.json"'})
    buf = io.StringIO()
    csv.writer(buf).writerows(rows)
    return Response("\ufeff" + buf.getvalue(), mimetype="text/csv",
                    headers={"Content-Disposition": f'attachment; filename="{sheet}.csv"'})


def stored_path(sheet):
    with lock:
        row = index_row(load_workbook(XLSX_PATH, data_only=True), sheet) if XLSX_PATH.exists() else None
    if not row or len(row) < 7 or not row[6]:
        return None, None
    for d in (FILE_DIR, PHOTO_DIR):
        if (d / row[6]).exists():
            return d / row[6], row
    return None, row


@app.get("/api/file/<sheet>")
def entry_file(sheet):
    path, row = stored_path(sheet)
    if not path:
        return jsonify(error="Original file not found"), 404
    return send_file(path, as_attachment=request.args.get("dl") == "1", download_name=row[2])


@app.get("/api/photo/<sheet>")
def entry_photo(sheet):
    path, row = stored_path(sheet)
    if not path or path.suffix.lower() not in IMAGE_EXT | {".jpg"}:
        return jsonify(error="No image"), 404
    return send_file(path)


@app.get("/api/search")
def search():
    q = request.args.get("q", "").strip().lower()
    if not q or not XLSX_PATH.exists():
        return jsonify(results=[])
    with lock:
        wb = load_workbook(XLSX_PATH, data_only=True)
        out = []
        for name in wb.sheetnames:
            if name == INDEX:
                continue
            rows = sheet_rows(wb[name])
            for i, r in enumerate(rows[1:], 1):
                if any(q in str(c).lower() for c in r):
                    out.append(dict(sheet=name, row=i, headers=rows[0], cells=r))
                    if len(out) >= 300:
                        return jsonify(results=out)
    return jsonify(results=out)


@app.get("/api/stats")
def stats():
    if not XLSX_PATH.exists():
        return jsonify(photos=0, rows=0, last="-", size_kb=0)
    with lock:
        ws = load_workbook(XLSX_PATH, data_only=True)[INDEX]
        rows = sheet_rows(ws)[1:]
    return jsonify(photos=len(rows), rows=sum(r[4] for r in rows if isinstance(r[4], int)),
                   last=rows[-1][1] if rows else "-", size_kb=round(XLSX_PATH.stat().st_size / 1024, 1))


@app.delete("/api/entries/<sheet>")
def entry_delete(sheet):
    with lock:
        wb = load_workbook(XLSX_PATH)
        if sheet not in wb.sheetnames or sheet == INDEX:
            return jsonify(error="Not found"), 404
        del wb[sheet]
        idx = wb[INDEX]
        for r in range(2, idx.max_row + 1):
            if idx.cell(row=r, column=4).value == sheet:
                stored = idx.cell(row=r, column=7).value
                idx.delete_rows(r)
                for d in (FILE_DIR, PHOTO_DIR):
                    if stored and (d / stored).exists():
                        (d / stored).unlink()
                break
        wb.save(XLSX_PATH)
    return jsonify(deleted=sheet)


def ocr_status():
    try:
        import rapidocr_onnxruntime
        return dict(ok=True, engine="rapidocr")
    except Exception as e1:
        try:
            import pytesseract
            pytesseract.get_tesseract_version()
            return dict(ok=True, engine="tesseract")
        except Exception:
            return dict(ok=False, error="Text reader not installed. In a terminal run:  pip install rapidocr-onnxruntime",
                        detail=str(e1))


@app.get("/api/health")
def health():
    return jsonify(server=True, version=3, ocr=ocr_status())


@app.get("/api/download")
def download():
    if not XLSX_PATH.exists():
        return jsonify(error="Nothing stored yet."), 404
    return send_file(XLSX_PATH, as_attachment=True, download_name="entries.xlsx")


def selftest(url):
    import urllib.request
    try:
        with urllib.request.urlopen(url + "/api/health", timeout=3) as r:
            print("  Self-test: OK - the server answers at", url, "\n")
    except Exception as exc:
        print("  Self-test FAILED:", exc)
        print("  Something on this computer is blocking the connection (firewall, antivirus or proxy).\n")


def port_in_use(port):
    with socket.socket() as probe:
        probe.settimeout(0.4)
        if probe.connect_ex(("127.0.0.1", port)) == 0:
            return True
    with socket.socket() as probe:
        try:
            probe.bind(("0.0.0.0", port))
            return False
        except OSError:
            return True


def lan_ip():
    sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        sock.connect(("10.255.255.255", 1))
        return sock.getsockname()[0]
    except OSError:
        return None
    finally:
        sock.close()


def show_phone_access(port):
    ip = lan_ip()
    if not ip:
        print("  Phone access: could not find this computer's Wi-Fi address.\n")
        return
    link = f"http://{ip}:{port}/?pin={PIN}"
    print("  PHONE ACCESS (phone and computer must be on the same Wi-Fi)")
    print(f"  Open on your phone:  {link}")
    print(f"  PIN: {PIN}   (anyone on your Wi-Fi with this PIN can use the app)")
    try:
        import qrcode
        qr = qrcode.QRCode(border=1)
        qr.add_data(link)
        qr.print_ascii(invert=True)
    except ImportError:
        print("  Tip: run  pip install qrcode  to also show a QR code to scan with the phone camera.")
    print()


def free_port(start=5000, tries=30):
    for port in range(start, start + tries):
        if not port_in_use(port):
            return port
    raise RuntimeError("No free port between 5000 and 5029.")


if __name__ == "__main__":
    try:
        phone = "--phone" in sys.argv
        if phone:
            custom = next((a.split("=", 1)[1] for a in sys.argv if a.startswith("--pin=")), None)
            PIN = custom or str(secrets.randbelow(900000) + 100000)
        port = free_port()
        url = f"http://127.0.0.1:{port}"
        st = ocr_status()
        print("\n  Smart Data Entry is running.")
        print(f"  Open in your browser:  {url}   (opening it for you now)")
        if port != 5000:
            print("  (port 5000 was busy - maybe an old copy is still running; to close old copies in")
            print("   PowerShell run:  taskkill /F /IM python.exe  )")
        print("  Keep this window open while you use the app. Press Ctrl+C to stop.")
        print("  Text reader:", st.get("engine") or "NOT INSTALLED -> " + st["error"], "\n")
        if phone:
            show_phone_access(port)
        else:
            print("  To use it from your phone too, stop this and run:  python app.py --phone\n")
        import webbrowser
        threading.Timer(1.2, lambda: webbrowser.open(url)).start()
        threading.Timer(2.0, lambda: selftest(url)).start()
        app.run(host="0.0.0.0" if phone else "127.0.0.1", port=port, debug=False)
    except Exception as exc:
        print("\n  ERROR:", exc)
        input("\n  Press Enter to close...")

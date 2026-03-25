"""
Masco Canada — Cycle Count Web App
Run: python app.py
Then open: http://localhost:5000  (Don's computer)
Counters on iPads: http://<Don's IP>:5000
"""

from flask import Flask, render_template, request, jsonify, send_file
import csv, io, json, os, openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max

# In-memory session store (keyed by batch number)
SESSIONS = {}

# ── CSV Parser ─────────────────────────────────────────────────────────────
def parse_jde_csv(file_bytes):
    text = file_bytes.decode("utf-8-sig", errors="replace")
    reader = csv.reader(io.StringIO(text))
    all_rows = list(reader)

    cycle_num, description = "UNKNOWN", "UNKNOWN"
    for r in all_rows[:7]:
        if r and r[0].strip() == "Cycle Count Number" and len(r) > 2:
            cycle_num = r[2].strip()
        if r and r[0].strip() == "Description" and len(r) > 2:
            description = r[2].strip()

    items = []
    for r in all_rows[7:]:
        if not r or not r[0].strip():
            continue
        try:
            qoh_raw = r[3].strip() if len(r) > 3 else "0"
            qoh = int(float(qoh_raw)) if qoh_raw else 0
        except:
            qoh = 0
        items.append({
            "item":     r[0].strip(),
            "qoh":      qoh,
            "desc":     r[7].strip() if len(r) > 7 else "",
            "location": r[14].strip() if len(r) > 14 else "",
            "abc":      r[10].strip() if len(r) > 10 else "",
            "branch":   r[12].strip().strip() if len(r) > 12 else "MC08",
        })
    return cycle_num, description, items

# ── Excel export for Don ───────────────────────────────────────────────────
def build_export(session):
    items   = session["items"]
    counts  = session.get("counts", {})
    batch   = session["batch"]
    desc    = session["description"]

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "JDE_PASTE"

    NAVY  = "1B3A6B"; TEAL = "1E7B74"; WHITE = "FFFFFF"
    GREEN = "1A7A3C"; LTGREEN = "E8F5EE"; GRAY = "4A4A4A"
    LTGRAY = "F2F2F2"; MID = "CCCCCC"

    def s(side_color, weight="thin"):
        x = Side(style=weight, color=side_color)
        return Border(left=x, right=x, top=x, bottom=x)

    # Title
    ws.merge_cells("A1:C1")
    t = ws["A1"]
    t.value = f"Batch {batch} — {desc} — Copy col A → Paste into JDE Quantity"
    t.font = Font(bold=True, color=WHITE, size=11, name="Calibri")
    t.fill = PatternFill("solid", fgColor=TEAL)
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 24

    # Headers
    for c, (h, w) in enumerate(zip(["QUANTITY", "Item Number", "Description"], [14, 20, 40]), 1):
        cell = ws.cell(2, c, h)
        cell.font = Font(bold=True, color=WHITE, size=10, name="Calibri")
        cell.fill = PatternFill("solid", fgColor=NAVY)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = s(MID)
        ws.column_dimensions[get_column_letter(c)].width = w

    # Data
    for i, item in enumerate(items, 3):
        key = str(i - 3)
        c_data = counts.get(key, {})
        upb = c_data.get("upb", 0)
        boxes = c_data.get("boxes", 0)
        total = upb * boxes if upb and boxes else 0

        bg = LTGRAY if i % 2 == 0 else WHITE

        qty = ws.cell(i, 1, total)
        qty.font = Font(bold=True, size=12, name="Calibri",
                        color=GREEN if total > 0 else GRAY)
        qty.fill = PatternFill("solid", fgColor=LTGREEN if total > 0 else bg)
        qty.alignment = Alignment(horizontal="center", vertical="center")
        qty.border = s(TEAL if total > 0 else MID, "medium" if total > 0 else "thin")

        for c, val in enumerate([item["item"], item["desc"]], 2):
            cell = ws.cell(i, c, val)
            cell.font = Font(size=10, name="Calibri",
                             bold=(c == 2), color=NAVY if c == 2 else GRAY)
            cell.fill = PatternFill("solid", fgColor=bg)
            cell.alignment = Alignment(horizontal="left" if c > 1 else "center",
                                       vertical="center")
            cell.border = s(MID)

    ws.freeze_panes = "A3"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ── Routes ─────────────────────────────────────────────────────────────────
@app.route("/")
def index():
    return render_template("index.html")

@app.route("/upload", methods=["POST"])
def upload():
    f = request.files.get("file")
    if not f:
        return jsonify({"error": "No file"}), 400
    try:
        batch, description, items = parse_jde_csv(f.read())
        SESSIONS[batch] = {"batch": batch, "description": description,
                           "items": items, "counts": {}}
        return jsonify({"batch": batch, "description": description,
                        "count": len(items), "items": items})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/session/<batch>")
def get_session(batch):
    s = SESSIONS.get(batch)
    if not s:
        return jsonify({"error": "Session not found"}), 404
    return jsonify(s)

@app.route("/save_count", methods=["POST"])
def save_count():
    data = request.json
    batch = data.get("batch")
    idx   = str(data.get("idx"))
    upb   = int(data.get("upb", 0))
    boxes = int(data.get("boxes", 0))
    if batch not in SESSIONS:
        return jsonify({"error": "Session not found"}), 404
    SESSIONS[batch]["counts"][idx] = {"upb": upb, "boxes": boxes}
    total = upb * boxes
    done  = sum(1 for v in SESSIONS[batch]["counts"].values()
                if v.get("upb") and v.get("boxes"))
    return jsonify({"total": total, "done": done,
                    "total_items": len(SESSIONS[batch]["items"])})

@app.route("/export/<batch>")
def export(batch):
    s = SESSIONS.get(batch)
    if not s:
        return "Session not found", 404
    buf = build_export(s)
    return send_file(buf, as_attachment=True,
                     download_name=f"CycleCount_{batch}_JDE_Upload.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.route("/summary/<batch>")
def summary(batch):
    s = SESSIONS.get(batch)
    if not s:
        return jsonify({"error": "Not found"}), 404
    items  = s["items"]
    counts = s["counts"]
    result = []
    for i, item in enumerate(items):
        c = counts.get(str(i), {})
        upb = c.get("upb", 0)
        boxes = c.get("boxes", 0)
        result.append({"item": item["item"], "desc": item["desc"],
                       "location": item["location"], "upb": upb,
                       "boxes": boxes, "total": upb * boxes,
                       "counted": bool(upb and boxes)})
    done = sum(1 for r in result if r["counted"])
    return jsonify({"items": result, "done": done, "total": len(result)})

if __name__ == "__main__":
    import socket
    hostname = socket.gethostname()
    try:
        local_ip = socket.gethostbyname(hostname)
    except:
        local_ip = "your-computer-IP"
    print("\n" + "="*55)
    print("  MASCO CANADA — Cycle Count App")
    print("="*55)
    print(f"  Don's computer : http://localhost:5000")
    print(f"  Counters (iPad): http://{local_ip}:5000")
    print("="*55 + "\n")
    app.run(host="0.0.0.0", port=5000, debug=False)

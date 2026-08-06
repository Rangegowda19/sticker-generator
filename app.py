#!/usr/bin/env python3
"""
Sticker Generator Web App
=========================
Run:    python app.py
Open:   http://localhost:5000

Steps in the app:
  1. Select sticker type  : Solid  (Ratio coming soon)
  2. Upload the PO Excel file
  3. Select the phase     : Phase 1 / Phase 2  (4 lines)
                            Phase 3            (adds top line, e.g. 8411/SG)
  4. Click Generate -> the sticker PDF downloads automatically
"""

import os
import tempfile
from flask import Flask, request, send_file, render_template_string

import sticker_generator as sg

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 30 * 1024 * 1024  # 30 MB upload limit

PAGE = """
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Sticker Generator</title>
<style>
  body { font-family: Arial, sans-serif; background:#f4f5f7; margin:0;
         display:flex; justify-content:center; align-items:flex-start;
         min-height:100vh; padding-top:60px; }
  .card { background:#fff; border-radius:12px; padding:32px 36px;
          box-shadow:0 2px 12px rgba(0,0,0,.08); width:420px; }
  h1 { font-size:20px; margin:0 0 6px; }
  p.sub { color:#666; font-size:13px; margin:0 0 24px; }
  label { display:block; font-weight:bold; font-size:14px; margin:18px 0 6px; }
  select, input[type=file] { width:100%; padding:9px; font-size:14px;
          border:1px solid #ccc; border-radius:8px; background:#fff; }
  button { margin-top:26px; width:100%; padding:12px; font-size:15px;
           font-weight:bold; color:#fff; background:#1a73e8; border:none;
           border-radius:8px; cursor:pointer; }
  button:hover { background:#155ec1; }
  .msg { margin-top:16px; padding:10px 12px; border-radius:8px; font-size:13px; }
  .err { background:#fdecea; color:#b3261e; }
  .note { background:#eef3fd; color:#333; }
  #ratioNote { display:none; }
</style>
</head>
<body>
<div class="card">
  <h1>Carton Sticker Generator</h1>
  <p class="sub">Upload the company PO Excel and download the print-ready PDF.</p>

  {% if error %}<div class="msg err">{{ error }}</div>{% endif %}

  <form method="post" action="/generate" enctype="multipart/form-data">
    <label>1. Sticker type</label>
    <select name="stype" id="stype" onchange="onType()">
      <option value="solid">Solid</option>
      <option value="ratio">Ratio (coming soon)</option>
    </select>
    <div class="msg note" id="ratioNote">
      Ratio stickers are not ready yet. Please select Solid for now.
    </div>

    <label>2. PO Excel file</label>
    <input type="file" name="pofile" accept=".xlsx,.xlsm" required>

    <label>3. Phase</label>
    <select name="phase">
      <option value="1">Phase 1 &mdash; 4 lines</option>
      <option value="2">Phase 2 &mdash; 4 lines</option>
      <option value="3">Phase 3 &mdash; 5 lines (adds top line, e.g. 8411/SG)</option>
    </select>

    <button type="submit">Generate Sticker PDF</button>
  </form>
</div>
<script>
function onType() {
  var t = document.getElementById('stype').value;
  document.getElementById('ratioNote').style.display =
      (t === 'ratio') ? 'block' : 'none';
}
</script>
</body>
</html>
"""

@app.route("/")
def index():
    return render_template_string(PAGE, error=None)

@app.route("/generate", methods=["POST"])
def generate():
    stype = request.form.get("stype", "solid")
    phase = request.form.get("phase", "1")
    f = request.files.get("pofile")

    if stype == "ratio":
        return render_template_string(
            PAGE, error="Ratio stickers are coming soon. Please select Solid.")
    if not f or f.filename == "":
        return render_template_string(PAGE, error="Please choose an Excel file.")
    if not f.filename.lower().endswith((".xlsx", ".xlsm")):
        return render_template_string(
            PAGE, error="Please upload an Excel file (.xlsx).")

    tmpdir = tempfile.mkdtemp()
    in_path = os.path.join(tmpdir, f.filename)
    f.save(in_path)

    try:
        fmt, designs = sg.parse_po(in_path)
        if not designs:
            return render_template_string(
                PAGE, error="Could not read any sticker data from this file. "
                            "Please check it is the correct PO Excel.")
        po_label = next((d["po_no"] for d in designs if d.get("po_no")),
                        os.path.splitext(f.filename)[0])
        base = os.path.splitext(f.filename)[0]
        out_name = f"{base}_STICKERS_Phase{phase}.pdf"
        out_path = os.path.join(tmpdir, out_name)
        # circle letter (C/D/B) is auto-detected from each sheet's image
        sg.generate_pdf(designs, out_path, po_label,
                        show_dest=(phase == "3"))
    except Exception as e:
        return render_template_string(
            PAGE, error=f"Something went wrong while processing the file: {e}")

    return send_file(out_path, as_attachment=True, download_name=out_name)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    print()
    print("  Sticker Generator running.")
    print(f"  Open this address in your browser:  http://localhost:{port}")
    print("  (Press CTRL+C in this window to stop)")
    print()
    # host 0.0.0.0 so hosting services (Render etc.) can reach the app;
    # on your own PC it still works at localhost as before
    app.run(host="0.0.0.0", port=port, debug=False)
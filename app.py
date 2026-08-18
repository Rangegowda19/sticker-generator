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
import glob
import shutil
import platform
import subprocess
import tempfile
import threading
from flask import Flask, request, send_file, render_template_string

import sticker_generator as sg
import export_cdr_data as cdrexport

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 30 * 1024 * 1024  # 30 MB upload limit

# When True (and running on the same PC as CorelDraw), the generated PDF is
# opened in CorelDraw automatically after each generate. Set False to just
# download without launching Corel.
OPEN_IN_CORELDRAW = True

# Folder where generated PDFs are saved so CorelDraw can open them.
# Defaults to a "sticker_output" folder on the Desktop.
OUTPUT_DIR = os.path.join(os.path.expanduser("~"), "Desktop", "sticker_output")

def _find_coreldraw():
    """Locate the CorelDraw executable on Windows. Returns path or None."""
    if platform.system() != "Windows":
        return None
    # confirmed install path on this machine (CorelDRAW X8)
    known = r"C:\Program Files (x86)\Corel\CorelDRAW Graphics Suite X8\Programs\CorelDRW.exe"
    if os.path.exists(known):
        return known
    # common install locations across versions
    patterns = [
        # 64-bit and 32-bit program folders, all versions incl. X8
        r"C:\Program Files\Corel\CorelDRAW Graphics Suite *\Programs64\CorelDRW.exe",
        r"C:\Program Files\Corel\CorelDRAW Graphics Suite *\Programs\CorelDRW.exe",
        r"C:\Program Files\Corel\CorelDRAW Graphics Suite *\Draw\CorelDRW.exe",
        r"C:\Program Files (x86)\Corel\CorelDRAW Graphics Suite *\Programs64\CorelDRW.exe",
        r"C:\Program Files (x86)\Corel\CorelDRAW Graphics Suite *\Programs\CorelDRW.exe",
        r"C:\Program Files (x86)\Corel\CorelDRAW Graphics Suite *\Draw\CorelDRW.exe",
    ]
    found = []
    for p in patterns:
        found += glob.glob(p)
    if found:
        return sorted(found)[-1]           # newest version
    # last resort: walk the Corel folders for any CorelDRW.exe
    for base in (r"C:\Program Files\Corel", r"C:\Program Files (x86)\Corel"):
        if os.path.isdir(base):
            for root, _dirs, files in os.walk(base):
                for fn in files:
                    if fn.lower() == "coreldrw.exe":
                        return os.path.join(root, fn)
    return None

# If True, open the file in the CorelDRAW that is ALREADY running (no fresh
# instance, no sign-in screen). If no Corel is open, it will start one.
REUSE_OPEN_CORELDRAW = True

def _coreldraw_is_running():
    """True if a CorelDRW.exe process is already running (Windows)."""
    try:
        out = subprocess.run(["tasklist", "/FI", "IMAGENAME eq CorelDRW.exe"],
                             capture_output=True, text=True).stdout
        return "CorelDRW.exe" in out
    except Exception:
        return False

def _open_in_coreldraw(pdf_path):
    """Open THIS pdf in CorelDRAW by calling its exe directly.
    Does NOT change Windows defaults - other PDFs still open in Chrome,
    and there is no 'open with' prompt."""
    exe = _find_coreldraw()
    abspdf = os.path.abspath(pdf_path)
    if exe:
        try:
            print("  opening in CorelDRAW:", exe)
            subprocess.Popen([exe, abspdf], close_fds=True)
            return True
        except Exception as e:
            print("  launch failed:", repr(e))
    else:
        print("  CorelDRAW exe not found")
    return False

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

    <label>4. Output format</label>
    <select name="fmt">
      <option value="cdr">CorelDRAW data (for the StickerImport macro &rarr; native .cdr)</option>
      <option value="pdf">PDF</option>
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
    outfmt = request.form.get("fmt", "svg")
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

        if outfmt == "cdr":
            # ---- CorelDRAW data file for the VBA macro ----
            data_name = f"{base}_Phase{phase}_CDRDATA.txt"
            data_path = os.path.join(tmpdir, data_name)
            pages = cdrexport.export(designs, data_path,
                                     sg.sheets_for_design, sg.copies_for_height,
                                     show_dest=(phase == "3"))
            try:
                os.makedirs(OUTPUT_DIR, exist_ok=True)
                saved = os.path.join(OUTPUT_DIR, data_name)
                shutil.copy(data_path, saved)
                print("=" * 50)
                print(f"Wrote CorelDRAW data ({pages} pages) to: {saved}")
                print("In CorelDRAW: run the StickerImport macro and pick this file.")
                print("=" * 50)
                try:
                    subprocess.Popen(["explorer", OUTPUT_DIR])
                except Exception:
                    pass
            except Exception as e:
                print("save data ERROR:", repr(e))
            return send_file(data_path, as_attachment=True, download_name=data_name)

        # ---- PDF output ----
        out_name = f"{base}_STICKERS_Phase{phase}.pdf"
        out_path = os.path.join(tmpdir, out_name)
        sg.generate_pdf(designs, out_path, po_label,
                        show_dest=(phase == "3"))
        saved_path = out_path
        try:
            os.makedirs(OUTPUT_DIR, exist_ok=True)
            saved_path = os.path.join(OUTPUT_DIR, out_name)
            shutil.copy(out_path, saved_path)
            print("Saved PDF to:", saved_path)
            if OPEN_IN_CORELDRAW:
                threading.Thread(target=_open_in_coreldraw,
                                 args=(saved_path,), daemon=True).start()
        except Exception as e:
            print("Save step ERROR:", repr(e))
    except Exception as e:
        return render_template_string(
            PAGE, error=f"Something went wrong while processing the file: {e}")

    return send_file(saved_path, as_attachment=True, download_name=out_name)

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
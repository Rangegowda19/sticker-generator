#!/usr/bin/env python3
"""
Carton Sticker Generator
========================
Reads a company PO Excel file (Ratio or Solid format), extracts sticker data
from every sheet, and generates a print-ready PDF where each sticker design is
repeated N times per page based on carton Height (H):

    H <= 18  -> 8 copies per page
    H <= 26  -> 4 copies per page
    H <= 30  -> 2 copies per page
    otherwise -> 1 copy per page

Also writes a summary Excel listing every design, its H, copies, and qty.

Usage:
    python sticker_generator.py <po_file.xlsx> [more_po_files.xlsx ...]

Edit COPY_RULES below if the height thresholds change.
"""

import sys, re, os
import openpyxl
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas

# ----------------------------------------------------------------------
# CONFIG — edit these if company rules change
# ----------------------------------------------------------------------
# (max_height_cm, copies). Checked top-down; first match wins.
COPY_RULES = [
    (18, 8),
    (26, 4),
    (30, 2),
    (9999, 1),
]

# Exact size of ONE sticker for each copy count: copies -> (width_in, height_in)
STICKER_SIZE_IN = {
    8: (4.0, 2.5),
    4: (4.1, 5.0),
    2: (5.0, 5.0),
    1: (6.0, 6.0),
}

# Letter shown inside the circle on the volume line (changes per PO: D, C, ...)
CIRCLE_LETTER = "D"

# Horizontal width of the text as a percentage. 100 = normal Arial width.
# 82 makes the text narrower (same height) to match the original stickers,
# e.g. a line that printed 11 cm wide becomes 9 cm wide.
TEXT_HSCALE = 82

# Exact text heights in INCHES for the 2-copy (5" x 5") sticker.
# Change these numbers to resize each line.
# Exact text WIDTHS in INCHES for the 4-copy sticker. Each line is stretched
# or compressed horizontally to exactly this width (heights come from
# FOUR_COPY_LINE_H). "vol" is the width of the "(0.037M3)" part only —
# the rest of that line scales along with it.
FOUR_COPY_LINE_W = {
    "dest": 1.210,
    "item": 2.933,
    "code": 1.700,
    "vol":  1.512,
}

# Exact text heights in INCHES for the 4-copy (4.1" x 5") sticker.
FOUR_COPY_LINE_H = {
    "dest": 0.470,   # 1st line: 8465/EU (printed only in Phase 3)
    "item": 0.523,   # 2nd line: 341-489933(71-06)
    "code": 0.410,   # 3rd line: 09-004-000
    "vol":  0.525,   # 4th line: (0.059M3) D 16 PCS  (right-aligned)
    "desc": 0.650,   # 5th line: description — ends at the bracket of the
                     #           4th line, never goes before it
}

TWO_COPY_LINE_H = {
    "dest": 0.530,   # extra top line (e.g. 8411/SG) — printed only in Phase 3
    "item": 0.500,   # 1st line: 341-489933(71-06)
    "code": 0.450,   # 2nd line: 09-004-000
    "vol":  0.700,   # 3rd line: (0.059M3) D 16 PCS  (right-aligned)
    "desc": 0.700,   # 4th line: description, same height as 3rd line
                     #           (shrinks automatically if too wide)
}

# ---- Font: use real Arial Bold if available (Windows), else Helvetica-Bold ----
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

FONT = "Helvetica-Bold"
for _p in (r"C:\Windows\Fonts\arialbd.ttf",
           r"C:\Windows\Fonts\ARIALBD.TTF",
           "/usr/share/fonts/truetype/msttcorefonts/Arial_Bold.ttf",
           "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf"):
    if os.path.exists(_p):
        try:
            pdfmetrics.registerFont(TTFont("Arial-Bold", _p))
            FONT = "Arial-Bold"
            break
        except Exception:
            pass

PAGE_SIZE = A4  # print sheet size

def copies_for_height(h):
    for max_h, copies in COPY_RULES:
        if h <= max_h:
            return copies
    return 1

# ----------------------------------------------------------------------
# PARSING
# ----------------------------------------------------------------------
MEAS_RE = re.compile(r"L\s*(\d+(?:\.\d+)?)\s*X\s*W\s*(\d+(?:\.\d+)?)\s*X\s*H\s*(\d+(?:\.\d+)?)", re.I)
ITEM_RE = re.compile(r"^\d{3}-\d{6}\(\d+-\d+\)$")          # 341-489933(71-06)
VOL_RE  = re.compile(r"^\(([\d.]+)\s*M3\)$", re.I)          # (0.059M3)
PCS_RE  = re.compile(r"^(\d+)\s*PCS$", re.I)                # 16PCS
QTY_RE  = re.compile(r"Qty\s*Required\s*is\s*:?\s*(\d+)\s*box", re.I)
SIZE_NAMES = {"XS","S","M","L","XL","XXL","2XL","3XL","4XL","5XL","6XL","XXXL"}

def iter_cells(ws):
    for row in ws.iter_rows():
        for c in row:
            if c.value is not None and str(c.value).strip() != "":
                yield c

def cell_map(ws):
    """{(row,col): stripped string value}"""
    return {(c.row, c.column): str(c.value).strip() for c in iter_cells(ws)}

def find_first(cm, pred):
    for (r, col), v in sorted(cm.items()):
        if pred(v):
            return (r, col), v
    return None, None

def parse_measurement(cm):
    for _, v in cm.items():
        m = MEAS_RE.search(v)
        if m:
            return float(m.group(1)), float(m.group(2)), float(m.group(3))
    return None, None, None

def parse_common(cm):
    d = {}
    _, d["item_no"] = find_first(cm, lambda v: ITEM_RE.match(v))
    pos, pcs = find_first(cm, lambda v: PCS_RE.match(v))
    d["pcs"] = int(PCS_RE.match(pcs).group(1)) if pcs else None
    _, vol = find_first(cm, lambda v: VOL_RE.match(v))
    d["volume"] = vol
    d["L"], d["W"], d["H"] = parse_measurement(cm)
    _, po = find_first(cm, lambda v: v.upper().startswith(":P0"))
    d["po_no"] = po.lstrip(":") if po else None
    _, dest = find_first(cm, lambda v: re.match(r"^\d{3,5}/[A-Z]{2,3}$", v))
    d["dest"] = dest
    qty = None
    for _, v in cm.items():
        m = QTY_RE.search(v)
        if m:
            qty = int(m.group(1))
    d["qty_boxes"] = qty
    # description: longest text line that isn't a label/known field
    labels = ("width side","length side","margin","rfid","po no","gross","net",
              "measurement","country","qty required","carton qty","colors","pattern","colour")
    best = ""
    for _, v in cm.items():
        lv = v.lower()
        if any(k in lv for k in labels):
            continue
        if ITEM_RE.match(v) or PCS_RE.match(v) or VOL_RE.match(v) or v.startswith(":"):
            continue
        if len(v) > len(best) and re.search(r"[a-zA-Z]{4,}", v) and not re.match(r"^\d", v):
            best = v
    d["description"] = best or None
    return d

def parse_ratio_sheet(ws):
    """One sheet = one assortment carton design."""
    cm = cell_map(ws)
    d = parse_common(cm)
    d["type"] = "ratio"
    # assort code: cell directly below item_no (e.g. I017 / T008 / X053 / E031)
    pos, _ = find_first(cm, lambda v: ITEM_RE.match(v))
    d["code"] = None
    if pos:
        r, c = pos
        for dr in (1, 2):
            v = cm.get((r + dr, c))
            if v and re.match(r"^[A-Z]\d{3}$", v):
                d["code"] = v
                break
    if not d["code"]:
        d["code"] = ws.title.split()[0]
    # size table: rows where col X = size name, col X+1.. = numbers; "Total" row too
    table_rows = []
    first_size_pos = None
    for (r, c), v in sorted(cm.items()):
        if v.upper() in SIZE_NAMES or v.lower() == "total":
            nums = []
            cc = c + 1
            while (r, cc) in cm and re.match(r"^\d+$", cm[(r, cc)]):
                nums.append(int(cm[(r, cc)]))
                cc += 1
            if nums:
                if first_size_pos is None:
                    first_size_pos = (r, c)
                table_rows.append((v if v.lower() != "total" else "TOTAL", nums))
    # color header codes (e.g. 41, 50) sitting 1-4 rows above the first size row
    colors = []
    if first_size_pos:
        r0, c0 = first_size_pos
        for dr in range(1, 5):
            row_codes = []
            cc = c0 + 1
            while (r0 - dr, cc) in cm and re.match(r"^\d{2}$", cm[(r0 - dr, cc)]):
                row_codes.append(cm[(r0 - dr, cc)])
                cc += 1
            if row_codes:
                colors = row_codes
                break
    if colors:
        table_rows.insert(0, ("", [int(x) for x in colors]))
    d["size_table"] = table_rows
    return [d]

def parse_solid_sheet(ws):
    """One sheet -> many designs (one per size x color combo with qty > 0)."""
    cm = cell_map(ws)
    base = parse_common(cm)
    # colors from legend like '09-BLACK'
    colors = {}
    for _, v in cm.items():
        m = re.match(r"^(\d{2})[-:]\s*([A-Z ]+)$", v)
        if m:
            colors[m.group(1)] = m.group(2).strip()
    # sizes like 'XS-002' / 'XS:002'
    sizes = {}
    for _, v in cm.items():
        m = re.match(r"^([A-Z0-9]{1,4})[-:](\d{3})$", v)
        if m and m.group(1).upper() in SIZE_NAMES:
            sizes[m.group(1).upper()] = m.group(2)
    # qty grid: header row containing color codes, size rows below
    hdr = None
    for (r, c), v in sorted(cm.items()):
        if re.match(r"^\d{2}:", v):  # e.g. 09:BLACK header cell
            hdr = r
            break
    grid = {}  # (size, color_code) -> qty
    if hdr:
        col_color = {}
        for (r, c), v in cm.items():
            if r == hdr:
                m = re.match(r"^(\d{2}):", v)
                if m:
                    col_color[c] = m.group(1)
        for (r, c), v in sorted(cm.items()):
            if r <= hdr:
                continue
            m = re.match(r"^([A-Z0-9]{1,4}):(\d{3})$", v)
            if m and m.group(1).upper() in SIZE_NAMES:
                size = m.group(1).upper()
                for cc, ccode in col_color.items():
                    q = cm.get((r, cc))
                    if q and re.match(r"^\d+$", q):
                        grid[(size, ccode)] = int(q)
    designs = []
    for (size, ccode), qty in sorted(grid.items()):
        d = dict(base)
        d["type"] = "solid"
        d["size"] = size
        d["color_code"] = ccode
        d["color_name"] = colors.get(ccode, "")
        d["code"] = f"{ccode}-{sizes.get(size,'???')}-000"
        d["qty_boxes"] = qty
        d["size_table"] = []
        designs.append(d)
    return designs

def detect_format(wb):
    for sn in wb.sheetnames:
        if "solid" in sn.lower():
            return "solid"
    return "ratio"

def parse_po(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    fmt = detect_format(wb)
    designs = []
    for sn in wb.sheetnames:
        ws = wb[sn]
        try:
            if fmt == "solid":
                designs += parse_solid_sheet(ws)
            else:
                designs += parse_ratio_sheet(ws)
        except Exception as e:
            print(f"  WARNING: could not parse sheet '{sn}': {e}")
    return fmt, designs

# ----------------------------------------------------------------------
# PDF DRAWING
# ----------------------------------------------------------------------
def _sw(c, txt, size):
    """Scaled string width (accounts for TEXT_HSCALE)."""
    return c.stringWidth(txt, FONT, size) * TEXT_HSCALE / 100.0

def _text(c, anchor_x, y, txt, size, align="c"):
    """Draw horizontally-scaled text. align: c=center, r=right, l=left.
       Returns (left_x, drawn_width)."""
    wds = _sw(c, txt, size)
    if align == "c":
        x0 = anchor_x - wds / 2
    elif align == "r":
        x0 = anchor_x - wds
    else:
        x0 = anchor_x
    t = c.beginText(x0, y)
    t.setFont(FONT, size)
    t.setHorizScale(TEXT_HSCALE)
    t.textOut(txt)
    c.drawText(t)
    return x0, wds

def draw_sticker(c, x, y, w, h, d, fixed_h=None, show_dest=False,
                 fixed_w=None, draw_box=False):
    """Draw one sticker at fixed size, matching the final print layout.
       fixed_h: dict of exact line heights in inches.
       fixed_w: dict of exact line widths in inches (text is stretched or
                compressed horizontally to hit these exactly).
       show_dest: Phase 3 — print the destination line (e.g. 8411/SG) on top."""
    pad = 2 * mm
    c.saveState()
    if draw_box:
        c.setLineWidth(0.5)
        c.rect(x, y, w, h)      # cut box
    cx = x + w / 2

    n_table = len(d.get("size_table") or [])
    if fixed_h:
        F = {k: v * 72.0 for k, v in fixed_h.items()}   # inches -> points
    else:
        big = min(h / ((7.2 if show_dest else 6.2) + n_table * 0.95), w / 12, 26)
        F = {"dest": big, "item": big, "code": big * 0.95,
             "vol": big, "desc": big}
    W = {k: v * 72.0 for k, v in (fixed_w or {}).items()}

    def shrink(txt, size, maxw):
        while size > 5 and c.stringWidth(txt, FONT, size) > maxw:
            size -= 0.5
        return size

    def scaled_text(anchor_x, ypos, txt, size, scale, align="c"):
        """Draw text at exact horizontal scale (percent). Returns (left, width)."""
        wds = c.stringWidth(txt, FONT, size) * scale / 100.0
        if align == "c":
            x0 = anchor_x - wds / 2
        elif align == "r":
            x0 = anchor_x - wds
        else:
            x0 = anchor_x
        t = c.beginText(x0, ypos)
        t.setFont(FONT, size)
        t.setHorizScale(scale)
        t.textOut(txt)
        c.drawText(t)
        return x0, wds

    def scale_for(key, txt, size):
        """Percent scale so txt at size hits the exact width W[key]."""
        if key in W and txt:
            nat = c.stringWidth(txt, FONT, size)
            if nat > 0:
                return max(20, min(160, W[key] / nat * 100.0))
        return TEXT_HSCALE

    yy = y + h - pad - F["item"] * 1.2

    # Phase 3 only: destination line (e.g. 8411/SG) — centered on top
    if show_dest and d.get("dest"):
        fd = F.get("dest", F["item"])
        sc = scale_for("dest", d["dest"], fd)
        scaled_text(cx, yy, d["dest"], fd, sc, "c")
        yy -= F.get("dest", F["item"]) * 1.35

    # Line 1: item number — centered, exact width
    item = d["item_no"] or ""
    f1 = F["item"]
    sc1 = scale_for("item", item, f1)
    if not fixed_w:
        f1 = shrink(item, f1, w - 2 * pad)
    _, iw = scaled_text(cx, yy, item, f1, sc1, "c")
    item_right = cx + iw / 2
    yy -= F["code"] * 1.35

    # Line 2: code — right-aligned to the item number's right edge, exact width
    f2 = F["code"]
    sc2 = scale_for("code", d["code"] or "", f2)
    if not fixed_w:
        f2 = shrink(d["code"] or "", f2, w - 2 * pad)
    scaled_text(item_right, yy, d["code"] or "", f2, sc2, "r")
    yy -= F["vol"] * 1.35

    # Ratio size table (left side), if any
    if d.get("size_table"):
        th_row = F["code"] * 1.3
        base = min(w * 0.15, 16 * mm, th_row * 2.2)
        widths = [base * 1.5] + [base] * (max(len(n) for _, n in d["size_table"]))
        xs = [x + pad]
        for wd in widths[:-1]:
            xs.append(xs[-1] + wd)
        tfont = min(F["code"] * 0.8, th_row * 0.6)
        ty = yy + F["vol"] * 0.5
        c.setLineWidth(0.7)
        for name, nums in d["size_table"]:
            vals = [name] + [str(nn) for nn in nums]
            for j in range(len(widths)):
                c.rect(xs[j], ty - th_row, widths[j], th_row)
                txt = vals[j] if j < len(vals) else ""
                _text(c, xs[j] + widths[j] / 2, ty - th_row + th_row * 0.3,
                      txt, tfont, "c")
            ty -= th_row
        yy = ty - F["vol"] * 1.1

    # Line 3: (volume) circled-letter PCS — right-aligned, exact width of the
    # "(0.037M3)" part sets the scale for the whole line
    vol = f"({d['volume'].strip('()') if d['volume'] else ''})"
    pcs = f"{d['pcs']} PCS" if d["pcs"] else ""
    L = CIRCLE_LETTER
    mid = f"{vol}  {L}  {pcs}"
    f3 = F["vol"]
    sc3 = scale_for("vol", vol, f3)
    if not fixed_w:
        f3 = shrink(mid, f3, w - 2 * pad)
        right_edge = item_right if fixed_h else             cx + c.stringWidth(mid, FONT, f3) * sc3 / 100 / 2
    else:
        right_edge = item_right
    left_x, total_w = scaled_text(right_edge, yy, mid, f3, sc3, "r")
    vw = c.stringWidth(vol + "  ", FONT, f3) * sc3 / 100
    d_center = left_x + vw + c.stringWidth(L, FONT, f3) * sc3 / 100 / 2
    c.setLineWidth(1.2)
    c.circle(d_center, yy + f3 * 0.35, f3 * 0.55)
    yy -= F["desc"] * 1.5

    # Line 4: description — FIXED height, kept WITHIN line 3's span:
    # never starts before the "(" bracket; compressed in width if too long
    desc = d["description"] or ""
    f4 = F["desc"]
    natural = c.stringWidth(desc, FONT, f4)
    span_left = left_x
    span_right = right_edge
    maxw = span_right - span_left
    scale = min(TEXT_HSCALE, (maxw / natural) * 100 if natural else 100)
    wds = natural * scale / 100
    mid_x = (span_left + span_right) / 2
    t = c.beginText(mid_x - wds / 2, yy)
    t.setFont(FONT, f4)
    t.setHorizScale(scale)
    t.textOut(desc)
    c.drawText(t)
    c.restoreState()

GRID = {8: (2, 4), 4: (2, 2), 2: (1, 2), 1: (1, 1)}  # copies -> (cols, rows)
inch = 72.0

def generate_pdf(designs, out_path, po_label, show_dest=False):
    c = canvas.Canvas(out_path, pagesize=PAGE_SIZE)
    pw, ph = PAGE_SIZE
    header_h = 8 * mm
    for d in designs:
        n = copies_for_height(d["H"] or 999)
        cols, rows = GRID.get(n, (1, 1))
        sw_in, sh_in = STICKER_SIZE_IN.get(n, (6.0, 6.0))
        sw, sh = sw_in * inch, sh_in * inch
        # page header (info strip — not part of stickers)
        c.setFont("Helvetica", 8)
        c.setFillGray(0.35)
        hdr = (f"PO {po_label}  |  {d['code']}  |  "
               f"L{d['L']:g} x W{d['W']:g} x H{d['H']:g} cm  ->  {n} copies "
               f"@ {sw_in:g}\" x {sh_in:g}\" each"
               + (f"  |  Qty required: {d['qty_boxes']} boxes" if d.get('qty_boxes') else ""))
        c.drawString(6 * mm, ph - 6 * mm, hdr)
        c.setFillGray(0)
        # stickers tile edge-to-edge (cut on shared lines, like the print)
        gap = 0
        grid_w = cols * sw + (cols - 1) * gap
        grid_h = rows * sh + (rows - 1) * gap
        x0 = pw - grid_w - 2 * mm          # boxes at the RIGHT of the sheet
        avail_h = ph - header_h
        y0 = header_h and (avail_h - grid_h) / 2  # center in area below header
        y_top = ph - header_h - max(0, (avail_h - grid_h) / 2)
        i = 0
        for r in range(rows):
            for col in range(cols):
                x = x0 + col * (sw + gap)
                y = y_top - (r + 1) * sh - r * gap
                fh, fw = None, None
                if n == 2:
                    fh = TWO_COPY_LINE_H
                elif n == 4:
                    fh, fw = FOUR_COPY_LINE_H, FOUR_COPY_LINE_W
                draw_sticker(c, x, y, sw, sh, d,
                             fixed_h=fh, show_dest=show_dest, fixed_w=fw,
                             draw_box=(n == 4))
                i += 1
                if i >= n:
                    break
        c.showPage()
    c.save()

def write_summary(all_rows, out_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Summary"
    hdrs = ["File", "Type", "Code", "Item No", "Description", "PCS",
            "L", "W", "H", "Volume", "Copies/Sheet", "Qty (boxes)", "Size/Color"]
    ws.append(hdrs)
    from openpyxl.styles import Font
    for c in ws[1]:
        c.font = Font(name="Arial", bold=True)
    for r in all_rows:
        ws.append(r)
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.font = Font(name="Arial")
    widths = [26, 8, 12, 20, 34, 6, 6, 6, 6, 10, 12, 11, 16]
    for i, wdt in enumerate(widths, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = wdt
    wb.save(out_path)

# ----------------------------------------------------------------------
def main(paths, outdir="."):
    summary = []
    for path in paths:
        name = os.path.splitext(os.path.basename(path))[0]
        fmt, designs = parse_po(path)
        print(f"{name}: format={fmt}, {len(designs)} sticker designs")
        po_label = next((d["po_no"] for d in designs if d.get("po_no")), name)
        out_pdf = os.path.join(outdir, f"{name}_STICKERS.pdf")
        generate_pdf(designs, out_pdf, po_label)
        print(f"  -> {out_pdf}")
        for d in designs:
            n = copies_for_height(d["H"] or 999)
            sc = (f"{d.get('size','')} / {d.get('color_code','')}-{d.get('color_name','')}"
                  if d["type"] == "solid" else
                  ", ".join(f"{s}:{'/'.join(map(str,nums))}" for s, nums in d["size_table"] if s != "TOTAL"))
            summary.append([os.path.basename(path), d["type"], d["code"], d["item_no"],
                            d["description"], d["pcs"], d["L"], d["W"], d["H"],
                            d["volume"], n, d.get("qty_boxes"), sc])
    out_x = os.path.join(outdir, "STICKER_SUMMARY.xlsx")
    write_summary(summary, out_x)
    print(f"  -> {out_x}")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    main(sys.argv[1:])

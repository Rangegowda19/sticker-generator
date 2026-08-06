"""
layout_four.py — everything for the 4-COPY sheet (sticker 4.1" x 5").
Edit ONLY this file to change the 4-piece stickers.
All measurements are in INCHES.
"""
from reportlab.lib.units import mm
from common import (FONT, TEXT_HSCALE, VOL_SEP,
                    scaled_text, scale_for, draw_oval, vol_line_parts)

# ----------------------------------------------------------- sticker size
STICKER_W_IN = 4.1
STICKER_H_IN = 5.0
GRID = (2, 2)              # columns x rows on the sheet
PLACEMENT = "center"       # centered so nothing is clipped

# Right shift of the text block inside each sticker (inches). 0 = centered.
SHIFT_RIGHT = 0.25

# Move the whole text block UP inside the sticker (inches). 0 = normal,
# negative moves it down. Keep below ~0.25 or the first line pokes out.
SHIFT_UP = 0.15

# Space between lines in the block (inches).
# Set a number like 0.35 for fixed spacing, or None for automatic
# (automatic = leftover space divided equally; currently about 0.40").
LINE_GAP = 0.20

# ----------------------------------------------------------- text sizes
LINE_W = {                 # exact line WIDTHS (inches)
    "dest": 1.410,         # 8465/EU  (Phase 3 only)
    "item": 3.133,         # 341-489933(71-06)
    "code": 1.800,         # 09-004-000
    "vol":  1.412,         # the "(0.059M3)" part; D + PCS scale with it
}
LINE_H = {                 # exact line HEIGHTS (inches)
    "dest": 0.670,
    "item": 0.623,
    "code": 0.510,
    "vol":  0.625,
    "desc": 0.750,
}

# top/bottom margin inside the sticker, as a fraction of sticker height
# (0.064 x 5" = 0.32")
PAD_V_FRACTION = 0.064

# Oval around the circle letter — specific to THIS sheet.
OVAL_LINE_WIDTH = 3.0   # stroke thickness (points)
OVAL_RX = 0.46          # half-WIDTH  (x font size)
OVAL_RY = 0.52          # half-HEIGHT (x font size)

# ----------------------------------------------------------- drawing
def draw(c, x, y, w, h, d, show_dest=False):
    c.saveState()
    cx = x + w / 2 + SHIFT_RIGHT * 72

    F = {k: v * 72.0 for k, v in LINE_H.items()}
    W = {k: v * 72.0 for k, v in LINE_W.items()}

    # which lines this sticker has
    keys = (["dest"] if (show_dest and d.get("dest")) else []) \
           + ["item", "code", "vol", "desc"]

    # vertical layout: fixed gap if LINE_GAP is set, else spread evenly
    padV = h * PAD_V_FRACTION
    total_txt = sum(F[k] for k in keys)
    if LINE_GAP is None:
        gap = (h - 2 * padV - total_txt) / max(1, len(keys) - 1)
    else:
        gap = LINE_GAP * 72

    baselines = {}
    yy = y + h - padV - F[keys[0]] + SHIFT_UP * 72
    for i, k in enumerate(keys):
        baselines[k] = yy
        if i + 1 < len(keys):
            yy -= gap + F[keys[i + 1]]

    # destination (Phase 3)
    if "dest" in keys:
        fd = F["dest"]
        sc = scale_for(c, W, "dest", d["dest"], fd)
        scaled_text(c, cx, baselines["dest"], d["dest"], fd, sc, "c")

    # item number — centered
    item = d["item_no"] or ""
    f1 = F["item"]
    sc1 = scale_for(c, W, "item", item, f1)
    _, iw = scaled_text(c, cx, baselines["item"], item, f1, sc1, "c")
    item_right = cx + iw / 2

    # code — right-aligned to the item's right edge
    f2 = F["code"]
    sc2 = scale_for(c, W, "code", d["code"] or "", f2)
    scaled_text(c, item_right, baselines["code"], d["code"] or "", f2, sc2, "r")

    # (volume) OVAL-letter PCS — right-aligned to item's right edge
    vol, pcs, L, midtxt = vol_line_parts(d)
    f3 = F["vol"]
    sc3 = scale_for(c, W, "vol", vol, f3)
    left_x, _ = scaled_text(c, item_right, baselines["vol"], midtxt, f3, sc3, "r")
    vw = c.stringWidth(vol + VOL_SEP, FONT, f3) * sc3 / 100
    d_center = left_x + vw + c.stringWidth(L, FONT, f3) * sc3 / 100 / 2
    draw_oval(c, d_center, baselines["vol"] + f3 * 0.35, f3,
              line_width=OVAL_LINE_WIDTH, rx=OVAL_RX, ry=OVAL_RY)

    # description — fixed height, within vol line's span (never before "(")
    desc = d["description"] or ""
    f4 = F["desc"]
    natural = c.stringWidth(desc, FONT, f4)
    maxw = item_right - left_x
    scale = min(TEXT_HSCALE, (maxw / natural) * 100 if natural else 100)
    wds = natural * scale / 100
    mid_x = (left_x + item_right) / 2
    t = c.beginText(mid_x - wds / 2, baselines["desc"])
    t.setFont(FONT, f4)
    t.setHorizScale(scale)
    t.textOut(desc)
    c.drawText(t)
    c.restoreState()

def page_lines(c, pw, ph, x0, y_top, sw, sh):
    """No cutting lines on the 4-piece sheet (removed as requested)."""
    pass
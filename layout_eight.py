"""
layout_eight.py — everything for the 8-COPY sheet (sticker 4" x 2.5").
Edit ONLY this file to change these stickers.
All measurements are in INCHES. Same structure/knobs as layout_four.py.
"""
from reportlab.lib.units import mm
from common import (FONT, TEXT_HSCALE,
                    scaled_text, scale_for, draw_oval, vol_line_parts)

# ----------------------------------------------------------- sticker size
STICKER_W_IN = 4.0
STICKER_H_IN = 2.5
GRID = (2, 4)              # columns x rows on the sheet
PLACEMENT = "right"

# Right shift of the text block inside each sticker (inches). 0 = centered.
SHIFT_RIGHT = 0.25

# Move the whole text block UP inside the sticker (inches). 0 = normal,
# negative moves it down. Separate value per phase.
SHIFT_UP_P3 = 0        # Phase 3 (5 lines)
SHIFT_UP_P2 = 0        # Phase 2 (4 lines)

# Keep the 4-line (Phase 2) spacing IDENTICAL to the 5-line (Phase 3) sticker:
# the item/code/vol/desc lines sit in their Phase-3 positions and the top
# (where the dest line would be) is simply left empty. Set False for the old
# behaviour (4 lines spread to fill the height on their own).
# NOTE: turn this OFF if you want Phase 2 to use its own heights/spacing below.
MATCH_5LINE_SPACING = False

# ----------------------------------------------------------- text sizes
# Line WIDTHS are shared by both phases.
LINE_W = {                 # exact line WIDTHS (inches)
    "dest": 1.38,         # 8465/EU  (Phase 3 only)
    "item": 3.06,         # 341-489933(71-06)
    "code": 1.86,         # 09-004-000
    "vol":  1.675,         # the "(0.059M3)" part; D + PCS scale with it
}

# ---- PHASE 3 (5 lines) : heights + line spacing ----
LINE_H_P3 = {              # exact line HEIGHTS (inches)
    "dest": 0.335,
    "item": 0.262,
    "code": 0.205,
    "vol":  0.263,
    "desc": 0.325,
}
LINE_GAP_P3 = 0.165         # space between lines (inches); None = auto

# ---- PHASE 2 (4 lines) : its OWN heights + line spacing ----
LINE_H_P2 = {              # exact line HEIGHTS (inches)
    "item": 0.385,
    "code": 0.322,
    "vol":  0.323,
    "desc": 0.385,
}
LINE_GAP_P2 = 0.15         # space between lines (inches); None = auto

# top/bottom margin inside the sticker, as a fraction of sticker height
PAD_V_FRACTION = 0.064

# Oval around the circle letter — specific to THIS sheet.
OVAL_LINE_WIDTH = 1.5   # stroke thickness (points)
OVAL_RX = 0.66          # half-WIDTH  (x font size)
OVAL_RY = 0.64          # half-HEIGHT (x font size)

# Space on EACH side of the D/C oval on the volume line — THIS sheet only.
# More spaces = more gap between the oval and the bracket / "16".
VOL_SEP = " "          # 1 space


# ---- PRODUCTION FONT SIZES (points) measured from 4_page_solid.cdr ----
# These set the actual TEXT size. Line spacing still uses LINE_H_* above.
FONT_PT_P3 = {"dest": 22.728, "item": 19.975, "code": 19.797, "vol": 19.975, "desc": 27.656}
FONT_PT_P2 = {"item": 19.975, "code": 19.797, "vol": 19.975, "desc": 27.656}
LETTER_PT = 20.66   # the D/C letter size (slightly bigger than vol)


# ---- PRODUCTION BASELINE GAPS (inches) reverse-engineered from the real sheet.
# Exact vertical gap between each pair of line baselines (Phase 3).
# When set, these OVERRIDE the even LINE_GAP spacing for pixel-exact match.
BASELINE_GAPS_P3 = {"dest_item": 0.395, "item_code": 0.337, "code_vol": 0.368, "vol_desc": 0.596}

# ----------------------------------------------------------- drawing
def draw(c, x, y, w, h, d, show_dest=False):
    c.saveState()
    cx = x + w / 2 + SHIFT_RIGHT * 72

    # which lines this sticker actually has
    has_dest = bool(show_dest and d.get("dest"))
    keys = (["dest"] if has_dest else []) + ["item", "code", "vol", "desc"]

    # pick this phase's own heights + gap + up-shift
    if has_dest:
        LINE_H = LINE_H_P3
        LINE_GAP = LINE_GAP_P3
        shift_up = SHIFT_UP_P3
    else:
        LINE_H = LINE_H_P2
        LINE_GAP = LINE_GAP_P2
        shift_up = SHIFT_UP_P2

    F = {k: v * 72.0 for k, v in LINE_H.items()}
    W = {k: v * 72.0 for k, v in LINE_W.items()}

    # spacing keys: when matching, Phase 2 also reserves the dest slot so the
    # four lines sit in Phase-3 positions; otherwise it uses its own layout.
    if MATCH_5LINE_SPACING and not has_dest:
        # use Phase-3 heights + gap to reproduce exact P3 positions
        F = {k: v * 72.0 for k, v in LINE_H_P3.items()}
        LINE_GAP = LINE_GAP_P3
        layout_keys = ["dest", "item", "code", "vol", "desc"]
    else:
        layout_keys = keys

    padV = h * PAD_V_FRACTION
    total_txt = sum(F[k] for k in layout_keys)
    if LINE_GAP is None:
        gap = (h - 2 * padV - total_txt) / max(1, len(layout_keys) - 1)
    else:
        gap = LINE_GAP * 72

    # assign a baseline to every layout position, then draw only real keys
    baselines = {}
    yy = y + h - padV - F[layout_keys[0]] + shift_up * 72
    for i, k in enumerate(layout_keys):
        baselines[k] = yy
        if i + 1 < len(layout_keys):
            yy -= gap + F[layout_keys[i + 1]]

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
    vol, pcs, L, midtxt = vol_line_parts(d, sep=VOL_SEP)
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
    """No cutting lines on the 8-piece sheet."""
    pass
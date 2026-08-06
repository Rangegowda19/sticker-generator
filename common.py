"""
common.py — shared settings + helpers used by ALL sticker layouts.
Edit here only for things that should change EVERYWHERE at once.
Layout-specific sizes live in layout_eight.py / layout_four.py /
layout_two.py / layout_one.py.
"""
import os

# ---------------------------------------------------------------- settings
# Letter shown inside the oval on the volume line (changes per PO: D, C, ...)
CIRCLE_LETTER = "D"

# Default horizontal narrowness of text (percent). 100 = normal Arial width.
TEXT_HSCALE = 82

# Move the text block RIGHT inside every sticker (inches) — printer's left
# edge does not print properly. 0 = centered.
TEXT_SHIFT_RIGHT = 0.55

# Oval around the letter
OVAL_LINE_WIDTH = 3.0   # stroke thickness (points)
OVAL_RX = 0.46           # half-WIDTH of oval  (x font size)
OVAL_RY = 0.52            # half-HEIGHT of oval (x font size)

# Space on EACH side of the oval letter in the volume line.
# More spaces = more air between the oval and the bracket / number.
VOL_SEP = "   "          # 3 spaces

# ---------------------------------------------------------------- font
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

# ---------------------------------------------------------------- helpers
def scaled_text(c, anchor_x, ypos, txt, size, scale, align="c"):
    """Draw horizontally-scaled text. align: c=center, r=right, l=left.
       Returns (left_x, drawn_width)."""
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

def scale_for(c, W, key, txt, size):
    """Percent scale so txt at size hits the exact width W[key] (points)."""
    if key in W and txt:
        nat = c.stringWidth(txt, FONT, size)
        if nat > 0:
            return max(20, min(160, W[key] / nat * 100.0))
    return TEXT_HSCALE

def shrink(c, txt, size, maxw):
    """Reduce font size until unscaled text fits maxw."""
    while size > 5 and c.stringWidth(txt, FONT, size) > maxw:
        size -= 0.5
    return size

def draw_oval(c, d_center, cy_o, f, line_width=None, rx=None, ry=None):
    """Bold oval around the circle letter.
       Each layout can pass its own line_width / rx / ry; otherwise the
       shared defaults (OVAL_LINE_WIDTH / OVAL_RX / OVAL_RY) are used."""
    c.setLineWidth(line_width if line_width is not None else OVAL_LINE_WIDTH)
    c.setStrokeGray(0)
    rxv = f * (rx if rx is not None else OVAL_RX)
    ryv = f * (ry if ry is not None else OVAL_RY)
    c.ellipse(d_center - rxv, cy_o - ryv, d_center + rxv, cy_o + ryv)

def vol_line_parts(d, sep=None):
    """Build the '(0.059M3)   D   16 PCS' pieces for a design.
       sep = spaces on each side of the letter; defaults to shared VOL_SEP."""
    s = sep if sep is not None else VOL_SEP
    vol = f"({d['volume'].strip('()') if d['volume'] else ''})"
    pcs = f"{d['pcs']} PCS" if d["pcs"] else ""
    L = CIRCLE_LETTER
    return vol, pcs, L, f"{vol}{s}{L}{s}{pcs}"
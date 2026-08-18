"""
ratio_generator.py — sheet counting + orchestration for RATIO stickers.
Ratio has only TWO layouts: 8-up and 2-up (same box sizes as solid).
Uses ratio_parser.py to read the Excel.
"""
import math
import ratio_parser

STICKERS_PER_CARTON = 2

# Same height->copies rule as solid, but ratio only uses 8-up and 2-up.
# 8 copies for short cartons, 2 copies for tall. (No 4-up/1-up in ratio.)
def copies_for_height(h):
    if h is None:
        return 8
    if h <= 18:
        return 8
    return 2          # ratio uses only 8-up or 2-up

def layout_for_height(h):
    return 8 if copies_for_height(h) == 8 else 2

def sheets_for_design(d):
    """(#copies-per-sheet, total stickers, sheets) using cartons x 2."""
    n = copies_for_height(d.get("H"))
    # 'total' in ratio = pieces; but stickers needed = boxes x 2.
    boxes = d.get("boxes")
    if not boxes:
        # if boxes unknown, fall back to 1 sheet
        boxes = 1
    stickers = boxes * STICKERS_PER_CARTON
    sheets = math.ceil(stickers / n)
    return n, stickers, sheets

def parse_po(path):
    return ratio_parser.parse_ratio_po(path)

if __name__ == "__main__":
    ds = parse_po("/mnt/user-data/uploads/P0727-489933-001_-_DF74_-_Ratio.xlsx")
    print(f"{len(ds)} ratio designs")
    for d in ds[:5]:
        n = copies_for_height(d["H"])
        print(f"  {d['sheet']}: H={d['H']} -> {n}-up, table={d['table_type']}, "
              f"sizes={len(d['sizes'])}, total={d['total']}")
# Run this on YOUR PC:  python diagnose.py "P0826-486111-013 - CO59 - Solid.xlsx"
import sys
print("--- dependency check ---")
for mod in ["numpy", "PIL", "openpyxl", "reportlab", "flask"]:
    try:
        __import__(mod); print(f"  {mod}: OK")
    except Exception as e:
        print(f"  {mod}: MISSING  ({e})")

try:
    import pytesseract, shutil
    print("  tesseract program:", "found" if shutil.which("tesseract") else "NOT installed (ok, not needed)")
except Exception:
    print("  pytesseract: not installed (ok, not needed)")

print("--- detection check ---")
import sticker_generator as sg
path = sys.argv[1] if len(sys.argv) > 1 else None
if path:
    res = sg.detect_sheet_letters(path)
    print("  per-sheet letters:", res)
    fmt, designs = sg.parse_po(path)
    from collections import Counter
    print("  design letters:", Counter(d.get('circle_letter') for d in designs))
else:
    print("  (pass the Excel filename as an argument to test detection)")
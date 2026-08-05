# Carton Sticker Generator

Web app that converts a company PO Excel file into a print-ready carton
sticker PDF (Solid stickers; Ratio coming soon).

## Run locally
```
pip install -r requirements.txt
python app.py
```
Open http://localhost:5000

## Files
- `app.py`               — Flask web app (upload page + PDF download)
- `sticker_generator.py` — parsing + PDF layout engine (all sizes in inches,
                           configurable at the top of the file)

## Usage
1. Select sticker type (Solid)
2. Upload the PO Excel (.xlsx)
3. Select Phase (1/2 = 4 lines, 3 = adds destination line e.g. 8411/SG)
4. Download the generated PDF and print at 100% / Actual size

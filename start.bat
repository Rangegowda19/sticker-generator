@echo off
REM ---- One-click launcher for the Sticker Generator ----
cd /d "%~dp0"
echo Starting Sticker Generator...
echo Your browser will open automatically. Keep this window open while using it.
echo (Close this window or press Ctrl+C to stop.)
start "" http://localhost:5000
python app.py
pause

@echo off
echo ============================================
echo   Masco Canada - Cycle Count App
echo ============================================
echo.
echo Installing requirements...
pip install flask openpyxl --quiet
echo.
echo Starting app...
echo.
python app.py
pause

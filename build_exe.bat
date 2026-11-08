@echo off
rem Builds dist\PDF_Binder.exe (one file, no console window)
python -m pip install -r requirements.txt pyinstaller
python -m PyInstaller --noconfirm --onefile --windowed --name PDF_Binder pdf_binder.py
echo.
echo Done: dist\PDF_Binder.exe
pause

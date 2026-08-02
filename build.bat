@echo off
REM Build DocxToPdf.exe from source.
REM Requires: pip install --user pyinstaller pillow
REM Output: dist\DocxToPdf.exe (standalone, no Python needed to run)

cd /d "%~dp0"

echo Building DocxToPdf.exe ...
python -m PyInstaller ^
    --onefile ^
    --windowed ^
    --name DocxToPdf ^
    --icon assets\icon.ico ^
    --noconfirm ^
    docx_to_pdf.py

if errorlevel 1 (
    echo.
    echo BUILD FAILED.
    pause
    exit /b 1
)

echo.
echo Done. The program is at: dist\DocxToPdf.exe
pause

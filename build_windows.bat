@echo off
setlocal enabledelayedexpansion

REM Build Windows executable with PyInstaller
REM Usage: double-click or run from cmd

where python >nul 2>nul
if errorlevel 1 (
  echo Python not found in PATH. Install Python 3.9+ and try again.
  exit /b 1
)

python -m pip install --upgrade pip
python -m pip install -r requirements.txt

python -m PyInstaller ^
  --onefile ^
  --windowed ^
  --name "BlueBook Counter Upper" ^
  bluebook_count.py

echo.
echo Build complete. App is at: dist\BlueBook Counter Upper.exe
pause



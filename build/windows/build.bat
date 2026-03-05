@echo off
setlocal enabledelayedexpansion

REM ==========================================
REM Maintenance Reporter - Windows Build Script
REM Output: dist\MaintenanceReporter\MaintenanceReporter.exe (onedir)
REM ==========================================

cd /d "%~dp0\..\.."
set PROJECT_ROOT=%cd%
set SPEC_FILE=%PROJECT_ROOT%\build\pyinstaller\maintenance_reporter.spec

echo [1/5] Create/activate venv...
if not exist ".venv" (
  python -m venv .venv
)
call ".venv\Scripts\activate.bat"

echo [2/5] Install dependencies...
python -m pip install --upgrade pip
pip install -r requirements.txt
pip install pyinstaller

echo [3/5] Clean previous build artifacts...
if exist "build\pyinstaller\dist" rmdir /s /q "build\pyinstaller\dist"
if exist "build\pyinstaller\build" rmdir /s /q "build\pyinstaller\build"

echo [4/5] Build with PyInstaller (onedir)...
pyinstaller ^
  --clean ^
  --distpath "build\pyinstaller\dist" ^
  --workpath "build\pyinstaller\build" ^
  "%SPEC_FILE%"

echo [5/5] Done.
echo Output folder:
echo   %PROJECT_ROOT%\build\pyinstaller\dist\MaintenanceReporter
pause
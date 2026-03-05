@echo off
setlocal enabledelayedexpansion

REM ==========================================
REM Maintenance Reporter - Windows Build Script
REM Output: dist\MaintenanceReporter.exe (onefile)
REM ==========================================

cd /d "%~dp0\..\.."
set PROJECT_ROOT=%cd%
set SPEC_FILE=%PROJECT_ROOT%\build\pyinstaller\maintenance_reporter_onefile.spec

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
if exist "build\pyinstaller\dist_onefile" rmdir /s /q "build\pyinstaller\dist_onefile"
if exist "build\pyinstaller\build_onefile" rmdir /s /q "build\pyinstaller\build_onefile"

echo [4/5] Build with PyInstaller (onefile)...
pyinstaller ^
  --clean ^
  --distpath "build\pyinstaller\dist_onefile" ^
  --workpath "build\pyinstaller\build_onefile" ^
  "%SPEC_FILE%"

echo [5/5] Done.
echo Output file:
echo   %PROJECT_ROOT%\build\pyinstaller\dist_onefile\MaintenanceReporter.exe
pause

@echo off
cd /d "%~dp0\..\.."
if exist "build\pyinstaller\dist" rmdir /s /q "build\pyinstaller\dist"
if exist "build\pyinstaller\build" rmdir /s /q "build\pyinstaller\build"
if exist "*.spec" del /q "*.spec"
echo Cleaned.
pause

@echo off
setlocal

REM ==========================================
REM Build EXE (PyInstaller) + Installer (Inno Setup)
REM ==========================================

cd /d "%~dp0\..\.."
set PROJECT_ROOT=%cd%

REM 1) Build EXE
call build\windows\build.bat
if errorlevel 1 (
  echo.
  echo [ERROR] EXE build failed.
  pause
  exit /b 1
)

REM 2) Find Inno Setup Compiler (ISCC.exe)
set ISCC=
for %%I in (ISCC.exe) do set ISCC=%%~$PATH:I

if "%ISCC%"=="" (
  if exist "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" set ISCC="C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
)

if "%ISCC%"=="" (
  echo.
  echo [ERROR] Inno Setup Compiler(ISCC.exe) not found.
  echo Please install Inno Setup 6 and ensure ISCC.exe is on PATH.
  echo Download: https://jrsoftware.org/isdl.php
  pause
  exit /b 2
)

echo.
echo [Installer] Compiling setup.iss ...
%ISCC% "%PROJECT_ROOT%\build\installer\setup.iss"

if errorlevel 1 (
  echo.
  echo [ERROR] Installer build failed.
  pause
  exit /b 3
)

echo.
echo [OK] Installer created at:
echo   %PROJECT_ROOT%\build\installer_output
pause

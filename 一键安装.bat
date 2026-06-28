@echo off
chcp 65001 >nul
title Job Assistant Installer (offline)
cd /d "%~dp0"
echo ================================================
echo   Job Assistant - One-Click Installer (offline)
echo ================================================
echo.

rem --- find Python ---
set "PYEXE="
where py >nul 2>nul && set "PYEXE=py"
if not defined PYEXE ( where python >nul 2>nul && set "PYEXE=python" )

if not defined PYEXE (
  echo Python not found. Installing bundled Python...
  for %%f in ("%~dp0bundle\python-*.exe") do set "PYSETUP=%%f"
  if defined PYSETUP (
    "%PYSETUP%" /quiet InstallAllUsers=0 PrependPath=1 Include_pip=1
    set "PYEXE=python"
  ) else (
    echo [ERROR] No bundled Python found. Install Python from python.org then rerun.
    pause
    exit /b 1
  )
)

echo Using Python: %PYEXE%
echo Running installer (this may take a few minutes)...
echo.
%PYEXE% "%~dp0install.py"

pause

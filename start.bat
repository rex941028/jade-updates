@echo off
cd /d "%~dp0"
title Starting...

set "UV=%USERPROFILE%\.local\bin\uv.exe"
if not exist "%UV%" set "UV=%LOCALAPPDATA%\Programs\uv\uv.exe"
if not exist "%UV%" (
    where uv >nul 2>&1
    if not errorlevel 1 ( set "UV=uv" ) else ( goto :install_uv )
)
goto :check_venv

:install_uv
echo [Setup] Installing runtime, please wait...
powershell -NoProfile -ExecutionPolicy Bypass -Command "try { irm https://astral.sh/uv/install.ps1 | iex } catch { exit 1 }" >nul 2>&1
set "UV=%USERPROFILE%\.local\bin\uv.exe"
if not exist "%UV%" set "UV=%LOCALAPPDATA%\Programs\uv\uv.exe"
if not exist "%UV%" goto :error_uv

:check_venv
if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" -c "import openpyxl, tkcalendar" >nul 2>&1
    if errorlevel 1 (
        echo [Setup] Repairing environment...
        rd /s /q .venv >nul 2>&1
    )
)

if not exist ".venv\Scripts\python.exe" (
    echo [Setup] Installing packages, please wait ~1 min...
    "%UV%" venv .venv
    if errorlevel 1 goto :error_pkg
    "%UV%" pip install --python ".venv\Scripts\python.exe" openpyxl tkcalendar
    if errorlevel 1 goto :error_pkg
    echo [Setup] Done.
)

start "" ".venv\Scripts\pythonw.exe" app.py
exit /b 0

:error_uv
echo.
echo [ERROR] Cannot install runtime. Check internet and try again.
pause
exit /b 1

:error_pkg
echo.
echo [ERROR] Package install failed. Check internet and try again.
pause
exit /b 1

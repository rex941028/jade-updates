@echo off
chcp 65001 > nul
cd /d "%~dp0"
title 瑾琰上品 - 啟動中...

:: ══════════════════════════════════════════════════════
::  Step 1：尋找 uv，找不到就自動安裝
:: ══════════════════════════════════════════════════════
set "UV=%USERPROFILE%\.local\bin\uv.exe"
if not exist "%UV%" set "UV=%LOCALAPPDATA%\Programs\uv\uv.exe"
if not exist "%UV%" (
    where uv >nul 2>&1
    if not errorlevel 1 (
        set "UV=uv"
    ) else (
        echo 首次啟動：正在安裝必要組件，請稍候 30 秒...
        powershell -NoProfile -ExecutionPolicy Bypass -Command ^
            "try { irm https://astral.sh/uv/install.ps1 | iex } catch { exit 1 }" >nul 2>&1
        set "UV=%USERPROFILE%\.local\bin\uv.exe"
        if not exist "%UV%" set "UV=%LOCALAPPDATA%\Programs\uv\uv.exe"
    )
)

if not exist "%UV%" (
    echo.
    echo  ╔══════════════════════════════════════════════╗
    echo  ║  [錯誤] 無法安裝執行組件                      ║
    echo  ║  請確認網路連線正常後重新開啟程式。            ║
    echo  ║  若持續失敗，請聯絡系統管理員。                ║
    echo  ╚══════════════════════════════════════════════╝
    pause
    exit /b 1
)

:: ══════════════════════════════════════════════════════
::  Step 2：驗證現有 .venv；損壞或跨機搬移時自動重建
:: ══════════════════════════════════════════════════════
if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" -c "import openpyxl, tkcalendar" >nul 2>&1
    if errorlevel 1 (
        echo 偵測到執行環境異常，正在自動修復，請稍候...
        rmdir /s /q .venv >nul 2>&1
    )
)

if not exist ".venv\Scripts\python.exe" (
    echo 安裝必要套件中，請稍候約 1 分鐘...
    "%UV%" venv .venv
    if errorlevel 1 goto :install_error
    "%UV%" pip install --python ".venv\Scripts\python.exe" openpyxl tkcalendar
    if errorlevel 1 goto :install_error
    echo 安裝完成！
)

:: ══════════════════════════════════════════════════════
::  Step 3：啟動主程式（無命令列視窗）
:: ══════════════════════════════════════════════════════
start "" ".venv\Scripts\pythonw.exe" app.py
exit /b 0

:install_error
echo.
echo  ╔══════════════════════════════════════════════╗
echo  ║  [錯誤] 套件安裝失敗                          ║
echo  ║  請確認網路連線正常後重新開啟程式。            ║
echo  ║  若持續失敗，請聯絡系統管理員。                ║
echo  ╚══════════════════════════════════════════════╝
pause
exit /b 1

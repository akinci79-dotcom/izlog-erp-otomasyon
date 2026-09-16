@echo off
chcp 65001 >nul
title IZLOG KPI Kurulum
echo.
echo === IZLOG KPI KURULUM (tek tik) ===
echo.

REM Bu .bat dosyasinin bulundugu klasor = KPI klasoru
set "KPI_DIR=%~dp0"
set "KPI_DIR=%KPI_DIR:~0,-1%"
cd /d "%KPI_DIR%"

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0kpi_kur.ps1"
if errorlevel 1 (
    echo.
    echo HATA olustu. Yukaridaki mesaji okuyun.
    pause
    exit /b 1
)

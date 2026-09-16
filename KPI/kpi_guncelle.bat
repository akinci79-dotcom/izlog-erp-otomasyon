@echo off
chcp 65001 >nul
title IZLOG KPI Kod Guncelleme
echo.
echo === IZLOG KPI KOD GUNCELLEME ===
echo Sadece Python kodu guncellenir; sablon ve ayarlar korunur.
echo.

cd /d "%~dp0"

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0kpi_guncelle.ps1"
if errorlevel 1 (
    echo.
    echo HATA olustu.
    pause
    exit /b 1
)

echo.
pause

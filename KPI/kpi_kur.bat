@echo off
chcp 65001 >nul
title IZLOG KPI Guncelleme
echo.
echo === IZLOG KPI GUNCELLEME (tek tik) ===
echo GitHub'dan son kod cekilir. ayarlar.py ve SQL korunur.
echo Eski sablon varsa otomatik yeni formullu sablon alinir.
echo.
echo Rapor uretmek icin: kpi_rapor_olustur.bat
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

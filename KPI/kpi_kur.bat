@echo off
chcp 65001 >nul
title IZLOG KPI Guncelleme
echo.
echo === IZLOG KPI GUNCELLEME (tek tik) ===
echo GitHubdan son kod cekilir. ayarlar.py ve SQL korunur.
echo Eski sablon varsa otomatik yeni formullu sablon alinir.
echo.
echo Rapor uretmek icin: kpi_rapor_olustur.bat
echo.

cd /d "%~dp0"

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0kpi_guncelle.ps1" -KurModu
if errorlevel 1 (
    echo.
    echo HATA olustu. Yukaridaki mesaji okuyun.
    pause
    exit /b 1
)

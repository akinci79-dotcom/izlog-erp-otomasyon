@echo off
title KPI Rapor Olustur
cd /d "%~dp0"
echo.
echo KPI raporu baslatiliyor...
echo ONEMLI: Siyah pencereye tiklamayin - islem durur. Durduysa Enter veya Esc basin.
echo.
python -u kpi_rapor_olustur.py
if errorlevel 1 (
    echo.
    echo HATA: Rapor olusturulamadi. Yukaridaki mesaji okuyun.
)
echo.
pause

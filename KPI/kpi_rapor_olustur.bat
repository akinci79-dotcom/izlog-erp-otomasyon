@echo off
chcp 65001 >nul
title IZLOG KPI Rapor
echo.
echo === IZLOG KPI RAPORU ===
echo.

set "KPI_DIR=%~dp0"
set "KPI_DIR=%KPI_DIR:~0,-1%"
cd /d "%KPI_DIR%"

if not exist "ayarlar.py" (
    echo HATA: ayarlar.py bulunamadi.
    echo Once ayarlar.example.py dosyasini ayarlar.py olarak kopyalayin.
    echo Tarih araligi ve DB_SIFRE alanlarini doldurun.
    echo.
    pause
    exit /b 1
)

if not exist "referans\kpi_sablon.xlsx" (
    echo HATA: referans\kpi_sablon.xlsx bulunamadi.
    echo Temmuz KPI dosyanizi bu konuma kpi_sablon.xlsx adiyla kopyalayin.
    echo.
    pause
    exit /b 1
)

python kpi_rapor_olustur.py
set "EXIT_CODE=%ERRORLEVEL%"

echo.
if "%EXIT_CODE%"=="0" (
    echo Rapor hazir: %KPI_DIR%\raporlar\
    if exist "raporlar\kpi_rapor.xlsx" start "" "raporlar\kpi_rapor.xlsx"
    if exist "raporlar\kpi_rapor.xlsm" start "" "raporlar\kpi_rapor.xlsm"
) else (
    echo HATA olustu. Yukaridaki mesaji okuyun.
)

echo.
pause
exit /b %EXIT_CODE%

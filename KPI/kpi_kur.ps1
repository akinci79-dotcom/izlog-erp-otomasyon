# KPI modulu — tek seferde kurulum
# Sag tik -> "PowerShell ile Calistir"  VEYA  PowerShell'de:
#   cd "C:\Users\hakinci\Desktop\Kodlarım\Cursor ERP Otomasyon\KPI"
#   powershell -ExecutionPolicy Bypass -File kpi_kur.ps1

$ErrorActionPreference = "Stop"

# Bu script KPI klasorunun icinde veya ust klasorde calistirilabilir
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if (Test-Path (Join-Path $ScriptDir "kpi_analiz.py")) {
    $KpiDir = $ScriptDir
} else {
    $KpiDir = Join-Path $ScriptDir "KPI"
}

$Base = Split-Path -Parent $KpiDir
$Temp = Join-Path $Base "izlog-kpi-temp"
$Repo = "https://github.com/akinci79-dotcom/izlog-erp-otomasyon.git"
$Branch = "cursor/kpi-analiz-rapor-0bd3"

Write-Host ""
Write-Host "=== IZLOG KPI KURULUM ===" -ForegroundColor Cyan
Write-Host "Hedef: $KpiDir"
Write-Host ""

# Dosyalar yoksa GitHub'dan indir
if (-not (Test-Path (Join-Path $KpiDir "kpi_analiz.py"))) {
    Write-Host "[1/4] KPI dosyalari indiriliyor..." -ForegroundColor Yellow
    if (Test-Path $Temp) { Remove-Item $Temp -Recurse -Force }
    Set-Location $Base
    git clone -b $Branch --depth 1 $Repo izlog-kpi-temp
    New-Item -ItemType Directory -Force -Path $KpiDir | Out-Null
    Copy-Item "$Temp\KPI\*" $KpiDir -Recurse -Force
    Remove-Item $Temp -Recurse -Force
    Write-Host "      Dosyalar kopyalandi." -ForegroundColor Green
} else {
    Write-Host "[1/4] KPI dosyalari zaten mevcut, atlaniyor." -ForegroundColor Green
}

Set-Location $KpiDir

# ayarlar.py
Write-Host "[2/4] ayarlar.py hazirlaniyor..." -ForegroundColor Yellow
if (-not (Test-Path "ayarlar.py")) {
    Copy-Item "ayarlar.example.py" "ayarlar.py"
    Write-Host "      ayarlar.py olusturuldu — DB_SIFRE alanini doldurmayi unutmayin!" -ForegroundColor Magenta
} else {
    Write-Host "      ayarlar.py zaten var." -ForegroundColor Green
}

# pip
Write-Host "[3/4] Python paketleri kuruluyor..." -ForegroundColor Yellow
pip install -r requirements.txt

# test raporu
Write-Host "[4/4] Ornek rapor uretiliyor..." -ForegroundColor Yellow
python kpi_rapor_olustur.py --ornek

Write-Host ""
Write-Host "=== TAMAMLANDI ===" -ForegroundColor Green
Write-Host "Rapor: $KpiDir\raporlar\kpi_rapor_ORNEK.xlsx"
Write-Host ""
Write-Host "Sonraki adimlar:"
Write-Host "  1. ayarlar.py icinde DB_SIFRE doldurun"
Write-Host "  2. python kpi_rapor_olustur.py   (gercek veri)"
Write-Host ""
Read-Host "Kapatmak icin Enter'a basin"

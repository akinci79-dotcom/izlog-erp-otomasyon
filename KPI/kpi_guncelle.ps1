# KPI modulu — GitHub'dan guncelle (ayarlar.py ve sablon korunur)
# KPI klasorunden:
#   powershell -ExecutionPolicy Bypass -File kpi_guncelle.ps1

$ErrorActionPreference = "Stop"

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if (Test-Path (Join-Path $ScriptDir "kpi_rapor_olustur.py")) {
    $KpiDir = $ScriptDir
} else {
    $KpiDir = Join-Path $ScriptDir "KPI"
}

$Base = Split-Path -Parent $KpiDir
$Temp = Join-Path $Base "izlog-kpi-temp"
$Repo = "https://github.com/akinci79-dotcom/izlog-erp-otomasyon.git"
$Branch = "cursor/kpi-analiz-rapor-0bd3"

Write-Host ""
Write-Host "=== IZLOG KPI GUNCELLEME ===" -ForegroundColor Cyan
Write-Host "Hedef: $KpiDir"
Write-Host ""

if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
    throw "git bulunamadi. Git for Windows kurulu olmali."
}

if (Test-Path $Temp) { Remove-Item $Temp -Recurse -Force }
Set-Location $Base
git clone -b $Branch --depth 1 $Repo izlog-kpi-temp

$Kaynak = Join-Path $Temp "KPI"
if (-not (Test-Path $Kaynak)) {
    throw "Klonlanan repoda KPI klasoru yok: $Kaynak"
}

New-Item -ItemType Directory -Force -Path $KpiDir | Out-Null

# Korunacak dosyalar
$AyarlarYedek = Join-Path $env:TEMP "izlog_kpi_ayarlar_yedek.py"
$SablonYedek = Join-Path $env:TEMP "izlog_kpi_sablon_yedek.xlsx"
$VeriSqlYedek = Join-Path $env:TEMP "izlog_kpi_veri_rapor_yedek.sql"
if (Test-Path (Join-Path $KpiDir "ayarlar.py")) {
    Copy-Item (Join-Path $KpiDir "ayarlar.py") $AyarlarYedek -Force
}
if (Test-Path (Join-Path $KpiDir "referans\kpi_sablon.xlsx")) {
    Copy-Item (Join-Path $KpiDir "referans\kpi_sablon.xlsx") $SablonYedek -Force
}
if (Test-Path (Join-Path $KpiDir "referans\kpi_veri_rapor.sql")) {
    Copy-Item (Join-Path $KpiDir "referans\kpi_veri_rapor.sql") $VeriSqlYedek -Force
}

Copy-Item "$Kaynak\*" $KpiDir -Recurse -Force

if (Test-Path $AyarlarYedek) {
    Copy-Item $AyarlarYedek (Join-Path $KpiDir "ayarlar.py") -Force
    Remove-Item $AyarlarYedek -Force
}
if (Test-Path $SablonYedek) {
    New-Item -ItemType Directory -Force -Path (Join-Path $KpiDir "referans") | Out-Null
    Copy-Item $SablonYedek (Join-Path $KpiDir "referans\kpi_sablon.xlsx") -Force
    Remove-Item $SablonYedek -Force
}
if (Test-Path $VeriSqlYedek) {
    New-Item -ItemType Directory -Force -Path (Join-Path $KpiDir "referans") | Out-Null
    Copy-Item $VeriSqlYedek (Join-Path $KpiDir "referans\kpi_veri_rapor.sql") -Force
    Remove-Item $VeriSqlYedek -Force
}

Remove-Item $Temp -Recurse -Force
Set-Location $KpiDir

Write-Host "Guncelleme tamamlandi." -ForegroundColor Green
Write-Host ""
Write-Host "Sonraki adim:"
Write-Host "  python kpi_rapor_olustur.py"
Write-Host ""
Write-Host "Basarili ciktida su satiri gormelisiniz:"
Write-Host "  BAŞARILI: KPI şablon raporu -> ..."
Write-Host ""

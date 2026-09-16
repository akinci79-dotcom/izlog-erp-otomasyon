# KPI modulu — tek tik guncelleme (kpi_kur.bat bunu cagirir)
# GitHub'dan son kodu ceker; ayarlar.py + kpi_veri_rapor.sql korunur.
# Eski / eksik sablon algilanirsa -YeniSablon ile formullu sablon alinir.

$ErrorActionPreference = "Stop"

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if (Test-Path (Join-Path $ScriptDir "kpi_rapor_olustur.py")) {
    $KpiDir = $ScriptDir
} else {
    $KpiDir = Join-Path $ScriptDir "KPI"
}

Write-Host ""
Write-Host "=== IZLOG KPI GUNCELLEME ===" -ForegroundColor Cyan
Write-Host "Hedef: $KpiDir"
Write-Host ""

$YeniSablon = $false
$SablonYolu = Join-Path $KpiDir "referans\kpi_sablon.xlsx"
if (-not (Test-Path $SablonYolu)) {
    Write-Host "Sablon bulunamadi -> yeni formullu sablon indirilecek." -ForegroundColor Yellow
    $YeniSablon = $true
} elseif ((Get-Item $SablonYolu).Length -lt 900000) {
    Write-Host "Eski sablon algilandi -> yeni formullu sablon indirilecek." -ForegroundColor Yellow
    $YeniSablon = $true
}

$GuncelleScript = Join-Path $ScriptDir "kpi_guncelle.ps1"
if (-not (Test-Path $GuncelleScript)) {
    throw "kpi_guncelle.ps1 bulunamadi: $GuncelleScript"
}

$GuncelleArgs = @(
    "-NoProfile",
    "-ExecutionPolicy", "Bypass",
    "-File", $GuncelleScript
)
if ($YeniSablon) {
    $GuncelleArgs += "-YeniSablon"
}

& powershell @GuncelleArgs
if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
}

Set-Location $KpiDir

Write-Host ""
Write-Host "Python paketleri kontrol ediliyor..." -ForegroundColor Yellow
pip install -r requirements.txt -q
if ($LASTEXITCODE -ne 0) {
    Write-Host "pip uyarisi: paket kurulumu basarisiz olabilir." -ForegroundColor Yellow
}

Write-Host ""
Write-Host "=== TAMAMLANDI ===" -ForegroundColor Green
Write-Host ""
Write-Host "Sonraki adim: kpi_rapor_olustur.bat ile rapor uretin"
Write-Host "  veya: python kpi_rapor_olustur.py"
Write-Host ""
Read-Host "Kapatmak icin Enter'a basin"

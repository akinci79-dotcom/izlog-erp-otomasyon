# KPI modulu — GitHubdan guncelle
#
# Normal guncelleme (mevcut sablon + ayarlar korunur):
#   powershell -ExecutionPolicy Bypass -File kpi_guncelle.ps1
#
# Yeni formullu sablon + guncel ayar bayraklari (ilk gecis / Agustos 2026+ sablon):
#   powershell -ExecutionPolicy Bypass -File kpi_guncelle.ps1 -YeniSablon
#
# kpi_kur.bat ile ayni is (pip + bekleme):
#   powershell -ExecutionPolicy Bypass -File kpi_guncelle.ps1 -KurModu
#
param(
    [string]$Branch = "cursor/ozet-manuel-tablo-guncelle-0bd3",
    [switch]$YeniSablon,
    [switch]$KurModu
)

$ErrorActionPreference = "Stop"

function Get-SablonSurumu {
    param([string]$Klasor)
    $dosya = Join-Path $Klasor "referans\sablon_surumu.txt"
    if (-not (Test-Path $dosya)) { return $null }
    $size = $null
    $sha256 = $null
    foreach ($line in Get-Content $dosya) {
        if ($line -match "^size=(\d+)$") { $size = [long]$matches[1] }
        if ($line -match "^sha256=([0-9a-fA-F]+)$") { $sha256 = $matches[1].ToLower() }
    }
    if ($null -eq $size -or $null -eq $sha256) { return $null }
    return @{ Size = $size; Sha256 = $sha256 }
}

function Test-SablonRepodaGuncel {
    param(
        [string]$SablonYolu,
        $Beklenen
    )
    if ($null -eq $Beklenen) { return $false }
    if (-not (Test-Path $SablonYolu)) { return $false }
    $dosya = Get-Item $SablonYolu
    if ($dosya.Length -ne $Beklenen.Size) { return $false }
    $hash = (Get-FileHash -Path $SablonYolu -Algorithm SHA256).Hash.ToLower()
    return ($hash -eq $Beklenen.Sha256)
}

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if (Test-Path (Join-Path $ScriptDir "kpi_rapor_olustur.py")) {
    $KpiDir = $ScriptDir
} else {
    $KpiDir = Join-Path $ScriptDir "KPI"
}

$Base = Split-Path -Parent $KpiDir
$Temp = Join-Path $Base "izlog-kpi-temp"
$Repo = "https://github.com/akinci79-dotcom/izlog-erp-otomasyon.git"
$SablonYolu = Join-Path $KpiDir "referans\kpi_sablon.xlsx"

Write-Host ""
Write-Host "=== IZLOG KPI GUNCELLEME ===" -ForegroundColor Cyan
Write-Host "Hedef : $KpiDir"
Write-Host "Branch: $Branch"
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

$RepoSablonSurumu = Get-SablonSurumu -Klasor $Kaynak
if (-not $YeniSablon) {
    if (-not (Test-SablonRepodaGuncel -SablonYolu $SablonYolu -Beklenen $RepoSablonSurumu)) {
        if ($null -eq $RepoSablonSurumu) {
            Write-Host "sablon_surumu.txt yok - eski sablon boyutu kontrolu yapiliyor." -ForegroundColor Yellow
            if (-not (Test-Path $SablonYolu) -or ((Get-Item $SablonYolu).Length -lt 900000)) {
                Write-Host "Eski veya eksik sablon - YeniSablon modu otomatik acildi." -ForegroundColor Yellow
                $YeniSablon = $true
            }
        } else {
            Write-Host "Repodaki sablon surumu ile yerel dosya farkli - YeniSablon modu otomatik acildi." -ForegroundColor Yellow
            $YeniSablon = $true
        }
    }
}

if ($YeniSablon) {
    Write-Host "Mod   : YeniSablon (repodaki kpi_sablon.xlsx + ayar bayraklari guncellenir)" -ForegroundColor Yellow
} else {
    Write-Host "Mod   : Standart (ayarlar.py + kpi_sablon.xlsx korunur)"
}
Write-Host ""

New-Item -ItemType Directory -Force -Path $KpiDir | Out-Null

# Korunacak dosyalar
$AyarlarYedek = Join-Path $env:TEMP "izlog_kpi_ayarlar_yedek.py"
$SablonYedek = Join-Path $env:TEMP "izlog_kpi_sablon_yedek.xlsx"
$VeriSqlYedek = Join-Path $env:TEMP "izlog_kpi_veri_rapor_yedek.sql"

if (Test-Path (Join-Path $KpiDir "ayarlar.py")) {
    Copy-Item (Join-Path $KpiDir "ayarlar.py") $AyarlarYedek -Force
}
if (-not $YeniSablon) {
    if (Test-Path (Join-Path $KpiDir "referans\kpi_sablon.xlsx")) {
        Copy-Item (Join-Path $KpiDir "referans\kpi_sablon.xlsx") $SablonYedek -Force
    }
}
if (Test-Path (Join-Path $KpiDir "referans\kpi_veri_rapor.sql")) {
    Copy-Item (Join-Path $KpiDir "referans\kpi_veri_rapor.sql") $VeriSqlYedek -Force
}

# Repodan kaldirilmis eski dosyalari temizle
$Korunacaklar = @(
    "ayarlar.py",
    "raporlar",
    "referans\kpi_veri_rapor.sql",
    "__pycache__",
    "izlog-kpi-temp"
)
if (-not $YeniSablon) {
    $Korunacaklar += "referans\kpi_sablon.xlsx"
}

$YeniDosyalar = Get-ChildItem -Path $Kaynak -Recurse -File | ForEach-Object {
    $_.FullName.Substring($Kaynak.Length + 1)
}
Get-ChildItem -Path $KpiDir -Recurse -File -ErrorAction SilentlyContinue | ForEach-Object {
    $goreli = $_.FullName.Substring($KpiDir.Length + 1)
    $korunan = $false
    foreach ($k in $Korunacaklar) {
        if ($goreli -eq $k -or $goreli.StartsWith("$k\")) { $korunan = $true; break }
    }
    if (-not $korunan -and ($YeniDosyalar -notcontains $goreli)) {
        Write-Host "  Eski dosya siliniyor: $goreli" -ForegroundColor DarkYellow
        Remove-Item $_.FullName -Force -ErrorAction SilentlyContinue
    }
}

Copy-Item "$Kaynak\*" $KpiDir -Recurse -Force

if (Test-Path $AyarlarYedek) {
    Copy-Item $AyarlarYedek (Join-Path $KpiDir "ayarlar.py") -Force
    Remove-Item $AyarlarYedek -Force
} elseif (-not (Test-Path (Join-Path $KpiDir "ayarlar.py"))) {
    Copy-Item (Join-Path $KpiDir "ayarlar.example.py") (Join-Path $KpiDir "ayarlar.py") -Force
    Write-Host "  ayarlar.py olusturuldu (ayarlar.example.py dosyasindan). DB sifresini doldurun." -ForegroundColor Yellow
}

if ($YeniSablon) {
    $SablonHedef = Join-Path $KpiDir "referans\kpi_sablon.xlsx"
    if (-not (Test-Path $SablonHedef)) {
        throw "Repoda yeni sablon bulunamadi: $SablonHedef"
    }
    Write-Host "  Yeni kpi_sablon.xlsx repodan alindi." -ForegroundColor Green

    $AyarlarDosya = Join-Path $KpiDir "ayarlar.py"
    if (Test-Path $AyarlarDosya) {
        $icerik = Get-Content $AyarlarDosya -Raw -Encoding UTF8
        $bayraklar = @{
            "KPI_ZARAR_DETAY_GUNCELLE" = "False"
            "KPI_ARAC_TIPI_PERFORMANS_GUNCELLE" = "False"
            "KPI_OZET_MANUEL_TABLOLAR_GUNCELLE" = "False"
        }
        foreach ($anahtar in $bayraklar.Keys) {
            $desen = "($anahtar\s*=\s*)(True|False)"
            if ($icerik -match $desen) {
                $icerik = $icerik -replace $desen, "`${1}$($bayraklar[$anahtar])"
            } else {
                $icerik += "`n$anahtar = $($bayraklar[$anahtar])`n"
            }
        }
        $icerik = $icerik -replace '(?m)^\s*KPI_RAPOR_DOSYASI\s*=\s*["'']kpi_rapor\.xls[xm]["'']\s*\r?\n', ''
        Set-Content -Path $AyarlarDosya -Value $icerik -Encoding UTF8 -NoNewline
        Write-Host "  ayarlar.py: formullu sablon icin 3 bayrak False yapildi." -ForegroundColor Green
        Write-Host "    KPI_ZARAR_DETAY_GUNCELLE = False"
        Write-Host "    KPI_ARAC_TIPI_PERFORMANS_GUNCELLE = False"
        Write-Host "    KPI_OZET_MANUEL_TABLOLAR_GUNCELLE = False"
    }
} elseif (Test-Path $SablonYedek) {
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

Write-Host ""
Write-Host "Guncelleme tamamlandi." -ForegroundColor Green
Write-Host ""
Write-Host "Sonraki adimlar:"
Write-Host "  1. ayarlar.py icinde KPI_DONEM ve DB sifresini kontrol edin"
Write-Host "  2. python kpi_rapor_olustur.py"
Write-Host ""
Write-Host "Basarili ciktida su satiri gormelisiniz:"
Write-Host "  BASARILI: KPI sablon raporu -> ..."
Write-Host ""

if ($KurModu) {
    Write-Host "Python paketleri kontrol ediliyor..." -ForegroundColor Yellow
    pip install -r requirements.txt -q
    if ($LASTEXITCODE -ne 0) {
        Write-Host "pip uyarisi: paket kurulumu basarisiz olabilir." -ForegroundColor Yellow
    }
    Write-Host ""
    Write-Host "=== TAMAMLANDI ===" -ForegroundColor Green
    Write-Host "Sonraki adim: kpi_rapor_olustur.bat veya python kpi_rapor_olustur.py"
    Write-Host ""
    Read-Host "Kapatmak icin Enter tusuna basin"
}

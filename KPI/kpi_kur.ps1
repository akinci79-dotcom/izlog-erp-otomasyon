# Geriye uyumluluk: kpi_kur.bat artik dogrudan kpi_guncelle.ps1 cagirir.
# Bu dosyayi elle calistirirsaniz ayni isi yapar.

$ErrorActionPreference = "Stop"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$GuncelleScript = Join-Path $ScriptDir "kpi_guncelle.ps1"

if (-not (Test-Path $GuncelleScript)) {
    throw "kpi_guncelle.ps1 bulunamadi: $GuncelleScript"
}

& powershell -NoProfile -ExecutionPolicy Bypass -File $GuncelleScript -KurModu
exit $LASTEXITCODE

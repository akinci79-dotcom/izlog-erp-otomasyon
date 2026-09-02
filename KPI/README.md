# İzlog Lojistik — KPI Rapor Otomasyonu

Oracle ERP verilerini **mevcut KPI Excel şablonunuza** yapıştırır (VERİ + Filo Detay), pivot sayfalarını günceller.

**Otomasyon projesinden bağımsızdır** — yalnızca `KPI/` klasörü kullanılır.

## Hızlı kurulum (Windows)

**En kolay yol** — KPI klasöründen:

```powershell
cd "C:\Users\hakinci\Desktop\Kodlarım\Cursor ERP Otomasyon\KPI"
powershell -ExecutionPolicy Bypass -File kpi_kur.ps1
```

**Güncelleme** (yeni kodu çekmek için — `ayarlar.py` ve şablon korunur):

```powershell
$KpiDir = "C:\Users\hakinci\Desktop\Kodlarım\Cursor ERP Otomasyon\KPI"
$Base = "C:\Users\hakinci\Desktop\Kodlarım"
$Temp = "$Base\izlog-kpi-temp"
Set-Location $Base
if (Test-Path $Temp) { Remove-Item $Temp -Recurse -Force }
git clone -b cursor/kpi-analiz-rapor-0bd3 --depth 1 https://github.com/akinci79-dotcom/izlog-erp-otomasyon.git izlog-kpi-temp
if (Test-Path "$KpiDir\ayarlar.py") { Copy-Item "$KpiDir\ayarlar.py" "$env:TEMP\izlog_kpi_ayarlar_yedek.py" -Force }
if (Test-Path "$KpiDir\referans\kpi_sablon.xlsx") { Copy-Item "$KpiDir\referans\kpi_sablon.xlsx" "$env:TEMP\izlog_kpi_sablon_yedek.xlsx" -Force }
Copy-Item "$Temp\KPI\*" $KpiDir -Recurse -Force
if (Test-Path "$env:TEMP\izlog_kpi_ayarlar_yedek.py") { Copy-Item "$env:TEMP\izlog_kpi_ayarlar_yedek.py" "$KpiDir\ayarlar.py" -Force }
if (Test-Path "$env:TEMP\izlog_kpi_sablon_yedek.xlsx") { New-Item -ItemType Directory -Force -Path "$KpiDir\referans" | Out-Null; Copy-Item "$env:TEMP\izlog_kpi_sablon_yedek.xlsx" "$KpiDir\referans\kpi_sablon.xlsx" -Force }
Remove-Item $Temp -Recurse -Force
Set-Location $KpiDir
python kpi_rapor_olustur.py
```

`kpi_guncelle.ps1` dosyası geldikten sonra kısa yol:

```powershell
cd "C:\Users\hakinci\Desktop\Kodlarım\Cursor ERP Otomasyon\KPI"
powershell -ExecutionPolicy Bypass -File ".\kpi_guncelle.ps1"
python kpi_rapor_olustur.py
```

**Manuel kurulum** (tüm değişkenleri tek seferde yapıştırın; satır satır değil):

```powershell
$Base = "C:\Users\hakinci\Desktop\Kodlarım"
$KpiDir = "$Base\Cursor ERP Otomasyon\KPI"
$Temp = "$Base\izlog-kpi-temp"
Set-Location $Base
if (Test-Path $Temp) { Remove-Item $Temp -Recurse -Force }
git clone -b cursor/kpi-analiz-rapor-0bd3 --depth 1 https://github.com/akinci79-dotcom/izlog-erp-otomasyon.git izlog-kpi-temp
Copy-Item "$Temp\KPI\*" $KpiDir -Recurse -Force
Remove-Item $Temp -Recurse -Force
Set-Location $KpiDir
if (-not (Test-Path ayarlar.py)) { Copy-Item ayarlar.example.py ayarlar.py }
pip install -r requirements.txt
```

## Şablonu ve VERİ SQL'ini yerleştirin

Temmuz KPI dosyanızı şuraya kopyalayın:

```
KPI\referans\kpi_sablon.xlsx
```

Uyumsoft VERİ raporu SQL'i repoda `referans/kpi_veri_rapor.sql` olarak kayıtlıdır (LojistikYükSevkKalemRaporu). Kod SQL'e alan eklemez; çalıştırma anında `@CoCode@`, `@BranchCodes@`, `@DocDateF@`, `@DocDateL@` parametreleri `ayarlar.py` değerleriyle doldurulur.

## ayarlar.py

```python
KPI_BASLANGIC_TARIHI = "01.07.2026"
KPI_BITIS_TARIHI = "31.07.2026"
CO_CODE = "IZLOG"
BRANCH_CODE = "MERKEZ"
DB_SIFRE = "..."
```

## Çalıştırma

**Çift tık:** `kpi_rapor_olustur.bat` — yalnızca `python kpi_rapor_olustur.py` çalıştırır (pencere kapanmasın diye sonunda Enter bekler).

Bekleme süresi bat dosyasından değil; Oracle veri çekimi + Excel pivot yenileme + sütun genişletmeden gelir (2500+ satırda birkaç dakika normal).

**Komut satırı:**

```powershell
cd KPI
python kpi_rapor_olustur.py
```

**Ne yapar:**
1. `referans/kpi_sablon.xlsx` kopyalanır → `raporlar/kpi_rapor.xlsx` (veya şablon `.xlsm` ise `.xlsm`)
2. **VERİ** ve **Filo Detay** sayfalarına Oracle verisi **Excel COM** ile yazılır (pivot şablonu bozulmaz)
3. Pivotlar yenilenir, sütunlar genişletilir

## Eski analiz raporu (isteğe bağlı)

```powershell
python kpi_rapor_olustur.py --analiz
python kpi_rapor_olustur.py --ornek
```

## Pivot yenileme ve sütun genişliği

Otomatik pivot yenileme ve sütun AutoFit için `pip install pywin32` ve yüklü Microsoft Excel gerekir.
`ayarlar.py` içinde `KPI_SUTUN_AUTOFIT = True` (varsayılan) — pivot özet sayfalarındaki tutar sütunları da dahil tüm sayfalar genişletilir.
VERİ ve Filo Detay sayfaları Excel olmasa bile openpyxl ile önceden genişletilir.
Başarısız olursa rapor yine oluşur; Excel'de manuel yenileyin.

## Sorun giderme

**Copy-Item: izlog-kpi-temp\KPI bulunamadı:** `$Base`, `$KpiDir`, `$Temp` tanımlanmadan sadece alt satırlar çalıştırılmış demektir. Yukarıdaki bloğu **baştan sona tek parça** yapıştırın veya `kpi_guncelle.ps1` kullanın.

**Eski rapor modu çalışıyor:** Konsolda `KPI raporu oluşturuldu` + `Tespit edilen problem` görüyorsanız kod güncellenmemiştir. Güncelleme sonrası `BAŞARILI: KPI şablon raporu` yazmalı.

**#BAŞV! / #REF! pivot hatası:** Genelde VERİ tablosunun sütunları daraltıldığında oluşur (güncel kod bunu engeller). Hâlâ görürseniz:

```powershell
python kpi_sablon_kolon_kesif.py
```

Eşleşmeyen kolonları `ayarlar.py` → `KPI_KOLON_ESLEME` ile tanımlayın.

**Şablon bulunamadı:** `KPI\referans\kpi_sablon.xlsx` dosyasını oluşturun.

**VERİ sayfası bulunamadı:** Şablondaki gizli sayfa adını `KPI_VERI_SAYFA_ADLARI` ile ayarlayın.

**ORA-00933:** Güncel KPI kodunu git clone ile alın (Oracle 11g uyumlu).

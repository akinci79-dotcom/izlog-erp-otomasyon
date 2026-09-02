# İzlog Lojistik — KPI Rapor Otomasyonu

Oracle ERP verilerini **mevcut KPI Excel şablonunuza** yapıştırır (VERİ + Filo Detay), pivot sayfalarını günceller.

**Otomasyon projesinden bağımsızdır** — yalnızca `KPI/` klasörü kullanılır.

## Hızlı kurulum (Windows)

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

## Şablonu yerleştirin

Temmuz KPI dosyanızı şuraya kopyalayın:

```
KPI\referans\kpi_sablon.xlsx
```

## ayarlar.py

```python
KPI_BASLANGIC_TARIHI = "01.07.2026"
KPI_BITIS_TARIHI = "31.07.2026"
DB_SIFRE = "..."
# KPI_SABLON_DOSYASI = "kpi_sablon.xlsx"   # varsayılan: referans/kpi_sablon.xlsx
# KPI_VERI_SAYFA_ADLARI = ["VERİ", "VERI"]
# KPI_FILO_SAYFA_ADLARI = ["Filo Detay"]
```

## Çalıştırma

```powershell
cd KPI
python kpi_rapor_olustur.py
```

**Ne yapar:**
1. `referans/kpi_sablon.xlsx` kopyalanır → `raporlar/kpi_rapor.xlsx`
2. **VERİ** sayfasına yük bazında detay Oracle sorgusu yazılır
3. **Filo Detay** sayfasına tedarikçi hesaplaşma raporu yazılır
4. Windows + Excel varsa pivotlar yenilenir ve tüm sayfalarda sütunlar içeriğe göre genişletilir (`#######` görünmez); yoksa dosyayı açıp **Verileri Yenile** + sütun başlıklarını çift tıklayarak genişletin

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

**Şablon bulunamadı:** `KPI\referans\kpi_sablon.xlsx` dosyasını oluşturun.

**VERİ sayfası bulunamadı:** Şablondaki gizli sayfa adını `KPI_VERI_SAYFA_ADLARI` ile ayarlayın.

**ORA-00933:** Güncel KPI kodunu git clone ile alın (Oracle 11g uyumlu).

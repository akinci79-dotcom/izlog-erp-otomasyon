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
KPI_DONEM = "07.2026"   # AA.YYYY — ayın kaç gün çektiğini sistem hesaplar
CO_CODE = "IZLOG"
BRANCH_CODE = "MERKEZ"
DB_SIFRE = "..."
```

`KPI_DONEM` üç formatı destekler:

| Format | Örnek | Anlamı |
|---|---|---|
| `AA.YYYY` | `"08.2026"` | Tek ay (Ağustos 2026) |
| `AA.YYYY-AA.YYYY` | `"01.2026-04.2026"` | Ay aralığı (Ocak–Nisan 2026, 4 aylık) |
| `YYYY` | `"2026"` | Tam yıl (01.01.2026 – 31.12.2026) |

Ayın/yılın kaç gün çektiğine hiç bakmanıza gerek yok, otomatik hesaplanır. Bu kalıplara uymayan özel bir aralık gerekiyorsa (`KPI_DONEM` yerine) `KPI_BASLANGIC_TARIHI` / `KPI_BITIS_TARIHI` kullanılabilir.

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
3. **Zarar Detay** sayfası (varsa) — VERİ'den o ayın zarar eden sevkleri otomatik hesaplanıp `ZararTedarikci`/`ZararKiralik` tablolarına yazılır (bkz. aşağıdaki bölüm)
4. Pivotlar yenilenir, sütunlar genişletilir

## "Zarar Detay" sayfası otomatik tazeleme

"Zarar Detay" sayfasındaki `ZararTedarikci`/`ZararKiralik` tabloları eskiden **elle** dolduruluyordu
(geçmiş ayın zarar eden sevkleri tek tek bulunup Sevk No/Müşteri/Rota bilgileri yapıştırılıyordu).
Bu tablolardaki Alış/Satış/Kâr-Zarar sütunları VERİ sayfasına (`Tablo5`) bakan formüller olduğu için,
otomasyon her ay `Tablo5`'i o ayın verisiyle sıfırdan yazınca eski ayda elle girilmiş Sevk No'lar
artık `Tablo5`'te bulunmuyor ve bu formüller sessizce **0** dönüyordu.

**Önemli:** Bu durum sadece "Zarar Detay" sayfasıyla sınırlı kalmıyor — "Özet" sayfasındaki
**YÖNETİM ALARMLARI** kutusu (`Toplam zarar büyüklüğü`, `Kiralık araç zararı`, `Tedarikçi araç zararı`,
`Zarar Eden Sevkiyat Oranı`) doğrudan bu iki tablonun toplamını okuyor; yani sayfa güncellenmeyince
o alarmlar da sessizce 0/yanlış görünüyordu.

Artık `kpi_rapor_olustur.py` her çalıştığında:
1. VERİ satırlarını **Sevk No**'ya göre gruplar (aynı sevkteki birden fazla yük satırı birleştirilir),
2. Toplam Kâr/Zarar'ı negatif olan (zarar eden) sevkleri bulur,
3. `PLAKA_MULKIYET` alanına göre **Tedarikçi**/**Kiralık** olarak ikiye ayırır, en büyük zarardan küçüğe sıralar,
4. `ZararTedarikci`/`ZararKiralik` tablolarının **mevcut satır kapasitesi içinde** üstten yazar
   (tablo boyutu büyütülüp küçültülmez — hemen altındaki "Ara Toplam" satırı ve bir sonraki bölümün
   yeri bozulmasın diye); kapasiteden fazla zarar eden sevk varsa konsolda uyarı verir (tabloyu Excel'de
   büyütmeniz gerekir).

Bu davranış `ayarlar.py` → `KPI_ZARAR_DETAY_GUNCELLE = False` ile kapatılabilir; sayfa/tablo adları
farklıysa `KPI_ZARAR_DETAY_SAYFA_ADLARI` / `KPI_ZARAR_TEDARIKCI_TABLO_ADI` / `KPI_ZARAR_KIRALIK_TABLO_ADI`
ile ayarlanabilir. Şablonda bu sayfa hiç yoksa adım sessizce atlanır.

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

### Sütun genişliğini kalıcı sabitle (her ay elle düzeltmeyin)

Bir sayfada (örn. **Özet**) genişlikleri elle mükemmel hale getirdiyseniz, bir daha
hiç değişmesin diye kalıcı yapabilirsiniz:

1. Ayarladığınız `kpi_rapor.xlsx`'i KPI klasörüne kopyalayın (veya `raporlar\` altındaysa oradan kullanın).
2. Genişlikleri okuyun:
   ```powershell
   python kpi_sutun_genislik_oku.py "raporlar\kpi_rapor.xlsx" "Özet"
   ```
3. Çıktıyı olduğu gibi `ayarlar.py`'ye yapıştırın:
   ```python
   KPI_SABIT_SUTUN_GENISLIKLERI = {
       "Özet": {"A": 14.0, "B": 32.5, "C": 28.0, ...},
   }
   ```

Bu ayar tanımlıysa o sayfada AutoFit tamamen devre dışı kalır, sütunlar her raporda
birebir bu genişliklerle açılır. Birden fazla sayfa için sözlüğe ek sayfa girebilirsiniz.

## Boş alanlara otomatik varsayılan değer (pivot'ta "(boş)" kategorisini önler)

Personel ERP'ye veri girerken bazı alanları atlayabiliyor; bu durumda pivot tablolarda
"(boş)" diye bir kategori oluşuyor. Otomasyon, VERİ Oracle'dan çekildikten SONRA (SQL'e
dokunmadan) şu iki alanı varsayılan bir değerle dolduruyor:

- **MÜLKİYET** (VERİ sayfası Z sütunu) boşsa → **Tedarikçi**
- **Proje Kodu = "Konya"** VE **Yük Fiyat Tipi Kodu** (VERİ sayfası K sütunu) boşsa → **ŞARKÜTERİ**

`ayarlar.py` içinde ayarlanabilir:

```python
KPI_BOS_ALAN_VARSAYILARI = True                       # tamamen kapatmak için False
KPI_MULKIYET_BOS_VARSAYILAN = "Tedarikçi"
KPI_KONYA_YUK_FIYAT_TIPI_BOS_VARSAYILAN = "ŞARKÜTERİ"
```

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

**"Zarar Detay" sayfasındaki rakamlar / Özet'teki "Toplam zarar büyüklüğü" hep 0:** Bu, eski
(elle doldurulan) tasarımın bilinen bir sorunuydu — bkz. yukarıdaki "Zarar Detay sayfası otomatik
tazeleme" bölümü. Güncel koddan sonra hâlâ 0 görüyorsanız: (1) `KPI_ZARAR_DETAY_GUNCELLE` yanlışlıkla
`False` yapılmış olabilir, (2) konsolda `Zarar Detay güncellendi: ...` satırını arayın — hiç
görünmüyorsa sayfa adı eşleşmiyordur (`KPI_ZARAR_DETAY_SAYFA_ADLARI`'nı kontrol edin), (3) "tabloda X
satırlık yer var" uyarısı varsa `ZararTedarikci`/`ZararKiralik` tablosunun satır kapasitesi bu ayki
zarar eden sevk sayısına yetmiyor demektir — Excel'de tabloyu (Ara Toplam satırından önce) birkaç yüz
satır büyütüp tekrar deneyin.

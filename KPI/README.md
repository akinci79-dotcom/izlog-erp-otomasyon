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
1. `referans/kpi_sablon.xlsx` kopyalanır → `raporlar/<Dönem> İzlog Lojistik Raporları.xlsx` (örn. tek ay için
   `Ağustos 2026 İzlog Lojistik Raporları.xlsx`, tam yıl için `2026 İzlog Lojistik Raporları.xlsx`, aksi
   (ay aralığı) durumda `01.01.2026 – 30.04.2026 İzlog Lojistik Raporları.xlsx` — şablonun uzantısı `.xlsm`
   ise çıktı da `.xlsm` olur). Sabit bir isim istiyorsanız `ayarlar.py`'ye `KPI_RAPOR_DOSYASI = "..."` ekleyin
   (bu, otomatik adlandırmayı devre dışı bırakır); dönem etiketinin sonundaki metni değiştirmek isterseniz
   `KPI_RAPOR_ADI_SONEKI = "..."` ile özelleştirebilirsiniz (varsayılan: `"İzlog Lojistik Raporları"`).
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
4. `ZararTedarikci`/`ZararKiralik` tablolarına üstten yazar — **tablonun satır kapasitesi yetersizse
   otomatik olarak büyütülür** (aşağıdaki "Tablo kapasitesi artık otomatik büyüyor" bölümüne bakın),
   yani zarar eden sevk sayısı ay ay değişse bile elle bir şey yapmanız gerekmez.

Bu davranış `ayarlar.py` → `KPI_ZARAR_DETAY_GUNCELLE = False` ile kapatılabilir; sayfa/tablo adları
farklıysa `KPI_ZARAR_DETAY_SAYFA_ADLARI` / `KPI_ZARAR_TEDARIKCI_TABLO_ADI` / `KPI_ZARAR_KIRALIK_TABLO_ADI`
ile ayarlanabilir. Şablonda bu sayfa hiç yoksa adım sessizce atlanır.

### Tablo kapasitesi artık otomatik büyüyor (elle satır eklemeniz gerekmiyor)

Zarar eden sevk sayısı her ay farklı olacağı için (bir ay 89, başka bir ay 300 olabilir), tabloyu
Excel'de bir kere elle büyütmek kalıcı bir çözüm değildi. Artık `ZararTedarikci`/`ZararKiralik`
tablolarının satır kapasitesi o ayki zarar eden sevk sayısına yetmezse, kod eksik kadar satırı
**gerçek bir Excel satır ekleme işlemiyle** tablonun içine (mevcut son veri satırının tam üzerine)
otomatik olarak ekler.

**Neden gerçek bir PivotTable değil de bu yöntem?** PivotTable'lar satır sayısına göre doğal olarak
büyür/küçülür, ama sıfırdan bir PivotCache/PivotTable kurmak (alan yerleşimi, sadece zarar eden
(negatif) sevkleri gösteren filtre, gizli bir ham veri alanının senkronizasyonu) çok daha karmaşık
bir COM inşası gerektiriyor ve mevcut şablondaki elle ayarlanmış görünümü/biçimlendirmeyi bozma
riski taşıyor — üstelik bu sandboxta gerçek Excel'de hiç doğrulanamaz. Bunun yerine mevcut sabit
tablo + "Ara Toplam" formülü tasarımı korunuyor, sadece güvenilirliği artırıldı (aşağıya bakın).

**Ara Toplam formülü artık TAHMİN EDİLMİYOR, DOĞRULANIYOR:** Satır ekleme sonrası Excel'in "bir
aralığın içine satır eklenirse ona bakan formüller otomatik genişler" davranışına güvenmek yerine
(bu davranış bu sandboxta gerçek Excel'de doğrulanamadığı için), kod artık satır eklendikten SONRA
hemen altındaki "Ara Toplam" satırının `SUM(...)` formülünü GERÇEKTEN OKUYUP yeni veri aralığını
kapsayıp kapsamadığını kontrol ediyor; kapsamıyorsa formülün SADECE satır numarasını (formülün geri
kalanını olduğu gibi koruyarak) açıkça düzeltiyor. Yani sonuç artık "Excel'in umulan davranışına"
değil, kodun kendi doğrulama adımına dayanıyor — hem "Excel otomatik genişletti" hem "genişletmedi"
senaryosunda doğru sonucu garanti eder. Yeni eklenen satırların Alış/Satış/Kâr-Zarar/Zarar %/Zarar
Payı (J-N) formülleri de tablonun ilk veri satırından açıkça kopyalanır (Excel'in kendiliğinden
kopyalamasına güvenilmez).

Konsolda böyle bir büyütme olduğunda şu satırları görürsünüz:

```
[Excel] ZararTedarikci kapasitesi yetersiz (77 satır var, 89 gerekiyor) — 12 satır otomatik ekleniyor...
[Excel] ZararTedarikci kapasitesi 89 satıra büyütüldü ('Ara Toplam' formülü doğrulandı/gerekirse düzeltildi).
```

Formülün açıkça düzeltilmesi gerekirse (yani Excel kendiliğinden genişletmediyse) ayrıca şu satırı
da görürsünüz — bu bir HATA DEĞİL, sistemin tam olarak tasarlandığı gibi (doğrulayıp gerekirse
düzelterek) çalıştığının kanıtıdır:

```
[Excel] ZararTedarikci 'Ara Toplam' J sütunu Excel tarafından otomatik genişletilmemişti — açıkça düzeltildi: =SUM(J2:J77) → =SUM(J2:J89)
```

Otomatik büyütme (çok nadir — ör. korumalı sayfa, birleştirilmiş hücre gibi beklenmedik bir Excel
kısıtlaması) başarısız olursa, eski davranışa (kapasiteyi aşan sevkler gösterilmez, konsolda ve
raporun uyarı mesajında bu belirtilir) geri dönülür; rapor yine de başarıyla tamamlanır. Bu durumda
konsol logundaki hata mesajına bakıp tabloyu Excel'de elle büyütmeniz gerekebilir — ama normal
şartlarda bu hiç gerekmemeli.

### VERİ / Filo Detay sayfaları da artık kapasiteye göre otomatik büyüyor (altındaki içerik korunur)

**Geçmişte olan sorun:** VERİ ve Filo Detay sayfalarına yazan `_com_sayfaya_yaz` fonksiyonu, tablonun
sınırını `ListObject.Resize` ile doğrudan büyütüyordu. `Resize` gerçek bir satır **eklemez/kaydırmaz** —
sadece tablonun kapladığı alanı yeniden tanımlar. Bu ay gelen araç/sevk satırı sayısı önceki ay
kaydedilmiş kapasiteyi aşarsa (ör. Filo Detay'da 30 → 38 araca çıkması), tablonun eski sınırının
**hemen altında** fiziksel olarak var olan herhangi bir içerik (ör. şablonda elle konmuş bir
**"Genel Toplam"** satırı) tabloya "yutuluyor" ve hemen ardından gerçek veriyle **üzerine yazılıyordu**
— bu, "Filo Detay sayfasını da bozmuş, genişletmek yerine toplam satırına bilgi basmış" şeklinde
bildirilen sorunun kök nedeniydi.

**Düzeltme:** Artık `_com_sayfaya_yaz`, `Resize` çağırmadan ÖNCE yazılacak satır sayısını
(`ListObject.ListRows.Count` — native Toplam Satırını saymaz) mevcut kapasiteyle karşılaştırıyor;
kapasite yetersizse (Zarar Detay'daki KANITLANMIŞ desenle **paylaşılan** `_com_tablo_satir_ekle`
yardımcısıyla) eksik kadar satırı gerçek bir Excel satır ekleme işlemiyle (`Rows.Insert`, başarısız
olursa `ListRows.Add(AlwaysInsert=True)` yedeği) tablonun mevcut son veri satırının tam üzerine ekler.
Bu, tablonun **altındaki** her ne varsa (Genel Toplam satırı dahil) **kaybolmadan aşağı kaymasını**
sağlar; ancak SONRA `Resize` + toplu veri yazımı yapılır, artık gerçekten "boş" satırların üzerine.

Bu mantık VERİ sayfası için de otomatik uygulanır (aynı fonksiyonu paylaştığı için) — VERİ'nin altında
korunması gereken bir içerik yoksa satır ekleme zararsızdır (sadece kapasite gerçekten yetersiz
kaldığında devreye girer). Tablo **küçülürse** (bu ay geçen aydan az satır) bu mantık hiç tetiklenmez,
eski `Resize` davranışı değişmeden kalır.

Konsolda böyle bir büyütme olduğunda (Zarar Detay'dakine benzer) şu satırı görürsünüz, bu normaldir:

```
[Excel] Filo Detay: tablo kapasitesi yetersiz (30 satır var, 38 gerekiyor) — 8 satır otomatik ekleniyor (altındaki içerik korunacak)...
```

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

1. Ayarladığınız rapor dosyasını (`raporlar\` klasöründeki, artık dönem adıyla üretilen `.xlsx`) kullanın.
2. Genişlikleri okuyun (dosya adını kendi çıktınıza göre değiştirin):
   ```powershell
   python kpi_sutun_genislik_oku.py "raporlar\Ağustos 2026 İzlog Lojistik Raporları.xlsx" "Özet"
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
dokunmadan) şu alanları varsayılan bir değerle dolduruyor:

- **MÜLKİYET** (VERİ sayfası Z sütunu) boşsa → **Tedarikçi**
- **Proje Kodu = "Konya"** VE **Yük Fiyat Tipi Kodu** (VERİ sayfası K sütunu) boşsa → **ŞARKÜTERİ**
- **Araç Tipi** (VERİ sayfası X sütunu) boşsa → **Tır Frigorifik**

`ayarlar.py` içinde ayarlanabilir:

```python
KPI_BOS_ALAN_VARSAYILARI = True                       # tamamen kapatmak için False
KPI_MULKIYET_BOS_VARSAYILAN = "Tedarikçi"
KPI_KONYA_YUK_FIYAT_TIPI_BOS_VARSAYILAN = "ŞARKÜTERİ"
KPI_ARAC_TIPI_BOS_VARSAYILAN = "Tır Frigorifik"
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

**`[WinError 32] The process cannot access the file because it is being used by another process`:**
`raporlar\` klasöründeki çıktı dosyası (veya şablon) başka bir işlem tarafından açık tutuluyor. Kod artık
bunu birkaç kez kısa aralıkla otomatik tekrar deniyor; hâlâ başarısızsa: (1) çıktı dosyası Excel'de açıksa kapatın,
(2) Görev Yöneticisi'nde (`Ctrl+Shift+Esc`) arkada kalmış bir `EXCEL.EXE` süreci varsa (görünür pencere
olmasa bile — Excel COM otomasyonu gizli/`Visible=False` çalışır, önceki bir çalıştırma çökmüşse arkada
kalabilir) sonlandırın, (3) `python kpi_rapor_olustur.py`'yi tekrar çalıştırın.

**Konsol "`[Excel] Zarar Detay güncelleniyor...`" satırında uzun süre donuyor/takılıyor gibi
görünüyor:** Bu, eski koddaki bir performans hatasıydı — Zarar Detay tablosundaki her hücreye TEK TEK
(satır satır, hücre hücre) yazılıyordu; her biri ayrı bir Excel COM çağrısı olduğu için yüzlerce satırlık
bir tabloda bu binlerce ayrı çağrıya (ve dakikalarca sürebilen bir beklemeye) yol açabiliyordu. Kod artık
diğer sayfalarda (VERİ/Filo) olduğu gibi tüm bloğu TEK bir toplu (bulk) yazma işlemiyle dolduruyor — bu
adım artık saniyeler içinde bitmeli. Güncel kodu çektiğiniz halde hâlâ uzun sürüyorsa, muhtemelen normal
Excel/Oracle gecikmesidir; sabırla bekleyin veya konsoldaki en son satırın hangi adımda kaldığına bakın.

**"Zarar Detay" sayfasındaki rakamlar / Özet'teki "Toplam zarar büyüklüğü" hep 0:** Bu, eski
(elle doldurulan) tasarımın bilinen bir sorunuydu — bkz. yukarıdaki "Zarar Detay sayfası otomatik
tazeleme" bölümü. Güncel koddan sonra hâlâ 0 görüyorsanız: (1) `KPI_ZARAR_DETAY_GUNCELLE` yanlışlıkla
`False` yapılmış olabilir, (2) konsolda `Zarar Detay güncellendi: ...` satırını arayın — hiç
görünmüyorsa sayfa adı eşleşmiyordur (`KPI_ZARAR_DETAY_SAYFA_ADLARI`'nı kontrol edin).

**"...otomatik satır ekleme denendi ama X tanesi yine de sığmadı" uyarısı / "Range sınıfının Insert
yöntemi başarısız" hatası:** Tablo kapasitesi artık otomatik büyütülüyor (bkz. yukarıdaki "Tablo
kapasitesi artık otomatik büyüyor" bölümü); bu uyarı SADECE otomatik büyütmenin kendisi başarısız
olduğunda çıkar. **Bilinen kök neden** [WebSearch ile teyit edildi]: ZararTedarikci/ZararKiralik
tablolarından biri Excel'in NATİF "Toplam Satırı" (Table Style Options → **Total Row**,
`ListObject.ShowTotals`) özelliğiyle kurulmuş olabilir — bu AÇIKKEN tablonun kendi "Ara Toplam" satırı
aslında bu native Toplam Satırıdır. Kod artık satır sayısını `ListObject.ListRows.Count` ile hesaplıyor
(bu, native Toplam Satırını HİÇ saymaz), bu yüzden yeni satırlar her zaman GERÇEK son veri satırının
üzerine ekleniyor, Toplam Satırının kendisine değil — "Insert method of Range class failed" hatasının
asıl nedeni buydu (bir tablonun Toplam Satırının üzerine/içine satır eklenemez). Ayrıca bir güvenlik ağı
var: `Rows.Insert` yine de başarısız olursa kod otomatik olarak `ListObject.ListRows.Add(AlwaysInsert=True)`
(UI'daki "Tablo Satırlarını Üstte Ekle" ile birebir aynı davranış) yöntemine geçer. Bu ikisi de
başarısız olursa (çok nadir — ör. sayfa korumalı, tabloda birleştirilmiş hücre var), eski davranışa
(kapasiteyi aşan sevkler gösterilmez, konsolda ve raporun uyarı mesajında bu belirtilir) geri dönülür.
Konsol logunda `kapasitesi otomatik büyütülemedi (...)` satırını arayıp asıl hatayı görün; gerekirse
tabloyu (Ara Toplam/Toplam Satırından önce) Excel'de elle büyütüp tekrar deneyin.

**VERİ / Zarar Detay sayfasındaki tarih sütunları yanlış görünüyor (ör. "08.mm.2026" gibi):** Bu,
hücrenin Excel'de önceden "Metin" ya da bozuk bir özel tarih biçimiyle kilitli kalmasından
kaynaklanıyordu. Güncel kod artık (1) hedef hücreleri veri yazılmadan ÖNCE "General"e sıfırlıyor,
(2) float Excel seri numarasını yazıyor, (3) hem `NumberFormat` ("dd.mm.yyyy") hem
`NumberFormatLocal` ("gg.aa.yyyy") olarak biçimi uyguluyor ve sonucu doğruluyor. Rapor sonunda hâlâ
bozuk görünüyorsa konsoldaki `Uyarı: ... tarih biçimi doğrulanamadı` satırını arayın — hangi kolonun
sorunlu olduğunu doğrudan gösterir.

**"Uyarı: ... tarih biçimi doğrulanamadı" — TÜM tarih sütunları AYNI ANDA başarısız görünüyor
(veri doğru ama uyarı çıkıyor):** Bu, [WebSearch ile teyit edilen] bilinen bir Excel COM
tuhaflığından kaynaklanan bir **YANLIŞ ALARM** idi, artık düzeltildi. Microsoft'un dokümantasyonu
`.NumberFormat`'ın locale-bağımsız olacağını (her zaman `dd`/`mm`/`yyyy` gibi İngilizce kodlarla
döneceğini) söylüyor, ama harici bir COM istemcisinden (bu projedeki Python/pywin32 gibi) sürülen
otomasyonda, özellikle Excel arayüz dili Türkçe olduğunda, bu okuma bazen Türkçeleştirilmiş bir
karşılığı (`gg`/`aa`/`yyyy`) döndürüyor. Eski doğrulama SADECE literal `"dd.mm.yyyy"` dizisine tam
eşitlik arıyordu; format GERÇEKTE doğru uygulanmış olsa bile (Excel'de doğru görünüyor) bu Türkçe
varyant döndüğünde HER ZAMAN "başarısız" sanılıyordu — 4 farklı tarih sütununun HEPSİNİN TEK SEFERDE
başarısız olması da (veri/hücre bazlı bir sorun değil, doğrulama mantığı hatası olduğu için) bu
teoriyle uyumluydu. Doğrulama artık `.NumberFormat`/`.NumberFormatLocal`'ın "gün.ay.yıl" YAPISINA
uyup uymadığını kontrol ediyor (İngilizce **veya** Türkçe kod harfleri, `.`/`/`/`-` ayraçları, kaçışlı
ayraçlar hepsi kabul ediliyor) — ama bilinen GERÇEKTEN bozuk `mm/\m\m/yyyy` kalıbı (gün bilgisi
tamamen kayıp, ekranda "08.mm.2026" gibi görünen) hâlâ YAKALANIYOR (bu kalıp gün bileşeniyle
BAŞLAMADIĞI için yapısal kontrolden geçemiyor). Güncel kodu çektiğiniz halde bu uyarı hâlâ çıkıyorsa,
bu artık gerçek bir sorunun işaretidir — hücreyi Excel'de elle kontrol edin (bkz. yukarıdaki madde).

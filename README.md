# İzlog Lojistik — Uyumsoft ERP Yük/Sevk Otomasyonu

Oracle'dan okunan "yük" verilerini, Excel listesi (`islem_listesi.xlsx`) üzerinden
sırayla işleyip Uyumsoft ERP arayüzünde **Playwright** ile:

1. Kaynak yükü bulur,
2. Kopyalayıp yeni yük olarak kaydeder (tarih, satış fiyatları, fatura bağlantıları),
3. Yeni yük üzerinden sevk (nakliye) kaydı oluşturur.

İşlem durumu her satır için Excel'e (checkpoint olarak) yazılır; script yarıda
kesilse veya hata alsa bile bir sonraki çalıştırmada kalınan yerden devam eder.

> ⚠️ Bu otomasyon **Windows sunucusunda** çalışır (Oracle Instant Client Thick
> Mode yolu `ayarlar.py`/`oracle_okuyucu.py` içinde `C:\instantclient\...`
> olarak sabit). Linux/macOS'ta doğrudan çalışmaz.

## Dosyalar

| Dosya | Görevi |
|---|---|
| `izlog_yuk_otomasyon.py` | Ana otomasyon (Playwright + Excel döngüsü) |
| `ayarlar.py` | ERP/Oracle bağlantı bilgileri, `DRY_RUN` anahtarı |
| `oracle_okuyucu.py` | Oracle'dan yük/satış/fatura verisini çeken ve başarılı kayıtların "iz"ini temizleyen katman |
| `excel_olustur.py` | `islem_listesi.xlsx` şablonunu sıfırdan oluşturur |
| `fatura_bul.py` | Oracle şemasında fatura tablolarını keşfetmek için tek seferlik yardımcı script |

## Kurulum (Windows sunucusu)

```powershell
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
playwright install chromium
```

Oracle Instant Client'ın `ayarlar.py`/`oracle_okuyucu.py` içinde belirtilen yolda
(`C:\instantclient\instantclient_19_32`) kurulu olması gerekir.

`ayarlar.py` içindeki `DB_SIFRE` ve `ERP_SIFRE` alanlarını gerçek şifrelerinizle
doldurun. **Bu dosyayı gerçek şifrelerle git'e commit etmeyin.**

## Excel şablonu

```bash
python excel_olustur.py
```

`islem_listesi.xlsx` şu kolonlarla oluşturulur:

| Sütun | İçerik |
|---|---|
| A | KAYNAK_YUK_NO |
| B | PLAKA |
| C | SEVK_ALIS_FIYATI |
| D | PROJE_KODU (otomatik doldurulur) |
| E | TARIH (otomatik doldurulur) |
| F | FATURA_NO (otomatik doldurulur) |
| G | FATURA_TARIHI (otomatik doldurulur) |
| H | YENI_YUK_NO (otomatik / checkpoint) |
| I | YENI_SEVK_NO (otomatik doldurulur) |
| J | DURUM: `BAŞARILI`, `DRY_RUN BAŞARILI`, `YÜK OLUŞTU`, `HATA_YUK`, `HATA_SEVK`, `HATA_BİLİNMEYEN` |
| K | HATA_ACIKLAMASI |

İlk 3 sütuna test verisi girin (örn. `Y-575733`, `06DFZ463`, `73000`).

## Çalıştırma

```powershell
python izlog_yuk_otomasyon.py
```

- **Önce `ayarlar.py` içinde `DRY_RUN = True` ile test edin.** Bu modda script
  sadece ERP'ye giriş yapar, yükü arar ve satırı seçer; `Kopya`, `Kaydet` veya
  `Sevk Oluştur` butonlarına **kesinlikle basmaz** — hiçbir kayıt/veritabanı
  değişikliği yapılmaz.
- Gerçek kayıt/veritabanı işlemleri için `DRY_RUN = False` yapın.
- Hata durumunda ilgili satırın ekran görüntüsü (`hata_<yükno>_<saat>.png`)
  alınır ve hata mesajı Excel'e yazılır; otomasyon güvenli şekilde bir
  sonraki satıra geçer.
- Başarılı Yük/Sevk kayıtları sonunda `oracle_okuyucu.yeni_kayitlari_veritabaninda_guncelle()`
  ile veritabanında "iz temizliği" (create/update kullanıcı ve tarih alanlarının
  Uyumsoft standardına sıfırlanması) yapılır. `DRY_RUN` aktifken bu adım da atlanır.

## Bu sürümde yapılan düzeltme

Fatura seçim penceresinde (3 nokta ile açılan LOV), doğru satırı bulmak için
tutarı grid formatına (`78.279,00`) birebir string eşleştirmesiyle (`has_text`)
aramak kırılgandı — ERP'nin gösterdiği format ile üretilen string arasında ufak
bir fark olduğunda satır bulunamıyor ve akış hata veriyordu. Bu sürümde:

- Fatura no'ya göre (kelime sınırı ile) filtrelenen satırlar arasında, tutar
  artık **satır metninden ayrıştırılıp** (`Decimal`, kuruş toleranslı) gerçek
  tutarla karşılaştırılıyor — format farklılıklarına karşı dayanıklı.
- Aynı fatura no + aynı tutara sahip birden fazla satır olsa da sorun değil:
  ilk eşleşen satır seçiliyor.
- Eşleşme bulunamazsa, incelenen satırların metnini içeren açıklayıcı bir hata
  fırlatılıyor (teşhis için).
- `wait_for_load_state("networkidle")` çağrıları, ERP'nin sürekli arka plan
  isteği attığı durumlarda sonsuz beklemeye/timeout'a düşmesin diye güvenli bir
  yedek beklemeyle sarmalandı (`_agsakinligini_bekle`).

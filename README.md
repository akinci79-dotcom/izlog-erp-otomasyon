# Uyumsoft ERP Yük/Sevk Otomasyonu

Oracle'dan okunan "yük" verilerini, Excel listesi (`islem_listesi.xlsx`) üzerinden
sırayla işleyip Uyumsoft ERP arayüzünde **Playwright** ile:

1. Kaynak yükü bulur,
2. Kopyalayıp yeni yük olarak kaydeder (tarih, satış fiyatları, fatura bağlantıları),
3. Yeni yük üzerinden sevk (nakliye) kaydı oluşturur.

İşlem durumu her satır için Excel'e (checkpoint olarak) yazılır; script yarıda
kesilse veya hata alsa bile bir sonraki çalıştırmada kalınan yerden devam eder.

## Kurulum

```bash
python -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate
pip install -r requirements.txt
playwright install chromium
```

`.env.example` dosyasını `.env` olarak kopyalayıp ERP ve Oracle bilgilerinizi girin:

```bash
cp .env.example .env
```

> `oracle_okuyucu.py` bir **şablondur**: `TODO` işaretli sorguları kendi Oracle
> şemanıza (tablo/kolon adları) göre doldurmanız gerekir.

## Excel formatı (`islem_listesi.xlsx`)

| Sütun | İçerik |
|---|---|
| A | Kaynak Yük No |
| B | Plaka |
| C | Sevk Alış Fiyatı |
| D | Proje (otomatik doldurulur) |
| E | Yük Tarihi (otomatik doldurulur) |
| F | Fatura No (otomatik doldurulur) |
| G | Fatura Tarihi (otomatik doldurulur) |
| H | Yeni Yük No (otomatik doldurulur / checkpoint) |
| I | Yeni Sevk No (otomatik doldurulur) |
| J | Durum (`BAŞARILI`, `YÜK OLUŞTU`, `HATA_YUK`, `HATA_SEVK`, `DRY_RUN BAŞARILI`, ...) |
| K | Hata Mesajı |

## Çalıştırma

```bash
python main.py
```

- `DRY_RUN=1` ortam değişkeni ile hiçbir kayıt yapılmadan sadece doğrulama yapılabilir.
- `HEADLESS=1` ile tarayıcı görünmez modda çalışır (ilk denemelerde `0` önerilir).
- Hata durumunda ilgili satırın ekran görüntüsü (`hata_<yükno>_<saat>.png`) alınır
  ve hata mesajı Excel'e yazılır; otomasyon güvenli şekilde bir sonraki satıra geçer.

## Bu sürümde yapılan düzeltme

Önceki sürümde, fatura seçim penceresinde (3 nokta ile açılan LOV) tutarı arama/filtre
kutusuna (`#myListPage_DXFREditorcol6_I`) yazmaya çalışmak, bu alanın DevExpress'e özel
maskeli/formatlı bir sayısal editör olması nedeniyle hataya yol açıyordu. Bu sürümde:

- Tutar hiçbir zaman bu kutuya yazılmıyor; sadece fatura no ile filtreleniyor.
- Filtrelenen satırların metni Python tarafında ayrıştırılıp (`Decimal` ile) gerçek
  tutarla karşılaştırılarak doğru satır bulunuyor (biçim farklılıklarına dayanıklı).
- Eşleşme bulunamazsa, hangi satırların incelendiğini gösteren açıklayıcı bir hata
  fırlatılıyor (teşhis kolaylığı için).

Diğer küçük sağlamlaştırmalar: tarih alanı temizleme kodu tekilleştirildi, `wait_for_function`
içindeki değer enjeksiyonu güvenli hale getirildi, eksik/`None` fiyat verileri için
anlamlı hata mesajları eklendi, `browser.close()` `try/finally` ile garantiye alındı,
kimlik bilgileri ortam değişkenlerine taşındı.

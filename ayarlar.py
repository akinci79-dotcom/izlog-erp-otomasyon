"""
Ortam ayarları / konfigürasyon.

Tüm hassas bilgiler (kullanıcı adı, şifre, bağlantı bilgileri) kaynak koda
GÖMÜLMEZ; ortam değişkenlerinden (.env dosyası veya sistem env) okunur.

Kullanım:
    1. `.env.example` dosyasını `.env` olarak kopyalayın.
    2. Kendi ERP / Oracle bilgilerinizi `.env` içine yazın.
    3. `.env` dosyasını ASLA git'e eklemeyin (.gitignore içinde zaten hariç tutuluyor).
"""
import os

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    # python-dotenv kurulu değilse sorun değil; sistem ortam değişkenleri
    # zaten os.environ üzerinden okunabilir.
    pass


def _zorunlu(anahtar, varsayilan=None):
    deger = os.environ.get(anahtar, varsayilan)
    if deger is None:
        raise RuntimeError(
            f"HATA: '{anahtar}' ortam değişkeni tanımlı değil. "
            f".env dosyanızı kontrol edin (bkz. .env.example)."
        )
    return deger


# --- ERP Bağlantı Bilgileri ---
ERP_LOGIN_URL = _zorunlu("ERP_LOGIN_URL", "https://erp.ornek-firma.com/Login.aspx")
ERP_YUK_LISTESI_URL = _zorunlu("ERP_YUK_LISTESI_URL", "https://erp.ornek-firma.com/LGoods/List.aspx")
ERP_KULLANICI = _zorunlu("ERP_KULLANICI", "kullanici.adi")
ERP_SIFRE = _zorunlu("ERP_SIFRE", "degistiriniz")

# --- Oracle Bağlantı Bilgileri (oracle_okuyucu.py tarafından kullanılır) ---
ORACLE_DSN = os.environ.get("ORACLE_DSN", "localhost:1521/ORCLPDB1")
ORACLE_KULLANICI = os.environ.get("ORACLE_KULLANICI", "oracle_kullanici")
ORACLE_SIFRE = os.environ.get("ORACLE_SIFRE", "degistiriniz")

# --- Çalışma Modu ---
# DRY_RUN=1 -> Hiçbir kayıt/kopyalama yapılmaz, sadece veri doğrulaması yapılır.
DRY_RUN = os.environ.get("DRY_RUN", "0") == "1"

# HEADLESS=1 -> Tarayıcı görünmez modda çalışır (sunucu/CI ortamları için).
# Not: ERP'nin bazı DevExpress bileşenleri headless modda farklı davranabilir;
# ilk denemelerde HEADLESS=0 (görünür) kullanmanız önerilir.
HEADLESS = os.environ.get("HEADLESS", "0") == "1"

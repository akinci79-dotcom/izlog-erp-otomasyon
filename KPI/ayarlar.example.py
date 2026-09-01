# ==========================================
# İZLOG KPI MODÜLÜ — AYARLAR (otomasyondan bağımsız)
# ==========================================
#
# Bu dosya yalnızca KPI/ klasöründeki scriptler tarafından kullanılır.
# Üst klasördeki otomasyon ayarlarıyla (ayarlar.py) HİÇBİR bağlantısı yoktur.
#
# Kurulum: copy ayarlar.example.py ayarlar.py  →  şifreleri doldurun.

# Oracle bağlantı
DB_KULLANICI = "uyumsoft"
DB_SIFRE = "DEĞİŞTİRİNİZ"
DB_DSN = "172.17.8.11:1521/UYUMSOFT"
ORACLE_CLIENT_LIB_DIR = r"C:\instantclient\instantclient_19_32"

# KPI rapor dönemi (DD.MM.YYYY)
KPI_BASLANGIC_TARIHI = "01.01.2026"
KPI_BITIS_TARIHI = "31.01.2026"

# Çıktı dosyası — raporlar/ alt klasöründe oluşur
KPI_RAPOR_DOSYASI = "kpi_rapor.xlsx"

# Opsiyonel: sevk tablosunda yük eşleme kolonu (otomatik bulunamazsa elle girin)
# KPI_SEVK_YUK_KOLONU = "L_GOODS_ID"

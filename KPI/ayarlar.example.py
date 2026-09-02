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

# Referans KPI şablonu (Temmuz raporunuzun kopyası)
KPI_SABLON_DOSYASI = "kpi_sablon.xlsx"

# VERİ sayfası SQL — boş bırakılırsa referans/kpi_veri_rapor.sql aranır, yoksa varsayılan sorgu
# Uyumsoft LojistikYükSevkKalemRaporu SQL'ini referans/kpi_veri_rapor.sql olarak kaydedin
# KPI_VERI_SQL_DOSYASI = "kpi_veri_rapor.sql"

# Şablondaki sayfa adları (gizli VERİ sayfası dahil)
KPI_VERI_SAYFA_ADLARI = ["VERİ", "VERI", "Veri"]
KPI_FILO_SAYFA_ADLARI = ["Filo Detay", "Filo detay"]

# VERİ / Filo Detay başlık satırı (genelde 1)
KPI_VERI_BASLIK_SATIRI = 1
KPI_FILO_BASLIK_SATIRI = 1

# Windows'ta Excel ile veri yazma + pivot yenileme + sütun AutoFit (pywin32 + Excel gerekir)
KPI_EXCEL_KULLAN = True
KPI_PIVOT_YENILE = True
KPI_SUTUN_AUTOFIT = True
# Hızlı mod: CalculateFullRebuild yerine Calculate, VERİ/Filo AutoFit atlanır (varsayılan: True)
KPI_HIZLI_MOD = True
# Pivot özet sayfalarında tutar sütunları (Alış/Satış) — #### önleme
KPI_PARA_SUTUN_MIN_GENISLIK = 24
# Özet sayfasında Alış sütunu (genelde C)
KPI_ALIS_SUTUN_HARFI = "C"
KPI_ALIS_SUTUN_GENISLIK = 28
# İsteğe bağlı: KPI_SUTUN_GENISLIK = {"D": 26, "E": 26}

# Çıktı dosyası — boş bırakılırsa şablon uzantısı kullanılır (.xlsx veya .xlsm)
# KPI_RAPOR_DOSYASI = "kpi_rapor.xlsx"

# Şablon başlığı ↔ Oracle kolon eşlemesi (gerekirse)
# KPI_KOLON_ESLEME = {"Yük No": "YUK_NO", "Satış Tutar": "SATIS_TUTAR"}

# Opsiyonel firma/şube filtresi (Uyumsoft yük detay raporu ile aynı mantık)
# Boş bırakılırsa tüm firma/şubeler dahil edilir.
# CO_CODE = "IZLOG"
# BRANCH_CODE = "MERKEZ"

# Kapıdan kapıya yükleri hariç tut (rapor SQL'inde IS_DOOR_TO_DOOR = 0)
KPI_KAPI_KAPI_HARIC = True

# Fatura beklenmeyen operasyon kodları (faturasız problem sayımına dahil edilmez)
# KPI_FATURA_MUAF_OPERASYONLAR = ["BAŞKA_KOD"]

# Kalem detay sayfası satır üst sınırı (sevk + yük kalemleri birlikte)
KPI_KALEM_DETAY_LIMIT = 10000

# Fatura detay sayfası satır üst sınırı
KPI_FATURA_DETAY_LIMIT = 10000

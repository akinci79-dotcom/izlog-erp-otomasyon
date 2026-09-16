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

# KPI rapor dönemi — ayın/yılın kaç gün çektiğini sistem otomatik hesaplar.
# Üç format desteklenir:
#   KPI_DONEM = "01.2026"             -> tek ay (Ocak 2026)
#   KPI_DONEM = "01.2026-04.2026"     -> ay aralığı (Ocak-Nisan 2026, 4 ay)
#   KPI_DONEM = "2026"                -> tam yıl (01.01.2026 - 31.12.2026)
KPI_DONEM = "01.2026"

# Ay/yıl kalıbına uymayan özel bir tarih aralığı gerekiyorsa (örn. ay ortası
# kesim) KPI_DONEM yerine bu ikisini kullanın — ikisi de tanımlıysa KPI_DONEM
# önceliklidir:
# KPI_BASLANGIC_TARIHI = "01.01.2026"
# KPI_BITIS_TARIHI = "31.01.2026"

# Referans KPI şablonu (Temmuz raporunuzun kopyası)
KPI_SABLON_DOSYASI = "kpi_sablon.xlsx"

# VERİ sayfası SQL — Uyumsoft rapor SQL'iniz (değiştirilmeden çalıştırılır)
# KPI/referans/kpi_veri_rapor.sql dosyasına kaydedin
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

# Belirli sayfalarda sütun genişliklerini SABİT tutun (AutoFit tamamen atlanır).
# Excel'de elle ayarladığınız genişlikleri buraya yazarsanız her ay aynı kalır.
# Sayfa adı: {Sütun Harfi: Genişlik}
# KPI_SABIT_SUTUN_GENISLIKLERI = {
#     "Özet": {"A": 14, "B": 32, "C": 28, "D": 18},
# }

# "Zarar Detay" sayfası (ZararTedarikci/ZararKiralik tabloları) — VERİ'den
# hesaplanan zarar eden sevkleri her ay otomatik tazeler. Bu sayfa yoksa
# (şablonunuzda tanımlı değilse) sessizce atlanır. "Özet" sayfasındaki
# YÖNETİM ALARMLARI kutusu (Toplam zarar büyüklüğü vb.) bu tablolara bağlı
# olduğu için kapatmanız önerilmez — sadece test amaçlı False yapın.
KPI_ZARAR_DETAY_GUNCELLE = True
# Sayfa/tablo adları farklıysa (varsayılan: "Zarar Detay" / "ZararTedarikci" / "ZararKiralik")
# KPI_ZARAR_DETAY_SAYFA_ADLARI = ["Zarar Detay"]
# KPI_ZARAR_TEDARIKCI_TABLO_ADI = "ZararTedarikci"
# KPI_ZARAR_KIRALIK_TABLO_ADI = "ZararKiralik"

# "Filo Analizi" sayfasındaki "Araç Tipi Performansı" bloğu (statik araç tipi
# kategori listesi + COUNTIF/SUMIF formülleri, Excel Tablosu DEĞİL düz hücre
# aralığı) — VERİ'de o ay GERÇEKTEN görülen araç tiplerini bu listeyle
# karşılaştırıp EKSİK olanları (ör. "Lowbed", "Panelvan") otomatik satır
# ekleyerek tamamlar. Sayfa/blok bulunamazsa sessizce atlanır.
KPI_ARAC_TIPI_PERFORMANS_GUNCELLE = True
# Sayfa adı / blok başlık metni farklıysa (varsayılan: "Filo Analizi" / "Araç Tipi")
# KPI_FILO_ANALIZ_SAYFA_ADLARI = ["Filo Analizi"]
# KPI_ARAC_TIPI_PERFORMANS_BASLIK_METNI = "Araç Tipi"

# Firma / şube (Uyumsoft VERİ raporu @CoCode@ / @BranchCodes@ — zorunlu)
CO_CODE = "IZLOG"
BRANCH_CODE = "MERKEZ"

# Personel veri girişinde atlanan alanlara varsayılan değer atanır (pivot
# tablolarda '(boş)' kategorisi oluşmasını önler). VERİ Oracle'dan çekildikten
# SONRA Python tarafında uygulanır — kpi_veri_rapor.sql'e dokunulmaz.
KPI_BOS_ALAN_VARSAYILARI = True
# MÜLKİYET (Z sütunu) boşsa:
KPI_MULKIYET_BOS_VARSAYILAN = "Tedarikçi"
# Proje Kodu "Konya" VE Yük Fiyat Tipi Kodu (K sütunu) boşsa:
KPI_KONYA_YUK_FIYAT_TIPI_BOS_VARSAYILAN = "ŞARKÜTERİ"
# Araç Tipi (X sütunu) boşsa:
KPI_ARAC_TIPI_BOS_VARSAYILAN = "Tır Frigorifik"

# Kapıdan kapıya yükleri hariç tut (rapor SQL'inde IS_DOOR_TO_DOOR = 0)
KPI_KAPI_KAPI_HARIC = True

# Fatura beklenmeyen operasyon kodları (faturasız problem sayımına dahil edilmez)
# KPI_FATURA_MUAF_OPERASYONLAR = ["BAŞKA_KOD"]

# Kalem detay sayfası satır üst sınırı (sevk + yük kalemleri birlikte)
KPI_KALEM_DETAY_LIMIT = 10000

# Fatura detay sayfası satır üst sınırı
KPI_FATURA_DETAY_LIMIT = 10000

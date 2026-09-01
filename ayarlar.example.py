# ==========================================
# İZLOG LOJİSTİK - SİSTEM AYARLARI VE ŞİFRELER
# ==========================================
#
# ⚠️ GÜVENLİK NOTU: Bu dosyadaki DB_SIFRE ve ERP_SIFRE alanlarını gerçek
# şifrelerinizle DOLDURUN, ancak gerçek şifrelerle bu dosyayı tekrar git'e
# commit ETMEYİN. Windows sunucusundaki gerçek (şifreleri dolu) kopya,
# repodaki bu şablon kopyadan bağımsız tutulmalı.

# 1. ORACLE VERİTABANI BAĞLANTI BİLGİLERİ
DB_KULLANICI = "uyumsoft"
DB_SIFRE = "DEĞİŞTİRİNİZ"
DB_DSN = "172.17.8.11:1521/UYUMSOFT"

# 2. UYUMSOFT ERP GİRİŞ BİLGİLERİ
ERP_LOGIN_URL = "https://erp.izlog.com.tr/login.aspx"  # ERP'nin giriş adresi
ERP_KULLANICI = "DEĞİŞTİRİNİZ"                          # Uyumsoft'a girerken kullandığınız ad
ERP_SIFRE = "DEĞİŞTİRİNİZ"                               # Uyumsoft şifreniz

# 3. OTOMASYON ÇALIŞMA KURALLARI VE LİNKLER
# DRY_RUN = True ise sistem sadece ekranı açar, verileri çeker ve tıklanabilirliği test eder.
# KESİNLİKLE "Kaydet", "Kopya" veya "Sevk Oluştur" butonlarına basmaz.
DRY_RUN = True

# DERIN_TEST_MODU = True (ve DRY_RUN = True iken) sistem Kopyalama + TÜM satış
# satırlarının veri girişini + fatura LOV eşleştirmesini (asıl düzeltilen hata
# burada) GERÇEKTEN yapar (satır bazlı "Kaydet" dahil -- bu güvenlidir), ama
# ana "Kaydet" (#btnSave_CD) butonuna KESİNLİKLE basmaz; pencereyi kayıt
# yapılmadan kapatır. "Sevk Oluştur" adımına hiç girmez (gerçekten kaydedilmiş
# bir Yük'e ihtiyaç duyar).
#
# NOT: Bilinen bir ERP hatası var -- ANA KAYDET İLE KAYDEDİLMİŞ bir Yük'e daha
# sonra geri dönüp faturalı bir fiyat satırında tekrar işlem yapmaya
# çalışıldığında Yük kilitleniyor (satır SQL'den boşaltılmadan silinemiyor).
# Bu, satır bazlı Kaydet'i veya bu test modunu ETKİLEMEZ; sadece "zaten
# kaydedilmiş bir Yük'ü tekrar açıp düzenleme" senaryosunda geçerli, ki
# otomasyon bunu hiç yapmıyor.
DERIN_TEST_MODU = False

ERP_YUK_LISTESI_URL = "https://erp.izlog.com.tr/MainList.aspx?CommandName=LGoodsCollection.Show&M=1&MenuId=654&WinId=01"

# Veritabanı İz Temizliğinde Kullanılacak Hedef Kullanıcı ID
HEDEF_KULLANICI_ID = 11310

# 4. TEŞHİS (DEBUG) EKRAN GÖRÜNTÜLERİ
# False (varsayılan): SADECE gerçek bir hata/istisna oluştuğunda ekran
# görüntüsü alınır (bunlar zaten nadirdir ve teşhis için gereklidir).
# True: Her BAŞARILI satırda da (Ücret Tipi/Operasyon Kodu/Tutar yazıldıktan
# ve Kaydet'e basıldıktan sonra) ek "sağlama" ekran görüntüleri alınır --
# bu, aktif geliştirme/hata ayıklama sırasında faydalıydı ama normal
# çalışmada klasörü hızla debug_*.png dosyalarıyla dolduruyor. Yeni bir
# hata türü araştırılırken geçici olarak True yapılabilir.
TESHIS_EKRAN_GORUNTUSU_AL = False

# 5. KPI ANALİZ VE ÜST YÖNETİM RAPORU AYARLARI
# Tarih formatı: DD.MM.YYYY
KPI_BASLANGIC_TARIHI = "01.01.2026"
KPI_BITIS_TARIHI = "31.01.2026"
KPI_RAPOR_DOSYASI = "kpi_rapor.xlsx"  # KPI/ klasörü içinde oluşur

# 6. KLASÖR YAPISI
# Proje kökü: C:\Users\hakinci\Desktop\Kodlarım\Cursor ERP Otomasyon
#
# Üç bağımsız iş hattı:
#   Test/   — Yük/sevk otomasyon geliştirme ve sevk testleri (UAT)
#   CANLI/  — Onaylanmış otomasyon; sevk testleri geçince buraya geçilir
#   KPI/    — KPI analiz ve raporlar; otomasyondan TAMAMEN bağımsız
#
# Sevk testleri devam ederken:
OTOMASYON_KLASORU = "Test"
# Canlıya geçince yalnızca bu satırı değiştirin (kod taşımaya gerek yok):
# OTOMASYON_KLASORU = "CANLI"
#
KPI_KLASORU = "KPI"                   # Her zaman KPI/ — Test/CANLI ile karışmaz
ISLEM_LISTESI_DOSYASI = "islem_listesi.xlsx"

# Windows sunucusunda Oracle Instant Client yolu (oracle_baglanti.py tarafından kullanılır)
ORACLE_CLIENT_LIB_DIR = r"C:\instantclient\instantclient_19_32"

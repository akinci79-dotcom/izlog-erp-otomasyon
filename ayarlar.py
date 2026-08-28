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

# DERIN_TEST_MODU = True (ve DRY_RUN = True iken) sistem Kopyalama + TEK bir
# satış satırının veri girişini + fatura LOV eşleştirmesini (asıl düzeltilen
# hata burada) GERÇEKTEN yapar, ama:
#   - Satır bazlı "Kaydet"e (a[id*='editnew']) KESİNLİKLE basmaz -- bilinen
#     ERP hatası: fatura seçildikten sonra satırı kaydetmek Yükü kilitliyor
#     ve satır SQL'den boşaltılmadan silinemiyor hale geliyor.
#   - Ana "Kaydet" (#btnSave_CD) butonuna da basmaz.
#   - "Sevk Oluştur" adımına hiç girmez (gerçekten kaydedilmiş bir Yük'e
#     ihtiyaç duyar).
# Bu yüzden birden fazla satır işlenmez; faturalı ilk satırda (fatura
# bulunup seçildikten sonra) test güvenle sonlandırılıp pencere kapatılır.
DERIN_TEST_MODU = False

ERP_YUK_LISTESI_URL = "https://erp.izlog.com.tr/MainList.aspx?CommandName=LGoodsCollection.Show&M=1&MenuId=654&WinId=01"

# Veritabanı İz Temizliğinde Kullanılacak Hedef Kullanıcı ID
HEDEF_KULLANICI_ID = 11310

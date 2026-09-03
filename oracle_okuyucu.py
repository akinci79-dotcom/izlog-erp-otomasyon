import oracledb
from decimal import Decimal
import ayarlar

# ==========================================
# ORACLE KALIN MOD (THICK MODE) BAŞLATICI
# ==========================================
oracledb.init_oracle_client(lib_dir=r"C:\instantclient\instantclient_19_32")

# ✅ "Ücret Tipi" KAYNAĞI KESİN OLARAK BULUNDU [DOĞRULANMIŞ, iki bağımsız
# veri noktasıyla teyitli]: `ucret_tipi_kesif.py` çıktısında
# INVD_BRANCH_EXPENSE tablosunun INVD_EXPENSE.EXPENSE_ID ile eşleşen
# satırlarında bir `OPERATION_PRICE_TYPE` (sayısal) kolonu bulundu:
#   - NAVLUN (EXPENSE_ID=732)  -> OPERATION_PRICE_TYPE = 1
#   - UĞRAMA (EXPENSE_ID=745)  -> OPERATION_PRICE_TYPE = 5
# ERP ekranındaki "Ücret Tipi" dropdown'ının sırası (Navlun, Hizmet, Diğer,
# Avans, Ek_Navlun, Yakit) 1'den başlayarak numaralandırılırsa:
#   1=Navlun, 2=Hizmet, 3=Diğer, 4=Avans, 5=Ek_Navlun, 6=Yakit
# Bu, HER İKİ bilinen değeri de (NAVLUN->1->"Navlun", UĞRAMA->5->"Ek_Navlun")
# birebir doğruluyor -- rastlantı olma ihtimali yok denecek kadar düşük.
# Böylece "Ücret Tipi" artık elle bakılan bir sözlük yerine gerçek bir Oracle
# sorgusuyla (aşağıdaki JOIN) otomatik ve güvenilir şekilde bulunuyor.
#
# **[VARSAYIM/TODO]**: INVD_BRANCH_EXPENSE çoklu şube (BRANCH_ID) desteği
# içeriyor; aynı EXPENSE_ID için farklı şubelerde farklı OPERATION_PRICE_TYPE
# tanımlı olma ihtimaline karşı SQL'de MIN() ile tek bir değere indirgeniyor
# (İzlog'un tüm bilinen kayıtlarında BRANCH_ID=6364/CO_ID=2371 sabit
# görünüyor, yani pratikte tek bir değer olması bekleniyor). Eğer ileride
# yanlış bir Ücret Tipi yazıldığı görülürse, bu JOIN'e ayrıca YK'nin kendi
# BRANCH_ID'siyle eşleşme şartı eklenmesi gerekebilir.
UCRET_TIPI_ENUM_ESLEME = {
    1: "Navlun",
    2: "Hizmet",
    3: "Diğer",
    4: "Avans",
    5: "Ek_Navlun",
    6: "Yakit",
}


def yuk_goods_id_getir(yuk_no):
    """
    Verilen Yük referans no'sunun (Y-...) Oracle GOODS_ID'sini (LMST_L_GOODS
    tablosunun birincil anahtarı) döndürür.

    NEDEN: Uyumsoft'un İncele ekranı `GeneralCard.aspx?CommandName=
    LGoodsCollection.Analyze&ObjectId={GOODS_ID}&WinId=01` URL kalıbıyla
    DOĞRUDAN açılabiliyor [DOĞRULANMIŞ, kullanıcı canlı ERP'de bu URL'yi
    ekran görüntüsüyle paylaştı -- `ayarlar.example.py`'deki
    `ERP_YUK_LISTESI_URL` içindeki `CommandName=LGoodsCollection.Show` ile
    AYNI aile, sadece komut adı "Show" değil "Analyze"]. Bu, Yük Listesi'nde
    satırı arayıp seçip "İncele" butonunu bulup tıklamaktan (id/metin
    tahmini, yeni pencere/aynı sayfa belirsizliği dahil) ÇOK daha güvenilir
    -- `izlog_yuk_otomasyon.py`'deki devam (resume) akışında birincil
    yöntem olarak kullanılıyor, buton tıklama sadece yedek (fallback).
    """
    KULLANICI = ayarlar.DB_KULLANICI
    SIFRE = ayarlar.DB_SIFRE
    DSN = ayarlar.DB_DSN

    baglanti = None
    try:
        baglanti = oracledb.connect(user=KULLANICI, password=SIFRE, dsn=DSN)
        cursor = baglanti.cursor()
        cursor.execute(
            "SELECT GOODS_ID FROM LMST_L_GOODS WHERE REFERENCE_NO = :yuk_no",
            {"yuk_no": yuk_no}
        )
        satir = cursor.fetchone()
        if not satir or satir[0] is None:
            raise ValueError(f"Yük veritabanında bulunamadı (GOODS_ID sorgusu): {yuk_no}")
        return satir[0]
    finally:
        if baglanti:
            baglanti.close()


def kaynak_yuk_verilerini_getir(kaynak_yuk_no):
    KULLANICI = ayarlar.DB_KULLANICI
    SIFRE = ayarlar.DB_SIFRE
    DSN = ayarlar.DB_DSN

    baglanti = None
    try:
        baglanti = oracledb.connect(user=KULLANICI, password=SIFRE, dsn=DSN)
        cursor = baglanti.cursor()

        # --- 1. AŞAMA: Ana Yük Bilgileri ---
        sql_ana_yuk = """
            SELECT
                TO_CHAR(Y.DOC_DATE, 'DD.MM.YYYY') as YUK_TARIHI,
                P.PROJECT_CODE as PROJE_KODU
            FROM LMST_L_GOODS Y
            LEFT JOIN LMSD_L_AGR_PROJ_TYPE P ON Y.PROJECT_ID = P.PROJECT_ID
            WHERE Y.REFERENCE_NO = :yuk_no
        """
        cursor.execute(sql_ana_yuk, {"yuk_no": kaynak_yuk_no})
        ana_kayit = cursor.fetchone()

        if not ana_kayit:
            raise ValueError(f"Kaynak yük veritabanında bulunamadı: {kaynak_yuk_no}")

        yuk_tarihi, proje_kodu = ana_kayit

        # --- 2. AŞAMA: FATURA BAĞLANTISI, OPERASYON KODU VE ÜCRET TİPİ ---
        # NOT: Önceki varsayımlar (LMSD_L_OP_DEFINITION / LMSD_L_GOODSPRICE_TYPE)
        # kullanıcının paylaştığı GERÇEK, ÇALIŞAN bir Uyumsoft raporunun
        # (LojistikYükSevkKalemRaporu) SQL'i ve Excel çıktısıyla KANITLANMIŞ
        # şekilde YANLIŞ çıktı:
        #   - "Operasyon Kodu" (NAVLUN, UĞRAMA, GENSET, MESAİ, BEKLEME) aslında
        #     INVD_EXPENSE tablosundan geliyor: INVD_EXPENSE.EXPENSE_CODE,
        #     join: INVD_EXPENSE.EXPENSE_ID = LMST_L_GOODS_OP_DET.OPERATION_ID.
        #     Gerçek rapor verisinde bu değerler birebir eşleşti.
        #   - LMSD_L_GOODSPRICE_TYPE (eski "Ücret Tipi" varsayımımız) aslında
        #     "Ücret Tipi" DEĞİL -- YÜK'ün (tüm yükün, satır bazlı değil) KARGO
        #     KATEGORİSİ (SAKARYA, KURUYÜK, ŞARKÜTERİ gibi). Bu yüzden bu join
        #     tamamen kaldırıldı.
        #   - "Ücret Tipi" [DOĞRULANMIŞ]: INVD_BRANCH_EXPENSE.OPERATION_PRICE_TYPE
        #     sayısal kolonundan geliyor (join: INVD_BRANCH_EXPENSE.EXPENSE_ID =
        #     INVD_EXPENSE.EXPENSE_ID). Bkz. yukarıdaki UCRET_TIPI_ENUM_ESLEME
        #     yorumu -- iki bağımsız kayıtla (NAVLUN->1, UĞRAMA->5) teyitli.
        sql_satis_satirlari = """
            SELECT
                HK.EXPENSE_CODE AS OPERATION_CODE,
                OPDET.AMT AS SATIS_FIYATI,
                INV.DOC_NO AS FATURA_NO,
                TO_CHAR(INV.DOC_DATE, 'DD.MM.YYYY') AS FATURA_TARIHI,
                BE.OPERATION_PRICE_TYPE AS OPERATION_PRICE_TYPE
            FROM LMST_L_GOODS_OP_DET OPDET
            LEFT JOIN LMST_L_GOODS YK ON YK.GOODS_ID = OPDET.GOODS_ID
            LEFT JOIN INVD_EXPENSE HK ON HK.EXPENSE_ID = OPDET.OPERATION_ID
            LEFT JOIN PSMT_INVOICE_M INV ON INV.INVOICE_M_ID = OPDET.INVOICE_M_ID
            LEFT JOIN (
                SELECT EXPENSE_ID, MIN(OPERATION_PRICE_TYPE) AS OPERATION_PRICE_TYPE
                FROM INVD_BRANCH_EXPENSE
                WHERE ISPASSIVE = 0
                GROUP BY EXPENSE_ID
            ) BE ON BE.EXPENSE_ID = HK.EXPENSE_ID

            WHERE YK.REFERENCE_NO = :yuk_no
              AND OPDET.PURCHASE_SALES_TYPE IN (2,4)
        """
        cursor.execute(sql_satis_satirlari, {"yuk_no": kaynak_yuk_no})
        sutunlar = [col[0] for col in cursor.description]

        satis_satirlari = []
        for satir in cursor.fetchall():
            veri = dict(zip(sutunlar, satir))

            ham_fiyat = veri.get("SATIS_FIYATI", 0)
            guvenli_fiyat = Decimal(str(ham_fiyat)) if ham_fiyat is not None else Decimal('0.00')

            op_metni_ham = veri.get("OPERATION_CODE")
            op_metni = str(op_metni_ham).strip() if op_metni_ham and str(op_metni_ham).strip() else "NAVLUN"

            # Bkz. yukarıdaki NOT: Ücret Tipi artık gerçek bir Oracle enum
            # kolonundan (INVD_BRANCH_EXPENSE.OPERATION_PRICE_TYPE) okunuyor.
            enum_kodu = veri.get("OPERATION_PRICE_TYPE")
            tip_metni = UCRET_TIPI_ENUM_ESLEME.get(enum_kodu) if enum_kodu is not None else None
            if tip_metni is None:
                raise ValueError(
                    f"[{kaynak_yuk_no}] HATA: '{op_metni}' operasyon kodu (EXPENSE_CODE) için "
                    f"INVD_BRANCH_EXPENSE.OPERATION_PRICE_TYPE bulunamadı ya da bilinmeyen bir "
                    f"sayısal değer döndü (ham değer: {enum_kodu!r}, bilinen değerler: "
                    f"{UCRET_TIPI_ENUM_ESLEME}). ERP'de bu satırın Ücret Tipi alanını gözle kontrol "
                    f"edip gerekirse oracle_okuyucu.py'deki UCRET_TIPI_ENUM_ESLEME sözlüğüne yeni "
                    f"sayı->metin eşlemesini ekleyin (yanlış tahmin ile ERP'ye hatalı veri girmemek "
                    f"için otomasyon burada durduruldu)."
                )

            print(f"[Bilgi] Satış satırı: OPERASYON_KODU='{op_metni}' (ham Oracle değeri: {op_metni_ham!r}), "
                  f"OPERATION_PRICE_TYPE={enum_kodu!r} -> UCRET_TIPI='{tip_metni}', AMT='{ham_fiyat}', "
                  f"FATURA_NO='{veri.get('FATURA_NO')}'")

            satis_satirlari.append({
                "OPERASYON_KODU": op_metni,
                "UCRET_TIPI": tip_metni,
                "SATIS_FIYATI": guvenli_fiyat,
                "FATURA_NO": veri.get("FATURA_NO"),
                "FATURA_TARIHI": veri.get("FATURA_TARIHI")
            })

        return {
            "YUK_TARIHI": yuk_tarihi,
            "PROJE_KODU": proje_kodu,
            "SATIS_SATIRLARI": satis_satirlari
        }

    except Exception as e:
        print(f"Oracle Veritabanı Hatası: {str(e)}")
        raise e
    finally:
        if baglanti:
            baglanti.close()

# ==========================================
# VERİTABANI İZ TEMİZLİĞİ (TOPLU GÜNCELLEME)
# ==========================================
def yeni_kayitlari_veritabaninda_guncelle(yuk_listesi, sevk_listesi):
    if ayarlar.DRY_RUN:
        print("DRY_RUN AKTİF: Veritabanı toplu güncellemeleri simüle edildi (Sorgu çalıştırılmadı).")
        return

    KULLANICI = ayarlar.DB_KULLANICI
    SIFRE = ayarlar.DB_SIFRE
    DSN = ayarlar.DB_DSN

    hedef_kullanici = getattr(ayarlar, "HEDEF_KULLANICI_ID", None)

    if not hedef_kullanici:
        print("UYARI: ayarlar.py içinde 'HEDEF_KULLANICI_ID' bulunamadı, güncelleme atlanıyor.")
        return

    baglanti = None
    try:
        baglanti = oracledb.connect(user=KULLANICI, password=SIFRE, dsn=DSN)
        cursor = baglanti.cursor()

        for yuk in yuk_listesi:
            cursor.execute("""
                UPDATE LMST_L_GOODS
                SET CREATE_USER_ID = :kullanici_id
                WHERE REFERENCE_NO = :yuk_no
            """, {"kullanici_id": hedef_kullanici, "yuk_no": yuk})

            # KRİTİK DÜZELTME: NULL yerine Uyumsoft'un orijinal sıfırlama değerleri
            cursor.execute("""
                UPDATE LMST_L_GOODS
                SET UPDATE_DATE = TO_DATE('01.01.0001', 'DD.MM.YYYY'), UPDATE_USER_ID = 0
                WHERE REFERENCE_NO = :yuk_no
            """, {"yuk_no": yuk})

        for sevk in sevk_listesi:
            cursor.execute("""
                UPDATE LMST_L_TRANSPORT
                SET CREATE_USER_ID = :kullanici_id
                WHERE TRANSPORT_NO = :sevk_no
            """, {"kullanici_id": hedef_kullanici, "sevk_no": sevk})

            cursor.execute("""
                UPDATE LMST_L_TRANSPORT
                SET UPDATE_DATE = TO_DATE('01.01.0001', 'DD.MM.YYYY'), UPDATE_USER_ID = 0
                WHERE TRANSPORT_NO = :sevk_no
            """, {"sevk_no": sevk})

        baglanti.commit()
        print(f"BAŞARILI: {len(yuk_listesi)} Yük ve {len(sevk_listesi)} Sevk için veritabanı ayak izi Uyumsoft standartlarında temizlendi.")

    except Exception as e:
        print(f"Toplu Güncelleme Oracle Hatası: {str(e)}")
        if baglanti:
            baglanti.rollback()
    finally:
        if baglanti:
            baglanti.close()

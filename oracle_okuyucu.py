import oracledb
from decimal import Decimal
import ayarlar

# ==========================================
# ORACLE KALIN MOD (THICK MODE) BAŞLATICI
# ==========================================
oracledb.init_oracle_client(lib_dir=r"C:\instantclient\instantclient_19_32")

def oto_kolon_kesfi(cursor, tablo_adi):
    """
    Sistemin hata vermesini engellemek için, Uyumsoft sözlük tablolarındaki
    metin/açıklama kolonunun adını tahmin etmeden Oracle'dan dinamik olarak okur.

    NOT: Bu fonksiyon TEK bir kolon adı döndürür (geriye dönük uyumluluk için
    tutuluyor). Canlı testte, tek kolona güvenmenin riskli olduğu görüldü:
    örn. "DESCRIPTION" kolonu bazı standart dışı operasyon kodlarında
    ("UĞRAMA" gibi) NULL/boş kalabiliyor, oysa aynı tablodaki "OPERATION_CODE"
    kolonu dolu olabiliyor. Bu yüzden asıl veri çekme sorgusunda artık
    `oto_kolon_listesi_kesfi()` kullanılıyor (aşağıda) -- birden fazla aday
    kolonu SQL'de COALESCE ile birleştirip "hangisi doluysa onu kullan"
    mantığı kuruyor.
    """
    cursor.execute(f"SELECT * FROM {tablo_adi} WHERE 1=0")
    kolonlar = [col[0].upper() for col in cursor.description]

    if 'DESCRIPTION' in kolonlar: return 'DESCRIPTION'
    if 'OPERATION_CODE' in kolonlar: return 'OPERATION_CODE'
    if 'GOODS_PRICE_TYPE_CODE' in kolonlar: return 'GOODS_PRICE_TYPE_CODE'

    for k in kolonlar:
        if 'CODE' in k or 'DESC' in k or 'NAME' in k or 'AD' in k:
            return k
    return kolonlar[1]


def oto_kolon_listesi_kesfi(cursor, tablo_adi):
    """
    `oto_kolon_kesfi()`'nin tek kolon yerine, en olası "okunabilir kod/açıklama"
    kolonlarının TAMAMINI öncelik sırasına göre bir liste olarak döndüren hali.

    NEDEN BÖYLE: Canlı testte, kaynak yükün fiyat satırında Operasyon Kodu
    "UĞRAMA" olmasına rağmen otomasyon bunu hiç göremedi ve varsayılan
    "NAVLUN" değerine düştü. Kök neden: tek kolona (örn. "DESCRIPTION")
    güvenmek -- bu kolon "UĞRAMA" gibi standart dışı kodlarda NULL/boş
    kalabiliyor, oysa aynı tablodaki başka bir kolon (örn. "OPERATION_CODE")
    her zaman dolu olabiliyor. Çağıran taraf bu listeyi SQL'de COALESCE ile
    birleştirerek hangi kolon doluysa onu kullanabiliyor.
    """
    cursor.execute(f"SELECT * FROM {tablo_adi} WHERE 1=0")
    kolonlar = [col[0].upper() for col in cursor.description]

    oncelik_sirasi = []
    for aday in ("DESCRIPTION", "OPERATION_CODE", "GOODS_PRICE_TYPE_CODE"):
        if aday in kolonlar and aday not in oncelik_sirasi:
            oncelik_sirasi.append(aday)

    for k in kolonlar:
        if ('CODE' in k or 'DESC' in k or 'NAME' in k or 'AD' in k) and k not in oncelik_sirasi:
            oncelik_sirasi.append(k)

    if not oncelik_sirasi:
        oncelik_sirasi = [kolonlar[1]]

    return oncelik_sirasi


def _coalesce_ifadesi(tablo_alias, kolon_listesi):
    """Birden fazla kolonu, boş string'leri de NULL sayarak COALESCE eden bir SQL ifadesi üretir."""
    parcalar = [f"NULLIF(TRIM({tablo_alias}.{kolon}), '')" for kolon in kolon_listesi]
    return f"COALESCE({', '.join(parcalar)})"

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

        # --- 2. AŞAMA: Dinamik Yapay Zeka (Sözlük Kolonlarını Bulma) ---
        # NOT: Artık TEK bir kolon tahmin edilmiyor -- birden fazla aday kolon
        # bulunup SQL'de COALESCE ile birleştiriliyor (bkz. yukarıdaki
        # oto_kolon_listesi_kesfi() açıklaması, "UĞRAMA" hatası).
        op_kolonlari = oto_kolon_listesi_kesfi(cursor, "LMSD_L_OP_DEFINITION")
        price_kolonlari = oto_kolon_listesi_kesfi(cursor, "LMSD_L_GOODSPRICE_TYPE")
        print(f"[Bilgi] Operasyon Kodu için denenecek kolonlar (öncelik sırasıyla): {op_kolonlari}")
        print(f"[Bilgi] Ücret Tipi için denenecek kolonlar (öncelik sırasıyla): {price_kolonlari}")

        op_ifadesi = _coalesce_ifadesi("OD", op_kolonlari)
        price_ifadesi = _coalesce_ifadesi("GT", price_kolonlari)

        # --- 3. AŞAMA: GERÇEK IT SQL'İ İLE FATURA BAĞLANTISI ---
        sql_satis_satirlari = f"""
            SELECT
                {op_ifadesi} as OPERATION_CODE,
                {price_ifadesi} as PRICE_TYPE,
                OPDET.AMT AS SATIS_FIYATI,
                INV.DOC_NO AS FATURA_NO,
                TO_CHAR(INV.DOC_DATE, 'DD.MM.YYYY') AS FATURA_TARIHI
            FROM LMST_L_GOODS_OP_DET OPDET
            LEFT JOIN LMST_L_GOODS YK ON YK.GOODS_ID = OPDET.GOODS_ID
            LEFT JOIN LMSD_L_OP_DEFINITION OD ON OD.OPERATION_ID = OPDET.OPERATION_ID
            LEFT JOIN LMSD_L_GOODSPRICE_TYPE GT ON GT.GOODS_PRICE_TYPE_ID = OPDET.OPERATION_PRICE_TYPE

            -- GİZEM ÇÖZÜLDÜ: Fatura ID'si zaten OP_DET'in içindeymiş ve tablo PSMT_INVOICE_M imiş!
            LEFT JOIN PSMT_INVOICE_M INV ON INV.INVOICE_M_ID = OPDET.INVOICE_M_ID

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

            # TEŞHİS: Oracle'dan varsayılan değere düşülmeden ÖNCE ham gelen
            # değerleri konsola yaz -- "UĞRAMA" gibi kodların neden NAVLUN'a
            # düştüğünü net görebilmek için.
            print(f"[Bilgi] Ham satış satırı verisi: OPERATION_CODE='{veri.get('OPERATION_CODE')}', "
                  f"PRICE_TYPE='{veri.get('PRICE_TYPE')}', AMT='{ham_fiyat}', "
                  f"FATURA_NO='{veri.get('FATURA_NO')}'")

            op_metni = veri.get("OPERATION_CODE")
            if not op_metni: op_metni = "NAVLUN"

            tip_metni = veri.get("PRICE_TYPE")
            if not tip_metni: tip_metni = "NAVLUN"

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

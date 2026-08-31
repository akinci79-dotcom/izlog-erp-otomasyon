"""
ucret_tipi_kesif.py

AMAÇ: "Ücret Tipi" (Navlun, Hizmet, Diğer, Avans, Ek_Navlun, Yakit) alanının
GERÇEKTE Oracle'da nereden geldiğini bulmak.

Bugüne kadarki varsayım YANLIŞ: kod, Operasyon Kodu "NAVLUN" değilse Ücret
Tipini sabit olarak "Ek_Navlun" yazıyordu -- ama ERP'deki dropdown'da 6 farklı
gerçek seçenek olduğu görüldü (Navlun, Hizmet, Diğer, Avans, Ek_Navlun,
Yakit). Bu script SADECE OKUMA yapar (SELECT), hiçbir veri değiştirmez.

KULLANIM: Aşağıdaki KAYNAK_YUK_NO değerini incelemek istediğiniz gerçek bir
yük numarasıyla değiştirip çalıştırın: python ucret_tipi_kesif.py
Çıktının TAMAMINI (konsoldan kopyalayıp) paylaşın.
"""
import oracledb
import ayarlar

oracledb.init_oracle_client(lib_dir=r"C:\instantclient\instantclient_19_32")

KAYNAK_YUK_NO = "Y-575733"  # İncelemek istediğiniz kaynak yük numarasını buraya yazın

baglanti = oracledb.connect(user=ayarlar.DB_KULLANICI, password=ayarlar.DB_SIFRE, dsn=ayarlar.DB_DSN)
cursor = baglanti.cursor()

print("=" * 70)
print("1. AŞAMA: INVD_EXPENSE tablosunun TÜM kolonları")
print("=" * 70)
cursor.execute("SELECT * FROM INVD_EXPENSE WHERE 1=0")
kolonlar = [col[0] for col in cursor.description]
print("Kolonlar:", kolonlar)

print()
print("=" * 70)
print(f"2. AŞAMA: '{KAYNAK_YUK_NO}' için satış satırlarındaki OPERATION_ID'lerin")
print("   INVD_EXPENSE tablosundaki TÜM ham kolonları")
print("=" * 70)
cursor.execute("""
    SELECT DISTINCT E.*
    FROM LMST_L_GOODS_OP_DET OPDET
    LEFT JOIN LMST_L_GOODS YK ON YK.GOODS_ID = OPDET.GOODS_ID
    LEFT JOIN INVD_EXPENSE E ON E.EXPENSE_ID = OPDET.OPERATION_ID
    WHERE YK.REFERENCE_NO = :yuk_no
      AND OPDET.PURCHASE_SALES_TYPE IN (2,4)
""", {"yuk_no": KAYNAK_YUK_NO})
satirlar = cursor.fetchall()
for satir in satirlar:
    for k, v in zip(kolonlar, satir):
        print(f"  {k}: {v}")
    print("-" * 40)

print()
print("=" * 70)
print("3. AŞAMA: INVD_EXPENSE tablosunun TÜM içeriği (küçükse tam liste)")
print("=" * 70)
cursor.execute("SELECT COUNT(*) FROM INVD_EXPENSE")
adet = cursor.fetchone()[0]
print(f"Toplam satır sayısı: {adet}")
if adet <= 200:
    cursor.execute("SELECT * FROM INVD_EXPENSE")
    for satir in cursor.fetchall():
        print(dict(zip(kolonlar, satir)))
else:
    print("Tablo çok büyük, tam liste atlanıyor -- yukarıdaki 2. aşama zaten yeterli olmalı.")

print()
print("=" * 70)
print("3b. AŞAMA: INVD_BRANCH_EXPENSE tablosu (rapor SQL'inde JOIN ediliyor")
print("    ama kolonu hiç SELECT edilmiyor -- 'Ücret Tipi' burada olabilir)")
print("=" * 70)
try:
    cursor.execute("SELECT * FROM INVD_BRANCH_EXPENSE WHERE 1=0")
    kolonlar_be = [col[0] for col in cursor.description]
    print("Kolonlar:", kolonlar_be)

    cursor.execute("""
        SELECT DISTINCT FHK.*
        FROM LMST_L_GOODS_OP_DET OPDET
        LEFT JOIN LMST_L_GOODS YK ON YK.GOODS_ID = OPDET.GOODS_ID
        LEFT JOIN INVD_BRANCH_EXPENSE FHK ON FHK.EXPENSE_ID = OPDET.OPERATION_ID
            AND FHK.BRANCH_ID = YK.BRANCH_ID
        WHERE YK.REFERENCE_NO = :yuk_no
          AND OPDET.PURCHASE_SALES_TYPE IN (2,4)
    """, {"yuk_no": KAYNAK_YUK_NO})
    for satir in cursor.fetchall():
        for k, v in zip(kolonlar_be, satir):
            print(f"  {k}: {v}")
        print("-" * 40)
except Exception as e:
    print(f"HATA: {e}")

print()
print("=" * 70)
print("4. AŞAMA: 'Ek_Navlun', 'Hizmet', 'Avans', 'Yakit' metinlerini içeren")
print("   TÜM tablo/kolon kombinasyonlarını ara (biraz sürebilir)")
print("=" * 70)
cursor.execute("""
    SELECT table_name, column_name
    FROM user_tab_columns
    WHERE data_type IN ('VARCHAR2', 'CHAR', 'NVARCHAR2', 'NCHAR')
    ORDER BY table_name, column_name
""")
metin_kolonlari = cursor.fetchall()
print(f"Taranacak metin kolonu sayısı: {len(metin_kolonlari)}")

bulunanlar = []
for aranan in ["EK_NAVLUN", "HİZMET", "HIZMET", "AVANS", "YAKIT"]:
    for tablo, kolon in metin_kolonlari:
        try:
            cursor.execute(f"""
                SELECT COUNT(*) FROM {tablo}
                WHERE UPPER({kolon}) LIKE :p
            """, {"p": f"%{aranan}%"})
            adet2 = cursor.fetchone()[0]
            if adet2 > 0:
                bulunanlar.append((aranan, tablo, kolon, adet2))
                print(f"  BULUNDU ('{aranan}'): {tablo}.{kolon} -> {adet2} satır")
        except Exception:
            continue

print()
print("=" * 70)
print("SONUÇ ÖZETİ:")
print(bulunanlar)
print("=" * 70)

baglanti.close()

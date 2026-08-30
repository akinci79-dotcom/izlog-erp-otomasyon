"""
operasyon_ucret_kesif.py

AMAÇ: "Ücret Tipi" (Navlun, Ek_Navlun, ...) ve "Operasyon Kodu" (UĞRAMA, ...)
alanlarının GERÇEKTE hangi Oracle tablosunda/kolonunda tutulduğunu bulmak.

`oracle_okuyucu.py`'deki mevcut varsayım (LMSD_L_OP_DEFINITION /
LMSD_L_GOODSPRICE_TYPE tablolarının bu alanları tuttuğu) canlı testte YANLIŞ
çıktı -- "ŞARKÜTERİ" gibi gerçek bir operasyon/ücret kodu OLMAYAN bir metin
döndü. Bu script SADECE OKUMA yapar (SELECT), hiçbir veri değiştirmez,
güvenle çalıştırılabilir.

KULLANIM: Aşağıdaki KAYNAK_YUK_NO değerini incelemek istediğiniz gerçek bir
yük numarasıyla değiştirip çalıştırın: python operasyon_ucret_kesif.py
Çıktının TAMAMINI (konsoldan kopyalayıp) paylaşın.
"""
import oracledb
import ayarlar

oracledb.init_oracle_client(lib_dir=r"C:\instantclient\instantclient_19_32")

KAYNAK_YUK_NO = "Y-575733"  # İncelemek istediğiniz kaynak yük numarasını buraya yazın

baglanti = oracledb.connect(user=ayarlar.DB_KULLANICI, password=ayarlar.DB_SIFRE, dsn=ayarlar.DB_DSN)
cursor = baglanti.cursor()

print("=" * 70)
print(f"1. AŞAMA: '{KAYNAK_YUK_NO}' için LMST_L_GOODS_OP_DET satırlarının TÜM ham kolonları")
print("=" * 70)
cursor.execute("""
    SELECT OPDET.*
    FROM LMST_L_GOODS_OP_DET OPDET
    LEFT JOIN LMST_L_GOODS YK ON YK.GOODS_ID = OPDET.GOODS_ID
    WHERE YK.REFERENCE_NO = :yuk_no
      AND OPDET.PURCHASE_SALES_TYPE IN (2,4)
""", {"yuk_no": KAYNAK_YUK_NO})
kolonlar = [col[0] for col in cursor.description]
print("Kolonlar:", kolonlar)
print()
satirlar = cursor.fetchall()
for satir in satirlar:
    for k, v in zip(kolonlar, satir):
        print(f"  {k}: {v}")
    print("-" * 40)

print()
print("=" * 70)
print("2. AŞAMA: LMSD_L_GOODSPRICE_TYPE tablosunun TÜM içeriği (küçük bir sözlük tablosu olmalı)")
print("=" * 70)
try:
    cursor.execute("SELECT * FROM LMSD_L_GOODSPRICE_TYPE")
    kolonlar2 = [col[0] for col in cursor.description]
    print("Kolonlar:", kolonlar2)
    for satir in cursor.fetchall():
        print(dict(zip(kolonlar2, satir)))
except Exception as e:
    print(f"HATA: {e}")

print()
print("=" * 70)
print("3. AŞAMA: LMSD_L_OP_DEFINITION tablosunun TÜM içeriği (küçük bir sözlük tablosu olmalı)")
print("=" * 70)
try:
    cursor.execute("SELECT * FROM LMSD_L_OP_DEFINITION")
    kolonlar3 = [col[0] for col in cursor.description]
    print("Kolonlar:", kolonlar3)
    for satir in cursor.fetchall():
        print(dict(zip(kolonlar3, satir)))
except Exception as e:
    print(f"HATA: {e}")

print()
print("=" * 70)
print("4. AŞAMA: 'NAVLUN' metnini içeren TÜM tablo/kolon kombinasyonlarını ara")
print("   (Kendi şemanızdaki tüm metin kolonlarını tarar, biraz sürebilir)")
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
for tablo, kolon in metin_kolonlari:
    try:
        cursor.execute(f"""
            SELECT COUNT(*) FROM {tablo}
            WHERE UPPER({kolon}) LIKE '%NAVLUN%'
        """)
        adet = cursor.fetchone()[0]
        if adet > 0:
            bulunanlar.append((tablo, kolon, adet))
            print(f"  BULUNDU: {tablo}.{kolon} -> {adet} satır")
    except Exception:
        continue

print()
print("=" * 70)
print("SONUÇ ÖZETİ ('NAVLUN' metni bulunan tablo.kolon listesi):")
print(bulunanlar)
print("=" * 70)

baglanti.close()

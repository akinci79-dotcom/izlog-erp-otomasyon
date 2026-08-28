import oracledb
import ayarlar

try:
    # Oracle bağlantısını başlat
    oracledb.init_oracle_client(lib_dir=r"C:\instantclient\instantclient_19_32")
    baglanti = oracledb.connect(user=ayarlar.DB_KULLANICI, password=ayarlar.DB_SIFRE, dsn=ayarlar.DB_DSN)
    cursor = baglanti.cursor()

    print("=== İZLOG UYUMSOFT FATURA DEDEKTİFİ ===")

    # 1. Sistemdeki Fatura tablolarını bul
    cursor.execute("""
        SELECT object_name
        FROM all_objects
        WHERE object_type IN ('TABLE', 'VIEW')
          AND (object_name LIKE '%INVOICE_M' OR object_name LIKE 'FIND_INVOICE%')
    """)
    print("\n[Bulunan Fatura Ana Tabloları]:")
    for row in cursor.fetchall():
        print(" ->", row[0])

    # 2. Fatura Detay tablolarını bul
    cursor.execute("""
        SELECT object_name
        FROM all_objects
        WHERE object_type IN ('TABLE', 'VIEW')
          AND (object_name LIKE '%INVOICE_D' OR object_name LIKE 'FIND_INVOICE_D%')
    """)
    print("\n[Bulunan Fatura Detay Tabloları]:")
    for row in cursor.fetchall():
        print(" ->", row[0])

except Exception as e:
    print("HATA:", str(e))
finally:
    if 'baglanti' in locals() and baglanti:
        baglanti.close()

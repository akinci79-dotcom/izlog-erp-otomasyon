"""
KPI modülü — Oracle bağlantı yardımcısı (otomasyondan bağımsız).
"""
from __future__ import annotations

import os
from contextlib import contextmanager

import ayarlar

_ORACLE_CLIENT_BASLATILDI = False


def oracle_client_baslat():
    global _ORACLE_CLIENT_BASLATILDI
    if _ORACLE_CLIENT_BASLATILDI:
        return

    import oracledb

    lib_dir = getattr(
        ayarlar,
        "ORACLE_CLIENT_LIB_DIR",
        r"C:\instantclient\instantclient_19_32",
    )
    if os.path.isdir(lib_dir):
        oracledb.init_oracle_client(lib_dir=lib_dir)

    _ORACLE_CLIENT_BASLATILDI = True


def baglanti_ac():
    import oracledb

    oracle_client_baslat()
    return oracledb.connect(
        user=ayarlar.DB_KULLANICI,
        password=ayarlar.DB_SIFRE,
        dsn=ayarlar.DB_DSN,
    )


@contextmanager
def baglanti_yonet():
    baglanti = baglanti_ac()
    try:
        yield baglanti
    finally:
        baglanti.close()


def tablo_var_mi(tablo_adi):
    with baglanti_yonet() as baglanti:
        cursor = baglanti.cursor()
        cursor.execute(
            """
            SELECT COUNT(*)
            FROM all_tables
            WHERE UPPER(table_name) = UPPER(:tablo)
            """,
            {"tablo": tablo_adi},
        )
        return cursor.fetchone()[0] > 0


def satir_limit_sql(ic_sorgu: str, limit: int) -> str:
    """
    Oracle 11g uyumlu satır sınırı.

    FETCH FIRST (12c+) yerine sıralı alt sorgu + ROWNUM kullanır.
    """
    n = int(limit)
    ic = ic_sorgu.strip().rstrip(";")
    return f"""SELECT * FROM (
    {ic}
) WHERE ROWNUM <= {n}"""


def tablo_kolonlari(cursor, tablo_adi: str) -> list[str]:
    """Verilen tablonun kolon adlarını döner (Oracle ALL_TAB_COLUMNS)."""
    cursor.execute(
        """
        SELECT COLUMN_NAME
        FROM ALL_TAB_COLUMNS
        WHERE UPPER(TABLE_NAME) = UPPER(:tablo)
        ORDER BY COLUMN_ID
        """,
        {"tablo": tablo_adi},
    )
    return [satir[0] for satir in cursor.fetchall()]

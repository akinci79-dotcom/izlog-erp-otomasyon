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

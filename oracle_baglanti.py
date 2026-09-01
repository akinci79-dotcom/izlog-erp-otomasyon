"""
Paylaşımlı Oracle bağlantı yardımcısı.

Windows sunucusunda Thick Mode (Instant Client) kullanır.
Linux/macOS geliştirme ortamında Thin Mode'a düşer (yalnızca ağ erişimi varsa).
"""
from __future__ import annotations

import os
from contextlib import contextmanager

import ayarlar

_ORACLE_CLIENT_BASLATILDI = False


def oracle_client_baslat():
    """Oracle istemcisini bir kez başlatır (Windows Thick Mode)."""
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
    """Yeni Oracle bağlantısı açar."""
    import oracledb

    oracle_client_baslat()
    return oracledb.connect(
        user=ayarlar.DB_KULLANICI,
        password=ayarlar.DB_SIFRE,
        dsn=ayarlar.DB_DSN,
    )


@contextmanager
def baglanti_yonet():
    """Bağlantıyı otomatik kapatır."""
    baglanti = baglanti_ac()
    try:
        yield baglanti
    finally:
        baglanti.close()


def sorgu_calistir(sql, parametreler=None):
    """Tek sorgu çalıştırır; (sütun_adları, satırlar) döner."""
    parametreler = parametreler or {}
    with baglanti_yonet() as baglanti:
        cursor = baglanti.cursor()
        cursor.execute(sql, parametreler)
        sutunlar = [col[0] for col in cursor.description]
        satirlar = cursor.fetchall()
    return sutunlar, satirlar


def tablo_var_mi(tablo_adi):
    """Verilen tablonun mevcut olup olmadığını kontrol eder."""
    sql = """
        SELECT COUNT(*)
        FROM all_tables
        WHERE UPPER(table_name) = UPPER(:tablo)
    """
    _, satirlar = sorgu_calistir(sql, {"tablo": tablo_adi})
    return satirlar[0][0] > 0

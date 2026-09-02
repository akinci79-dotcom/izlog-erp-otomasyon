"""VERİ raporu için Oracle tablo/kolon keşfi — yükleme, sürücü, arşiv alanları.

Kullanım (KPI klasöründen, Windows + Oracle):
  python kpi_veri_tablo_kesif.py
"""
from __future__ import annotations

from oracle_baglanti import baglanti_yonet

ANAHTAR_KELIMELER = (
    "GOODS_POINT",
    "POINT",
    "ROUTE",
    "DRIVER",
    "SURUCU",
    "ARCHIVE",
    "ARSIV",
    "ORDER",
    "SIPARIS",
    "LOADING",
    "UNLOAD",
    "YUKLEME",
    "BOSALTMA",
)


def main() -> None:
    print("=== VERİ raporu tablo keşfi ===\n")
    with baglanti_yonet() as baglanti:
        cursor = baglanti.cursor()
        for kelime in ANAHTAR_KELIMELER:
            cursor.execute(
                """
                SELECT object_name, object_type
                FROM all_objects
                WHERE owner = USER
                  AND object_type IN ('TABLE', 'VIEW')
                  AND UPPER(object_name) LIKE :p
                ORDER BY object_name
                ) WHERE ROWNUM <= 30
                """,
                {"p": f"%{kelime.upper()}%"},
            )
            satirlar = cursor.fetchall()
            if satirlar:
                print(f"[{kelime}]")
                for ad, tip in satirlar:
                    print(f"  {ad} ({tip})")
                print()

        print("LMST_L_GOODS önemli kolonlar:")
        cursor.execute(
            """
            SELECT column_name
            FROM all_tab_columns
            WHERE table_name = 'LMST_L_GOODS'
              AND (
                UPPER(column_name) LIKE '%ENTITY%'
                OR UPPER(column_name) LIKE '%ARCH%'
                OR UPPER(column_name) LIKE '%ORDER%'
                OR UPPER(column_name) LIKE '%WEIGHT%'
                OR UPPER(column_name) LIKE '%TYPE%'
                OR UPPER(column_name) LIKE '%DRIVER%'
              )
            ORDER BY column_id
            """
        )
        for (kolon,) in cursor.fetchall():
            print(f"  {kolon}")


if __name__ == "__main__":
    main()

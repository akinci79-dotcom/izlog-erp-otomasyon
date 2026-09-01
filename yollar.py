"""
Proje klasör yapısı ve dosya yolları.

Otomasyon çıktıları (Excel, ekran görüntüleri) `otomasyon/` altında;
KPI raporları `raporlar/` altında tutulur. İki iş hattının çıktıları karışmaz.
"""
from __future__ import annotations

import os

import ayarlar

_PROJE_KOKU = os.path.dirname(os.path.abspath(__file__))


def _klasor_yolu(ayar_adi: str, varsayilan: str) -> str:
    deger = getattr(ayarlar, ayar_adi, varsayilan)
    if os.path.isabs(deger):
        return deger
    return os.path.join(_PROJE_KOKU, deger)


def otomasyon_klasoru() -> str:
    return _klasor_yolu("OTOMASYON_KLASORU", "otomasyon")


def raporlar_klasoru() -> str:
    return _klasor_yolu("RAPORLAR_KLASORU", "raporlar")


def klasorleri_olustur():
    """Gerekli alt klasörleri oluşturur (yoksa)."""
    os.makedirs(otomasyon_klasoru(), exist_ok=True)
    os.makedirs(raporlar_klasoru(), exist_ok=True)


def islem_listesi_yolu() -> str:
    dosya = getattr(ayarlar, "ISLEM_LISTESI_DOSYASI", "islem_listesi.xlsx")
    if os.path.isabs(dosya):
        return dosya
    return os.path.join(otomasyon_klasoru(), dosya)


def yedek_excel_yolu(zaman_damgasi: str) -> str:
    return os.path.join(otomasyon_klasoru(), f"yedek_islem_listesi_{zaman_damgasi}.xlsx")


def kpi_rapor_yolu(dosya_adi: str | None = None) -> str:
    dosya = dosya_adi or getattr(ayarlar, "KPI_RAPOR_DOSYASI", "kpi_rapor.xlsx")
    if os.path.isabs(dosya):
        return dosya
    return os.path.join(raporlar_klasoru(), dosya)


def ekran_goruntusu_yolu(dosya_adi: str) -> str:
    return os.path.join(otomasyon_klasoru(), dosya_adi)

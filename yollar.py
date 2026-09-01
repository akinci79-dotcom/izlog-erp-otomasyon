"""
Proje klasör yapısı ve dosya yolları.

Varsayılan yapı (Cursor ERP Otomasyon kök klasörü):
  CANLI/  — yük/sevk otomasyonu (Excel, ekran görüntüleri)
  KPI/    — KPI analiz raporları
"""
from __future__ import annotations

import os

import ayarlar

_PROJE_KOKU = os.path.dirname(os.path.abspath(__file__))


def _klasor_yolu(ayar_adi: str, varsayilan: str) -> str:
    deger = getattr(ayarlar, ayar_adi, None)
    if deger is None and ayar_adi == "KPI_KLASORU":
        # Eski ayar adı geriye dönük uyumluluk
        deger = getattr(ayarlar, "RAPORLAR_KLASORU", varsayilan)
    if deger is None:
        deger = varsayilan
    if os.path.isabs(deger):
        return deger
    return os.path.join(_PROJE_KOKU, deger)


def otomasyon_klasoru() -> str:
    return _klasor_yolu("OTOMASYON_KLASORU", "Test")


def kpi_klasoru() -> str:
    return _klasor_yolu("KPI_KLASORU", "KPI")


def raporlar_klasoru() -> str:
    """Geriye dönük uyumluluk — kpi_klasoru() ile aynı."""
    return kpi_klasoru()


def klasorleri_olustur():
    """Test, CANLI ve KPI klasörlerini oluşturur (yoksa)."""
    for ad in ("Test", "CANLI", "KPI"):
        os.makedirs(os.path.join(_PROJE_KOKU, ad), exist_ok=True)
    os.makedirs(otomasyon_klasoru(), exist_ok=True)
    os.makedirs(kpi_klasoru(), exist_ok=True)


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
    return os.path.join(kpi_klasoru(), dosya)


def ekran_goruntusu_yolu(dosya_adi: str) -> str:
    return os.path.join(otomasyon_klasoru(), dosya_adi)


def alt_klasor_etiketi(klasor_yolu: str | None = None) -> str:
    """Hata mesajlarında kısa klasör adı (örn. CANLI, KPI)."""
    yol = klasor_yolu or otomasyon_klasoru()
    return os.path.basename(os.path.normpath(yol)) or "CANLI"

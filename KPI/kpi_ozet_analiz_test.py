"""Özet sayfası manuel tablo analizi birim testleri — Oracle/Excel olmadan."""
from __future__ import annotations

from decimal import Decimal

from kpi_ozet_analiz import (
    donus_yuku_katki,
    en_karli_musteriler,
    en_karli_rotalar,
    kritik_zarar_rotalari,
    mulkiyet_kategorileri,
    sube_kategorileri,
)


def _assert(cond: bool, msg: str) -> None:
    if not cond:
        raise AssertionError(msg)


def _ornek_veri() -> list[dict]:
    return [
        {
            "PROJE_KODU": "Konya",
            "MUSTERI_ADI": "Müşteri A",
            "PLAKA_MULKIYET": "Tedarikçi",
            "SEFER_TURU": "GİDİŞ YÜKÜ",
            "YUKLEME_SEHIR_ADI": "İstanbul",
            "BOSALTMA_SEHIR_ADI": "Ankara",
            "TOPLAM_ALIS": Decimal("100"),
            "TOPLAM_SATIS": Decimal("150"),
            "TOPLAM_KAR_ZARAR": Decimal("50"),
        },
        {
            "PROJE_KODU": "Sakarya",
            "MUSTERI_ADI": "Müşteri B",
            "PLAKA_MULKIYET": "Kiralık",
            "SEFER_TURU": "DÖNÜŞ YÜKÜ",
            "YUKLEME_SEHIR_ADI": "Ankara",
            "BOSALTMA_SEHIR_ADI": "İstanbul",
            "TOPLAM_ALIS": Decimal("200"),
            "TOPLAM_SATIS": Decimal("180"),
            "TOPLAM_KAR_ZARAR": Decimal("-20"),
        },
        {
            "PROJE_KODU": "Konya",
            "MUSTERI_ADI": "Müşteri A",
            "PLAKA_MULKIYET": "Tedarikçi",
            "SEFER_TURU": "GİDİŞ YÜKÜ",
            "YUKLEME_SEHIR_ADI": "İstanbul",
            "BOSALTMA_SEHIR_ADI": "Ankara",
            "TOPLAM_ALIS": Decimal("100"),
            "TOPLAM_SATIS": Decimal("200"),
            "TOPLAM_KAR_ZARAR": Decimal("100"),
        },
        {
            "PROJE_KODU": "İzmir",
            "MUSTERI_ADI": "Müşteri C",
            "PLAKA_MULKIYET": "Tedarikçi",
            "SEFER_TURU": "DÖNÜŞ YÜKÜ",
            "YUKLEME_SEHIR_ADI": "İzmir",
            "BOSALTMA_SEHIR_ADI": "Bursa",
            "TOPLAM_ALIS": Decimal("50"),
            "TOPLAM_SATIS": Decimal("120"),
            "TOPLAM_KAR_ZARAR": Decimal("70"),
        },
    ]


def test_sube_kategorileri() -> None:
    veri = _ornek_veri()
    subeler = sube_kategorileri(veri)
    _assert(subeler == ["İzmir", "Konya", "Sakarya"], f"Beklenmeyen şubeler: {subeler}")


def test_mulkiyet_kategorileri() -> None:
    veri = _ornek_veri()
    mulkiyet = mulkiyet_kategorileri(veri)
    _assert("Kiralık" in mulkiyet and "Tedarikçi" in mulkiyet, mulkiyet)


def test_en_karli_musteriler() -> None:
    veri = _ornek_veri()
    top = en_karli_musteriler(veri, n=2)
    _assert(len(top) == 2, "2 müşteri bekleniyordu")
    _assert(top[0]["etiket"] == "Müşteri A", "En kârlı Müşteri A olmalı")
    _assert(top[0]["kar_zarar"] == 150.0, f"Toplam kâr 150 olmalı, {top[0]['kar_zarar']}")


def test_donus_yuku_katki() -> None:
    veri = _ornek_veri()
    top = donus_yuku_katki(veri, n=5)
    _assert(len(top) == 2, "2 dönüş yükü müşterisi olmalı")
    etiketler = {t["etiket"] for t in top}
    _assert("Müşteri B" in etiketler and "Müşteri C" in etiketler, etiketler)


def test_en_karli_rotalar() -> None:
    veri = _ornek_veri()
    top = en_karli_rotalar(veri, n=3)
    _assert(top[0]["etiket"] == "İstanbul → Ankara", f"En kârlı rota: {top[0]}")
    _assert(top[0]["sefer"] == 2, "İstanbul-Ankara 2 sefer")


def test_kritik_zarar_rotalari() -> None:
    veri = _ornek_veri()
    top = kritik_zarar_rotalari(veri, n=3)
    _assert(len(top) == 1, "Tek zararlı rota olmalı")
    _assert(top[0]["etiket"] == "Ankara → İstanbul", top[0])
    _assert(top[0]["kar_zarar"] == -20.0, top[0])


def main() -> None:
    test_sube_kategorileri()
    test_mulkiyet_kategorileri()
    test_en_karli_musteriler()
    test_donus_yuku_katki()
    test_en_karli_rotalar()
    test_kritik_zarar_rotalari()
    print("OK — tüm Özet analiz testleri geçti")


if __name__ == "__main__":
    main()

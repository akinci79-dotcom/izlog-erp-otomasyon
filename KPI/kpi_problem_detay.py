"""
Yönetici özetindeki problemlerin satır bazında doğrulanabilir detay listesi.
"""
from __future__ import annotations

from datetime import datetime
from decimal import Decimal
from typing import Any

from kpi_fatura_detay import _tarih_gun_farki, fatura_probleme_dahil
from kpi_kalem_detay import _decimal

PROBLEM_DETAY_SUTUNLARI = [
    "PROBLEM_TIPI",
    "TIP",
    "SEVK_YUK_NO",
    "YUK_NO",
    "BELGE_TARIHI",
    "FATURA_NO",
    "FATURA_TARIHI",
    "GUN_FARKI",
    "OPERASYON_KODU",
    "PROJE_KODU",
    "PLAKA",
    "TUTAR",
]


def _tarih_metin(deger) -> str | None:
    if deger is None:
        return None
    if isinstance(deger, datetime):
        return deger.strftime("%d.%m.%Y")
    return str(deger)


def erken_fatura_kalemleri(fatura_detay: list[dict]) -> list[dict]:
    sonuc = []
    for satir in fatura_detay:
        if not satir.get("FATURA_NO"):
            continue
        if not fatura_probleme_dahil(satir.get("OPERASYON_KODU")):
            continue
        gun = _tarih_gun_farki(satir.get("BELGE_TARIHI"), satir.get("FATURA_TARIHI"))
        if gun is not None and gun < 0:
            sonuc.append(
                {
                    "PROBLEM_TIPI": "Fatura tarihi belge tarihinden önce",
                    "TIP": satir.get("TIP"),
                    "SEVK_YUK_NO": satir.get("SEVK_YUK_NO"),
                    "YUK_NO": "",
                    "BELGE_TARIHI": _tarih_metin(satir.get("BELGE_TARIHI")),
                    "FATURA_NO": satir.get("FATURA_NO"),
                    "FATURA_TARIHI": _tarih_metin(satir.get("FATURA_TARIHI")),
                    "GUN_FARKI": int(gun),
                    "OPERASYON_KODU": satir.get("OPERASYON_KODU"),
                    "PROJE_KODU": "",
                    "PLAKA": "",
                    "TUTAR": None,
                }
            )
    return sonuc


def faturasiz_kalem_satirlari(kalem_detay: list[dict]) -> list[dict]:
    sonuc = []
    for satir in kalem_detay:
        alis_satis = satir.get("ALIS_SATIS") or ""
        if satir.get("FATURA_NO") or alis_satis not in ("Satış", "Alış"):
            continue
        if not fatura_probleme_dahil(satir.get("OPERASYON_KODU")):
            continue
        belge = satir.get("SEVK_TARIHI") if satir.get("TIP") == "Sevk Kalemi" else satir.get("YUK_TARIHI")
        sonuc.append(
            {
                "PROBLEM_TIPI": "Faturasız kalem",
                "TIP": satir.get("TIP"),
                "SEVK_YUK_NO": satir.get("SEVK_NO"),
                "YUK_NO": satir.get("YUK_NO"),
                "BELGE_TARIHI": _tarih_metin(belge),
                "FATURA_NO": "",
                "FATURA_TARIHI": "",
                "GUN_FARKI": None,
                "OPERASYON_KODU": satir.get("OPERASYON_KODU"),
                "PROJE_KODU": satir.get("PROJE_KODU"),
                "PLAKA": satir.get("PLAKA"),
                "TUTAR": satir.get("TUTAR"),
            }
        )
    return sonuc


def sevksiz_yuk_satirlari(sevksiz_yukler: list[dict]) -> list[dict]:
    sonuc = []
    for satir in sevksiz_yukler:
        sonuc.append(
            {
                "PROBLEM_TIPI": "Sevki olmayan yük",
                "TIP": "Yük",
                "SEVK_YUK_NO": satir.get("REFERENCE_NO"),
                "YUK_NO": satir.get("REFERENCE_NO"),
                "BELGE_TARIHI": _tarih_metin(satir.get("DOC_DATE")),
                "FATURA_NO": "",
                "FATURA_TARIHI": "",
                "GUN_FARKI": None,
                "OPERASYON_KODU": "",
                "PROJE_KODU": satir.get("PROJE"),
                "PLAKA": "",
                "TUTAR": None,
            }
        )
    return sonuc


def zararli_sevk_satirlari(kirilim: list[dict]) -> list[dict]:
    sonuc = []
    for satir in kirilim:
        kz = _decimal(satir.get("NET_KAR_ZARAR"))
        if kz >= 0:
            continue
        sonuc.append(
            {
                "PROBLEM_TIPI": "Çoklu yüklü sevkte zarar",
                "TIP": "Sevk",
                "SEVK_YUK_NO": satir.get("SEVK_NO"),
                "YUK_NO": satir.get("YUK_LISTESI"),
                "BELGE_TARIHI": _tarih_metin(satir.get("SEVK_TARIHI")),
                "FATURA_NO": "",
                "FATURA_TARIHI": "",
                "GUN_FARKI": None,
                "OPERASYON_KODU": "",
                "PROJE_KODU": "",
                "PLAKA": satir.get("PLAKA"),
                "TUTAR": kz,
            }
        )
    return sonuc


def problem_detay_olustur(
    fatura_detay: list[dict],
    kalem_detay: list[dict],
    kalem_sevk_kirilim: list[dict],
    sevksiz_yukler: list[dict] | None = None,
) -> list[dict]:
    """Tüm problem tiplerini tek listede birleştirir."""
    satirlar: list[dict] = []
    satirlar.extend(erken_fatura_kalemleri(fatura_detay))
    satirlar.extend(faturasiz_kalem_satirlari(kalem_detay))
    if sevksiz_yukler:
        satirlar.extend(sevksiz_yuk_satirlari(sevksiz_yukler))
    satirlar.extend(zararli_sevk_satirlari(kalem_sevk_kirilim))
    return satirlar


def ornek_problem_detay() -> list[dict]:
    return [
        {
            "PROBLEM_TIPI": "Fatura tarihi belge tarihinden önce",
            "TIP": "Yük",
            "SEVK_YUK_NO": "Y-6003",
            "YUK_NO": "",
            "BELGE_TARIHI": "15.01.2026",
            "FATURA_NO": "SAT2026002",
            "FATURA_TARIHI": "10.01.2026",
            "GUN_FARKI": -5,
            "OPERASYON_KODU": "NAVLUN",
            "PROJE_KODU": "",
            "PLAKA": "",
            "TUTAR": None,
        },
    ]

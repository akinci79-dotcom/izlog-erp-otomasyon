"""
Özet sayfasındaki manuel tablolar için VERİ satırlarından saf Python analizi.

Bu modül Excel/COM'a bağımlı değildir — test edilebilir hesaplama katmanı.
Excel'e yazma: kpi_sablon_rapor._com_ozet_manuel_tablolari_guncelle
"""
from __future__ import annotations

import unicodedata
from collections import defaultdict
from decimal import Decimal, InvalidOperation
from typing import Any

_BOS_ETIKET = "(boş)"


def _decimal(deger: Any) -> Decimal:
    if deger is None:
        return Decimal("0")
    if isinstance(deger, Decimal):
        return deger
    try:
        return Decimal(str(deger))
    except (InvalidOperation, ValueError):
        return Decimal("0")


def _metin(deger: Any) -> str:
    if deger is None:
        return _BOS_ETIKET
    metin = str(deger).strip()
    return metin if metin else _BOS_ETIKET


def _normalize(metin: str) -> str:
    metin = str(metin or "").strip().upper()
    metin = unicodedata.normalize("NFKD", metin)
    return "".join(c for c in metin if not unicodedata.combining(c))


def _rota_etiketi(yukleme: Any, bosaltma: Any) -> str:
    return f"{_metin(yukleme)} → {_metin(bosaltma)}"


def _sefer_turu_donus_mu(deger: Any) -> bool:
    return "DONUS" in _normalize(str(deger or ""))


def sube_kategorileri(veri_satirlari: list[dict[str, Any]]) -> list[str]:
    """ŞUBE PERFORMANSI bloğu — benzersiz PROJE_KODU, alfabetik."""
    gorulen: set[str] = set()
    sonuc: list[str] = []
    for satir in veri_satirlari:
        ad = _metin(satir.get("PROJE_KODU"))
        norm = _normalize(ad)
        if norm in gorulen:
            continue
        gorulen.add(norm)
        sonuc.append(ad)
    return sorted(sonuc, key=lambda x: _normalize(x))


def mulkiyet_kategorileri(veri_satirlari: list[dict[str, Any]]) -> list[str]:
    """MÜLKİYET PERFORMANSI — benzersiz PLAKA_MULKIYET, alfabetik."""
    gorulen: set[str] = set()
    sonuc: list[str] = []
    for satir in veri_satirlari:
        ad = _metin(satir.get("PLAKA_MULKIYET"))
        norm = _normalize(ad)
        if norm in gorulen:
            continue
        gorulen.add(norm)
        sonuc.append(ad)
    return sorted(sonuc, key=lambda x: _normalize(x))


def _musteri_grupla(
    veri_satirlari: list[dict[str, Any]],
    *,
    sadece_donus: bool = False,
) -> list[dict[str, Any]]:
    gruplar: dict[str, dict[str, Any]] = {}

    for satir in veri_satirlari:
        if sadece_donus and not _sefer_turu_donus_mu(satir.get("SEFER_TURU")):
            continue

        musteri = _metin(satir.get("MUSTERI_ADI"))
        norm = _normalize(musteri)
        grup = gruplar.get(norm)
        if grup is None:
            grup = {
                "etiket": musteri,
                "sefer": 0,
                "alis": Decimal("0"),
                "satis": Decimal("0"),
                "kar_zarar": Decimal("0"),
            }
            gruplar[norm] = grup

        grup["sefer"] += 1
        grup["alis"] += _decimal(satir.get("TOPLAM_ALIS"))
        grup["satis"] += _decimal(satir.get("TOPLAM_SATIS"))
        grup["kar_zarar"] += _decimal(satir.get("TOPLAM_KAR_ZARAR"))

    return list(gruplar.values())


def en_karli_musteriler(
    veri_satirlari: list[dict[str, Any]], n: int = 5
) -> list[dict[str, Any]]:
    """EN ÇOK KÂR YARATAN MÜŞTERİLER — TOPLAM_KAR_ZARAR'a göre azalan top-N."""
    gruplar = _musteri_grupla(veri_satirlari)
    gruplar.sort(key=lambda g: g["kar_zarar"], reverse=True)
    return [_satir_dict_cevir(g) for g in gruplar[: max(n, 0)]]


def donus_yuku_katki(
    veri_satirlari: list[dict[str, Any]], n: int = 5
) -> list[dict[str, Any]]:
    """DÖNÜŞ YÜKÜ KATKISI — SEFER_TURU='DÖNÜŞ YÜKÜ' satırları, kâra göre top-N."""
    gruplar = _musteri_grupla(veri_satirlari, sadece_donus=True)
    gruplar.sort(key=lambda g: g["kar_zarar"], reverse=True)
    return [_satir_dict_cevir(g) for g in gruplar[: max(n, 0)]]


def _rota_grupla(veri_satirlari: list[dict[str, Any]]) -> list[dict[str, Any]]:
    gruplar: dict[str, dict[str, Any]] = {}

    for satir in veri_satirlari:
        etiket = _rota_etiketi(
            satir.get("YUKLEME_SEHIR_ADI"), satir.get("BOSALTMA_SEHIR_ADI")
        )
        norm = _normalize(etiket)
        grup = gruplar.get(norm)
        if grup is None:
            grup = {
                "etiket": etiket,
                "sefer": 0,
                "alis": Decimal("0"),
                "satis": Decimal("0"),
                "kar_zarar": Decimal("0"),
            }
            gruplar[norm] = grup

        grup["sefer"] += 1
        grup["alis"] += _decimal(satir.get("TOPLAM_ALIS"))
        grup["satis"] += _decimal(satir.get("TOPLAM_SATIS"))
        grup["kar_zarar"] += _decimal(satir.get("TOPLAM_KAR_ZARAR"))

    return list(gruplar.values())


def en_karli_rotalar(
    veri_satirlari: list[dict[str, Any]], n: int = 5
) -> list[dict[str, Any]]:
    """EN KÂRLI ROTALAR — şehir çifti, pozitif kâr öncelikli top-N."""
    gruplar = _rota_grupla(veri_satirlari)
    gruplar.sort(key=lambda g: g["kar_zarar"], reverse=True)
    return [_satir_dict_cevir(g) for g in gruplar[: max(n, 0)]]


def kritik_zarar_rotalari(
    veri_satirlari: list[dict[str, Any]], n: int = 5
) -> list[dict[str, Any]]:
    """KRİTİK ZARAR ROTALARI — en negatif kâr/zarar top-N."""
    gruplar = [g for g in _rota_grupla(veri_satirlari) if g["kar_zarar"] < 0]
    gruplar.sort(key=lambda g: g["kar_zarar"])
    return [_satir_dict_cevir(g) for g in gruplar[: max(n, 0)]]


def _satir_dict_cevir(grup: dict[str, Any]) -> dict[str, Any]:
    alis = grup["alis"]
    satis = grup["satis"]
    kar = grup["kar_zarar"]
    kar_yuzde = (kar / alis) if alis else Decimal("0")
    return {
        "etiket": grup["etiket"],
        "sefer": grup["sefer"],
        "alis": float(alis),
        "satis": float(satis),
        "kar_zarar": float(kar),
        "kar_yuzde": float(kar_yuzde),
    }

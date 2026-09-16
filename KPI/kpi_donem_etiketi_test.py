"""Dönem etiketi ve çıktı dosya adı testleri."""
from __future__ import annotations

import shutil
from pathlib import Path

_KPI = Path(__file__).resolve().parent
_AYARLAR = _KPI / "ayarlar.py"


def _hazirla_ayarlar() -> None:
    if not _AYARLAR.exists():
        shutil.copy(_KPI / "ayarlar.example.py", _AYARLAR)


def _assert(cond: bool, msg: str) -> None:
    if not cond:
        raise AssertionError(msg)


def test_tek_ay_etiketi() -> None:
    from kpi_sablon_rapor import _donem_etiketi

    _assert(
        _donem_etiketi("01.08.2026", "31.08.2026") == "8- Ağustos 2026",
        "Agustos etiketi ay numarasi ile uretilmeli",
    )
    _assert(
        _donem_etiketi("01.09.2026", "30.09.2026") == "9- Eylül 2026",
        "Eylul etiketi ay numarasi ile uretilmeli",
    )


def test_cikti_dosya_adi() -> None:
    import ayarlar
    from kpi_sablon_rapor import _cikti_yolu, sablon_yolu

    ayarlar.KPI_DONEM = "09.2026"
    ayarlar.KPI_RAPOR_DOSYASI = "kpi_rapor.xlsx"
    hedef = _cikti_yolu(sablon_yolu(), "01.09.2026", "30.09.2026")
    _assert(
        hedef.name == "9- Eylül 2026 İzlog Lojistik Raporları.xlsx",
        f"Beklenen donem adli dosya, gelen: {hedef.name}",
    )


def main() -> None:
    _hazirla_ayarlar()
    test_tek_ay_etiketi()
    test_cikti_dosya_adi()
    print("OK — donem etiketi testleri gecti")


if __name__ == "__main__":
    main()

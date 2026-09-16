"""VERİ kolon eşleme birim testleri — Excel COM olmadan çalışır."""
from __future__ import annotations

from kpi_sablon_rapor import (
    _com_veri_matrisi_hazirla,
    _kolon_esleme,
    _sutun1_arteakt_mi,
)


def _assert(cond: bool, msg: str) -> None:
    if not cond:
        raise AssertionError(msg)


def test_sutun1_arteakt() -> None:
    _assert(_sutun1_arteakt_mi("Sütun1"), "Sütun1 tanınmalı")
    _assert(_sutun1_arteakt_mi("Column1"), "Column1 tanınmalı")
    _assert(not _sutun1_arteakt_mi("PROJE_KODU"), "PROJE_KODU artefakt değil")


def test_esleme_a_sutunundan() -> None:
    basliklar = ["PROJE_KODU", "YUK_NO", "ALIS_TUTARI"]
    anahtarlar = ["PROJE_KODU", "YUK_NO", "ALIS_TUTAR"]
    esleme = _kolon_esleme(basliklar, anahtarlar, tablo_sol=1)
    _assert(esleme[1] == "PROJE_KODU", "PROJE_KODU A sütununda")
    _assert(esleme[2] == "YUK_NO", "YUK_NO B sütununda")
    _assert(esleme[3] == "ALIS_TUTAR", "ALIS_TUTARI → ALIS_TUTAR C sütununda")


def test_esleme_sutun1_kaydirilmis_sablon() -> None:
    """Bozuk şablonda Sütun1=A, PROJE_KODU=B — eşleme mutlak sütun numarası kullanır."""
    basliklar = ["Sütun1", "PROJE_KODU", "YUK_NO"]
    anahtarlar = ["PROJE_KODU", "YUK_NO", "SUBE"]
    esleme = _kolon_esleme(basliklar, anahtarlar, tablo_sol=1)
    _assert(1 not in esleme, "Sütun1 eşleşmemeli")
    _assert(esleme[2] == "PROJE_KODU", "PROJE_KODU B sütununda")
    _assert(esleme[3] == "YUK_NO", "YUK_NO C sütununda")


def test_matris_dogru_hucreye_yazar() -> None:
    basliklar = ["PROJE_KODU", "SUBE", "YUK_NO"]
    satirlar = [{"PROJE_KODU": "PRJ1", "SUBE": "Konya", "YUK_NO": "Y-1"}]
    matris = _com_veri_matrisi_hazirla(basliklar, satirlar, None, 3, tablo_sol=1)
    _assert(matris[0][0] == "PRJ1", "PROJE_KODU ilk hücre")
    _assert(matris[0][1] == "Konya", "SUBE ikinci hücre")
    _assert(matris[0][2] == "Y-1", "YUK_NO üçüncü hücre")


def test_matris_tablo_b_sutunundan_basliyorsa() -> None:
    """Tablo B sütunundan başlıyorsa veri B'ye yazılır, A'ya değil."""
    basliklar = ["PROJE_KODU", "SUBE"]
    satirlar = [{"PROJE_KODU": "PRJ1", "SUBE": "Konya"}]
    matris = _com_veri_matrisi_hazirla(basliklar, satirlar, None, 2, tablo_sol=2)
    _assert(len(matris[0]) == 2, "2 kolonluk matris")
    _assert(matris[0][0] == "PRJ1", "PROJE_KODU matris[0]")
    _assert(matris[0][1] == "Konya", "SUBE matris[1]")


def main() -> None:
    test_sutun1_arteakt()
    test_esleme_a_sutunundan()
    test_esleme_sutun1_kaydirilmis_sablon()
    test_matris_dogru_hucreye_yazar()
    test_matris_tablo_b_sutunundan_basliyorsa()
    print("OK — tüm kolon eşleme testleri geçti")


if __name__ == "__main__":
    main()

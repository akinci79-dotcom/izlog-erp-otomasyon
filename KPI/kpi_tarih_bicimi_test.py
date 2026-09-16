"""Tarih NumberFormat doğrulama testleri — Excel COM olmadan çalışır.

`_tarih_bicimi_gecerli_mi` (kpi_sablon_rapor.py), Excel'den GERİ OKUNAN
`.NumberFormat`/`.NumberFormatLocal` metninin GERÇEKTEN bir 'gün.ay.yıl'
tarih biçimi olup olmadığını YAPISAL olarak denetler — eski koddaki TAM
literal string eşitliği ('== "dd.mm.yyyy"') YANLIŞ ALARM üretiyordu (bkz.
fonksiyonun docstring'i ve README "Sorun giderme" bölümü).

Çalıştırma (ayarlar.py gerektirir, kpi_sablon_rapor.py modül seviyesinde
`import ayarlar` yaptığı için):
    cp ayarlar.example.py ayarlar.py
    python3 kpi_tarih_bicimi_test.py
    rm ayarlar.py
"""
from __future__ import annotations

from kpi_sablon_rapor import _tarih_bicimi_gecerli_mi


def _assert(cond: bool, msg: str) -> None:
    if not cond:
        raise AssertionError(msg)


def test_ingilizce_kod_gecerli() -> None:
    _assert(_tarih_bicimi_gecerli_mi("dd.mm.yyyy"), "'dd.mm.yyyy' geçerli olmalı")


def test_turkce_kod_gecerli() -> None:
    """✅ ASIL DÜZELTME — eskiden bu YANLIŞLIKLA 'başarısız' sayılıyordu, çünkü
    tek karşılaştırma literal 'dd.mm.yyyy' string'ine tam eşitlikti."""
    _assert(
        _tarih_bicimi_gecerli_mi("gg.aa.yyyy"),
        "'gg.aa.yyyy' (Türkçe COM okuma tuhaflığı) artık GEÇERLİ sayılmalı",
    )


def test_kacisli_ayrac_ikisi_de_gecerli() -> None:
    _assert(_tarih_bicimi_gecerli_mi(r"dd\.mm\.yyyy"), "kaçışlı ayraçlı 'dd\\.mm\\.yyyy' geçerli olmalı")
    _assert(_tarih_bicimi_gecerli_mi(r"gg\.aa\.yyyy"), "kaçışlı ayraçlı 'gg\\.aa\\.yyyy' geçerli olmalı")


def test_farkli_ayraclar_gecerli() -> None:
    _assert(_tarih_bicimi_gecerli_mi("dd/mm/yyyy"), "'/' ayraçlı biçim de geçerli olmalı")
    _assert(_tarih_bicimi_gecerli_mi("gg-aa-yyyy"), "'-' ayraçlı biçim de geçerli olmalı")
    _assert(_tarih_bicimi_gecerli_mi("DD.MM.YYYY"), "büyük harfli kod da geçerli olmalı (case-insensitive)")


def test_bilinen_bozuk_kalip_hala_yakalaniyor() -> None:
    """⚠️ KRİTİK — validasyonu gevşetirken GERÇEK bir hatayı GİZLEMEMELİYİZ.
    Kullanıcının canlı ekran görüntüsüyle teyit ettiği bozuk kalıp (gerçek ay +
    literal 'mm' metni + gerçek yıl, GÜN bilgisi tamamen kayıp — ekranda
    '08.mm.2026' gibi görünüyor) HÂLÂ False dönmeli."""
    _assert(
        not _tarih_bicimi_gecerli_mi(r"mm/\m\m/yyyy"),
        "Bilinen bozuk kalıp 'mm/\\m\\m/yyyy' HÂLÂ geçersiz sayılmalı (gün bileşeniyle başlamıyor)",
    )


def test_diger_gecersiz_degerler() -> None:
    _assert(not _tarih_bicimi_gecerli_mi("General"), "'General' bir tarih biçimi değil")
    _assert(not _tarih_bicimi_gecerli_mi("0.00"), "sayı biçimi tarih değil")
    _assert(not _tarih_bicimi_gecerli_mi(""), "boş string geçersiz olmalı")
    _assert(not _tarih_bicimi_gecerli_mi(None), "None geçersiz olmalı")
    _assert(not _tarih_bicimi_gecerli_mi("hh:mm:ss"), "saat biçimi bir tarih biçimi değil")


def main() -> None:
    test_ingilizce_kod_gecerli()
    test_turkce_kod_gecerli()
    test_kacisli_ayrac_ikisi_de_gecerli()
    test_farkli_ayraclar_gecerli()
    test_bilinen_bozuk_kalip_hala_yakalaniyor()
    test_diger_gecersiz_degerler()
    print("OK — tüm tarih biçimi doğrulama testleri geçti")


if __name__ == "__main__":
    main()

"""'Filo Analizi' sayfasındaki 'Araç Tipi Performansı' bloğu için otomatik
kategori tamamlama testleri — Excel COM olmadan, basit Python nesneleriyle
(Range/Cells taklidi) çalışır.

Çalıştırma (ayarlar.py gerektirir, kpi_sablon_rapor.py modül seviyesinde
`import ayarlar` yaptığı için):
    cp ayarlar.example.py ayarlar.py
    python3 kpi_arac_tipi_performans_test.py
    rm ayarlar.py

KÖK NEDEN (kullanıcı openpyxl ile gerçek rapor dosyasını inceleyip
doğruladı): "Filo Analizi" sayfasındaki "Araç Tipi Performansı" bloğu
(A9:G19 civarı) TAMAMEN STATİK bir kategori listesi — bir Excel Tablosu
(ListObject) DEĞİL, düz hücre aralığı + COUNTIF/SUMIF formülleri. VERİ
sayfasında (Tablo5) YENİ bir araç tipi (ör. "Lowbed", "Panelvan") ortaya
çıktığında, bu blok bunu hiç İÇERMEDİĞİ için o araç tipinin TÜM verisi
tablodan (ve varsa altındaki dip toplamdan) TAMAMEN GÖRÜNMEZ kalıyordu.

Bu testler `_com_arac_tipi_performans_guncelle`'in şu senaryoları doğru
ele aldığını doğrular:
  (a) Mevcut listede TÜM araç tipleri zaten varsa — hiçbir satır eklenmemeli,
      hiçbir şey değişmemeli (regresyon yok).
  (b) 1-2 yeni araç tipi varsa — doğru sayıda satır eklenmeli, doğru isimler
      A sütununa yazılmalı, B-G formülleri komşu satırdan doğru kopyalanmalı
      (göreli referans doğru KAYMIŞ olmalı — örn. yeni satır 20 için formül
      A20'yi göstermeli, A19 DEĞİL), altındaki dip toplam formülü
      genişletilmiş/doğrulanmış olmalı.
  (c) Başlık bloğu (veya sayfa) bulunamıyorsa — sessizce atlanmalı, hata
      fırlatılmamalı.
"""
from __future__ import annotations

import re
from typing import Any

from openpyxl.utils import get_column_letter

from kpi_sablon_rapor import _com_arac_tipi_performans_guncelle


def _assert(cond: bool, msg: str) -> None:
    if not cond:
        raise AssertionError(msg)


# --------------------------------------------------------------------------
# Basit Excel COM taklidi — bu dosyanın ihtiyaç duyduğu minimum yüzey
# (Cells/Range/Rows, ListObject YOK — bu blok bir ListObject değil).
# --------------------------------------------------------------------------


class _FakeCell:
    def __init__(self, sheet: "_FakeSheet", row: int, col: int) -> None:
        self.sheet = sheet
        self.row = row
        self.col = col

    def _k(self):
        return (self.row, self.col)

    @property
    def Value(self):
        return self.sheet.values.get(self._k())

    @Value.setter
    def Value(self, v):
        self.sheet.values[self._k()] = v

    @property
    def Formula(self):
        return self.sheet.formulas.get(self._k(), "")

    @Formula.setter
    def Formula(self, v):
        self.sheet.formulas[self._k()] = v

    @property
    def FormulaR1C1(self):
        return self.sheet.formula_r1c1.get(self._k(), "")

    @FormulaR1C1.setter
    def FormulaR1C1(self, v):
        self.sheet.formula_r1c1[self._k()] = v


class _FakeRange:
    def __init__(self, sheet: "_FakeSheet", r1: int, c1: int, r2: int, c2: int) -> None:
        self.sheet = sheet
        self.r1, self.c1, self.r2, self.c2 = r1, c1, r2, c2

    @property
    def Value(self):
        rows = [
            tuple(self.sheet.values.get((r, c)) for c in range(self.c1, self.c2 + 1))
            for r in range(self.r1, self.r2 + 1)
        ]
        return rows[0] if len(rows) == 1 else tuple(rows)

    @Value.setter
    def Value(self, v):
        if self.r2 == self.r1:
            for i, val in enumerate(v):
                self.sheet.values[(self.r1, self.c1 + i)] = val
        else:
            for ri, satir in enumerate(v):
                for ci, val in enumerate(satir):
                    self.sheet.values[(self.r1 + ri, self.c1 + ci)] = val

    @property
    def FormulaR1C1(self):
        return self.sheet.formula_r1c1.get((self.r1, self.c1), "")

    @FormulaR1C1.setter
    def FormulaR1C1(self, v):
        # Gerçek Excel COM davranışı: bir SKALER FormulaR1C1 dizesi çok
        # hücreli bir Range'e atanırsa, HER hücreye AYNI göreli (RC[...])
        # metin yazılır — Excel bunu her hücrenin KENDİ konumuna göre
        # otomatik yeniden yorumlar (bkz. bu dosyadaki `_efektif_a1` testi).
        for r in range(self.r1, self.r2 + 1):
            for c in range(self.c1, self.c2 + 1):
                self.sheet.formula_r1c1[(r, c)] = v


class _FakeRowsHandle:
    def __init__(self, sheet: "_FakeSheet", spec: str) -> None:
        self.sheet = sheet
        a, b = spec.split(":")
        self.a, self.b = int(a), int(b)

    def Insert(self) -> None:
        if self.sheet.insert_hatasi:
            raise RuntimeError(self.sheet.insert_hatasi)
        n = self.b - self.a + 1
        self.sheet._satirlari_kaydir(self.a, n)
        self.sheet.insert_calls.append((self.a, self.b))


class _FakeSheet:
    def __init__(self, name: str = "Filo Analizi") -> None:
        self.Name = name
        self.values: dict[tuple[int, int], Any] = {}
        self.formulas: dict[tuple[int, int], str] = {}
        self.formula_r1c1: dict[tuple[int, int], str] = {}
        self.insert_calls: list[tuple[int, int]] = []
        self.insert_hatasi: str | None = None

    def Cells(self, r: int, c: int) -> _FakeCell:
        return _FakeCell(self, r, c)

    def Range(self, cell1: _FakeCell, cell2: _FakeCell) -> _FakeRange:
        return _FakeRange(self, cell1.row, cell1.col, cell2.row, cell2.col)

    def Rows(self, spec: str) -> _FakeRowsHandle:
        return _FakeRowsHandle(self, spec)

    def _satirlari_kaydir(self, baslangic: int, n: int) -> None:
        for depo in (self.values, self.formulas, self.formula_r1c1):
            anahtarlar = sorted(
                (k for k in depo if k[0] >= baslangic), key=lambda k: -k[0]
            )
            for (r, c) in anahtarlar:
                depo[(r + n, c)] = depo.pop((r, c))


class _FakeWorkbook:
    def __init__(self, worksheets: list[_FakeSheet]) -> None:
        self.Worksheets = worksheets


# --------------------------------------------------------------------------
# Yardımcı: gerçekçi bir "Araç Tipi Performansı" bloğu kur (başlık satırı +
# N kategori satırı + isteğe bağlı bir dip toplam satırı).
# --------------------------------------------------------------------------

_ARAC_TIPLERI_ORIJINAL = [
    "10 Tkr Kamyon Açık", "10 Tkr Kamyon Frigorifik", "10 Tkr Kamyon Tenteli",
    "8 Tkr Kamyon Frigorifik", "Kamyonet", "Tır Açık", "Tır Frigorifik",
    "Tır Katlı Frigorifik", "Tır Kayar Perdeli", "Tır Tenteli",
]

_BASLIKLAR = ["Araç Tipi", "Sefer", "Alış", "Satış", "Kâr/Zarar", "Kâr %", "Sefer Başı Kâr"]

# B..G sütunlarının kaynak satırdaki (göreli, RC[...]) formül şablonları —
# hepsi AYNI satırdaki A sütununa (araç tipi adı) veya birbirine bakıyor.
_FORMUL_SABLONLARI = {
    2: "=COUNTIF(Tablo5[ARAC_TIPI],RC[-1])",                              # Sefer
    3: "=SUMIF(Tablo5[ARAC_TIPI],RC[-2],Tablo5[TOPLAM_ALIS])",            # Alış
    4: "=SUMIF(Tablo5[ARAC_TIPI],RC[-3],Tablo5[TOPLAM_SATIS])",          # Satış
    5: "=SUMIF(Tablo5[ARAC_TIPI],RC[-4],Tablo5[TOPLAM_KAR_ZARAR])",      # Kâr/Zarar
    6: "=IFERROR(RC[-1]/RC[-3],0)",                                       # Kâr %
    7: "=IFERROR(RC[-2]/RC[-4],0)",                                       # Sefer Başı Kâr
}

_RC_TOKEN = re.compile(r"RC\[(-?\d+)\]")


def _efektif_a1(formula_r1c1: str, row: int, col: int) -> str:
    """Test yardımcısı: bir R1C1 şablonunu, belirli bir (row, col) hücresi
    İÇİN gerçek Excel'in üreteceği A1-stili formüle çevirir — SADECE bu
    dosyadaki basit 'RC[n]' (aynı satır, göreli sütun) deseni için."""

    def _degistir(m: "re.Match[str]") -> str:
        offset = int(m.group(1))
        hedef_col = col + offset
        return f"{get_column_letter(hedef_col)}{row}"

    return _RC_TOKEN.sub(_degistir, formula_r1c1)


def _blok_kur(
    header_row: int = 9,
    tipler: list[str] | None = None,
    dip_toplam: bool = True,
) -> _FakeSheet:
    tipler = list(tipler if tipler is not None else _ARAC_TIPLERI_ORIJINAL)
    sheet = _FakeSheet()

    for i, baslik in enumerate(_BASLIKLAR):
        sheet.values[(header_row, 1 + i)] = baslik

    veri_ilk_satir = header_row + 1
    for i, tip in enumerate(tipler):
        row = veri_ilk_satir + i
        sheet.values[(row, 1)] = tip
        for col, sablon in _FORMUL_SABLONLARI.items():
            sheet.formula_r1c1[(row, col)] = sablon

    if dip_toplam and tipler:
        son_satir = veri_ilk_satir + len(tipler) - 1
        dip_satir = son_satir + 1
        for col in _FORMUL_SABLONLARI:
            col_harf = get_column_letter(col)
            sheet.formulas[(dip_satir, col)] = f"=SUM({col_harf}{veri_ilk_satir}:{col_harf}{son_satir})"

    return sheet


def _veri_satirlari_olustur(arac_tipleri: list[str]) -> list[dict[str, Any]]:
    return [{"ARAC_TIPI": t, "YUK_NO": f"Y-{i}"} for i, t in enumerate(arac_tipleri, start=1)]


# --------------------------------------------------------------------------
# (a) Mevcut listede TÜM araç tipleri zaten varsa — regresyon yok.
# --------------------------------------------------------------------------


def test_a_tum_tipler_mevcutsa_hicbir_sey_degismiyor() -> None:
    sheet = _blok_kur()
    wb = _FakeWorkbook([sheet])
    onceki_values = dict(sheet.values)
    onceki_formul_r1c1 = dict(sheet.formula_r1c1)

    # Veri sayfasındaki değerler AYNI 10 tip (fazladan tekrar + boşluk/case
    # farklılığıyla — normalize karşılaştırmanın çalıştığını da doğrular).
    veri = _veri_satirlari_olustur(_ARAC_TIPLERI_ORIJINAL + [_ARAC_TIPLERI_ORIJINAL[0]])
    veri.append({"ARAC_TIPI": "  tır frigorifik  ", "YUK_NO": "Y-99"})

    uyarilar = _com_arac_tipi_performans_guncelle(wb, veri)

    _assert(uyarilar == [], f"Uyarı olmamalıydı, {uyarilar} döndü")
    _assert(len(sheet.insert_calls) == 0, "Hiçbir yeni tip yokken Insert() çağrılmamalı")
    _assert(sheet.values == onceki_values, "Hiçbir hücre değeri değişmemeliydi")
    _assert(sheet.formula_r1c1 == onceki_formul_r1c1, "Hiçbir formül değişmemeliydi")


# --------------------------------------------------------------------------
# (b) 1-2 yeni araç tipi varsa — doğru satır ekleniyor, formüller/dip toplam
#     doğru kayıyor.
# --------------------------------------------------------------------------


def test_b_yeni_arac_tipleri_dogru_ekleniyor() -> None:
    header_row = 9
    sheet = _blok_kur(header_row=header_row, dip_toplam=True)
    wb = _FakeWorkbook([sheet])

    veri_ilk_satir = header_row + 1  # 10
    eski_veri_son_satir = veri_ilk_satir + len(_ARAC_TIPLERI_ORIJINAL) - 1  # 19
    eski_dip_toplam_satir = eski_veri_son_satir + 1  # 20

    # Sağlama: dip toplam eski konumda gerçekten var.
    _assert(
        sheet.formulas.get((eski_dip_toplam_satir, 2)) == f"=SUM(B{veri_ilk_satir}:B{eski_veri_son_satir})",
        "Test kurulumu hatalı: dip toplam beklenen eski konumda değil",
    )

    veri = _veri_satirlari_olustur(_ARAC_TIPLERI_ORIJINAL + ["Lowbed", "Panelvan"])

    uyarilar = _com_arac_tipi_performans_guncelle(wb, veri)
    _assert(uyarilar == [], f"Uyarı olmamalıydı, {uyarilar} döndü")

    eksik_sayisi = 2
    _assert(len(sheet.insert_calls) == 1, "TEK bir bulk Insert() çağrısı olmalı")
    _assert(
        sheet.insert_calls[0] == (eski_veri_son_satir, eski_veri_son_satir + eksik_sayisi - 1),
        f"Insert() eski son veri satırının ({eski_veri_son_satir}) TAM ÜZERİNE yapılmalı, "
        f"{sheet.insert_calls[0]} bulundu",
    )

    yeni_veri_son_satir = eski_veri_son_satir + eksik_sayisi  # 21
    yeni_lowbed_satir = eski_veri_son_satir  # 19 (yeni boş satırların İLKİ)
    yeni_panelvan_satir = eski_veri_son_satir + 1  # 20

    _assert(
        sheet.values.get((yeni_lowbed_satir, 1)) == "Lowbed",
        f"Lowbed satırı {yeni_lowbed_satir}'e yazılmalıydı, "
        f"bulunan: {sheet.values.get((yeni_lowbed_satir, 1))!r}",
    )
    _assert(
        sheet.values.get((yeni_panelvan_satir, 1)) == "Panelvan",
        f"Panelvan satırı {yeni_panelvan_satir}'e yazılmalıydı, "
        f"bulunan: {sheet.values.get((yeni_panelvan_satir, 1))!r}",
    )

    # 🔑 KRİTİK DOĞRULAMA: göreli formül referansı doğru KAYMIŞ olmalı —
    # Panelvan satırı (20) için Sefer formülü A20'yi göstermeli, A19 DEĞİL.
    panelvan_sefer_r1c1 = sheet.formula_r1c1.get((yeni_panelvan_satir, 2))
    _assert(
        panelvan_sefer_r1c1 is not None,
        "Panelvan satırının Sefer (B) formülü kopyalanmamış",
    )
    efektif = _efektif_a1(panelvan_sefer_r1c1, yeni_panelvan_satir, 2)
    _assert(
        efektif == f"=COUNTIF(Tablo5[ARAC_TIPI],A{yeni_panelvan_satir})",
        f"Panelvan Sefer formülü A{yeni_panelvan_satir}'ü göstermeliydi, "
        f"efektif formül: {efektif!r}",
    )
    _assert(
        f"A{eski_veri_son_satir}" not in efektif or yeni_panelvan_satir == eski_veri_son_satir,
        f"Formül hâlâ ESKİ satırı (A{eski_veri_son_satir}) gösteriyor, KAYMAMIŞ: {efektif!r}",
    )

    # Lowbed satırı (19) için de aynı doğrulama.
    lowbed_alis_r1c1 = sheet.formula_r1c1.get((yeni_lowbed_satir, 3))
    efektif_alis = _efektif_a1(lowbed_alis_r1c1, yeni_lowbed_satir, 3)
    _assert(
        efektif_alis == f"=SUMIF(Tablo5[ARAC_TIPI],A{yeni_lowbed_satir},Tablo5[TOPLAM_ALIS])",
        f"Lowbed Alış formülü A{yeni_lowbed_satir}'ü göstermeliydi, efektif: {efektif_alis!r}",
    )

    # Dip toplam formülü yeni aralığı kapsayacak şekilde düzeltilmiş olmalı.
    yeni_dip_toplam_satir = yeni_veri_son_satir + 1  # 22
    dip_formul = sheet.formulas.get((yeni_dip_toplam_satir, 2))
    _assert(
        dip_formul == f"=SUM(B{veri_ilk_satir}:B{yeni_veri_son_satir})",
        f"Dip toplam formülü yeni aralığı (B{veri_ilk_satir}:B{yeni_veri_son_satir}) "
        f"kapsayacak şekilde düzeltilmeliydi, bulunan: {dip_formul!r}",
    )

    # Eski dip toplam konumu artık YOK (kaymış) — orada eski (yanlış) formül kalmamalı.
    _assert(
        (eski_dip_toplam_satir, 2) not in sheet.formulas
        or sheet.formulas[(eski_dip_toplam_satir, 2)] != f"=SUM(B{veri_ilk_satir}:B{eski_veri_son_satir})",
        "Eski dip toplam konumunda hâlâ eski formül kalmış olmamalı (kaymalıydı)",
    )


def test_b2_tek_yeni_tip_ve_dip_toplam_yokken_hata_vermiyor() -> None:
    """Dip toplam satırı hiç YOKSA (blok en altta bitiyorsa) da sorunsuz
    çalışmalı — `_com_alt_toplam_formulu_dogrula_ve_duzelt` formül bulamadığında
    sessizce hiçbir şey yapmaz (hata değildir)."""
    header_row = 9
    sheet = _blok_kur(header_row=header_row, dip_toplam=False)
    wb = _FakeWorkbook([sheet])

    veri = _veri_satirlari_olustur(_ARAC_TIPLERI_ORIJINAL + ["Lowbed"])
    uyarilar = _com_arac_tipi_performans_guncelle(wb, veri)

    _assert(uyarilar == [], f"Uyarı olmamalıydı, {uyarilar} döndü")
    _assert(len(sheet.insert_calls) == 1, "Insert() çağrılmalıydı")
    veri_ilk_satir = header_row + 1
    eski_veri_son_satir = veri_ilk_satir + len(_ARAC_TIPLERI_ORIJINAL) - 1
    _assert(
        sheet.values.get((eski_veri_son_satir, 1)) == "Lowbed",
        "Lowbed satırı doğru konuma yazılmalıydı",
    )


# --------------------------------------------------------------------------
# (c) Başlık bloğu / sayfa bulunamıyorsa — sessizce atlanmalı.
# --------------------------------------------------------------------------


def test_c_sayfa_bulunamazsa_sessizce_atlaniyor() -> None:
    wb = _FakeWorkbook([_FakeSheet(name="Başka Bir Sayfa")])
    veri = _veri_satirlari_olustur(_ARAC_TIPLERI_ORIJINAL + ["Lowbed"])

    uyarilar = _com_arac_tipi_performans_guncelle(wb, veri)

    _assert(uyarilar == [], f"Sayfa yokken uyarı/hata olmamalı, {uyarilar} döndü")


def test_c2_baslik_bulunamazsa_sessizce_atlaniyor() -> None:
    sheet = _FakeSheet()
    sheet.values[(1, 1)] = "Bambaşka bir başlık"
    wb = _FakeWorkbook([sheet])
    veri = _veri_satirlari_olustur(_ARAC_TIPLERI_ORIJINAL + ["Lowbed"])

    uyarilar = _com_arac_tipi_performans_guncelle(wb, veri)

    _assert(uyarilar == [], f"Başlık yokken uyarı/hata olmamalı, {uyarilar} döndü")
    _assert(len(sheet.insert_calls) == 0, "Başlık bulunamadıysa Insert() hiç çağrılmamalı")


def main() -> None:
    test_a_tum_tipler_mevcutsa_hicbir_sey_degismiyor()
    test_b_yeni_arac_tipleri_dogru_ekleniyor()
    test_b2_tek_yeni_tip_ve_dip_toplam_yokken_hata_vermiyor()
    test_c_sayfa_bulunamazsa_sessizce_atlaniyor()
    test_c2_baslik_bulunamazsa_sessizce_atlaniyor()
    print("OK — tüm Araç Tipi Performansı testleri geçti")


if __name__ == "__main__":
    main()

"""Zarar Detay kapasite büyütme testleri — Excel COM olmadan, basit Python
nesneleriyle (Range/ListObject taklidi) çalışır.

Çalıştırma (ayarlar.py gerektirir, kpi_sablon_rapor.py modül seviyesinde
`import ayarlar` yaptığı için):
    cp ayarlar.example.py ayarlar.py
    python3 kpi_zarar_kapasite_test.py
    rm ayarlar.py

ZararTedarikci/ZararKiralik tablolarının satır kapasitesi ('Rows.Insert' ile
otomatik büyütme) — ÖZELLİKLE ListObject'in native "Toplam Satırı"
(`ShowTotals=True`) AÇIKKEN ekleme konumunun doğru hesaplandığını doğrular
(bkz. kpi_sablon_rapor.py `_com_zarar_detay_tablo_yaz`/`_com_zarar_detay_
kapasite_arttir` — "ZararTedarikci başarısız, ZararKiralik başarılı"
asimetrisinin kök nedeni). Tarih biçimi doğrulama testleri için ayrıca
`kpi_tarih_bicimi_test.py`'ye bakın.
"""
from __future__ import annotations

from typing import Any

from kpi_sablon_rapor import (
    _com_zarar_detay_kapasite_arttir,
    _com_zarar_detay_tablo_yaz,
)


def _assert(cond: bool, msg: str) -> None:
    if not cond:
        raise AssertionError(msg)


# --------------------------------------------------------------------------
# Basit Excel COM taklidi (Range/Cells/ListObject) — sadece bu test dosyasının
# ihtiyaç duyduğu özellik/metotları uygular, gerçek pywin32 COM nesnesi DEĞİL.
# --------------------------------------------------------------------------


class _FakeCount:
    def __init__(self, n: int) -> None:
        self.Count = n


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

    @property
    def NumberFormat(self):
        return self.sheet.numfmt.get(self._k(), "General")

    @NumberFormat.setter
    def NumberFormat(self, v):
        self.sheet.numfmt[self._k()] = v

    @property
    def NumberFormatLocal(self):
        return self.sheet.numfmt_local.get(self._k(), "General")

    @NumberFormatLocal.setter
    def NumberFormatLocal(self, v):
        self.sheet.numfmt_local[self._k()] = v


class _FakeRange:
    def __init__(self, sheet: "_FakeSheet", r1: int, c1: int, r2: int, c2: int) -> None:
        self.sheet = sheet
        self.r1, self.c1, self.r2, self.c2 = r1, c1, r2, c2

    @property
    def Row(self):
        return self.r1

    @property
    def Column(self):
        return self.c1

    @property
    def Rows(self):
        return _FakeCount(self.r2 - self.r1 + 1)

    @property
    def Columns(self):
        return _FakeCount(self.c2 - self.c1 + 1)

    def Cells(self, ro: int, co: int) -> _FakeCell:
        return _FakeCell(self.sheet, self.r1 + ro - 1, self.c1 + co - 1)

    @property
    def Value(self):
        rows = [
            tuple(self.sheet.values.get((r, c)) for c in range(self.c1, self.c2 + 1))
            for r in range(self.r1, self.r2 + 1)
        ]
        return rows[0] if len(rows) == 1 else tuple(rows)

    @Value.setter
    def Value(self, v):
        # Gerçek Excel COM davranışı: tek satırlık aralığa DÜZ bir tuple
        # atanabilir; çok satırlıya tuple-of-tuples gerekir (bkz.
        # `_com_araliga_yaz`, kpi_sablon_rapor.py).
        if self.r2 == self.r1:
            for i, val in enumerate(v):
                self.sheet.values[(self.r1, self.c1 + i)] = val
        else:
            for ri, satir in enumerate(v):
                for ci, val in enumerate(satir):
                    self.sheet.values[(self.r1 + ri, self.c1 + ci)] = val

    @property
    def NumberFormat(self):
        return self.sheet.numfmt.get((self.r1, self.c1), "General")

    @NumberFormat.setter
    def NumberFormat(self, v):
        for r in range(self.r1, self.r2 + 1):
            for c in range(self.c1, self.c2 + 1):
                self.sheet.numfmt[(r, c)] = v

    @property
    def NumberFormatLocal(self):
        return self.sheet.numfmt_local.get((self.r1, self.c1), "General")

    @NumberFormatLocal.setter
    def NumberFormatLocal(self, v):
        for r in range(self.r1, self.r2 + 1):
            for c in range(self.c1, self.c2 + 1):
                self.sheet.numfmt_local[(r, c)] = v

    @property
    def FormulaR1C1(self):
        return self.sheet.formula_r1c1.get((self.r1, self.c1), "")

    @FormulaR1C1.setter
    def FormulaR1C1(self, v):
        for r in range(self.r1, self.r2 + 1):
            for c in range(self.c1, self.c2 + 1):
                self.sheet.formula_r1c1[(r, c)] = v


class _FakeListRows:
    def __init__(self, lo: "_FakeListObject") -> None:
        self.lo = lo

    @property
    def Count(self):
        return self.lo.data_row_count

    def Add(self, Position=None, AlwaysInsert=None):
        self.lo.add_calls.append((Position, AlwaysInsert))
        self.lo.data_row_count += 1
        return None


class _FakeListObject:
    def __init__(
        self,
        sheet: "_FakeSheet",
        name: str,
        header_row: int,
        tablo_sol: int,
        kolon_sayisi: int,
        data_row_count: int,
        show_totals: bool,
    ) -> None:
        self.sheet = sheet
        self.Name = name
        self.header_row = header_row
        self.tablo_sol = tablo_sol
        self.kolon_sayisi = kolon_sayisi
        self.data_row_count = data_row_count
        self.ShowTotals = show_totals
        self.resize_calls: list[tuple[int, int]] = []
        self.add_calls: list[tuple[Any, Any]] = []

    @property
    def HeaderRowRange(self) -> _FakeRange:
        return _FakeRange(
            self.sheet, self.header_row, self.tablo_sol,
            self.header_row, self.tablo_sol + self.kolon_sayisi - 1,
        )

    @property
    def Range(self) -> _FakeRange:
        toplam_satir = 1 + self.data_row_count + (1 if self.ShowTotals else 0)
        return _FakeRange(
            self.sheet, self.header_row, self.tablo_sol,
            self.header_row + toplam_satir - 1, self.tablo_sol + self.kolon_sayisi - 1,
        )

    @property
    def ListRows(self) -> _FakeListRows:
        return _FakeListRows(self)

    def Resize(self, rng: _FakeRange) -> None:
        self.resize_calls.append((rng.r1, rng.r2))
        self.data_row_count = rng.r2 - self.header_row


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
    def __init__(self) -> None:
        self.values: dict[tuple[int, int], Any] = {}
        self.formulas: dict[tuple[int, int], str] = {}
        self.formula_r1c1: dict[tuple[int, int], str] = {}
        self.numfmt: dict[tuple[int, int], str] = {}
        self.numfmt_local: dict[tuple[int, int], str] = {}
        self._list_objects: dict[str, _FakeListObject] = {}
        self.insert_calls: list[tuple[int, int]] = []
        self.insert_hatasi: str | None = None

    def kayit_ekle_list_object(self, lo: _FakeListObject) -> None:
        self._list_objects[lo.Name] = lo

    def ListObjects(self, name: str) -> _FakeListObject:
        return self._list_objects[name]

    def Cells(self, r: int, c: int) -> _FakeCell:
        return _FakeCell(self, r, c)

    def Range(self, cell1: _FakeCell, cell2: _FakeCell) -> _FakeRange:
        return _FakeRange(self, cell1.row, cell1.col, cell2.row, cell2.col)

    def Rows(self, spec: str) -> _FakeRowsHandle:
        return _FakeRowsHandle(self, spec)

    def _satirlari_kaydir(self, baslangic: int, n: int) -> None:
        for depo in (
            self.values, self.formulas, self.formula_r1c1,
            self.numfmt, self.numfmt_local,
        ):
            anahtarlar = sorted(
                (k for k in depo if k[0] >= baslangic), key=lambda k: -k[0]
            )
            for (r, c) in anahtarlar:
                depo[(r + n, c)] = depo.pop((r, c))


# --------------------------------------------------------------------------
# Yardımcı: standart bir Zarar Detay tablosu (9 metin + 5 formül sütunu) kur.
# --------------------------------------------------------------------------

_METIN_KOLON_SAYISI = 9  # ZARAR_DETAY_METIN_SUTUNLARI ile birebir
_FORMUL_KOLON_SAYISI = 5  # J..N: Alış/Satış/Kâr-Zarar/Zarar %/Zarar Payı
_TOPLAM_KOLON = _METIN_KOLON_SAYISI + _FORMUL_KOLON_SAYISI  # 14


def _tablo_kur(
    header_row: int, veri_satir_sayisi: int, show_totals: bool
) -> tuple[_FakeSheet, _FakeListObject, int, int]:
    sheet = _FakeSheet()
    tablo_sol = 1
    veri_ilk = header_row + 1
    veri_son = veri_ilk + veri_satir_sayisi - 1

    lo = _FakeListObject(
        sheet, "TestTablo", header_row, tablo_sol, _TOPLAM_KOLON,
        veri_satir_sayisi, show_totals,
    )
    sheet.kayit_ekle_list_object(lo)

    # İlk veri satırının J..N formülleri (kaynak — yeni satırlara kopyalanacak).
    for offset in range(_FORMUL_KOLON_SAYISI):
        col = tablo_sol + _METIN_KOLON_SAYISI + offset
        sheet.formula_r1c1[(veri_ilk, col)] = f"=SUMIF(kaynak{offset})"

    # 'Ara Toplam' satırı — native ShowTotals AÇIKSA tablonun kendi Toplam
    # Satırı (veri_son+1), KAPALIYSA tablonun dışındaki elle eklenmiş satır
    # (fiziksel konum ikisinde de AYNI: veri_son+1).
    ara_toplam_satir = veri_son + 1
    for offset in range(_FORMUL_KOLON_SAYISI):
        col = tablo_sol + _METIN_KOLON_SAYISI + offset
        col_harf = chr(ord("A") + col - 1)  # yalnızca test kolonları için (A..N) yeterli
        sheet.formulas[(ara_toplam_satir, col)] = (
            f"=SUM({col_harf}{veri_ilk}:{col_harf}{veri_son})"
        )

    return sheet, lo, veri_ilk, ara_toplam_satir


def test_showtotals_acikken_dogru_konuma_ekleniyor() -> None:
    """✅ KRİTİK REGRESYON — ShowTotals=True'da eskiden Insert() native Toplam
    Satırının ÜZERİNE (veya ondan sonrasına) çağrılıp 'Insert method of Range
    class failed' hatası veriyordu. Artık `lo.ListRows.Count` kullanıldığı
    için ekleme konumu GERÇEK son veri satırının üzerinde olmalı."""
    header_row = 5
    sheet, lo, veri_ilk, eski_ara_toplam = _tablo_kur(
        header_row, veri_satir_sayisi=3, show_totals=True
    )
    eski_veri_son = veri_ilk + 3 - 1  # 8

    satirlar = [{"Tarih": f"0{i}.08.2026", "Sevk No": f"S-{i}"} for i in range(1, 6)]  # 5 satır > 3 kapasite

    yazilan, tasan = _com_zarar_detay_tablo_yaz(sheet, "TestTablo", satirlar)

    _assert(yazilan == 5 and tasan == 0, f"5 satırın hepsi yazılmalıydı, yazilan={yazilan} tasan={tasan}")
    _assert(lo.data_row_count == 5, f"ListObject veri satır sayısı 5 olmalı, {lo.data_row_count} çıktı")

    _assert(len(sheet.insert_calls) == 1, "Tam olarak bir bulk Insert() çağrısı olmalı")
    a, b = sheet.insert_calls[0]
    _assert(
        a == eski_veri_son,
        f"Insert() eski SON VERİ satırından ({eski_veri_son}) başlamalı, "
        f"native Toplam Satırından ({eski_veri_son + 1}) DEĞİL — {a} bulundu "
        "(ShowTotals düzeltmesi çalışmıyor demektir)",
    )
    _assert(b == eski_veri_son + 1, f"Insert() 2 satır eklemeli (8:9 gibi), {a}:{b} bulundu")

    yeni_veri_son = eski_veri_son + 2  # 10
    yeni_ara_toplam = yeni_veri_son + 1  # 11
    _assert(
        (yeni_ara_toplam, 10) in sheet.formulas,
        "Eski 'Ara Toplam' (native Toplam Satırı) satırı doğru şekilde aşağı kaymalı",
    )
    kaydirilmis_formul = sheet.formulas[(yeni_ara_toplam, 10)]
    _assert(
        str(eski_veri_son) not in kaydirilmis_formul,
        f"'Ara Toplam' formülü YENİ son satırı ({yeni_veri_son}) kapsayacak şekilde "
        f"düzeltilmeli, hâlâ eski satır ({eski_veri_son}) referansı içeriyor: {kaydirilmis_formul}",
    )
    _assert(
        str(yeni_veri_son) in kaydirilmis_formul,
        f"'Ara Toplam' formülü yeni son satırı ({yeni_veri_son}) içermeli: {kaydirilmis_formul}",
    )


def test_showtotals_kapaliyken_de_dogru_calisiyor() -> None:
    """Sanity: ShowTotals=False (ZararKiralik'in eski davranışı, hep başarılıydı)
    hâlâ doğru çalışmalı — regresyon YOK."""
    header_row = 5
    sheet, lo, veri_ilk, _ = _tablo_kur(header_row, veri_satir_sayisi=4, show_totals=False)
    eski_veri_son = veri_ilk + 4 - 1

    satirlar = [{"Sevk No": f"S-{i}"} for i in range(1, 7)]  # 6 satır > 4 kapasite

    yazilan, tasan = _com_zarar_detay_tablo_yaz(sheet, "TestTablo", satirlar)

    _assert(yazilan == 6 and tasan == 0, f"6 satırın hepsi yazılmalıydı, yazilan={yazilan} tasan={tasan}")
    _assert(len(sheet.insert_calls) == 1, "Tam olarak bir bulk Insert() çağrısı olmalı")
    a, _b = sheet.insert_calls[0]
    _assert(a == eski_veri_son, f"Insert() eski son veri satırından ({eski_veri_son}) başlamalı, {a} bulundu")


def test_kapasite_yeterliyse_insert_hic_cagrilmiyor() -> None:
    header_row = 5
    sheet, lo, veri_ilk, _ = _tablo_kur(header_row, veri_satir_sayisi=10, show_totals=True)
    satirlar = [{"Sevk No": f"S-{i}"} for i in range(1, 4)]  # 3 satır <= 10 kapasite

    yazilan, tasan = _com_zarar_detay_tablo_yaz(sheet, "TestTablo", satirlar)

    _assert(yazilan == 3, f"3 satır yazılmalıydı, {yazilan} çıktı")
    _assert(len(sheet.insert_calls) == 0, "Kapasite zaten yeterliyse Insert() çağrılmamalı")
    _assert(lo.data_row_count == 10, "Kapasite değişmemeli")


def test_rows_insert_basarisiz_olursa_listrows_add_yedegine_dusuyor() -> None:
    """GÜVENLİK AĞI: `Rows.Insert` (ShowTotals düzeltmesine RAĞMEN, beklenmedik
    bir nedenle) başarısız olursa kod `ListRows.Add(AlwaysInsert=True)`
    yedeğine düşmeli, sessizce/hiç hata fırlatmadan devam etmeli."""
    header_row = 5
    sheet, lo, veri_ilk, _ = _tablo_kur(header_row, veri_satir_sayisi=3, show_totals=True)
    veri_son = veri_ilk + 3 - 1
    sheet.insert_hatasi = (
        "(-2147352567, 'Exception occurred.', (0, 'Microsoft Excel', "
        "'Range sınıfının Insert yöntemi başarısız', 'xlmain11.chm', 0, "
        "-2146827284), None)"
    )

    yeni_kapasite, yeni_veri_son = _com_zarar_detay_kapasite_arttir(
        sheet, lo, "TestTablo", tablo_sol=1, veri_ilk_satir=veri_ilk,
        veri_son_satir=veri_son, kapasite=3, gereken_satir_sayisi=5,
    )

    _assert(yeni_kapasite == 5, f"Kapasite 5'e büyümeliydi (ListRows.Add yedeğiyle), {yeni_kapasite} çıktı")
    _assert(yeni_veri_son == veri_son + 2, "Yeni son veri satırı 2 artmalı")
    _assert(len(sheet.insert_calls) == 0, "Rows.Insert başarısız olduğu için hiç kayıt bırakmamalı")
    _assert(
        len(lo.add_calls) == 2,
        f"ListRows.Add TAM 2 kere (eksik satır sayısı) çağrılmalı, {len(lo.add_calls)} çıktı",
    )
    _assert(
        all(always_insert is True for _pos, always_insert in lo.add_calls),
        "Her ListRows.Add çağrısı AlwaysInsert=True ile yapılmalı (altındaki hücreleri kaydırsın)",
    )


def main() -> None:
    test_showtotals_acikken_dogru_konuma_ekleniyor()
    test_showtotals_kapaliyken_de_dogru_calisiyor()
    test_kapasite_yeterliyse_insert_hic_cagrilmiyor()
    test_rows_insert_basarisiz_olursa_listrows_add_yedegine_dusuyor()
    print("OK — tüm Zarar Detay kapasite testleri geçti")


if __name__ == "__main__":
    main()

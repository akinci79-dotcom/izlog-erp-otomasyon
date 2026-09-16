"""VERİ/Filo Detay sayfalarını yazan `_com_sayfaya_yaz()` için kapasite
büyütme testleri — Excel COM olmadan, basit Python nesneleriyle (Range/
ListObject taklidi) çalışır.

Çalıştırma (ayarlar.py gerektirir, kpi_sablon_rapor.py modül seviyesinde
`import ayarlar` yaptığı için):
    cp ayarlar.example.py ayarlar.py
    python3 kpi_sayfa_yazma_kapasite_test.py
    rm ayarlar.py

KÖK NEDEN (kullanıcı ekran görüntüsüyle bildirdi — "filo detay sayfasınıda
bozmuş genişletmek yerine toplam satırına bilgi basmış"): `_com_sayfaya_yaz()`
(HEM "VERİ" HEM "Filo Detay" sayfalarını yazan ORTAK fonksiyon) eskiden
`lo.Resize()`'ı doğrudan çağırıp ardından TÜM yeni aralığı toplu olarak
yazıyordu. `ListObject.Resize` gerçek bir satır EKLEMEZ/kaydırmaz — sadece
tablonun kapladığı ALANI yeniden tanımlar (bkz. Microsoft Learn:
"No cells are inserted or moved"). Bu ay gelen satır sayısı (`len(matris)`)
önceki kapasiteyi (`lo.ListRows.Count`) aşarsa, tablonun eski sınırının
HEMEN ALTINDAKİ herhangi bir içerik (örn. şablonda elle konmuş bir
'Genel Toplam' satırı) tabloya "yutuluyor" ve hemen ardından gerçek veriyle
EZİLİYORDU.

Bu testler, `_com_zarar_detay_kapasite_arttir` ile PAYLAŞILAN yeni
`_com_tablo_satir_ekle` yardımcısının, `_com_sayfaya_yaz` içinde de doğru
şekilde devreye girdiğini ve şu senaryoları doğru ele aldığını doğrular:
  (a) Altında GERÇEKTEN BOŞ satırlar varken büyüme — regresyon yok.
  (b) Altında ÖNCEDEN DOLU bir 'Genel Toplam' benzeri içerik varken büyüme —
      o içerik artık KORUNUYOR (aşağı kayıyor), üzerine yazılmıyor.
  (c) Tablo KÜÇÜLÜRKEN — regresyon yok (eski davranış değişmedi).
  (d) `Rows.Insert` başarısız olup `ListRows.Add` yedeğine düşme senaryosu.
"""
from __future__ import annotations

from typing import Any

from kpi_kiralk_arac import FILO_DETAY_SUTUNLARI
from kpi_sablon_rapor import _com_sayfaya_yaz


def _assert(cond: bool, msg: str) -> None:
    if not cond:
        raise AssertionError(msg)


# --------------------------------------------------------------------------
# Basit Excel COM taklidi (Range/Cells/ListObject) — kpi_zarar_kapasite_test.py
# ile AYNI desen, `_com_sayfaya_yaz`'ın ihtiyaç duyduğu ek özelliklerle
# (sheet.Name, sheet.ListObjects koleksiyonu — index VE ada göre erişim).
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
    ) -> None:
        self.sheet = sheet
        self.Name = name
        self.header_row = header_row
        self.tablo_sol = tablo_sol
        self.kolon_sayisi = kolon_sayisi
        self.data_row_count = data_row_count
        self.resize_calls: list[tuple[int, int]] = []
        self.add_calls: list[tuple[Any, Any]] = []

    @property
    def HeaderRowRange(self) -> _FakeRange:
        return _FakeRange(
            self.sheet, self.header_row, self.tablo_sol,
            self.header_row, self.tablo_sol + self.kolon_sayisi - 1,
        )

    @property
    def ListRows(self) -> _FakeListRows:
        return _FakeListRows(self)

    def Resize(self, rng: _FakeRange) -> None:
        self.resize_calls.append((rng.r1, rng.r2))
        self.data_row_count = rng.r2 - self.header_row


class _FakeListObjectsCollection:
    """`sheet.ListObjects.Count` (attribute) VE `sheet.ListObjects(i)` (index
    veya ad ile çağrı) — gerçek Excel COM koleksiyonunun davranışını taklit
    eder (bkz. `_com_listobject_bul`, kpi_sablon_rapor.py)."""

    def __init__(self) -> None:
        self._by_index: dict[int, _FakeListObject] = {}
        self._by_name: dict[str, _FakeListObject] = {}

    @property
    def Count(self):
        return len(self._by_index)

    def ekle(self, lo: _FakeListObject) -> None:
        idx = len(self._by_index) + 1
        self._by_index[idx] = lo
        self._by_name[lo.Name] = lo

    def __call__(self, key):
        if isinstance(key, int):
            return self._by_index[key]
        return self._by_name[key]


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
    def __init__(self, name: str = "TestSheet") -> None:
        self.Name = name
        self.values: dict[tuple[int, int], Any] = {}
        self.numfmt: dict[tuple[int, int], str] = {}
        self.numfmt_local: dict[tuple[int, int], str] = {}
        self.ListObjects = _FakeListObjectsCollection()
        self.insert_calls: list[tuple[int, int]] = []
        self.insert_hatasi: str | None = None

    def Cells(self, r: int, c: int) -> _FakeCell:
        return _FakeCell(self, r, c)

    def Range(self, cell1: _FakeCell, cell2: _FakeCell) -> _FakeRange:
        return _FakeRange(self, cell1.row, cell1.col, cell2.row, cell2.col)

    def Rows(self, spec: str) -> _FakeRowsHandle:
        return _FakeRowsHandle(self, spec)

    def _satirlari_kaydir(self, baslangic: int, n: int) -> None:
        for depo in (self.values, self.numfmt, self.numfmt_local):
            anahtarlar = sorted(
                (k for k in depo if k[0] >= baslangic), key=lambda k: -k[0]
            )
            for (r, c) in anahtarlar:
                depo[(r + n, c)] = depo.pop((r, c))


# --------------------------------------------------------------------------
# Yardımcı: Filo Detay benzeri bir tablo kur (header + N veri satırı kapasitesi,
# isteğe bağlı olarak tablonun HEMEN ALTINDA "Genel Toplam" benzeri içerik).
# --------------------------------------------------------------------------


def _tablo_kur(
    header_row: int,
    basliklar: list[str],
    mevcut_veri_satir_sayisi: int,
    tablo_sol: int = 1,
    alt_icerik: dict[int, Any] | None = None,
) -> tuple[_FakeSheet, _FakeListObject]:
    sheet = _FakeSheet()
    kolon_sayisi = len(basliklar)
    for i, baslik in enumerate(basliklar):
        sheet.values[(header_row, tablo_sol + i)] = baslik

    lo = _FakeListObject(
        sheet, "TestTablo", header_row, tablo_sol, kolon_sayisi, mevcut_veri_satir_sayisi
    )
    sheet.ListObjects.ekle(lo)

    if alt_icerik:
        alt_satir = header_row + 1 + mevcut_veri_satir_sayisi
        for col_offset, deger in alt_icerik.items():
            sheet.values[(alt_satir, tablo_sol + col_offset)] = deger

    return sheet, lo


def _satir_degerleri_oku(
    sheet: _FakeSheet, row: int, tablo_sol: int, kolon_sayisi: int
) -> tuple[Any, ...]:
    return tuple(sheet.values.get((row, tablo_sol + i)) for i in range(kolon_sayisi))


# --------------------------------------------------------------------------
# (a) Altında GERÇEKTEN BOŞ satırlar varken büyüme — regresyon yok.
# --------------------------------------------------------------------------


def test_a_bos_alt_satirlarla_buyume_regresyon_yok() -> None:
    header_row = 5
    basliklar = ["YUK_NO", "SEVK_NO", "SATIS_TUTAR"]
    sheet, lo = _tablo_kur(header_row, basliklar, mevcut_veri_satir_sayisi=3)

    satirlar = [
        {"YUK_NO": f"Y-{i}", "SEVK_NO": f"S-{i}", "SATIS_TUTAR": i * 1000}
        for i in range(1, 6)  # 5 satır > 3 kapasite
    ]

    yazilan, esleme, eslesmeyen, kolon_sayisi, tablo_sol = _com_sayfaya_yaz(
        sheet, header_row, satirlar
    )

    _assert(yazilan == 5, f"5 satır yazılmalıydı, {yazilan} çıktı")
    _assert(lo.data_row_count == 5, f"Tablo kapasitesi 5'e büyümeliydi, {lo.data_row_count} çıktı")
    _assert(len(sheet.insert_calls) == 1, "Kapasite yetersiz kaldığı için TEK bir Insert() çağrısı olmalı")

    veri_ilk = header_row + 1
    for i, satir in enumerate(satirlar):
        okunan = _satir_degerleri_oku(sheet, veri_ilk + i, tablo_sol, kolon_sayisi)
        beklenen = (satir["YUK_NO"], satir["SEVK_NO"], float(satir["SATIS_TUTAR"]))
        _assert(okunan == beklenen, f"Satır {i + 1} yanlış yazıldı: {okunan} != {beklenen}")

    # Tablonun hemen altındaki satır (önceden boştu) hâlâ boş olmalı.
    alt_satir = veri_ilk + 5
    _assert(
        all(v is None for v in _satir_degerleri_oku(sheet, alt_satir, tablo_sol, kolon_sayisi)),
        "Tablonun altında önceden hiçbir şey yoktu, hâlâ boş olmalı",
    )


# --------------------------------------------------------------------------
# (b) Altında ÖNCEDEN DOLU 'Genel Toplam' benzeri içerik varken büyüme —
#     o içerik ARTIK KORUNMALI (aşağı kaymalı), üzerine YAZILMAMALI.
# --------------------------------------------------------------------------


def test_b_genel_toplam_satirinin_ustune_yazilmiyor() -> None:
    header_row = 10
    basliklar = list(FILO_DETAY_SUTUNLARI)
    kapasite = 3
    sheet, lo = _tablo_kur(
        header_row,
        basliklar,
        mevcut_veri_satir_sayisi=kapasite,
        alt_icerik={0: "GENEL TOPLAM", 1: "", 4: 999999},  # ARAC_KODU, ARAC_TIPI, TOPLAM_KM sütunları
    )
    tablo_sol = 1
    kolon_sayisi = len(basliklar)
    veri_ilk = header_row + 1
    eski_veri_son = veri_ilk + kapasite - 1  # 12
    eski_alt_satir = eski_veri_son + 1  # 13 — 'Genel Toplam' başlangıçta burada

    # Önce eski konumda gerçekten "Genel Toplam" olduğunu doğrula (sağlama).
    _assert(
        sheet.values[(eski_alt_satir, tablo_sol)] == "GENEL TOPLAM",
        "Test kurulumu hatalı: 'Genel Toplam' beklenen eski konumda değil",
    )

    satirlar = [
        {**{k: None for k in FILO_DETAY_SUTUNLARI}, "ARAC_KODU": f"07ABG{i:03d}", "TOPLAM_KM": 1000 + i}
        for i in range(1, 6)  # 5 gerçek araç satırı > 3 kapasite
    ]

    yazilan, _, _, _, _ = _com_sayfaya_yaz(
        sheet, header_row, satirlar, sabit_kolonlar=FILO_DETAY_SUTUNLARI
    )
    eksik = len(satirlar) - kapasite  # 2

    _assert(yazilan == 5, f"5 araç satırı yazılmalıydı, {yazilan} çıktı")
    _assert(lo.data_row_count == 5, f"Tablo kapasitesi 5'e büyümeliydi, {lo.data_row_count} çıktı")
    _assert(len(sheet.insert_calls) == 1, "TEK bir bulk Insert() çağrısı olmalı")
    _assert(
        sheet.insert_calls[0] == (eski_veri_son, eski_veri_son + eksik - 1),
        f"Insert() eski son veri satırının ({eski_veri_son}) TAM ÜZERİNE yapılmalı, "
        f"{sheet.insert_calls[0]} bulundu",
    )

    yeni_veri_son = eski_veri_son + eksik  # 14
    yeni_alt_satir = yeni_veri_son + 1  # 15

    # 🔑 KRİTİK DOĞRULAMA: 'Genel Toplam' artık YENİ (kaymış) konumda, İÇERİĞİ
    # DEĞİŞMEDEN duruyor — gerçek araç verisiyle EZİLMEDİ.
    _assert(
        sheet.values.get((yeni_alt_satir, tablo_sol)) == "GENEL TOPLAM",
        f"'Genel Toplam' yeni konuma ({yeni_alt_satir}) KORUNARAK kaymalıydı, "
        f"bulunan: {sheet.values.get((yeni_alt_satir, tablo_sol))!r}",
    )
    _assert(
        sheet.values.get((yeni_alt_satir, tablo_sol + 4)) == 999999,
        "'Genel Toplam' satırının diğer hücreleri de (TOPLAM_KM sütunu) korunmalı",
    )

    # Eski konum (artık tablo İÇİNDE, kapasite. sıradaki araç satırı) artık
    # 'Genel Toplam' DEĞİL, gerçek araç verisi içermeli.
    beklenen_arac_kodu = satirlar[kapasite]["ARAC_KODU"]  # "07ABG004"
    _assert(
        sheet.values.get((eski_alt_satir, tablo_sol)) == beklenen_arac_kodu,
        f"Eski 'Genel Toplam' konumu artık gerçek araç verisiyle ({beklenen_arac_kodu}) "
        f"dolu olmalı, bulunan: {sheet.values.get((eski_alt_satir, tablo_sol))!r}",
    )

    # Tüm 5 araç satırı sırasıyla doğru yazılmış olmalı.
    arac_kodu_offset = FILO_DETAY_SUTUNLARI.index("ARAC_KODU")
    for i, satir in enumerate(satirlar):
        deger = sheet.values.get((veri_ilk + i, tablo_sol + arac_kodu_offset))
        _assert(
            deger == satir["ARAC_KODU"],
            f"Satır {i + 1} ARAC_KODU yanlış: {deger!r} != {satir['ARAC_KODU']!r}",
        )


# --------------------------------------------------------------------------
# (c) Tablo KÜÇÜLÜRKEN (bu ay geçen aydan az satır) — regresyon yok.
# --------------------------------------------------------------------------


def test_c_kucculme_regresyon_yok() -> None:
    header_row = 5
    basliklar = ["YUK_NO", "SEVK_NO", "SATIS_TUTAR"]
    kapasite = 6
    sheet, lo = _tablo_kur(header_row, basliklar, mevcut_veri_satir_sayisi=kapasite)

    tablo_sol = 1
    veri_ilk = header_row + 1
    # Eski (geçen ay) 6 satırlık veriyi elle döşe — küçülme sonrası bu eski
    # satırların ne olduğunu (temizlenip temizlenmediğini) gözlemleyeceğiz.
    for i in range(kapasite):
        sheet.values[(veri_ilk + i, tablo_sol)] = f"ESKI-{i + 1}"

    satirlar = [
        {"YUK_NO": f"Y-{i}", "SEVK_NO": f"S-{i}", "SATIS_TUTAR": i * 100}
        for i in range(1, 4)  # 3 satır < 6 kapasite -> küçülme
    ]

    yazilan, _, _, _, _ = _com_sayfaya_yaz(sheet, header_row, satirlar)

    _assert(yazilan == 3, f"3 satır yazılmalıydı, {yazilan} çıktı")
    _assert(len(sheet.insert_calls) == 0, "Küçülme senaryosunda HİÇ Insert() çağrılmamalı (regresyon)")
    _assert(lo.data_row_count == 3, f"Tablo kapasitesi 3'e küçülmeliydi (Resize), {lo.data_row_count} çıktı")

    for i, satir in enumerate(satirlar):
        deger = sheet.values.get((veri_ilk + i, tablo_sol))
        _assert(deger == satir["YUK_NO"], f"Satır {i + 1} yanlış yazıldı: {deger!r}")

    # Eski davranış (bu düzeltmenin KAPSAMI DIŞINDA): tablo dışında kalan
    # eski satırlar (4-6) fiziksel olarak SİLİNMEDEN sayfada kalır. Burada
    # SADECE bunun hâlâ eskisi gibi çalıştığını (bu değişiklikle KIRILMADIĞINI)
    # doğruluyoruz — davranışın kendisi bu görevin konusu değil.
    for i in range(3, kapasite):
        deger = sheet.values.get((veri_ilk + i, tablo_sol))
        _assert(
            deger == f"ESKI-{i + 1}",
            f"Küçülme sonrası tablo dışında kalan eski satır ({i + 1}) beklenmedik şekilde "
            f"değişti: {deger!r} — bu davranış bu düzeltmeyle DEĞİŞMEMELİYDİ",
        )


# --------------------------------------------------------------------------
# (d) `Rows.Insert` başarısız -> `ListRows.Add` yedeğine düşme.
# --------------------------------------------------------------------------


def test_d_rows_insert_basarisiz_olursa_listrows_add_yedegine_dusuyor() -> None:
    header_row = 10
    basliklar = list(FILO_DETAY_SUTUNLARI)
    kapasite = 3
    sheet, lo = _tablo_kur(
        header_row,
        basliklar,
        mevcut_veri_satir_sayisi=kapasite,
        alt_icerik={0: "GENEL TOPLAM"},
    )
    sheet.insert_hatasi = (
        "(-2147352567, 'Exception occurred.', (0, 'Microsoft Excel', "
        "'Range sınıfının Insert yöntemi başarısız', 'xlmain11.chm', 0, "
        "-2146827284), None)"
    )

    satirlar = [
        {**{k: None for k in FILO_DETAY_SUTUNLARI}, "ARAC_KODU": f"07ABG{i:03d}"}
        for i in range(1, 6)  # 5 satır > 3 kapasite
    ]

    yazilan, _, _, _, _ = _com_sayfaya_yaz(
        sheet, header_row, satirlar, sabit_kolonlar=FILO_DETAY_SUTUNLARI
    )
    eksik = len(satirlar) - kapasite  # 2

    _assert(yazilan == 5, f"5 satır yazılmalıydı (yedek yöntemle de olsa), {yazilan} çıktı")
    _assert(len(sheet.insert_calls) == 0, "Rows.Insert başarısız olduğu için hiç kayıt bırakmamalı")
    _assert(
        len(lo.add_calls) == eksik,
        f"ListRows.Add TAM {eksik} kere (eksik satır sayısı) çağrılmalı, {len(lo.add_calls)} çıktı",
    )
    _assert(
        all(always_insert is True for _pos, always_insert in lo.add_calls),
        "Her ListRows.Add çağrısı AlwaysInsert=True ile yapılmalı (altındaki hücreleri kaydırsın)",
    )
    _assert(lo.data_row_count == 5, f"Tablo kapasitesi 5'e büyümeliydi, {lo.data_row_count} çıktı")

    # Veri yine de doğru yazılmış olmalı (yedek yöntem devreye girse de akış
    # kesilmemeli — _com_sayfaya_yaz, _com_tablo_satir_ekle'nin RuntimeError
    # fırlatmadığı bu senaryoda sorunsuz devam etmeli).
    arac_kodu_offset = FILO_DETAY_SUTUNLARI.index("ARAC_KODU")
    tablo_sol = 1
    veri_ilk = header_row + 1
    for i, satir in enumerate(satirlar):
        deger = sheet.values.get((veri_ilk + i, tablo_sol + arac_kodu_offset))
        _assert(
            deger == satir["ARAC_KODU"],
            f"Satır {i + 1} ARAC_KODU yanlış: {deger!r} != {satir['ARAC_KODU']!r}",
        )


def main() -> None:
    test_a_bos_alt_satirlarla_buyume_regresyon_yok()
    test_b_genel_toplam_satirinin_ustune_yazilmiyor()
    test_c_kucculme_regresyon_yok()
    test_d_rows_insert_basarisiz_olursa_listrows_add_yedegine_dusuyor()
    print("OK — tüm VERİ/Filo Detay sayfa yazma kapasite testleri geçti")


if __name__ == "__main__":
    main()

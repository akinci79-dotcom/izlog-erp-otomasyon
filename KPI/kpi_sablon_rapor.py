"""
Referans KPI şablonunu doldurur: VERİ + Filo Detay → pivot sayfaları Excel'de güncellenir.

Pivotlu şablonlar openpyxl ile kaydedildiğinde bozulabildiği için (Excel açamaz),
Windows'ta veri yazımı doğrudan Excel COM ile yapılır.

Kullanım:
  1. Temmuz KPI dosyanızı KPI/referans/kpi_sablon.xlsx olarak kaydedin
  2. python kpi_rapor_olustur.py
"""
from __future__ import annotations

import re
import shutil
import unicodedata
from datetime import date, datetime
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

import ayarlar
from kpi_kiralk_arac import FILO_DETAY_SUTUNLARI, kiralk_arac_detay_getir, kiralk_arac_semasi_hazir
from kpi_veri import hucre_degeri, veri_satirlari_getir, veri_semasi_hazir
from oracle_baglanti import baglanti_yonet

_KPI_KOKU = Path(__file__).resolve().parent


def _referans_klasoru() -> Path:
    return _KPI_KOKU / "referans"


def sablon_yolu() -> Path:
    dosya = getattr(ayarlar, "KPI_SABLON_DOSYASI", "kpi_sablon.xlsx")
    yol = Path(dosya)
    if yol.is_absolute():
        return yol
    return _referans_klasoru() / dosya


def _raporlar_klasoru() -> Path:
    klasor = _KPI_KOKU / "raporlar"
    klasor.mkdir(exist_ok=True)
    return klasor


def _cikti_yolu(sablon: Path | None = None) -> Path:
    dosya = getattr(ayarlar, "KPI_RAPOR_DOSYASI", None)
    if dosya:
        yol = Path(dosya)
        if yol.is_absolute():
            return yol
        return _raporlar_klasoru() / dosya
    suffix = (sablon or sablon_yolu()).suffix or ".xlsx"
    return _raporlar_klasoru() / f"kpi_rapor{suffix}"


def _normalize_kolon(adi: str) -> str:
    if adi is None:
        return ""
    metin = str(adi).strip().upper()
    metin = unicodedata.normalize("NFKD", metin)
    metin = "".join(c for c in metin if not unicodedata.combining(c))
    metin = re.sub(r"[^A-Z0-9]+", "_", metin)
    return metin.strip("_")


def _sayfa_bul(wb, adlar: list[str]) -> Worksheet | None:
    ad_norm = {a.strip().upper() for a in adlar}
    for sheet in wb.worksheets:
        if sheet.title.strip().upper() in ad_norm:
            return sheet
    for ad in adlar:
        if ad in wb.sheetnames:
            return wb[ad]
    return None


def _basliklari_oku(ws: Worksheet, satir: int = 1) -> list[str | None]:
    max_col = ws.max_column or 1
    return [ws.cell(row=satir, column=c).value for c in range(1, max_col + 1)]


def _kolon_esleme(basliklar: list[str | None], veri_anahtarlari: list[str]) -> dict[int, str]:
    normalized_veri = {_normalize_kolon(k): k for k in veri_anahtarlari}
    ek = getattr(ayarlar, "KPI_KOLON_ESLEME", {}) or {}
    for sablon, oracle in ek.items():
        normalized_veri[_normalize_kolon(sablon)] = oracle

    esleme: dict[int, str] = {}
    for col_idx, baslik in enumerate(basliklar, 1):
        if baslik is None or str(baslik).strip() == "":
            continue
        norm = _normalize_kolon(str(baslik))
        if norm in normalized_veri:
            esleme[col_idx] = normalized_veri[norm]
    return esleme


def _sayfayi_temizle_yaz(
    ws: Worksheet,
    baslik_satiri: int,
    satirlar: list[dict[str, Any]],
    sabit_kolonlar: list[str] | None = None,
):
    basliklar = _basliklari_oku(ws, baslik_satiri)
    if not satirlar:
        max_row = ws.max_row or baslik_satiri
        for row in range(baslik_satiri + 1, max_row + 1):
            for col in range(1, len(basliklar) + 1):
                ws.cell(row=row, column=col, value=None)
        return 0

    anahtarlar = sabit_kolonlar or list(satirlar[0].keys())
    esleme = _kolon_esleme(basliklar, anahtarlar)

    if not esleme and sabit_kolonlar:
        esleme = {i + 1: k for i, k in enumerate(sabit_kolonlar) if i < len(basliklar)}

    max_row = ws.max_row or baslik_satiri
    for row in range(baslik_satiri + 1, max(max_row, baslik_satiri + len(satirlar)) + 1):
        for col in range(1, len(basliklar) + 1):
            ws.cell(row=row, column=col, value=None)

    for i, satir in enumerate(satirlar):
        row_no = baslik_satiri + 1 + i
        if sabit_kolonlar:
            for col_idx, kolon in enumerate(sabit_kolonlar, 1):
                ws.cell(row=row_no, column=col_idx, value=hucre_degeri(satir.get(kolon)))
        else:
            for col_idx, kolon in esleme.items():
                ws.cell(row=row_no, column=col_idx, value=hucre_degeri(satir.get(kolon)))

    return len(satirlar)


def _hucre_gorunen_uzunluk(deger: Any) -> int:
    if deger is None:
        return 0
    if isinstance(deger, bool):
        return 4
    if isinstance(deger, (datetime, date)):
        return len(deger.strftime("%d.%m.%Y"))
    if isinstance(deger, float):
        metin = f"{deger:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        return len(metin) + 2
    if isinstance(deger, int):
        metin = f"{deger:,}".replace(",", ".")
        return len(metin)
    return len(str(deger))


def _sayfa_sutunlarini_genislet(
    ws: Worksheet,
    baslik_satiri: int = 1,
    veri_satir_sayisi: int = 0,
    min_genislik: float = 10,
    max_genislik: float = 55,
    padding: float = 2,
):
    """Doldurulan sayfada sütun genişliklerini içeriğe göre ayarlar (####### önler)."""
    son_satir = baslik_satiri + max(veri_satir_sayisi, 0)
    if son_satir < baslik_satiri:
        return

    max_col = ws.max_column or 1
    for col in range(1, max_col + 1):
        en_uzun = 0
        for row in range(baslik_satiri, son_satir + 1):
            deger = ws.cell(row=row, column=col).value
            en_uzun = max(en_uzun, _hucre_gorunen_uzunluk(deger))
        if en_uzun <= 0:
            continue
        harf = get_column_letter(col)
        mevcut = ws.column_dimensions[harf].width or min_genislik
        yeni = min(max(en_uzun + padding, min_genislik), max_genislik)
        ws.column_dimensions[harf].width = max(mevcut, yeni)


def _excel_kullanilabilir() -> bool:
    if getattr(ayarlar, "KPI_EXCEL_KULLAN", True) is False:
        return False
    try:
        import win32com.client  # type: ignore  # noqa: F401
    except ImportError:
        return False
    return True


def _com_satir_oku(deger: Any) -> list[Any]:
    if deger is None:
        return []
    if not isinstance(deger, tuple):
        return [deger]
    if not deger:
        return []
    if not isinstance(deger[0], tuple):
        return list(deger)
    return list(deger[0])


def _com_basliklari_oku(sheet, baslik_satiri: int) -> list[Any]:
    kullanilan = sheet.UsedRange
    max_col = max(int(kullanilan.Columns.Count), 1)
    ham = sheet.Range(
        sheet.Cells(baslik_satiri, 1),
        sheet.Cells(baslik_satiri, max_col),
    ).Value
    basliklar = _com_satir_oku(ham)
    while basliklar and basliklar[-1] is None:
        basliklar.pop()
    return basliklar


def _com_sayfaya_yaz(
    sheet,
    baslik_satiri: int,
    satirlar: list[dict[str, Any]],
    sabit_kolonlar: list[str] | None = None,
) -> int:
    basliklar = _com_basliklari_oku(sheet, baslik_satiri)
    kolon_sayisi = len(basliklar) if basliklar else max(int(sheet.UsedRange.Columns.Count), 1)

    kullanilan = sheet.UsedRange
    eski_son = max(int(kullanilan.Rows.Count), baslik_satiri + len(satirlar))
    if eski_son > baslik_satiri:
        sheet.Range(
            sheet.Cells(baslik_satiri + 1, 1),
            sheet.Cells(eski_son, kolon_sayisi),
        ).ClearContents()

    if not satirlar:
        return 0

    anahtarlar = sabit_kolonlar or list(satirlar[0].keys())
    esleme = _kolon_esleme(basliklar, anahtarlar)
    if not esleme and sabit_kolonlar:
        esleme = {i + 1: k for i, k in enumerate(sabit_kolonlar) if i < len(basliklar)}

    if sabit_kolonlar:
        kolon_sayisi = len(sabit_kolonlar)
    elif esleme:
        kolon_sayisi = max(esleme.keys())
    else:
        kolon_sayisi = len(anahtarlar)

    matris: list[tuple[Any, ...]] = []
    for satir in satirlar:
        if sabit_kolonlar:
            satir_verisi = tuple(hucre_degeri(satir.get(k)) for k in sabit_kolonlar)
        else:
            satir_verisi = [None] * kolon_sayisi
            for col_idx, kolon in esleme.items():
                satir_verisi[col_idx - 1] = hucre_degeri(satir.get(kolon))
            satir_verisi = tuple(satir_verisi)
        matris.append(satir_verisi)

    hedef = sheet.Range(
        sheet.Cells(baslik_satiri + 1, 1),
        sheet.Cells(baslik_satiri + len(matris), kolon_sayisi),
    )
    if len(matris) == 1:
        hedef.Value = matris[0]
    else:
        hedef.Value = tuple(matris)
    return len(satirlar)


def _com_pivot_kaynak_guncelle(
    wb,
    sayfa_adi: str,
    baslik_satiri: int,
    satir_sayisi: int,
    kolon_sayisi: int,
):
    if satir_sayisi <= 0:
        return
    son_satir = baslik_satiri + satir_sayisi
    son_kolon = get_column_letter(max(kolon_sayisi, 1))
    yeni_kaynak = f"'{sayfa_adi}'!$A${baslik_satiri}:${son_kolon}${son_satir}"
    try:
        for i in range(1, int(wb.PivotCaches().Count) + 1):
            pc = wb.PivotCaches(i)
            src = str(pc.SourceData or "")
            if sayfa_adi.upper() in src.upper():
                pc.SourceData = yeni_kaynak
    except Exception:
        pass


def _excel_uygulama_ac():
    import win32com.client  # type: ignore

    try:
        excel = win32com.client.gencache.EnsureDispatch("Excel.Application")
    except Exception:
        excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.ScreenUpdating = False
    excel.EnableEvents = False
    return excel


def _excel_sayfa_bul(wb, adlar: list[str]):
    ad_norm = {a.strip().upper() for a in adlar}
    for sheet in wb.Worksheets:
        if str(sheet.Name).strip().upper() in ad_norm:
            return sheet
    return None


def _excel_sablon_doldur(
    dosya_yolu: Path,
    veri_satirlari: list[dict[str, Any]],
    filo_satirlari: list[dict[str, Any]],
    veri_sayfa_adlari: list[str],
    filo_sayfa_adlari: list[str],
    veri_baslik_satiri: int,
    filo_baslik_satiri: int,
    pivot_yenile: bool,
    sutun_autofit: bool,
) -> tuple[bool, str | None, int, int]:
    """Şablon kopyasına Excel COM ile veri yazar, pivot yeniler, sütunları genişletir."""
    excel = None
    wb = None
    uyarilar: list[str] = []
    dosya = str(dosya_yolu.resolve())

    try:
        excel = _excel_uygulama_ac()
        wb = excel.Workbooks.Open(
            Filename=dosya,
            UpdateLinks=0,
            ReadOnly=False,
            Notify=False,
        )

        ws_veri = _excel_sayfa_bul(wb, veri_sayfa_adlari)
        if ws_veri is None:
            return False, f"VERİ sayfası bulunamadı: {veri_sayfa_adlari}", 0, 0

        ws_filo = _excel_sayfa_bul(wb, filo_sayfa_adlari)
        if ws_filo is None:
            return False, f"Filo Detay sayfası bulunamadı: {filo_sayfa_adlari}", 0, 0

        veri_adet = _com_sayfaya_yaz(ws_veri, veri_baslik_satiri, veri_satirlari)
        veri_kolon = len(_com_basliklari_oku(ws_veri, veri_baslik_satiri)) or int(ws_veri.UsedRange.Columns.Count)
        _com_pivot_kaynak_guncelle(wb, ws_veri.Name, veri_baslik_satiri, veri_adet, veri_kolon)

        filo_yaz = [{k: r.get(k) for k in FILO_DETAY_SUTUNLARI} for r in filo_satirlari]
        filo_adet = _com_sayfaya_yaz(
            ws_filo, filo_baslik_satiri, filo_yaz, sabit_kolonlar=FILO_DETAY_SUTUNLARI
        )
        filo_kolon = len(FILO_DETAY_SUTUNLARI)
        _com_pivot_kaynak_guncelle(wb, ws_filo.Name, filo_baslik_satiri, filo_adet, filo_kolon)

        if sutun_autofit:
            for sheet in (ws_veri, ws_filo):
                try:
                    sheet.Columns.AutoFit()
                except Exception as exc:
                    uyarilar.append(f"{sheet.Name} AutoFit: {exc}")

        if pivot_yenile:
            try:
                wb.RefreshAll()
            except Exception as exc:
                uyarilar.append(f"pivot yenileme: {exc}")
            try:
                excel.CalculateUntilAsyncQueriesDone()
            except Exception:
                pass
            try:
                excel.CalculateFullRebuild()
            except Exception as exc:
                uyarilar.append(f"hesaplama: {exc}")

        if sutun_autofit:
            for sheet in wb.Worksheets:
                if int(sheet.Visible) != -1:
                    continue
                try:
                    sheet.Columns.AutoFit()
                except Exception as exc:
                    uyarilar.append(f"{sheet.Name} AutoFit: {exc}")

        wb.Save()
        wb.Close(SaveChanges=True)
        wb = None

        mesaj = f"kısmi uyarı: {'; '.join(uyarilar)}" if uyarilar else None
        return True, mesaj, veri_adet, filo_adet

    except Exception as exc:
        return False, str(exc), 0, 0

    finally:
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass


def _excel_islemleri(dosya_yolu: Path, pivot_yenile: bool = True) -> tuple[bool, str | None]:
    """Mevcut dosyada yalnızca pivot yenile + AutoFit (eski akış / test)."""
    try:
        import win32com.client  # type: ignore
    except ImportError:
        return False, "pywin32 kurulu değil"

    excel = None
    wb = None
    uyarilar: list[str] = []
    dosya = str(dosya_yolu.resolve())

    try:
        excel = _excel_uygulama_ac()
        wb = excel.Workbooks.Open(
            Filename=dosya,
            UpdateLinks=0,
            ReadOnly=False,
            Notify=False,
        )

        if pivot_yenile:
            try:
                wb.RefreshAll()
            except Exception as exc:
                uyarilar.append(f"pivot yenileme: {exc}")
            try:
                excel.CalculateUntilAsyncQueriesDone()
            except Exception:
                pass
            try:
                excel.CalculateFullRebuild()
            except Exception as exc:
                uyarilar.append(f"hesaplama: {exc}")

        autofit_sayisi = 0
        for sheet in wb.Worksheets:
            if int(sheet.Visible) != -1:
                continue
            try:
                sheet.Columns.AutoFit()
                autofit_sayisi += 1
            except Exception as exc:
                uyarilar.append(f"{sheet.Name} AutoFit: {exc}")

        wb.Save()
        wb.Close(SaveChanges=True)
        wb = None

        if autofit_sayisi == 0:
            mesaj = "; ".join(uyarilar) if uyarilar else "Görünür sayfa bulunamadı"
            return False, mesaj

        if uyarilar:
            return True, f"kısmi uyarı: {'; '.join(uyarilar)}"
        return True, None

    except Exception as exc:
        return False, str(exc)

    finally:
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass


def pivot_yenile(dosya_yolu: Path) -> tuple[bool, str | None]:
    """Windows + Excel kurulu ise pivotları yeniler ve sütunları genişletir."""
    return _excel_islemleri(dosya_yolu, pivot_yenile=True)


def _openpyxl_yedek_yaz(
    hedef: Path,
    veri_satirlari: list[dict[str, Any]],
    filo_satirlari: list[dict[str, Any]],
    veri_sayfa_adlari: list[str],
    filo_sayfa_adlari: list[str],
    veri_baslik_satiri: int,
    filo_baslik_satiri: int,
) -> tuple[int, int]:
    """Excel yoksa yedek — pivotlu şablon Excel'de açılamayabilir."""
    wb = load_workbook(hedef, keep_vba=True)

    ws_veri = _sayfa_bul(wb, veri_sayfa_adlari)
    if ws_veri is None:
        wb.close()
        raise ValueError(f"Şablonda VERİ sayfası bulunamadı: {veri_sayfa_adlari}")

    ws_filo = _sayfa_bul(wb, filo_sayfa_adlari)
    if ws_filo is None:
        wb.close()
        raise ValueError(f"Şablonda Filo Detay sayfası bulunamadı: {filo_sayfa_adlari}")

    veri_adet = _sayfayi_temizle_yaz(ws_veri, veri_baslik_satiri, veri_satirlari)
    filo_yaz = [{k: r.get(k) for k in FILO_DETAY_SUTUNLARI} for r in filo_satirlari]
    filo_adet = _sayfayi_temizle_yaz(
        ws_filo, filo_baslik_satiri, filo_yaz, sabit_kolonlar=FILO_DETAY_SUTUNLARI
    )
    _sayfa_sutunlarini_genislet(ws_veri, veri_baslik_satiri, veri_adet)
    _sayfa_sutunlarini_genislet(ws_filo, filo_baslik_satiri, filo_adet)
    wb.save(hedef)
    wb.close()
    return veri_adet, filo_adet


def _bind_olustur(bas: str, bit: str) -> dict:
    bind = {"bas": bas, "bit": bit}
    if getattr(ayarlar, "CO_CODE", None):
        bind["co_code"] = ayarlar.CO_CODE
    if getattr(ayarlar, "BRANCH_CODE", None):
        bind["branch_code"] = ayarlar.BRANCH_CODE
    return bind


def _tarih_araligi() -> tuple[str, str]:
    bugun = datetime.now().date()
    bas = getattr(ayarlar, "KPI_BASLANGIC_TARIHI", bugun.replace(day=1).strftime("%d.%m.%Y"))
    bit = getattr(ayarlar, "KPI_BITIS_TARIHI", bugun.strftime("%d.%m.%Y"))
    return bas, bit


def sablon_rapor_olustur(
    sablon: Path | None = None,
    cikti: Path | None = None,
    pivot_yenile_calistir: bool | None = None,
) -> str:
    kaynak = sablon or sablon_yolu()
    hedef = cikti or _cikti_yolu(kaynak)

    if not kaynak.exists():
        raise FileNotFoundError(
            f"KPI şablonu bulunamadı: {kaynak}\n"
            f"Temmuz KPI dosyanızı bu konuma 'kpi_sablon.xlsx' adıyla kopyalayın."
        )

    bas, bit = _tarih_araligi()
    bind = _bind_olustur(bas, bit)

    hedef.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(kaynak, hedef)

    veri_satirlari: list[dict] = []
    filo_satirlari: list[dict] = []

    with baglanti_yonet() as baglanti:
        cursor = baglanti.cursor()
        if veri_semasi_hazir()[0]:
            veri_satirlari = veri_satirlari_getir(cursor, bas, bit, bind)
        if kiralk_arac_semasi_hazir()[0]:
            filo_satirlari = kiralk_arac_detay_getir(cursor, bas, bit, bind)

    veri_sayfa_adlari = getattr(ayarlar, "KPI_VERI_SAYFA_ADLARI", ["VERİ", "VERI", "Veri"])
    filo_sayfa_adlari = getattr(ayarlar, "KPI_FILO_SAYFA_ADLARI", ["Filo Detay", "Filo detay"])
    veri_baslik_satiri = int(getattr(ayarlar, "KPI_VERI_BASLIK_SATIRI", 1))
    filo_baslik_satiri = int(getattr(ayarlar, "KPI_FILO_BASLIK_SATIRI", 1))

    if pivot_yenile_calistir is None:
        pivot_yenile_calistir = getattr(ayarlar, "KPI_PIVOT_YENILE", True)
    sutun_autofit = getattr(ayarlar, "KPI_SUTUN_AUTOFIT", True)

    excel_mesaj: str | None = None
    veri_adet = 0
    filo_adet = 0

    if _excel_kullanilabilir():
        excel_ok, excel_mesaj, veri_adet, filo_adet = _excel_sablon_doldur(
            hedef,
            veri_satirlari,
            filo_satirlari,
            veri_sayfa_adlari,
            filo_sayfa_adlari,
            veri_baslik_satiri,
            filo_baslik_satiri,
            pivot_yenile=pivot_yenile_calistir,
            sutun_autofit=sutun_autofit,
        )
        if not excel_ok:
            raise RuntimeError(
                f"Excel ile KPI şablonu doldurulamadı: {excel_mesaj}\n"
                "Pivotlu şablon openpyxl ile güvenle kaydedilemez; Excel kurulu ve dosya kapalı olmalı."
            )
    else:
        print(
            "  Uyarı: Excel COM kullanılamıyor — openpyxl yedek modu. "
            "Pivotlu şablon Excel'de açılamayabilir."
        )
        veri_adet, filo_adet = _openpyxl_yedek_yaz(
            hedef,
            veri_satirlari,
            filo_satirlari,
            veri_sayfa_adlari,
            filo_sayfa_adlari,
            veri_baslik_satiri,
            filo_baslik_satiri,
        )

    print(f"BAŞARILI: KPI şablon raporu → {hedef.resolve()}")
    print(f"  Dönem: {bas} — {bit}")
    print(f"  VERİ satırı: {veri_adet}")
    print(f"  Filo Detay satırı: {filo_adet}")
    if excel_mesaj:
        print(f"  Uyarı: {excel_mesaj}")

    return str(hedef.resolve())

"""
Referans KPI şablonunu doldurur: VERİ + Filo Detay → pivot sayfaları Excel'de güncellenir.

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


def _cikti_yolu() -> Path:
    dosya = getattr(ayarlar, "KPI_RAPOR_DOSYASI", "kpi_rapor.xlsx")
    yol = Path(dosya)
    if yol.is_absolute():
        return yol
    return _raporlar_klasoru() / dosya


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


def _pivot_kaynak_guncelle(wb, ws: Worksheet, baslik_satiri: int, satir_sayisi: int):
    """Pivot kaynak aralığını yeni satır sayısına göre genişletmeye çalışır."""
    if satir_sayisi <= 0:
        return
    son_satir = baslik_satiri + satir_sayisi
    son_kolon = get_column_letter(ws.max_column or 1)
    yeni_ref = f"'{ws.title}'!$A${baslik_satiri}:${son_kolon}${son_satir}"

    for sheet in wb.worksheets:
        if not hasattr(sheet, "_pivots") or not sheet._pivots:
            continue
        for pivot in sheet._pivots:
            try:
                if pivot.cache and pivot.cache.cacheSource:
                    pivot.cache.cacheSource.worksheetSource.ref = yeni_ref
            except Exception:
                pass


def pivot_yenile(dosya_yolu: Path) -> bool:
    """Windows + Excel kurulu ise tüm pivotları yeniler."""
    try:
        import win32com.client  # type: ignore
    except ImportError:
        return False

    excel = None
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(str(dosya_yolu.resolve()))
        wb.RefreshAll()
        excel.CalculateFullRebuild()
        wb.Save()
        wb.Close(SaveChanges=True)
        return True
    except Exception:
        return False
    finally:
        if excel is not None:
            excel.Quit()


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
    hedef = cikti or _cikti_yolu()

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

    wb = load_workbook(hedef, keep_vba=True)

    veri_sayfa_adlari = getattr(ayarlar, "KPI_VERI_SAYFA_ADLARI", ["VERİ", "VERI", "Veri"])
    filo_sayfa_adlari = getattr(ayarlar, "KPI_FILO_SAYFA_ADLARI", ["Filo Detay", "Filo detay"])
    veri_baslik_satiri = int(getattr(ayarlar, "KPI_VERI_BASLIK_SATIRI", 1))
    filo_baslik_satiri = int(getattr(ayarlar, "KPI_FILO_BASLIK_SATIRI", 1))

    ws_veri = _sayfa_bul(wb, veri_sayfa_adlari)
    if ws_veri is None:
        raise ValueError(
            f"Şablonda VERİ sayfası bulunamadı. Aranan adlar: {veri_sayfa_adlari}. "
            f"Mevcut sayfalar: {wb.sheetnames}"
        )

    ws_filo = _sayfa_bul(wb, filo_sayfa_adlari)
    if ws_filo is None:
        raise ValueError(
            f"Şablonda Filo Detay sayfası bulunamadı. Aranan adlar: {filo_sayfa_adlari}. "
            f"Mevcut sayfalar: {wb.sheetnames}"
        )

    veri_adet = _sayfayi_temizle_yaz(ws_veri, veri_baslik_satiri, veri_satirlari)
    _pivot_kaynak_guncelle(wb, ws_veri, veri_baslik_satiri, veri_adet)

    filo_yaz = [{k: r.get(k) for k in FILO_DETAY_SUTUNLARI} for r in filo_satirlari]
    filo_adet = _sayfayi_temizle_yaz(
        ws_filo, filo_baslik_satiri, filo_yaz, sabit_kolonlar=FILO_DETAY_SUTUNLARI
    )
    _pivot_kaynak_guncelle(wb, ws_filo, filo_baslik_satiri, filo_adet)

    wb.save(hedef)
    wb.close()

    if pivot_yenile_calistir is None:
        pivot_yenile_calistir = getattr(ayarlar, "KPI_PIVOT_YENILE", True)

    pivot_ok = False
    if pivot_yenile_calistir:
        pivot_ok = pivot_yenile(hedef)

    print(f"BAŞARILI: KPI şablon raporu → {hedef.resolve()}")
    print(f"  Dönem: {bas} — {bit}")
    print(f"  VERİ satırı: {veri_adet}")
    print(f"  Filo Detay satırı: {filo_adet}")
    if pivot_yenile_calistir and not pivot_ok:
        print("  Not: Pivot otomatik yenilenemedi — Excel'de dosyayı açıp 'Verileri Yenile' yapın.")

    return str(hedef.resolve())

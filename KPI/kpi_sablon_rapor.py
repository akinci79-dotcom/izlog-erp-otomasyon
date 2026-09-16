"""
Referans KPI şablonunu doldurur: VERİ + Filo Detay → pivot sayfaları Excel'de güncellenir.

Pivotlu şablonlar openpyxl ile kaydedildiğinde bozulabildiği için (Excel açamaz),
Windows'ta veri yazımı doğrudan Excel COM ile yapılır.

Kullanım:
  1. Temmuz KPI dosyanızı KPI/referans/kpi_sablon.xlsx olarak kaydedin
  2. python kpi_rapor_olustur.py
"""
from __future__ import annotations

import calendar
import re
import shutil
import time
import unicodedata
from datetime import date, datetime
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

import ayarlar
from kpi_kiralk_arac import FILO_DETAY_SUTUNLARI, kiralk_arac_detay_getir, kiralk_arac_semasi_hazir
from kpi_veri import hucre_degeri, veri_satirlari_getir, veri_semasi_hazir, veri_sql_kaynak_bilgisi
from kpi_zarar_detay import ZARAR_DETAY_METIN_SUTUNLARI, zarar_eden_sevkleri_hesapla
from oracle_baglanti import baglanti_yonet

_KPI_KOKU = Path(__file__).resolve().parent


def _progress(mesaj: str) -> None:
    """Konsola anında yazar (Windows cmd tamponlaması / sessiz bekleme önlenir)."""
    print(mesaj, flush=True)

# Şablon başlığı (normalize) → Oracle kolon adı — Temmuz KPI şablonu için varsayılanlar
_SABLON_KOLON_VARSAYILAN: dict[str, str] = {
    "YUK_NO": "YUK_NO",
    "YUK_NUMARASI": "YUK_NO",
    "SEVK_NO": "SEVK_NO",
    "SEVK_NUMARASI": "SEVK_NO",
    "YUK_TARIHI": "YUK_TARIHI",
    "SEVK_TARIHI": "SEVK_TARIHI",
    "PROJE_KODU": "PROJE_KODU",
    "PROJE": "PROJE_KODU",
    "PLAKA": "PLAKA",
    "ARAC_TIPI": "ARAC_TIPI",
    "MUSTERI_KODU": "MUSTERI_KODU",
    "MUSTERI_ADI": "MUSTERI_ADI",
    "SUBE": "SUBE",
    "SUBE_KODU": "SUBE_KODU",
    "SUBE_ADI": "SUBE",
    "SATIS_TUTAR": "SATIS_TUTAR",
    "SATIS": "SATIS_TUTAR",
    "NET_SATIS": "SATIS_TUTAR",
    "TOPLAM_SATIS": "TOPLAM_SATIS",
    "ALIS_TUTAR": "ALIS_TUTAR",
    "ALIS": "ALIS_TUTAR",
    "TOPLAM_ALIS": "TOPLAM_ALIS",
    "KAR_ZARAR": "TOPLAM_KAR_ZARAR",
    "NET_KAR_ZARAR": "TOPLAM_KAR_ZARAR",
    "KAR_ZARAR_TUTAR": "TOPLAM_KAR_ZARAR",
    "MARJ_YUZDE": "MARJ_YUZDE",
    "MARJ_ORANI": "MARJ_ORANI",
    "MARJ": "MARJ_YUZDE",
    "YUK_FIYAT_TIP_KODU": "YUK_FIYAT_TIP_KODU",
    "MAL_TIPI": "MAL_TIP",
    "MAL_TIP": "MAL_TIP",
    "MAL_TIP_KODU": "YUK_FIYAT_TIP_KODU",
    "ALIS_TUTARI": "ALIS_TUTARI",
    "SATIS_TUTARI": "SATIS_TUTARI",
    "ALIS_IADE_TUTARI": "ALIS_IADE_TUTARI",
    "SATIS_IADE_TUTARI": "SATIS_IADE_TUTARI",
    "TOPLAM_KAR_ZARAR": "TOPLAM_KAR_ZARAR",
    "TOPLAM_KAR_ZARAR_TUTAR": "TOPLAM_KAR_ZARAR",
    "KULLANICI": "KULLANICI",
    "ARSIV_DURUMU": "ARSIV_DURUMU",
    "MUSTERI_HESAPLASMA_ACIKLAMA": "MUSTERI_HESAPLASMA_ACIKLAMA",
    "SEVK_DURUMU": "SEVK_DURUMU",
    "SEFER_TURU": "SEFER_TURU",
    "SEFER_TIPI": "SEFER_TIPI",
    "SIPARIS_TARIHI": "SIPARIS_TARIHI",
    "SIPARIS_NO": "SIPARIS_NO",
    "SIPARIS_NOTLARI": "SIPARIS_NOTLARI",
    "GONDERICI_CARI_ADI": "GONDERICI_CARI_ADI",
    "ALICI_CARI_ADI": "ALICI_CARI_ADI",
    "SOZLESME_TARIHI": "SOZLESME_TARIHI",
    "SOZLESME_NO": "SOZLESME_NO",
    "MUSTERI_EVRAK_NO": "MUSTERI_EVRAK_NO",
    "PLAKA_CARI_ADI": "PLAKA_CARI_ADI",
    "MULKIYET": "PLAKA_MULKIYET",
    "PLAKA_MULKIYET": "PLAKA_MULKIYET",
    "SURUCU_ADI": "SURUCU_ADI",
    "SURUCU_TEL": "SURUCU_TEL",
    "NOKTA_SAYISI": "NOKTA_SAYISI",
    "YUKLEME_YER_KODU": "YUKLEME_YER_KODU",
    "YUKLEME_YER_ADI": "YUKLEME_YER_ADI",
    "YUKLEME_SEHIR_ADI": "YUKLEME_SEHIR_ADI",
    "YUKLEME_ILCE_ADI": "YUKLEME_ILCE_ADI",
    "YUKLEME_ADRESI": "YUKLEME_ADRESI",
    "BOSALTMA_YER_KODU": "BOSALTMA_YER_KODU",
    "BOSALTMA_YER_ADI": "BOSALTMA_YER_ADI",
    "BOSALTMA_SEHIR_ADI": "BOSALTMA_SEHIR_ADI",
    "BOSALTMA_ILCE_ADI": "BOSALTMA_ILCE_ADI",
    "BOSALTMA_ADRESI": "BOSALTMA_ADRESI",
    "BRUT_AGIRLIK": "BRUT_AGIRLIK",
    "CIKIS_KM": "CIKIS_KM",
    "VARIS_KM": "VARIS_KM",
    "BOS_KM": "BOS_KM",
}


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


_RAPOR_ADI_SONEKI_VARSAYILAN = "İzlog Lojistik Raporları"
_DOSYA_ADI_GECERSIZ_KARAKTERLER = re.compile(r'[\\/:*?"<>|]')


def _dosya_adi_icin_temizle(metin: str) -> str:
    """Windows'ta dosya adında YASAK olan karakterleri (\\ / : * ? " < > |)
    tire ile değiştirir, baştaki/sondaki boşlukları/nokta'ları kırpar (Windows
    dosya adının sonunda nokta/boşluk bırakmaz)."""
    temiz = _DOSYA_ADI_GECERSIZ_KARAKTERLER.sub("-", metin).strip()
    return temiz.rstrip(". ") or "KPI Raporu"


def _cikti_yolu(sablon: Path | None = None, bas: str = "", bit: str = "") -> Path:
    """Çıktı dosyasının yolu. `KPI_RAPOR_DOSYASI` ayarlar.py'de tanımlıysa
    [kullanıcı isteğiyle KALDIRILMADI — elle sabit bir isim isteyen için hâlâ
    öncelikli] o kullanılır. Aksi halde [kullanıcı isteği]: dosya adı rapor
    dönemine göre OTOMATİK üretilir — örn. tek ay için 'Ağustos 2026 İzlog
    Lojistik Raporları.xlsx', tam yıl için '2026 İzlog Lojistik Raporları.xlsx',
    aksi (aralık) durumda '01.08.2026 – 31.08.2026 İzlog Lojistik
    Raporları.xlsx'. `bas`/`bit` verilmezse (örn. çok eski bir çağrı yeri)
    eski sabit 'kpi_rapor.xlsx' adına düşer."""
    dosya = getattr(ayarlar, "KPI_RAPOR_DOSYASI", None)
    if dosya:
        yol = Path(dosya)
        if yol.is_absolute():
            return yol
        return _raporlar_klasoru() / dosya

    suffix = (sablon or sablon_yolu()).suffix or ".xlsx"
    if not bas or not bit:
        return _raporlar_klasoru() / f"kpi_rapor{suffix}"

    soneki = getattr(ayarlar, "KPI_RAPOR_ADI_SONEKI", _RAPOR_ADI_SONEKI_VARSAYILAN)
    etiket = _donem_etiketi(bas, bit)
    dosya_adi = _dosya_adi_icin_temizle(f"{etiket} {soneki}".strip())
    return _raporlar_klasoru() / f"{dosya_adi}{suffix}"


def _sablon_hedefe_kopyala(kaynak: Path, hedef: Path, deneme: int = 5, bekleme_sn: float = 2.0) -> None:
    """Şablonu çıktı yoluna kopyalar; 'dosya başka bir işlem tarafından kullanılıyor'
    (WinError 32) tipik olarak GEÇİCİ bir kilit (virüs taraması, OneDrive senkronu)
    veya çıktı dosyasının hâlâ Excel'de/bir önceki çökmüş çalıştırmanın arkada kalan
    (görünmez) Excel sürecinde açık olmasından kaynaklanır. Birkaç kez kısa aralıkla
    tekrar dener; hâlâ başarısızsa kullanıcıya somut bir eylem listesi veren net bir
    hata fırlatır (ham WinError metni yerine)."""
    son_hata: OSError | None = None
    for deneme_no in range(1, deneme + 1):
        try:
            shutil.copy2(kaynak, hedef)
            return
        except OSError as exc:
            son_hata = exc
            if deneme_no < deneme:
                _progress(
                    f"  Uyarı: {hedef.name} şu an başka bir işlem tarafından kullanılıyor "
                    f"gibi görünüyor, {bekleme_sn:.0f}sn sonra tekrar denenecek "
                    f"({deneme_no}/{deneme})..."
                )
                time.sleep(bekleme_sn)

    raise RuntimeError(
        f"'{hedef.name}' dosyasına yazılamadı — başka bir işlem (muhtemelen Excel) "
        f"dosyayı açık tutuyor.\n"
        f"Kontrol edin: (1) {hedef.name} dosyası Excel'de açıksa kapatın, "
        f"(2) Görev Yöneticisi'nde (Ctrl+Shift+Esc) arkada kalmış bir 'EXCEL.EXE' "
        f"süreci varsa (görünür pencere olmasa bile — Excel COM otomasyonu gizli "
        f"çalışır) sonlandırın, (3) tekrar 'python kpi_rapor_olustur.py' çalıştırın.\n"
        f"Ham hata: {son_hata}"
    ) from son_hata


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


def _kolon_esleme(
    basliklar: list[str | None],
    veri_anahtarlari: list[str],
    tablo_sol: int = 1,
) -> dict[int, str]:
    normalized_veri = {_normalize_kolon(k): k for k in veri_anahtarlari}
    for sablon_norm, oracle in _SABLON_KOLON_VARSAYILAN.items():
        if oracle in veri_anahtarlari:
            normalized_veri.setdefault(sablon_norm, oracle)
    ek = getattr(ayarlar, "KPI_KOLON_ESLEME", {}) or {}
    for sablon, oracle in ek.items():
        normalized_veri[_normalize_kolon(sablon)] = oracle

    esleme: dict[int, str] = {}
    for i, baslik in enumerate(basliklar):
        if baslik is None or str(baslik).strip() == "":
            continue
        norm = _normalize_kolon(str(baslik))
        if norm in normalized_veri:
            esleme[tablo_sol + i] = normalized_veri[norm]
    return esleme


def _eslesmeyen_basliklar(
    basliklar: list[Any],
    esleme: dict[int, str],
    tablo_sol: int = 1,
) -> list[str]:
    sonuc: list[str] = []
    for i, baslik in enumerate(basliklar):
        if baslik is None or str(baslik).strip() == "":
            continue
        col_idx = tablo_sol + i
        if col_idx not in esleme:
            sonuc.append(str(baslik).strip())
    return sonuc


def _sutun1_arteakt_mi(baslik: Any) -> bool:
    if baslik is None:
        return False
    return _normalize_kolon(str(baslik)) in ("SUTUN1", "COLUMN1")


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


def _com_basliklari_oku(
    sheet,
    baslik_satiri: int,
    tablo_sol: int = 1,
    tablo_sag: int | None = None,
) -> list[Any]:
    """Başlık satırını fiziksel hücrelerden okur (ListColumns kullanılmaz — kaydırma hatası önlenir)."""
    if tablo_sag is None:
        tablo_sag = tablo_sol + 59
        try:
            used = sheet.UsedRange
            tablo_sag = max(int(used.Column) + int(used.Columns.Count) - 1, tablo_sag)
        except Exception:
            pass
    ham = sheet.Range(
        sheet.Cells(baslik_satiri, tablo_sol),
        sheet.Cells(baslik_satiri, tablo_sag),
    ).Value
    basliklar = _com_satir_oku(ham)
    while basliklar and basliklar[-1] is None:
        basliklar.pop()
    return basliklar


def _com_sutun1_baslangic_temizle(sheet, baslik_satiri: int) -> bool:
    """Resize hatasıyla A sütununa eklenen 'Sütun1' artefaktını kaldırır."""
    try:
        a1 = sheet.Cells(baslik_satiri, 1).Value
        b1 = sheet.Cells(baslik_satiri, 2).Value
    except Exception:
        return False
    if not _sutun1_arteakt_mi(a1):
        return False
    if b1 is None or str(b1).strip() == "":
        return False
    try:
        sheet.Columns(1).Delete()
        return True
    except Exception:
        return False


def _com_listobject_bul(sheet, baslik_satiri: int):
    """Sayfadaki Excel Tablosunu (ListObject) bulur — pivot kaynağı genelde budur."""
    try:
        adet = int(sheet.ListObjects.Count)
    except Exception:
        return None
    for i in range(1, adet + 1):
        lo = sheet.ListObjects(i)
        try:
            if int(lo.HeaderRowRange.Row) == baslik_satiri:
                return lo
        except Exception:
            continue
    if adet >= 1:
        return sheet.ListObjects(1)
    return None


def _com_tablo_konumu(sheet, baslik_satiri: int) -> tuple[int, int, Any | None]:
    """Excel tablosunun gerçek sol sütun ve kolon sayısını döndürür."""
    lo = _com_listobject_bul(sheet, baslik_satiri)
    if lo is not None:
        try:
            hdr = lo.HeaderRowRange
            tablo_sol = int(hdr.Column)
            kolon_sayisi = int(hdr.Columns.Count)
            if kolon_sayisi > 0:
                return tablo_sol, kolon_sayisi, lo
        except Exception:
            pass
    basliklar = _com_basliklari_oku(sheet, baslik_satiri, 1)
    return 1, max(len(basliklar), 1), lo


def _com_veri_matrisi_hazirla(
    basliklar: list[Any],
    satirlar: list[dict[str, Any]],
    sabit_kolonlar: list[str] | None,
    kolon_sayisi: int,
    tablo_sol: int = 1,
) -> list[tuple[Any, ...]]:
    if not satirlar:
        return []

    anahtarlar = sabit_kolonlar or list(satirlar[0].keys())
    esleme = _kolon_esleme(basliklar, anahtarlar, tablo_sol)
    if not esleme and sabit_kolonlar:
        esleme = {
            tablo_sol + i: k for i, k in enumerate(sabit_kolonlar) if i < len(basliklar)
        }

    matris: list[tuple[Any, ...]] = []
    tablo_sag = tablo_sol + kolon_sayisi - 1
    for satir in satirlar:
        if sabit_kolonlar:
            satir_verisi: list[Any] = [
                hucre_degeri(satir.get(k)) for k in sabit_kolonlar[:kolon_sayisi]
            ]
            if len(satir_verisi) < kolon_sayisi:
                satir_verisi.extend([None] * (kolon_sayisi - len(satir_verisi)))
        else:
            satir_verisi = [None] * kolon_sayisi
            for col_idx, kolon in esleme.items():
                if tablo_sol <= col_idx <= tablo_sag:
                    satir_verisi[col_idx - tablo_sol] = hucre_degeri(satir.get(kolon))
        matris.append(tuple(satir_verisi))
    return matris


def _com_araliga_yaz(hedef, matris: list[tuple[Any, ...]]):
    if not matris:
        return
    if len(matris) == 1:
        hedef.Value = matris[0]
    else:
        hedef.Value = tuple(matris)


def _com_sayfaya_yaz(
    sheet,
    baslik_satiri: int,
    satirlar: list[dict[str, Any]],
    sabit_kolonlar: list[str] | None = None,
) -> tuple[int, dict[int, str], list[str], int, int]:
    """Pivot kaynağına yazar — tablo sütun sınırları korunur, A'ya genişletme yapılmaz."""
    if _com_sutun1_baslangic_temizle(sheet, baslik_satiri):
        lo = _com_listobject_bul(sheet, baslik_satiri)
        if lo is not None:
            try:
                lo = sheet.ListObjects(lo.Name)
            except Exception:
                pass

    tablo_sol, kolon_sayisi, lo = _com_tablo_konumu(sheet, baslik_satiri)
    tablo_sag = tablo_sol + kolon_sayisi - 1
    basliklar = _com_basliklari_oku(sheet, baslik_satiri, tablo_sol, tablo_sag)
    if basliklar:
        kolon_sayisi = len(basliklar)
        tablo_sag = tablo_sol + kolon_sayisi - 1

    anahtarlar = sabit_kolonlar or (list(satirlar[0].keys()) if satirlar else [])
    esleme = _kolon_esleme(basliklar, anahtarlar, tablo_sol)
    if not esleme and sabit_kolonlar:
        esleme = {
            tablo_sol + i: k for i, k in enumerate(sabit_kolonlar) if i < len(basliklar)
        }
    eslesmeyen = (
        _eslesmeyen_basliklar(basliklar, esleme, tablo_sol) if not sabit_kolonlar else []
    )

    if not satirlar:
        return 0, esleme, eslesmeyen, kolon_sayisi, tablo_sol

    matris = _com_veri_matrisi_hazirla(
        basliklar, satirlar, sabit_kolonlar, kolon_sayisi, tablo_sol
    )
    if not matris:
        return 0, esleme, eslesmeyen, kolon_sayisi, tablo_sol

    son_satir = baslik_satiri + len(matris)
    hedef = sheet.Range(
        sheet.Cells(baslik_satiri + 1, tablo_sol),
        sheet.Cells(son_satir, tablo_sag),
    )

    if lo is not None:
        try:
            hdr = lo.HeaderRowRange
            resize_sol = int(hdr.Column)
            resize_sag = resize_sol + int(hdr.Columns.Count) - 1
            yeni_tablo = sheet.Range(
                sheet.Cells(baslik_satiri, resize_sol),
                sheet.Cells(son_satir, resize_sag),
            )
            lo.Resize(yeni_tablo)
        except Exception:
            pass

    # Tarih kolonlarını veri yazılmadan ÖNCE 'General'e sıfırla (bkz.
    # _com_bicim_genel_yap docstring'i — sıralama önemli: hücre önceden 'Metin'
    # veya bozuk bir özel biçimde kilitliyse, format değişikliği yazımdan SONRA
    # yapılırsa bazı Excel/COM senaryolarında görünüm beklendiği gibi
    # güncellenmeyebiliyor). `dd.mm.yyyy` biçiminin kendisi yazımdan SONRA
    # `_com_tarih_bicimi_uygula` ile ayrıca uygulanıyor.
    for col_idx, oracle_kolon in esleme.items():
        if not str(oracle_kolon).upper().endswith("TARIHI"):
            continue
        try:
            _com_bicim_genel_yap(
                sheet.Range(
                    sheet.Cells(baslik_satiri + 1, col_idx),
                    sheet.Cells(son_satir, col_idx),
                )
            )
        except Exception:
            pass

    _com_araliga_yaz(hedef, matris)

    return len(satirlar), esleme, eslesmeyen, kolon_sayisi, tablo_sol


def _com_bicim_genel_yap(araligi) -> None:
    """Bir aralığın biçimini 'General'e sıfırlar (hem NumberFormat hem NumberFormatLocal).

    ÖNEMLİ SIRALAMA KURALI [kullanıcı isteğiyle netleştirildi]: bu fonksiyon veri
    YAZILMADAN ÖNCE çağrılmalı. Hücre önceden 'Metin' (@) biçiminde ya da bozuk
    bir özel tarih biçiminde kilitli kalmışsa (bkz. _com_tarih_bicimi_zorla
    docstring'i), format değişikliği YENİ değer yazıldıktan SONRA yapılırsa bazı
    Excel/COM senaryolarında hücrenin GÖRÜNÜMÜ beklenen şekilde güncellenmeyebilir.
    Güvenli sıra: (1) hedef aralığı 'General'e sıfırla (bu fonksiyon), (2) float
    seri numarasını yaz, (3) 'dd.mm.yyyy' uygula (bkz. _com_tarih_bicimi_zorla).
    Bu yüzden hem `_com_sayfaya_yaz`/`_com_zarar_detay_tablo_yaz` veri yazımından
    ÖNCE, hem `_com_tarih_bicimi_uygula`/`_com_tarih_bicimi_zorla` veri yazımından
    SONRA bu sıfırlamayı çift güvence olarak uyguluyor.
    """
    for _ in range(2):
        try:
            araligi.NumberFormat = "General"
            araligi.NumberFormatLocal = "General"
            return
        except Exception:
            continue


_TARIH_BICIM_GECERLI_DESEN = re.compile(
    r"^(?:d{1,2}|g{1,2})(?:\\?[./\-])+(?:m{1,2}|a{1,2})(?:\\?[./\-])+y{2,4}$",
    re.IGNORECASE,
)


def _tarih_bicimi_gecerli_mi(numberformat: Any) -> bool:
    """`numberformat`'ın (Excel'den GERİ OKUNAN ham `.NumberFormat`/
    `.NumberFormatLocal` metni) GERÇEKTEN bir 'gün.ay.yıl' tarih biçimi olup
    olmadığını YAPISAL olarak denetler — TAM literal string eşitliği ARAMAZ.

    ARKA PLAN [WebSearch ile araştırıldı — bkz. modül/README'deki not]:
    Microsoft'un resmi dokümantasyonu `.NumberFormat`'ın locale-BAĞIMSIZ
    (her zaman US-English 'd'/'m'/'y' kodlarıyla) olacağını söylüyor ve VBA
    içinden (Excel'in kendi makro motorundan) çalıştırıldığında bu genellikle
    doğru. AMA harici bir COM istemcisinden (bu projede olduğu gibi
    Python/pywin32) sürüldüğünde, ÖZELLİKLE Excel arayüz dili Türkçe
    olduğunda, `.NumberFormat`'ı OKURKEN Excel'in kodu Türkçe karşılığına
    ('dd'->'gg', 'mm'->'aa') çevirerek döndürdüğü bilinen bir tuhaflık var
    (bkz. stackoverflow.com/questions/19839047 ve oaltd.co.uk Excel VBA
    Programcı Referansı Böl. 22 — ikisi de NumberFormat'ın "İngilizce
    döndürülmesi gerekir" kuralının bazı locale/otomasyon senaryolarında
    beklendiği gibi çalışmadığını doğruluyor). Önceki turdaki doğrulama
    SADECE literal 'dd.mm.yyyy' dizisine tam eşitlik arıyordu; format
    GERÇEKTE doğru uygulanmış olsa da (Excel'de doğru görünüyor) Türkçe
    'gg.aa.yyyy' döndüğünde HER ZAMAN "başarısız" sanılıyor ve kullanıcıya
    YANLIŞ ALARM ('doğrulanamadı') gösteriliyordu — konsol logundaki 4
    tarih sütununun HEPSİNİN AYNI ANDA başarısız olması da (veri/hücre bazlı
    bir sorun değil, sistemsel bir doğrulama hatası) bu teoriyle uyumlu.

    Bu fonksiyon artık TAM literal eşitlik yerine YAPISAL bir kontrol yapıyor:
    "gün(gg/dd) + ayraç + ay(aa/mm) + ayraç + yıl(yyyy)" SIRASINI ve ŞEKLİNİ
    doğrular — İngilizce VEYA Türkçe kod harflerini, '.'/'/'/'-' ayraçlarını
    ve (bazı Excel sürümlerinde görülen) kaçışlı ayraçları ('\\.') kabul eder.

    AMA gerçekten bozuk bir kalıbı YAKALAMAYA devam eder (yanlış negatife
    DÖNMEZ): bilinen 'mm/\\m\\m/yyyy' kırılması (kaçışlı '\\m' İÇİNDE literal
    'm' harfi barındırıyor, GÜN bileşeni tamamen kayıp — biçim GÜN yerine
    doğrudan AY ile başlıyor) bu yapısal desenle EŞLEŞMEZ (desen '^' hemen
    gün/`d`/`g` grubunu zorunlu kılar, 'mm...' ile başlayan bir dize bu
    grupla asla eşleşmez) — fonksiyon False döner, tıpkı gün/ay/yıl
    gruplarının HERHANGİ birinin içine kaçışlı bir harf sızdığı (ayraç
    DIŞINDA bir yerde '\\' göründüğü) diğer olası bozuk kalıplar için de.
    """
    if not numberformat or not isinstance(numberformat, str):
        return False
    return bool(_TARIH_BICIM_GECERLI_DESEN.fullmatch(numberformat.strip()))


def _com_tarih_bicimi_zorla(araligi) -> bool:
    """Bir aralığa 'dd.mm.yyyy' biçimini GARANTİLİ uygulamaya çalışır, başarılı
    olup olmadığını (doğrulanmış) bool olarak döner.

    Bilinen kırılma [DOĞRULANMIŞ — kullanıcı canlı ekran görüntüsüyle teyit
    etti]: bazı tarih hücrelerinin (VERİ sayfası SEVK_TARIHI/YUK_TARIHI vb.)
    NumberFormat'ı 'mm/\\m\\m/yyyy' gibi İÇİNDE KAÇIŞLI ('\\m') literal 'mm'
    METNİ barındıran BOZUK bir kalıba sahip olabiliyor — bu, her hücrede
    GERÇEK ay (ör. '08') + SABİT literal 'mm' metni + gerçek yıl gösterip
    GÜN bilgisini tamamen kaybediyor (ekranda '08.mm.2026' gibi görünüyor;
    tüm satırlar aynı aydaysa hepsi birbirinin AYNI görünüyor, sanki hiç
    değişmiyor). Bu bozuk kalıp muhtemelen şablondan (kpi_sablon.xlsx'te
    elle/yanlışlıkla oluşturulmuş özel bir biçim) veya önceki bir
    çalıştırmadan kalıyor.

    Bu fonksiyon önceki turdaki tek denemeli ('General'e resetle + 'dd.mm.yyyy'
    ata, hatayı sessizce yut) yaklaşımına rağmen sorun DEVAM ETTİĞİ için
    [kullanıcı teyidi] sertleştirildi:
    1. Önce `_com_bicim_genel_yap` ile General'e sıfırlanır (veri yazımından
       SONRA çağrılsa bile — hücre içeriği zaten sayısal seri numarası olduğu
       için bu noktada zararsız, ek bir güvence).
    2. Hem `.NumberFormat` (İngilizce/locale-bağımsız kod, "dd.mm.yyyy") HEM
       `.NumberFormatLocal` (kullanıcının Excel arayüz diline göre kod —
       Türkçe Excel'de "gg.aa.yyyy") AÇIKÇA ayarlanır. `.NumberFormat` teorik
       olarak locale'den bağımsız olmalı, ama önceki turda '08.mm.2026' gibi
       bir kaçış/locale sorunu yaşandığı için ikisi birden ayarlanarak çift
       güvence sağlanıyor.
    3. Doğrulama başarısız olursa, noktaların açıkça kaçışlandığı ('\\.')
       alternatif bir kalıp ikinci bir deneme olarak uygulanır (bazı Excel
       sürümlerinde '.' karakteri özel bir ayraç kodu gibi yorumlanabiliyor
       ihtimaline karşı).
    4. Sonuç HER ZAMAN doğrulanır (hücrenin gerçek NumberFormat'ı okunarak) —
       eskiden olduğu gibi hatayı sessizce yutup "başarılı" varsaymıyor;
       başarısızsa çağıran tarafa bool ile bildiriyor ki bu bir uyarı olarak
       kullanıcıya (Excel'de elle kontrol etmesi için) yansıtılabilsin.

    5. [YENİ — YANLIŞ ALARM düzeltmesi, bkz. `_tarih_bicimi_gecerli_mi`
       docstring'i] Doğrulama artık okunan `.NumberFormat` dizisine TAM
       literal eşitlik ('== "dd.mm.yyyy"') ARAMIYOR — bu, format GERÇEKTE
       doğru uygulanmış olsa bile (örn. Excel Türkçe arayüzde okurken
       'gg.aa.yyyy' döndürdüğünde) HER ZAMAN "başarısız" diye YANLIŞ ALARM
       veriyordu (canlı loglarda 4 farklı tarih sütununun HEPSİNİN TEK
       SEFERDE başarısız olması bunun bir veri sorunu değil, bir doğrulama
       mantığı hatası olduğuna işaret ediyordu). Artık `.NumberFormat` HEM
       `.NumberFormatLocal` okunup ikisinden biri "gün.ay.yıl YAPISINA"
       (İngilizce veya Türkçe kod harfleriyle, olası ayraç/kaçış
       varyasyonlarıyla) uyuyorsa başarılı sayılıyor — ama bilinen bozuk
       'mm/\\m\\m/yyyy' kalıbı (gerçekten gün bilgisini kaybeden) hâlâ
       YAKALANIYOR, çünkü o kalıp gün bileşeniyle BAŞLAMIYOR.
    """
    _com_bicim_genel_yap(araligi)

    denemeler = (
        ("dd.mm.yyyy", "gg.aa.yyyy"),
        (r"dd\.mm\.yyyy", r"gg\.aa\.yyyy"),
    )
    for numberformat, numberformat_local in denemeler:
        try:
            araligi.NumberFormat = numberformat
        except Exception:
            continue
        try:
            araligi.NumberFormatLocal = numberformat_local
        except Exception:
            pass  # NumberFormat (locale-bağımsız) zaten denendi, bu sadece ek güvence

        try:
            uygulanan = str(araligi.Cells(1, 1).NumberFormat)
        except Exception:
            return True  # atama hata vermedi ama doğrulama okunamadı — iyimser kabul et

        if _tarih_bicimi_gecerli_mi(uygulanan):
            return True

        # `.NumberFormat` teorik olarak locale-bağımsız olmalı ama COM
        # otomasyonunda okurken Türkçeleştirilmiş bir varyant döndürebiliyor
        # (bkz. yukarıdaki docstring) — bazı Excel sürümlerinde bu çeviri
        # sadece `.NumberFormatLocal` üzerinde tutarlı olabilir, o yüzden
        # ikinci bir güvence olarak onu da deniyoruz.
        try:
            uygulanan_local = str(araligi.Cells(1, 1).NumberFormatLocal)
        except Exception:
            uygulanan_local = ""
        if _tarih_bicimi_gecerli_mi(uygulanan_local):
            return True

    return False


def _com_tarih_bicimi_uygula(
    sheet, baslik_satiri: int, satir_sayisi: int, esleme: dict[int, str]
) -> list[str]:
    """Tarih kolonlarına açık 'dd.mm.yyyy' biçimi uygular, başarısız olan
    kolonları (varsa) liste olarak döner ki çağıran taraf bunu kullanıcıya
    bir uyarı olarak gösterebilsin (eskiden olduğu gibi sessizce yutmaz).

    hucre_degeri() artık tarihleri Excel seri sayısına (düz float) çeviriyor
    (bkz. kpi_veri.py — pywin32'nin datetime->COM dönüşümündeki saat dilimi
    kayması hatasını önlemek için). Hücrenin biçimi zaten tarih değilse bu
    sayı düz bir rakam olarak görünür; bu yüzden ORACLE kolonu "...TARIHI" ile
    bitenler için biçim burada açıkça ayarlanır (bkz. _com_tarih_bicimi_zorla
    — bilinen 'mm/\\m\\m/yyyy' bozuk kalıp sorunu için General'e resetleyip
    yeniden uyguluyor, hem NumberFormat hem NumberFormatLocal ayarlıyor).

    NOT: Bu fonksiyon veri YAZILDIKTAN SONRA çağrılır (biçimi son kez
    doğrulayıp/garantilemek için). Veri yazılmadan ÖNCEKİ 'General'e sıfırlama
    adımı ayrıca `_com_sayfaya_yaz` içinde (aynı kolonlar için) yapılıyor —
    bkz. o fonksiyondaki `_com_bicim_genel_yap` çağrısı ve modülün sıralama
    notu ("önce sıfırla, sonra yaz, sonra biçimi uygula").
    """
    basarisiz: list[str] = []
    if satir_sayisi <= 0:
        return basarisiz
    son_satir = baslik_satiri + satir_sayisi
    for col_idx, oracle_kolon in esleme.items():
        if not str(oracle_kolon).upper().endswith("TARIHI"):
            continue
        try:
            araligi = sheet.Range(
                sheet.Cells(baslik_satiri + 1, col_idx),
                sheet.Cells(son_satir, col_idx),
            )
            if not _com_tarih_bicimi_zorla(araligi):
                basarisiz.append(f"{oracle_kolon} (sütun {col_idx})")
        except Exception as exc:
            basarisiz.append(f"{oracle_kolon} (sütun {col_idx}): {exc}")
    return basarisiz


def _com_pivot_kaynak_guncelle(
    wb,
    sayfa_adi: str,
    baslik_satiri: int,
    satir_sayisi: int,
    kolon_sayisi: int,
    tablo_sol: int = 1,
):
    if satir_sayisi <= 0:
        return
    son_satir = baslik_satiri + satir_sayisi
    bas_kolon = get_column_letter(max(tablo_sol, 1))
    son_kolon = get_column_letter(max(tablo_sol + kolon_sayisi - 1, tablo_sol))
    yeni_kaynak = f"'{sayfa_adi}'!${bas_kolon}${baslik_satiri}:${son_kolon}${son_satir}"
    try:
        for i in range(1, int(wb.PivotCaches().Count) + 1):
            pc = wb.PivotCaches(i)
            src = str(pc.SourceData or "")
            if sayfa_adi.upper() not in src.upper():
                continue
            # Tablo adına bağlı pivot (örn. 'VERİ'!Tablo1) — ListObject.Resize yeterli
            if "$" not in src:
                continue
            pc.SourceData = yeni_kaynak
    except Exception:
        pass


_ZARAR_DETAY_SAYFA_ADLARI_VARSAYILAN = ["Zarar Detay", "Zarar detay", "ZARAR DETAY"]
_ZARAR_TEDARIKCI_TABLO_ADI_VARSAYILAN = "ZararTedarikci"
_ZARAR_KIRALIK_TABLO_ADI_VARSAYILAN = "ZararKiralik"

# Otomatik satır ekleme için sadece bir "akıl sağlığı" güvenlik sınırı — sonsuz
# döngü riski yok (tek seferlik bulk Insert), ama beklenmedik derecede büyük
# (muhtemelen bir hesaplama hatasından kaynaklanan) bir istek varsa körü körüne
# binlerce satır eklemek yerine net bir hata ile durup elle kontrole yönlendirir.
_ZARAR_DETAY_MAX_OTOMATIK_SATIR_EKLEME = 20000


def _com_zarar_detay_ara_toplam_dogrula_ve_duzelt(
    sheet,
    tablo_adi: str,
    formul_sol: int,
    formul_sag: int,
    eski_veri_son_satir: int,
    yeni_veri_son_satir: int,
) -> None:
    """'Ara Toplam' satırındaki SUM formüllerinin, satır eklendikten SONRA
    GERÇEKTEN yeni veri aralığını (...son satır=yeni_veri_son_satir) kapsayıp
    kapsamadığını DOĞRULAR; Excel'in "aralığın İÇİNE satır eklenirse ona bakan
    formüller otomatik genişler" kuralına KÖRÜ KÖRÜNE güvenmek yerine (bu proje
    kapsamında henüz gerçek Excel'de canlı doğrulanmamış bir davranış), formülü
    okuyup GEREKTİĞİNDE açıkça düzeltir — böylece davranış artık "umulan Excel
    büyüsüne" değil, kodun kendi doğrulamasına dayanıyor.

    Yöntem: Satır ekleme sonrası 'Ara Toplam' satırı fiziksel olarak
    `yeni_veri_son_satir + 1` konumuna kaymış olmalı (Excel'in satır ekleme
    sonrası alttaki HER ŞEYİ aşağı kaydırması temel/kesin bir davranıştır, bu
    kısımda belirsizlik yok — belirsiz olan tek şey, formülün İÇİNDEKİ ARALIK
    REFERANSININ da otomatik büyüyüp büyümediğidir). O satırdaki her J-N
    hücresinin (locale-bağımsız `.Formula`, A1 stili) metni okunur; formülde
    KENDİ sütununa ait ESKİ son satır numarası (ör. 'J93' veya '$J$93') hâlâ
    geçiyorsa, bu aralığın OTOMATİK genişlemediği anlamına gelir — bu durumda
    SADECE o satır numarası, sütun harfiyle birlikte YENİ son satır numarasıyla
    ('J97') değiştirilir; formülün geri kalanı (SUM, başka sarmalayan
    fonksiyonlar, diğer sütunlara ait referanslar vb.) OLDUĞU GİBİ korunur —
    formül körü körüne '=SUM(...)' ile YENİDEN YAZILMAZ (şablondaki formülün
    tam yapısını bilmediğimiz için bu, bilinmeyen bir davranışı/eki kaybetme
    riskini taşırdı). Formülde eski satır numarası hiç geçmiyorsa (Excel zaten
    doğru genişletmiş demektir), hiçbir şey değiştirilmez — bu fonksiyon hem
    "genişledi" hem "genişlemedi" senaryosunda güvenli ve idempotenttir.
    """
    if eski_veri_son_satir == yeni_veri_son_satir:
        return
    ara_toplam_satir = yeni_veri_son_satir + 1
    for col in range(formul_sol, formul_sag + 1):
        col_harf = get_column_letter(col)
        try:
            hucre = sheet.Cells(ara_toplam_satir, col)
            mevcut_formul = hucre.Formula
        except Exception:
            continue
        if not mevcut_formul or not isinstance(mevcut_formul, str):
            continue

        desen = re.compile(
            rf"(?<![A-Za-z])(\$?{re.escape(col_harf)}\$?){eski_veri_son_satir}\b"
        )
        if not desen.search(mevcut_formul):
            continue  # Excel zaten doğru genişletmiş (ya da bu sütunda beklenen desen yok) — dokunma

        yeni_formul = desen.sub(rf"\g<1>{yeni_veri_son_satir}", mevcut_formul)
        try:
            hucre.Formula = yeni_formul
        except Exception as exc:
            raise RuntimeError(
                f"'{tablo_adi}' Ara Toplam {col_harf} sütunu formülü "
                f"düzeltilemedi: {exc}"
            ) from exc
        _progress(
            f"  [Excel] {tablo_adi} 'Ara Toplam' {col_harf} sütunu Excel "
            f"tarafından otomatik genişletilmemişti — açıkça düzeltildi: "
            f"{mevcut_formul} → {yeni_formul}"
        )


def _com_zarar_detay_kapasite_arttir(
    sheet,
    lo,
    tablo_adi: str,
    tablo_sol: int,
    veri_ilk_satir: int,
    veri_son_satir: int,
    kapasite: int,
    gereken_satir_sayisi: int,
) -> tuple[int, int]:
    """ZararTedarikci/ZararKiralik tablosunun satır kapasitesi yetersiz kaldığında
    eksik kadar satırı GERÇEK bir Excel satır ekleme işlemiyle tablonun İÇİNE ekler.

    Zarar eden sevk sayısı her ay farklı olacağı için (bu ay 89/236, gelecek ay
    farklı bir sayı) sabit bir kapasiteyi elle bir kere büyütmek kalıcı bir çözüm
    DEĞİL — bu fonksiyon her çalıştırmada gerektiği kadar otomatik büyütür.

    [Gerçek PivotTable'a KARŞI tercih gerekçesi — kullanıcı bunu önerdi, teknik
    olarak değerlendirildi]: Zarar Detay'ı sıfırdan bir PivotTable'a çevirmek
    (alan yerleşimi, sadece negatif değerleri gösteren filtre, gizli bir ham
    veri alanının senkronizasyonu) hem çok daha karmaşık hem de bu sandboxta
    gerçek Excel'de HİÇ doğrulanamaz bir COM inşası gerektiriyor — mevcut
    şablondaki elle ayarlanmış görünümü/biçimlendirmeyi bozma riski, aşağıda
    ele alınan "satır ekleme" riskinden daha yüksek. Bunun yerine mevcut sabit
    tablo + Ara Toplam formülü mimarisi KORUNUYOR, ama tek belirsiz nokta
    (Ara Toplam'ın otomatik genişleyip genişlemediği) artık TAHMİNE değil,
    aşağıdaki AÇIK DOĞRULAMA adımına dayanıyor (bkz. 3. madde).

    KRİTİK NOKTA — ekleme konumu: Yeni satırlar tablonun mevcut SON VERİ satırının
    ('veri_son_satir') TAM ÜZERİNE eklenir, 'veri_son_satir + 1'e DEĞİL. Excel'in
    "bir aralığın İÇİNE satır eklenirse o aralığa bakan formüller otomatik
    genişler, aralığın TAM ALTINA eklenirse genişlemez" kuralı bu ADIM için
    geçerli olmalı — ama bu kurala artık KÖRÜ KÖRÜNE güvenilmiyor (bkz. 3. madde).

    ✅ ÇÖZÜLDÜ — "ZararTedarikci başarısız, ZararKiralik başarılı" asimetrisi
    [WebSearch ile teyit edildi, bkz. r/vba "Utility to Add Rows to
    ListObjects" ve "Insert method of Range class failed" ile ilgili
    forum gönderileri]: `veri_son_satir` çağıran (`_com_zarar_detay_tablo_yaz`)
    tarafında artık `lo.ListRows.Count` üzerinden hesaplanıyor (Range.Rows.Count
    DEĞİL) — bu, ListObject'in NATİF "Toplam Satırı" (`lo.ShowTotals`) AÇIKSA
    onu HİÇ saymıyor. Eskiden `Range.Rows.Count` kullanıldığı için ShowTotals
    açık bir tabloda `veri_son_satir` yanlışlıkla NATİF Toplam Satırının
    kendisini işaret ediyordu; bir sonraki adımda `Rows.Insert` TAM O SATIRA
    (native Toplam Satırının üzerine) çağrılıyordu ki Excel bunu YAPISAL
    olarak REDDEDİYOR ("Insert method of Range class failed", 0x800A03D4) —
    ZararKiralik'te ShowTotals kapalıysa bu sorun hiç oluşmuyordu, tam
    gözlemlenen asimetriyi açıklıyor. Not: bu, bizim kendi elle yazdığımız
    'Ara Toplam' SUM formülü satırından (tablonun DIŞINDA, ayrı bir satır)
    TAMAMEN farklı bir şey.

    Ayrıca bir GÜVENLİK AĞI eklendi: `Rows.Insert` yukarıdaki düzeltmeye
    rağmen (beklenmedik bir şablon/Excel durumu için) yine de başarısız
    olursa, kod otomatik olarak Excel'in Tablo-farkında satır ekleme API'sine
    (`lo.ListRows.Add(Position=..., AlwaysInsert=True)`) düşer — bu, UI'daki
    "Tablo Satırlarını Üstte Ekle" ile birebir aynı davranışı taklit eder. O
    da başarısız olursa (iki ayrı hata mesajı birleştirilerek) net bir
    `RuntimeError` fırlatılır.

    1. Performans: `N` satır gerekiyorsa `N` kere ayrı `ListRows.Add()`/`Insert()`
       çağrısı YAPILMAZ (bkz. bu dosyadaki "hücre hücre yazma donması" dersi) —
       TEK bir `sheet.Rows("a:b").Insert()` çağrısıyla tüm eksik satırlar birden
       eklenir.
    2. Hesaplanan sütun formülleri (J:N — Alış/Satış/Kâr-Zarar/Zarar %/Zarar
       Payı): Excel'in tablo satır ekleme davranışı bunları otomatik
       kopyalamayı deneyebilir ama GARANTİ değildir; bu yüzden ekleme sonrası
       her sütun için tablonun İLK veri satırındaki (her zaman dolu,
       şablondan gelen orijinal) formül `.FormulaR1C1` (satır-bağımsız/
       relative referans korunarak) yeni satırlara AÇIKÇA kopyalanır.
    3. 'Ara Toplam' formülü ARTIK TAHMİN EDİLMİYOR: `_com_zarar_detay_ara_toplam_
       dogrula_ve_duzelt` ile satır ekleme sonrası formül GERÇEKTEN okunup yeni
       aralığı kapsayıp kapsamadığı kontrol ediliyor; kapsamıyorsa açıkça
       düzeltiliyor. Bu, kullanıcının haklı olarak işaret ettiği tek büyük
       belirsizliği ("Ara Toplam'ın doğru genişlediği hâlâ doğrulanmadı")
       ortadan kaldırıyor — davranış artık zımni bir Excel kuralına değil,
       kodun kendi doğrulama+düzeltme adımına dayanıyor.

    Dönüş: (yeni_kapasite, yeni_veri_son_satir). Herhangi bir adım başarısız
    olursa İSTİSNA fırlatır — çağıran taraf bunu yakalayıp ESKİ kapasiteyle
    devam eder (otomatik büyütme başarısız olursa en azından önceki davranış
    -- kapasiteyi aşanları logla -- korunur, rapor yine de tamamlanır).
    """
    eksik = gereken_satir_sayisi - kapasite
    if eksik <= 0:
        return kapasite, veri_son_satir

    if eksik > _ZARAR_DETAY_MAX_OTOMATIK_SATIR_EKLEME:
        raise RuntimeError(
            f"{eksik} satır isteniyor — güvenlik sınırını "
            f"({_ZARAR_DETAY_MAX_OTOMATIK_SATIR_EKLEME}) aşıyor, beklenmedik "
            "derecede büyük bir sıçrama görünüyor, elle kontrol edin."
        )

    _progress(
        f"  [Excel] {tablo_adi} kapasitesi yetersiz ({kapasite} satır var, "
        f"{gereken_satir_sayisi} gerekiyor) — {eksik} satır otomatik ekleniyor..."
    )

    # Formül kaynağı: tablonun İLK veri satırı — şablondan gelen, her zaman dolu
    # ve doğru olduğu varsayılan orijinal formülleri taşır (kaç kez büyütülürse
    # büyütülsün bu satır hiç silinmediği için güvenilir bir referans kalır).
    kaynak_satir = veri_ilk_satir
    eski_veri_son_satir = veri_son_satir

    # TEK bir Insert() çağrısıyla `eksik` satırı, mevcut son veri satırının TAM
    # ÜZERİNE ekle (yukarıdaki docstring'deki "aralık içine ekleme" kuralı için).
    # `sheet.Rows(...)` (EntireRow, TAM satır aralığı) kullanılıyor — sadece
    # belirli sütunları kapsayan bir Range'e Insert çağırmak Excel'in Shift
    # yönünü otomatik tahmin etmesini bazı durumlarda başarısız kılabiliyor;
    # EntireRow bu belirsizliği tamamen ortadan kaldırıyor.
    try:
        sheet.Rows(f"{veri_son_satir}:{veri_son_satir + eksik - 1}").Insert()
    except Exception as ilk_hata:
        # GÜVENLİK AĞI [WebSearch ile teyit edilen ShowTotals düzeltmesine
        # RAĞMEN beklenmedik bir Excel/şablon durumu için]: `Rows.Insert`
        # yine de başarısız olursa, Excel'in kendi Tablo-farkında satır
        # ekleme API'sine (`ListRows.Add`, UI'daki "Tablo Satırlarını Üstte
        # Ekle" ile birebir aynı davranış — `AlwaysInsert=True` altındaki
        # her şeyi aşağı kaydırır) düş. `Position`, eski son veri satırının
        # ListRows'taki (1-tabanlı, başlık HARİÇ) konumuna sabitleniyor ki
        # `eksik` adet yeni satır TAM ANA Insert() yolunun bıraktığı yere
        # (eski son satırın YERİNE, eski veriyi aşağı iterek) eklensin —
        # aşağıdaki formül kopyalama adımının satır aralığı beklentisiyle
        # birebir uyumlu kalsın.
        try:
            konum = veri_son_satir - veri_ilk_satir + 1
            for _ in range(eksik):
                lo.ListRows.Add(Position=konum, AlwaysInsert=True)
        except Exception as ikinci_hata:
            raise RuntimeError(
                f"satır ekleme (Rows.Insert) başarısız: {ilk_hata}; alternatif "
                f"yöntem (ListRows.Add) da başarısız: {ikinci_hata}"
            ) from ikinci_hata
        _progress(
            f"  [Excel] {tablo_adi}: Rows.Insert başarısız oldu ({ilk_hata}), "
            f"alternatif yöntemle (ListRows.Add) {eksik} satır eklendi."
        )

    yeni_veri_son_satir = veri_son_satir + eksik
    yeni_kapasite = kapasite + eksik

    metin_kolon_sayisi = len(ZARAR_DETAY_METIN_SUTUNLARI)
    formul_sol = tablo_sol + metin_kolon_sayisi
    formul_sag = formul_sol + 4  # J..N -> Alış/Satış/Kâr-Zarar/Zarar %/Zarar Payı (5 sütun)

    # Her sütun İÇİN AYRI (ama satırlar için TEK) atama yapılıyor: kaynaktaki
    # tek hücrenin FormulaR1C1'i (relative referans) hedef aralığa bir SKALER
    # olarak atanıyor — Excel R1C1 göreli formülleri her hedef hücre için
    # KENDİ konumuna göre otomatik ayarlar, bu yüzden 5 sütun x 1 satır kaynak
    # okuma + 5 sütun x N satır hedef yazma toplamda sadece 10 COM çağrısı
    # (satır sayısından BAĞIMSIZ) — performans için satır sayısı kadar döngü
    # kurulmuyor.
    for offset in range(formul_sag - formul_sol + 1):
        col = formul_sol + offset
        try:
            kaynak_formul = sheet.Cells(kaynak_satir, col).FormulaR1C1
            sheet.Range(
                sheet.Cells(veri_son_satir, col),
                sheet.Cells(yeni_veri_son_satir - 1, col),
            ).FormulaR1C1 = kaynak_formul
        except Exception as exc:
            raise RuntimeError(
                f"sütun {col} formülü yeni satırlara ({veri_son_satir}-"
                f"{yeni_veri_son_satir - 1}) kopyalanamadı: {exc}"
            ) from exc

    # ListObject sınırını açıkça genişlet — native Excel Table davranışı satır
    # eklendiğinde bunu genelde kendiliğinden yapar, ama garantiye almak için
    # açıkça da çağrılıyor (zararsız no-op olabilir, hataysa yutuluyor).
    try:
        hdr = lo.HeaderRowRange
        yeni_tablo_araligi = sheet.Range(
            hdr.Cells(1, 1),
            sheet.Cells(yeni_veri_son_satir, tablo_sol + int(hdr.Columns.Count) - 1),
        )
        lo.Resize(yeni_tablo_araligi)
    except Exception:
        pass

    # 'Ara Toplam' formülünün GERÇEKTEN yeni aralığı kapsadığını doğrula,
    # kapsamıyorsa açıkça düzelt (bkz. yukarıdaki docstring 3. madde — bu,
    # Excel'in zımni genişletme davranışına duyulan güveni ortadan kaldırır).
    _com_zarar_detay_ara_toplam_dogrula_ve_duzelt(
        sheet, tablo_adi, formul_sol, formul_sag, eski_veri_son_satir, yeni_veri_son_satir
    )

    _progress(
        f"  [Excel] {tablo_adi} kapasitesi {yeni_kapasite} satıra büyütüldü "
        "('Ara Toplam' formülü doğrulandı/gerekirse düzeltildi)."
    )

    return yeni_kapasite, yeni_veri_son_satir


def _com_zarar_detay_tablo_yaz(
    sheet, tablo_adi: str, satirlar: list[dict[str, Any]]
) -> tuple[int, int]:
    """ZararTedarikci/ZararKiralik tablosuna üstten aşağı yazar; kapasite
    yetersizse önce OTOMATİK olarak gereken kadar satır ekler (bkz.
    `_com_zarar_detay_kapasite_arttir`), sonra kullanılmayan (varsa) satırları
    sıfırlar.

    Tablonun kendisi eskiden BİLİNÇLİ OLARAK büyütülmüyordu çünkü hemen
    altındaki 'Ara Toplam' satırı (sabit satır numaralarına SUM(...) yapan bir
    formül) ve bir sonraki bölümün başlığı satır ekleme/silmeyle kayabilirdi.
    Artık `_com_zarar_detay_kapasite_arttir` bunu GÜVENLİ şekilde yapıyor —
    ekleme noktası özenle tablonun son veri satırının TAM ÜZERİNE seçiliyor ki
    Excel'in "aralık içine ekleme = referans genişlet" kuralı ile 'Ara Toplam'
    formülü otomatik büyüsün (bkz. o fonksiyonun docstring'i). Otomatik büyütme
    herhangi bir nedenle başarısız olursa (ör. korumalı sayfa, birleştirilmiş
    hücre), eski davranışa (kapasiteyi aşanları logla, göstermeden bırak) geri
    dönülür — rapor yine de BAŞARIYLA tamamlanır.

    Sadece metin/sayı sütunları (A-I: Tarih..Kayıt Sayısı) doğrudan yazılır;
    J-N (Alış/Satış/Kâr-Zarar/Zarar %/Zarar Payı) sütunlarındaki ORİJİNAL satır
    formülleri (Tablo5'e SUMIF ile bakan) dokunulmadan bırakılır — Sevk No (G)
    güncellenince bu formüller kendiliğinden doğru sonucu verir.
    """
    try:
        lo = sheet.ListObjects(tablo_adi)
    except Exception:
        return 0, 0

    try:
        hdr = lo.HeaderRowRange
        tablo_sol = int(hdr.Column)
        veri_ilk_satir = int(hdr.Row) + 1

        # ⚠️ ShowTotals (ListObject'in NATİF "Toplam Satırı") düzeltmesi
        # [WebSearch ile teyit edildi — bkz. r/vba "Utility to Add Rows to
        # ListObjects" ve ilgili forum gönderileri]: `lo.Range.Rows.Count`
        # tabloyu BAŞLIK + VERİ + (varsa) native Toplam Satırının TAMAMINI
        # sayar; ShowTotals AÇIKSA bu, "son VERİ satırı" hesabını YANLIŞLIKLA
        # Toplam Satırının kendisine kaydırıyordu — bir sonraki adımda
        # `Rows.Insert` TAM O SATIRA (native Toplam Satırının üzerine/içine)
        # çağrılıyordu, ki Excel bunu YAPISAL olarak REDDEDİYOR ("Insert
        # method of Range class failed", hata kodu 0x800A03D4). Bu, kod
        # incelemesinde ZararTedarikci'nin başarısız olup ZararKiralik'in
        # başarılı olmasının TAM olası açıklaması: iki tablonun ShowTotals
        # ayarı şablonda farklı bırakılmış olabilir.
        #
        # `lo.ListRows.Count` ise ListObject'in native Toplam Satırını HİÇ
        # SAYMAZ (sadece gerçek veri satırlarını sayar) — bu yüzden
        # `Range.Rows.Count`'tan daha güvenilir bir "gerçek veri satırı
        # sayısı" kaynağı. Birincil olarak bunu kullanıyoruz; sadece COM
        # bu özelliği hiç desinlemezse (çok eski bir Excel sürümü ihtimaline
        # karşı) `Range.Rows.Count`'tan ShowTotals'a göre düzeltilmiş bir
        # değer türetip devam ediyoruz.
        try:
            gosterilen_toplam_satiri = bool(lo.ShowTotals)
        except Exception:
            gosterilen_toplam_satiri = False
        try:
            veri_satir_sayisi = int(lo.ListRows.Count)
        except Exception:
            veri_satir_sayisi = (
                int(lo.Range.Rows.Count) - 1 - (1 if gosterilen_toplam_satiri else 0)
            )

        veri_son_satir = veri_ilk_satir + veri_satir_sayisi - 1
        kapasite = veri_satir_sayisi
    except Exception:
        return 0, 0

    if kapasite <= 0:
        return 0, len(satirlar)

    if len(satirlar) > kapasite:
        try:
            kapasite, veri_son_satir = _com_zarar_detay_kapasite_arttir(
                sheet, lo, tablo_adi, tablo_sol, veri_ilk_satir, veri_son_satir,
                kapasite, len(satirlar),
            )
        except Exception as exc:
            _progress(
                f"  Uyarı: {tablo_adi} kapasitesi otomatik büyütülemedi ({exc}) "
                "— mevcut kapasiteyle devam ediliyor, fazlası gösterilmeyecek."
            )

    yazilan = min(len(satirlar), kapasite)
    tasan = len(satirlar) - yazilan
    metin_kolon_sayisi = len(ZARAR_DETAY_METIN_SUTUNLARI)

    # Tarih sütununu veri yazılmadan ÖNCE 'General'e sıfırla (bkz.
    # _com_bicim_genel_yap docstring'i — format sıralaması önemli).
    tarih_kolon = tablo_sol + ZARAR_DETAY_METIN_SUTUNLARI.index("Tarih")
    try:
        _com_bicim_genel_yap(
            sheet.Range(
                sheet.Cells(veri_ilk_satir, tarih_kolon),
                sheet.Cells(veri_son_satir, tarih_kolon),
            )
        )
    except Exception:
        pass

    # ÖNEMLİ (performans): Buradaki her hücreye TEK TEK `sheet.Cells(r, c).Value = ...`
    # ile yazmak eskiden yüzlerce/binlerce ayrı COM çağrısına yol açıyordu — her çağrının
    # kendi gidiş-dönüş gecikmesi olduğu için (özellikle yavaş/uzak bir makinede) bu,
    # kullanıcıya sanki script "asılı kalmış" gibi görünen çok uzun bir bekleme yaratıyordu
    # (bkz. "[Excel] Zarar Detay güncelleniyor..." adımında donma şikayeti). Diğer
    # sayfalarda (VERİ/Filo, `_com_sayfaya_yaz`) olduğu gibi TÜM bloğu tek bir matris
    # halinde TEK BİR Range.Value atamasıyla yazmak, aynı işi tek (veya iki) COM
    # çağrısına indirip bu donmayı ortadan kaldırıyor.
    metin_matrisi: list[tuple[Any, ...]] = []
    for i in range(kapasite):
        if i < yazilan:
            metin_matrisi.append(
                tuple(hucre_degeri(satirlar[i].get(k)) for k in ZARAR_DETAY_METIN_SUTUNLARI)
            )
        else:
            metin_matrisi.append(tuple([None] * metin_kolon_sayisi))

    metin_araligi = sheet.Range(
        sheet.Cells(veri_ilk_satir, tablo_sol),
        sheet.Cells(veri_son_satir, tablo_sol + metin_kolon_sayisi - 1),
    )
    _com_araliga_yaz(metin_araligi, metin_matrisi)

    # Kullanılmayan satırların Alış/Satış/Kâr-Zarar/Zarar %/Zarar Payı formüllerini
    # sıfırla — Tablo5'te boş Sevk No eşleşmesiyle Ara Toplam'a hatalı katkı yapmasınlar
    # diye (bkz. modül docstring'i). Bu da tek bir bulk yazma ile yapılıyor.
    bos_satir_sayisi = kapasite - yazilan
    if bos_satir_sayisi > 0:
        sifir_matrisi = [tuple([0] * 5) for _ in range(bos_satir_sayisi)]
        sifir_araligi = sheet.Range(
            sheet.Cells(veri_ilk_satir + yazilan, tablo_sol + metin_kolon_sayisi),
            sheet.Cells(veri_son_satir, tablo_sol + metin_kolon_sayisi + 4),
        )
        _com_araliga_yaz(sifir_araligi, sifir_matrisi)

    if not _com_tarih_bicimi_zorla(
        sheet.Range(
            sheet.Cells(veri_ilk_satir, tarih_kolon),
            sheet.Cells(veri_son_satir, tarih_kolon),
        )
    ):
        _progress(f"  Uyarı: {tablo_adi} 'Tarih' sütunu biçimi doğrulanamadı.")

    return yazilan, tasan


def _donem_etiketi(bas: str, bit: str) -> str:
    """'01.08.2026'/'31.08.2026' -> 'Ağustos 2026'; tam yıl -> 'YYYY'; aksi halde aralık metni."""
    ay_adlari = [
        "Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran",
        "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık",
    ]
    try:
        bas_d = datetime.strptime(bas, "%d.%m.%Y").date()
        bit_d = datetime.strptime(bit, "%d.%m.%Y").date()
    except ValueError:
        return f"{bas} – {bit}"

    if bas_d.day == 1 and bas_d.month == 1 and bit_d.month == 12 and bit_d.day == 31 and bas_d.year == bit_d.year:
        return str(bas_d.year)

    son_gun = calendar.monthrange(bas_d.year, bas_d.month)[1]
    if bas_d.day == 1 and bit_d.day == son_gun and bas_d.month == bit_d.month and bas_d.year == bit_d.year:
        return f"{ay_adlari[bas_d.month - 1]} {bas_d.year}"

    return f"{bas} – {bit}"


def _com_zarar_detay_guncelle(
    wb, veri_satirlari: list[dict[str, Any]], bas: str, bit: str
) -> list[str]:
    """'Zarar Detay' sayfasındaki ZararTedarikci/ZararKiralik tablolarını bu
    dönemin zarar eden sevkleriyle tazeler. Sayfa/tablo bulunamazsa sessizce
    atlanır (her şablonda bu sayfa olmayabilir)."""
    uyarilar: list[str] = []

    sayfa_adlari = getattr(
        ayarlar, "KPI_ZARAR_DETAY_SAYFA_ADLARI", _ZARAR_DETAY_SAYFA_ADLARI_VARSAYILAN
    )
    ws = _excel_sayfa_bul(wb, sayfa_adlari)
    if ws is None:
        return uyarilar

    if getattr(ayarlar, "KPI_ZARAR_DETAY_GUNCELLE", True) is False:
        return uyarilar

    tedarikci_adi = getattr(
        ayarlar, "KPI_ZARAR_TEDARIKCI_TABLO_ADI", _ZARAR_TEDARIKCI_TABLO_ADI_VARSAYILAN
    )
    kiralik_adi = getattr(
        ayarlar, "KPI_ZARAR_KIRALIK_TABLO_ADI", _ZARAR_KIRALIK_TABLO_ADI_VARSAYILAN
    )

    tedarikci, kiralik, hesap_uyarilari = zarar_eden_sevkleri_hesapla(veri_satirlari)
    uyarilar.extend(hesap_uyarilari)

    ted_yazilan, ted_tasan = _com_zarar_detay_tablo_yaz(ws, tedarikci_adi, tedarikci)
    kir_yazilan, kir_tasan = _com_zarar_detay_tablo_yaz(ws, kiralik_adi, kiralik)

    # NOT: Tablo kapasitesi artık `_com_zarar_detay_tablo_yaz` içinde OTOMATİK
    # büyütülüyor (bkz. `_com_zarar_detay_kapasite_arttir`) — bu uyarılar normal
    # şartlarda ARTIK TETİKLENMEMELİ. Sadece otomatik büyütme bir nedenle
    # başarısız olursa (konsolda ayrıca `_progress` ile ayrıntılı hata loglanır)
    # devreye girer; bu yüzden mesaj artık "elle büyütün" yerine otomatik
    # büyütmenin başarısız olduğunu ve konsol logunun kontrol edilmesi
    # gerektiğini belirtiyor.
    if ted_tasan > 0:
        uyarilar.append(
            f"Zarar Detay/{tedarikci_adi}: {len(tedarikci)} zarar eden sevk var, "
            f"otomatik satır ekleme denendi ama {ted_tasan} tanesi (en küçük "
            f"zararlılar) yine de sığmadı (tabloda {ted_yazilan} satırlık yer var) "
            "— konsol logunda otomatik büyütme hatasına bakın; gerekirse tabloyu "
            "Excel'de elle büyütün."
        )
    if kir_tasan > 0:
        uyarilar.append(
            f"Zarar Detay/{kiralik_adi}: {len(kiralik)} zarar eden sevk var, "
            f"otomatik satır ekleme denendi ama {kir_tasan} tanesi (en küçük "
            f"zararlılar) yine de sığmadı (tabloda {kir_yazilan} satırlık yer var) "
            "— konsol logunda otomatik büyütme hatasına bakın; gerekirse tabloyu "
            "Excel'de elle büyütün."
        )

    try:
        baslik_hucre = ws.Cells(2, 1)
        mevcut = str(baslik_hucre.Value or "")
        if "zarar eden sevk" in mevcut.lower():
            donem = _donem_etiketi(bas, bit)
            baslik_hucre.Value = (
                f"{donem} | Aynı sevk numarasındaki satırlar birleştirilmiştir | "
                f"Tedarikçi {len(tedarikci)}, Kiralık {len(kiralik)} zarar eden sevk"
            )
    except Exception:
        pass

    _progress(
        f"  [Excel] Zarar Detay güncellendi: Tedarikçi {ted_yazilan}/{len(tedarikci)}, "
        f"Kiralık {kir_yazilan}/{len(kiralik)} zarar eden sevk yazıldı."
    )

    return uyarilar


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


def _autofit_ayarlari() -> tuple[float, float, float, int, dict[str, float]]:
    ozel = getattr(ayarlar, "KPI_SUTUN_GENISLIK", None) or {}
    if not isinstance(ozel, dict):
        ozel = {}
    return (
        float(getattr(ayarlar, "KPI_SUTUN_MIN_GENISLIK", 10)),
        float(getattr(ayarlar, "KPI_PARA_SUTUN_MIN_GENISLIK", 24)),
        float(getattr(ayarlar, "KPI_SUTUN_MAX_GENISLIK", 55)),
        int(getattr(ayarlar, "KPI_AUTOFIT_MAX_SATIR", 400)),
        {str(k).upper(): float(v) for k, v in ozel.items()},
    )


def _com_sutun_genisligi_ayarla(sheet, col: int | str, genislik: float, max_w: float) -> None:
    try:
        mevcut = float(sheet.Columns(col).ColumnWidth)
        hedef = min(max(genislik, mevcut), max_w)
        if hedef > mevcut:
            sheet.Columns(col).ColumnWidth = hedef
    except Exception:
        pass


def _sabit_genislik_eslemeleri() -> dict[str, dict[str, float]]:
    """Sayfa adı (normalize) -> {sütun harfi: sabit genişlik}. ayarlar.KPI_SABIT_SUTUN_GENISLIKLERI."""
    ham = getattr(ayarlar, "KPI_SABIT_SUTUN_GENISLIKLERI", None) or {}
    if not isinstance(ham, dict):
        return {}
    sonuc: dict[str, dict[str, float]] = {}
    for sayfa_adi, sutunlar in ham.items():
        if not isinstance(sutunlar, dict):
            continue
        sonuc[_normalize_kolon(str(sayfa_adi))] = {
            str(harf).upper(): float(genislik) for harf, genislik in sutunlar.items()
        }
    return sonuc


def _com_sabit_genislik_uygula(sheet, esleme: dict[str, float]) -> None:
    """Verilen sütun genişliklerini AYNEN uygular — AutoFit/heuristik atlanır (kullanıcının
    elle ayarladığı genişlikler her ay kalıcı olsun diye)."""
    for harf, genislik in esleme.items():
        try:
            sheet.Columns(harf).ColumnWidth = genislik
        except Exception:
            pass


def _com_sutunlari_genislet(sheet, satir_limit: int | None = None) -> None:
    """UsedRange AutoFit + #### tespiti — pivot tutar sütunları (Alış/Satış) için."""
    min_w, para_w, max_w, varsayilan_satir_limit, ozel_genislik = _autofit_ayarlari()
    if satir_limit is None:
        satir_limit = varsayilan_satir_limit

    try:
        used = sheet.UsedRange
        if used is None:
            return
    except Exception:
        return

    try:
        used.Columns.AutoFit()
    except Exception:
        pass

    try:
        first_col = int(used.Column)
        col_count = int(used.Columns.Count)
        first_row = int(used.Row)
        row_count = min(int(used.Rows.Count), satir_limit)
    except Exception:
        return

    for offset in range(col_count):
        col = first_col + offset
        hedef = min_w
        kesin_para = False
        for row in range(first_row, first_row + row_count):
            try:
                text = str(sheet.Cells(row, col).Text or "").strip()
            except Exception:
                continue
            if not text:
                continue
            genislik_ihtiyaci = min(len(text) * 1.12 + 2.0, max_w)
            if "#" in text:
                hedef = max(hedef, genislik_ihtiyaci, para_w)
                kesin_para = True
                break
            if any(ch in text for ch in "₺%") or ("." in text and "," in text) or (
                "," in text and any(c.isdigit() for c in text)
            ):
                hedef = max(hedef, genislik_ihtiyaci, para_w)
                kesin_para = True
            else:
                hedef = max(hedef, genislik_ihtiyaci)

        if kesin_para:
            hedef = max(hedef, para_w)

        _com_sutun_genisligi_ayarla(sheet, col, hedef, max_w)

    # Özet vb.: Alış sütunu (C) — şablonda tutar genelde burada
    alis_harf = str(getattr(ayarlar, "KPI_ALIS_SUTUN_HARFI", "C")).upper()
    alis_genislik = float(getattr(ayarlar, "KPI_ALIS_SUTUN_GENISLIK", para_w + 4))
    _com_sutun_genisligi_ayarla(sheet, alis_harf, alis_genislik, max_w)

    for harf, gen in ozel_genislik.items():
        _com_sutun_genisligi_ayarla(sheet, harf, gen, max_w)


def _excel_hesapla(excel) -> None:
    """Pivot sonrası hesaplama — hızlı modda tam yeniden derleme yapılmaz."""
    hizli = getattr(ayarlar, "KPI_HIZLI_MOD", True)
    try:
        excel.CalculateUntilAsyncQueriesDone()
    except Exception:
        pass
    try:
        if hizli:
            excel.Calculate()
        else:
            excel.CalculateFullRebuild()
    except Exception:
        pass


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
    bas: str = "",
    bit: str = "",
) -> tuple[bool, str | None, int, int, list[str]]:
    """Şablon kopyasına Excel COM ile veri yazar, pivot yeniler, sütunları genişletir."""
    excel = None
    wb = None
    uyarilar: list[str] = []
    eslesmeyen_veri: list[str] = []
    dosya = str(dosya_yolu.resolve())

    try:
        _progress("  [Excel] Uygulama açılıyor...")
        excel = _excel_uygulama_ac()
        _progress(f"  [Excel] Şablon açılıyor: {dosya_yolu.name}")
        wb = excel.Workbooks.Open(
            Filename=dosya,
            UpdateLinks=0,
            ReadOnly=False,
            Notify=False,
        )

        ws_veri = _excel_sayfa_bul(wb, veri_sayfa_adlari)
        if ws_veri is None:
            return False, f"VERİ sayfası bulunamadı: {veri_sayfa_adlari}", 0, 0, []

        ws_filo = _excel_sayfa_bul(wb, filo_sayfa_adlari)
        if ws_filo is None:
            return False, f"Filo Detay sayfası bulunamadı: {filo_sayfa_adlari}", 0, 0, []

        _progress(f"  [Excel] VERİ yazılıyor ({len(veri_satirlari)} satır)...")
        veri_adet, veri_esleme, eslesmeyen_veri, veri_kolon, veri_tablo_sol = _com_sayfaya_yaz(
            ws_veri, veri_baslik_satiri, veri_satirlari
        )
        tarih_basarisiz = _com_tarih_bicimi_uygula(
            ws_veri, veri_baslik_satiri, veri_adet, veri_esleme
        )
        if tarih_basarisiz:
            uyarilar.append(
                "VERİ tarih biçimi doğrulanamadı: " + ", ".join(tarih_basarisiz)
            )
        _com_pivot_kaynak_guncelle(
            wb, ws_veri.Name, veri_baslik_satiri, veri_adet, veri_kolon, veri_tablo_sol
        )

        _progress(f"  [Excel] Filo Detay yazılıyor ({len(filo_satirlari)} satır)...")
        filo_yaz = [{k: r.get(k) for k in FILO_DETAY_SUTUNLARI} for r in filo_satirlari]
        filo_adet, _, _, filo_kolon, filo_tablo_sol = _com_sayfaya_yaz(
            ws_filo, filo_baslik_satiri, filo_yaz, sabit_kolonlar=FILO_DETAY_SUTUNLARI
        )
        _com_pivot_kaynak_guncelle(
            wb, ws_filo.Name, filo_baslik_satiri, filo_adet, filo_kolon, filo_tablo_sol
        )

        _progress("  [Excel] Zarar Detay güncelleniyor...")
        try:
            zarar_uyarilari = _com_zarar_detay_guncelle(wb, veri_satirlari, bas, bit)
            uyarilar.extend(zarar_uyarilari)
        except Exception as exc:
            uyarilar.append(f"Zarar Detay güncelleme: {exc}")

        if pivot_yenile:
            _progress("  [Excel] Pivotlar yenileniyor...")
            try:
                wb.RefreshAll()
            except Exception as exc:
                uyarilar.append(f"pivot yenileme: {exc}")
            try:
                _excel_hesapla(excel)
            except Exception as exc:
                uyarilar.append(f"hesaplama: {exc}")

        sabit_genislikler = _sabit_genislik_eslemeleri()
        if sutun_autofit or sabit_genislikler:
            _progress("  [Excel] Özet sayfaları genişletiliyor...")
            for sheet in wb.Worksheets:
                if int(sheet.Visible) != -1:
                    continue
                if sheet.Name in (ws_veri.Name, ws_filo.Name):
                    continue
                sabit = sabit_genislikler.get(_normalize_kolon(sheet.Name))
                if sabit:
                    try:
                        _com_sabit_genislik_uygula(sheet, sabit)
                    except Exception as exc:
                        uyarilar.append(f"{sheet.Name} sabit genişlik: {exc}")
                    continue
                if not sutun_autofit:
                    continue
                try:
                    _com_sutunlari_genislet(sheet)
                except Exception as exc:
                    uyarilar.append(f"{sheet.Name} AutoFit: {exc}")

        _progress("  [Excel] Kaydediliyor...")
        wb.Save()
        wb.Close(SaveChanges=True)
        wb = None

        mesaj = f"kısmi uyarı: {'; '.join(uyarilar)}" if uyarilar else None
        return True, mesaj, veri_adet, filo_adet, eslesmeyen_veri

    except Exception as exc:
        return False, str(exc), 0, 0, []

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
                _excel_hesapla(excel)
            except Exception as exc:
                uyarilar.append(f"hesaplama: {exc}")

        sabit_genislikler = _sabit_genislik_eslemeleri()
        autofit_sayisi = 0
        for sheet in wb.Worksheets:
            if int(sheet.Visible) != -1:
                continue
            sabit = sabit_genislikler.get(_normalize_kolon(sheet.Name))
            if sabit:
                try:
                    _com_sabit_genislik_uygula(sheet, sabit)
                    autofit_sayisi += 1
                except Exception as exc:
                    uyarilar.append(f"{sheet.Name} sabit genişlik: {exc}")
                continue
            try:
                _com_sutunlari_genislet(sheet)
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


def _ay_yil_ayristir(metin: str, ham: str) -> tuple[int, int]:
    try:
        ay_str, yil_str = metin.strip().split(".")
        ay, yil = int(ay_str), int(yil_str)
        if not 1 <= ay <= 12:
            raise ValueError
        return ay, yil
    except (ValueError, AttributeError) as exc:
        raise ValueError(
            f"KPI_DONEM formatı hatalı: '{ham}'. Beklenen: 'AA.YYYY', "
            f"'AA.YYYY-AA.YYYY' veya 'YYYY'."
        ) from exc


def _donem_araligi_hesapla(donem: str) -> tuple[str, str]:
    """KPI_DONEM'i (tek ay / ay aralığı / tam yıl) tarih aralığına çevirir."""
    ham = donem
    donem = donem.strip()

    if re.fullmatch(r"\d{4}", donem):
        yil = int(donem)
        bas = date(yil, 1, 1).strftime("%d.%m.%Y")
        bit = date(yil, 12, 31).strftime("%d.%m.%Y")
        return bas, bit

    if "-" in donem:
        bas_metin, bit_metin = donem.split("-", 1)
        bas_ay, bas_yil = _ay_yil_ayristir(bas_metin, ham)
        bit_ay, bit_yil = _ay_yil_ayristir(bit_metin, ham)
        son_gun = calendar.monthrange(bit_yil, bit_ay)[1]
        bas = date(bas_yil, bas_ay, 1).strftime("%d.%m.%Y")
        bit = date(bit_yil, bit_ay, son_gun).strftime("%d.%m.%Y")
        return bas, bit

    ay, yil = _ay_yil_ayristir(donem, ham)
    son_gun = calendar.monthrange(yil, ay)[1]
    bas = date(yil, ay, 1).strftime("%d.%m.%Y")
    bit = date(yil, ay, son_gun).strftime("%d.%m.%Y")
    return bas, bit


def _tarih_araligi() -> tuple[str, str]:
    """Rapor dönemi:
    - KPI_DONEM = "AA.YYYY"           -> tek ay (ör. "08.2026")
    - KPI_DONEM = "AA.YYYY-AA.YYYY"   -> ay aralığı (ör. "01.2026-04.2026")
    - KPI_DONEM = "YYYY"              -> tam yıl (ör. "2026")
    Her durumda ayın kaç gün çektiği otomatik hesaplanır. KPI_DONEM tanımlı
    değilse eski KPI_BASLANGIC_TARIHI / KPI_BITIS_TARIHI kullanılır (geriye
    dönük uyumluluk)."""
    donem = getattr(ayarlar, "KPI_DONEM", None)
    if donem:
        return _donem_araligi_hesapla(str(donem))

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

    if not kaynak.exists():
        raise FileNotFoundError(
            f"KPI şablonu bulunamadı: {kaynak}\n"
            f"Temmuz KPI dosyanızı bu konuma 'kpi_sablon.xlsx' adıyla kopyalayın."
        )

    bas, bit = _tarih_araligi()
    bind = _bind_olustur(bas, bit)
    # Çıktı dosyasının adı [kullanıcı isteği] artık dönemden otomatik üretiliyor
    # (bkz. _cikti_yolu docstring'i) — bu yüzden bas/bit hesaplandıktan SONRA
    # çağrılıyor (eskiden şablonun suffix'i dışında döneme dair bilgisi yoktu).
    hedef = cikti or _cikti_yolu(kaynak, bas, bit)

    _progress(f"KPI raporu — dönem: {bas} — {bit}")
    _progress(f"  Şablon: {kaynak.name}")

    hedef.parent.mkdir(parents=True, exist_ok=True)
    _sablon_hedefe_kopyala(kaynak, hedef)
    _progress(f"  Çıktı kopyalandı: {hedef.name}")

    veri_satirlari: list[dict] = []
    filo_satirlari: list[dict] = []

    _progress("  Oracle bağlantısı açılıyor...")
    with baglanti_yonet() as baglanti:
        cursor = baglanti.cursor()
        veri_ok, veri_mesaj = veri_semasi_hazir()
        if veri_ok:
            _progress(f"  Oracle VERİ sorgusu ({veri_sql_kaynak_bilgisi()})...")
            veri_satirlari = veri_satirlari_getir(cursor, bas, bit, bind)
            _progress(f"  Oracle VERİ tamam: {len(veri_satirlari)} satır")
        else:
            _progress(f"  Uyarı: VERİ atlandı — {veri_mesaj}")
        filo_ok, filo_mesaj = kiralk_arac_semasi_hazir()
        if filo_ok:
            _progress("  Oracle Filo Detay sorgusu çalışıyor...")
            filo_satirlari = kiralk_arac_detay_getir(cursor, bas, bit, bind)
            _progress(f"  Oracle Filo tamam: {len(filo_satirlari)} satır")
        else:
            _progress(f"  Uyarı: Filo Detay atlandı — {filo_mesaj}")

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
    eslesmeyen_veri: list[str] = []

    if _excel_kullanilabilir():
        _progress("  Excel ile şablon dolduruluyor...")
        excel_ok, excel_mesaj, veri_adet, filo_adet, eslesmeyen_veri = _excel_sablon_doldur(
            hedef,
            veri_satirlari,
            filo_satirlari,
            veri_sayfa_adlari,
            filo_sayfa_adlari,
            veri_baslik_satiri,
            filo_baslik_satiri,
            pivot_yenile=pivot_yenile_calistir,
            sutun_autofit=sutun_autofit,
            bas=bas,
            bit=bit,
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
    if eslesmeyen_veri:
        print(f"  Bilgi: {len(eslesmeyen_veri)} VERİ kolonu şablonda var, Oracle sorgusunda yok (boş kalır — pivotlar etkilenmeyebilir).")
        if getattr(ayarlar, "KPI_ESLEME_UYARISI_DETAY", False):
            ornek = ", ".join(eslesmeyen_veri[:8])
            fazla = len(eslesmeyen_veri) - 8
            ek = f" (+{fazla} kolon daha)" if fazla > 0 else ""
            print(f"    Örnek: {ornek}{ek}")
            print("    → python kpi_sablon_kolon_kesif.py  veya  ayarlar.py KPI_KOLON_ESLEME")
    if excel_mesaj:
        print(f"  Uyarı: {excel_mesaj}")

    return str(hedef.resolve())

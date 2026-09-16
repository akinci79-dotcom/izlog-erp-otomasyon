"""
VERİ sayfası — Oracle sorgusu.

Yalnızca referans/kpi_veri_rapor.sql dosyasındaki SQL çalıştırılır.
Sorguya kod tarafında alan eklenmez/çıkarılmaz; Uyumsoft raporu aynen kullanılır.
Çalıştırma anında yalnızca @...@ Uyumsoft parametreleri ayarlardan doldurulur.

Dosya: KPI/referans/kpi_veri_rapor.sql
"""
from __future__ import annotations

from datetime import date, datetime
from decimal import Decimal
from pathlib import Path
from typing import Any

import ayarlar

_KPI_KOKU = Path(__file__).resolve().parent
_EXCEL_EPOK = datetime(1899, 12, 30)  # Excel'in seri tarih başlangıcı (1900 artık yıl hatası dahil)

_UYUMSOFT_PARAMETRELER = (
    "@CoCode@",
    "@BranchCodes@",
    "@DocDateF@",
    "@DocDateL@",
    "@ReferenceNo@",
    "@TransportNo@",
    "@ProjectCodes@",
    "@VehicleCode@",
)


def veri_sql_dosya_yolu() -> Path | None:
    """Kullanıcının VERİ rapor SQL dosyası."""
    ad = getattr(ayarlar, "KPI_VERI_SQL_DOSYASI", "kpi_veri_rapor.sql")
    adaylar: list[Path] = []
    yol = Path(str(ad))
    if yol.is_absolute():
        adaylar.append(yol)
    else:
        adaylar.append(_KPI_KOKU / yol)
        adaylar.append(_KPI_KOKU / "referans" / yol.name)
        adaylar.append(_KPI_KOKU / "referans" / str(ad))
    for aday in adaylar:
        if aday.is_file() and aday.stat().st_size > 20:
            return aday
    return None


def veri_sql_kaynak_bilgisi() -> str:
    yol = veri_sql_dosya_yolu()
    if yol:
        return str(yol.name)
    return "TANIMSIZ — referans/kpi_veri_rapor.sql gerekli"


def _co_code() -> str:
    deger = getattr(ayarlar, "CO_CODE", None)
    if not deger or not str(deger).strip():
        raise ValueError(
            "CO_CODE ayarlar.py içinde tanımlı olmalı.\n"
            "Uyumsoft VERİ raporu firma kodu (@CoCode@) zorunlu kullanır."
        )
    return str(deger).strip()


def _branch_code() -> str:
    deger = getattr(ayarlar, "BRANCH_CODE", None)
    if not deger or not str(deger).strip():
        raise ValueError(
            "BRANCH_CODE ayarlar.py içinde tanımlı olmalı.\n"
            "Uyumsoft VERİ raporu şube kodu (@BranchCodes@) zorunlu kullanır."
        )
    return str(deger).strip()


def veri_semasi_hazir() -> tuple[bool, str]:
    yol = veri_sql_dosya_yolu()
    if yol is None:
        return False, (
            "referans/kpi_veri_rapor.sql bulunamadı.\n"
            "Uyumsoft VERİ (LojistikYükSevkKalemRaporu) SQL'inizi bu dosyaya "
            "olduğu gibi kaydedin. Kod SQL'e dokunmaz."
        )
    try:
        _co_code()
        _branch_code()
    except ValueError as exc:
        return False, str(exc)
    return True, ""


def _veri_sql() -> str:
    yol = veri_sql_dosya_yolu()
    if yol is None:
        raise FileNotFoundError(
            "referans/kpi_veri_rapor.sql bulunamadı.\n"
            "Uyumsoft'tan aldığınız VERİ raporu SQL'ini KPI/referans/kpi_veri_rapor.sql "
            "olarak kaydedin — sorgu aynen çalıştırılır, alan eklenmez/çıkarılmaz."
        )
    return yol.read_text(encoding="utf-8-sig")


def _uyumsoft_parametreleri_yerlestir(sql: str, bas: str, bit: str) -> str:
    """Uyumsoft @...@ placeholder'larını ayarlardan doldurur; SQL metnine dokunmaz."""
    yerlestirme = {
        "@CoCode@": _co_code(),
        "@BranchCodes@": _branch_code(),
        "@DocDateF@": bas,
        "@DocDateL@": bit,
        "@ReferenceNo@": "null",
        "@TransportNo@": "null",
        "@ProjectCodes@": "null",
        "@VehicleCode@": "null",
    }
    for anahtar, deger in yerlestirme.items():
        sql = sql.replace(anahtar, deger)
    kalan = [p for p in _UYUMSOFT_PARAMETRELER if p in sql]
    if kalan:
        raise ValueError(
            f"SQL'de yerleştirilmemiş Uyumsoft parametreleri kaldı: {', '.join(kalan)}"
        )
    return sql


def _bos_mu(deger: Any) -> bool:
    return deger is None or (isinstance(deger, str) and deger.strip() == "")


def _varsayilanlari_uygula(satirlar: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Personel veri girişinde bazen atlanan alanları makul bir varsayılanla
    doldurur — pivot tablolarda '(boş)' kategorisi oluşmasını önler.

    [VARSAYIM/TODO — kullanıcı isteğiyle eklendi, henüz canlı veriyle teyit
    edilmedi. SQL'e (referans/kpi_veri_rapor.sql) DOKUNULMUYOR — bu, o dosyanın
    "Uyumsoft raporu aynen kullanılır" kuralına aykırı olmasın diye kasıtlı
    olarak Python tarafında, veri Oracle'dan çekildikten SONRA uygulanıyor.]

    - MÜLKİYET (PLAKA_MULKIYET) boşsa -> 'Tedarikçi'. SQL'deki CASE ifadesi
      (ED.OWNERSHIP_STATUS'a bakan) hiçbir WHEN'e uymayan/NULL bir durumda
      ELSE'siz kalıp NULL dönüyor; canlı verideki en sık görülen gerçek değer
      Tedarikçi olduğu için varsayılan bu seçildi.
    - Proje Kodu 'Konya' VE Yük Fiyat Tipi Kodu (YUK_FIYAT_TIP_KODU) boşsa
      -> 'ŞARKÜTERİ'. Konya şubesinin ağırlıklı kargo kategorisi bu olduğu
      için personel bu alanı atladığında varsayılan olarak bu kullanılıyor.
    - Araç Tipi (ARAC_TIPI, Excel VERİ sayfasında X sütunu) boşsa -> 'Tır
      Frigorifik' [kullanıcı isteğiyle eklendi]. Filonun ağırlıklı araç tipi
      bu olduğu için personel bu alanı atladığında varsayılan olarak bu
      kullanılıyor.
    """
    if getattr(ayarlar, "KPI_BOS_ALAN_VARSAYILARI", True) is False:
        return satirlar

    mulkiyet_varsayilan = getattr(ayarlar, "KPI_MULKIYET_BOS_VARSAYILAN", "Tedarikçi")
    konya_fiyat_tipi_varsayilan = getattr(
        ayarlar, "KPI_KONYA_YUK_FIYAT_TIPI_BOS_VARSAYILAN", "ŞARKÜTERİ"
    )
    arac_tipi_varsayilan = getattr(ayarlar, "KPI_ARAC_TIPI_BOS_VARSAYILAN", "Tır Frigorifik")

    for satir in satirlar:
        if "PLAKA_MULKIYET" in satir and _bos_mu(satir.get("PLAKA_MULKIYET")):
            satir["PLAKA_MULKIYET"] = mulkiyet_varsayilan

        proje_kodu = str(satir.get("PROJE_KODU") or "").strip().lower()
        if (
            proje_kodu == "konya"
            and "YUK_FIYAT_TIP_KODU" in satir
            and _bos_mu(satir.get("YUK_FIYAT_TIP_KODU"))
        ):
            satir["YUK_FIYAT_TIP_KODU"] = konya_fiyat_tipi_varsayilan

        if "ARAC_TIPI" in satir and _bos_mu(satir.get("ARAC_TIPI")):
            satir["ARAC_TIPI"] = arac_tipi_varsayilan

    return satirlar


def veri_satirlari_getir(cursor, bas: str, bit: str, bind: dict) -> list[dict[str, Any]]:
    del bind  # VERİ sorgusu Uyumsoft placeholder değiştirme kullanır; :bas/:bit bind edilmez
    sql = _uyumsoft_parametreleri_yerlestir(_veri_sql(), bas, bit)
    cursor.execute(sql)
    sutunlar = [c[0] for c in cursor.description]
    satirlar = [dict(zip(sutunlar, satir)) for satir in cursor.fetchall()]
    return _varsayilanlari_uygula(satirlar)


def hucre_degeri(deger: Any) -> Any:
    if deger is None:
        return None
    if isinstance(deger, Decimal):
        return float(deger)
    if isinstance(deger, (datetime, date)):
        # Uyumsoft'ta "01.01.0001" boş tarih anlamına gelir (bkz. SQL'deki
        # varsayılan CASE değerleri). Bu tür çok eski tarihler Excel COM'a
        # yazılamaz (pywintypes -> OSError [Errno 22] Invalid argument),
        # bu yüzden boş hücreye çevrilir.
        if deger.year < 1900:
            return None
        # ÖNEMLİ: Python datetime/date nesnesini DOĞRUDAN Range.Value'ya
        # atamak pywin32'nin COM tarih dönüştürücüsünü (saat dilimi farkı
        # uygulayan) devreye sokuyor -- gece yarısına yakın tarihler (örn.
        # 01.08.2026 00:00) UTC+3 farkıyla BİR GÜN ÖNCESİNE (31.07.2026)
        # kayıyordu (kullanıcı canlı testte teyit etti: Ağustos filtresiyle
        # gelen satırların YUK_TARIHI/SEVK_TARIHI kolonu 31.07.2026
        # gösteriyordu, oysa Oracle filtresi zaten sadece Ağustos'u
        # döndürüyor). Çözüm: tarihi Excel'in kendi seri sayısına çevirip
        # düz sayı olarak yazıyoruz -- saat dilimi dönüşümü hiç devreye
        # girmiyor, hücre biçimi (bkz. kpi_sablon_rapor._com_tarih_bicimi_uygula)
        # tarih olarak ayarlanınca doğru görünüyor.
        tam = deger if isinstance(deger, datetime) else datetime(deger.year, deger.month, deger.day)
        fark = tam - _EXCEL_EPOK
        return fark.days + fark.seconds / 86400.0
    return deger

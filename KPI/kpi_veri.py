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


def veri_satirlari_getir(cursor, bas: str, bit: str, bind: dict) -> list[dict[str, Any]]:
    del bind  # VERİ sorgusu Uyumsoft placeholder değiştirme kullanır; :bas/:bit bind edilmez
    sql = _uyumsoft_parametreleri_yerlestir(_veri_sql(), bas, bit)
    cursor.execute(sql)
    sutunlar = [c[0] for c in cursor.description]
    return [dict(zip(sutunlar, satir)) for satir in cursor.fetchall()]


def hucre_degeri(deger: Any) -> Any:
    if deger is None:
        return None
    if isinstance(deger, Decimal):
        return float(deger)
    if isinstance(deger, (datetime, date)):
        return deger
    return deger

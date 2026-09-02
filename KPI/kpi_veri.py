"""
VERİ sayfası — Oracle sorgusu.

Yalnızca referans/kpi_veri_rapor.sql dosyasındaki SQL çalıştırılır.
Sorguya kod tarafında alan eklenmez/çıkarılmaz; Uyumsoft raporu aynen kullanılır.

Dosya: KPI/referans/kpi_veri_rapor.sql
Tarih parametreleri: :bas ve :bit (DD.MM.YYYY) — ayarlar.py dönem tarihleri
"""
from __future__ import annotations

from datetime import date, datetime
from decimal import Decimal
from pathlib import Path
from typing import Any

import ayarlar

_KPI_KOKU = Path(__file__).resolve().parent


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


def veri_semasi_hazir() -> tuple[bool, str]:
    yol = veri_sql_dosya_yolu()
    if yol is None:
        return False, (
            "referans/kpi_veri_rapor.sql bulunamadı.\n"
            "Uyumsoft VERİ (LojistikYükSevkKalemRaporu) SQL'inizi bu dosyaya "
            "olduğu gibi kaydedin. Kod SQL'e dokunmaz."
        )
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


def veri_satirlari_getir(cursor, bas: str, bit: str, bind: dict) -> list[dict[str, Any]]:
    cursor.execute(_veri_sql(), bind)
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

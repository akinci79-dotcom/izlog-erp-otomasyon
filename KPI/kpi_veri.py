"""
Yük bazında detay — Uyumsoft VERİ sayfası kaynağı.

Şablondaki VERİ sayfasının 1. satır başlıkları ile Oracle kolon adları
eşleştirilir (büyük/küçük harf ve boşluk duyarsız).
"""
from __future__ import annotations

from datetime import date, datetime
from decimal import Decimal
from typing import Any

import ayarlar
from oracle_baglanti import tablo_var_mi


def _yuk_filtre_parcasi() -> tuple[str, str]:
    joins: list[str] = ["LEFT JOIN GNLD_BRANCH BR ON BR.BRANCH_ID = YK.BRANCH_ID"]
    wheres: list[str] = []
    if getattr(ayarlar, "CO_CODE", None):
        joins.append("JOIN GNLD_COMPANY CO ON CO.CO_ID = YK.CO_ID")
        wheres.append("CO.CO_CODE = :co_code")
    if getattr(ayarlar, "BRANCH_CODE", None):
        wheres.append("BR.BRANCH_CODE = :branch_code")
    if getattr(ayarlar, "KPI_KAPI_KAPI_HARIC", True):
        wheres.append("NVL(YK.IS_DOOR_TO_DOOR, 0) = 0")
    join_sql = "\n        ".join(joins)
    where_sql = (" AND " + " AND ".join(wheres)) if wheres else ""
    return join_sql, where_sql


def veri_semasi_hazir() -> tuple[bool, str]:
    if not tablo_var_mi("LMST_L_GOODS"):
        return False, "LMST_L_GOODS tablosu bulunamadı."
    if not tablo_var_mi("LMST_L_GOODS_OP_DET"):
        return False, "LMST_L_GOODS_OP_DET tablosu bulunamadı."
    return True, ""


def _veri_sql() -> str:
    yk_join, yk_where = _yuk_filtre_parcasi()
    return f"""
WITH YUK_SAYISI AS (
    SELECT TGD.TRANSPORT_ID, COUNT(*) AS YUK_ADEDI
    FROM LMST_L_TRANS_GOODS_DETAIL TGD
    WHERE TGD.GOODS_ID > 0
    GROUP BY TGD.TRANSPORT_ID
),
YFT AS (
    SELECT GOD.GOODS_ID,
           SUM(CASE WHEN GOD.PURCHASE_SALES_TYPE IN (2, 4) THEN GOD.AMT ELSE 0 END) SATIS,
           SUM(CASE WHEN GOD.PURCHASE_SALES_TYPE IN (1, 3) THEN GOD.AMT ELSE 0 END) SATIS_IADE
    FROM LMST_L_GOODS_OP_DET GOD
    GROUP BY GOD.GOODS_ID
),
SFT AS (
    SELECT TGD.GOODS_ID,
           SUM(CASE WHEN TOD.PURCHASE_SALES_TYPE IN (1, 3) THEN TOD.AMT ELSE 0 END)
               / NVL(YS.YUK_ADEDI, 1) ALIS,
           SUM(CASE WHEN TOD.PURCHASE_SALES_TYPE IN (2, 4) THEN TOD.AMT ELSE 0 END)
               / NVL(YS.YUK_ADEDI, 1) ALIS_IADE
    FROM LMST_L_TRANS_GOODS_DETAIL TGD
    JOIN LMST_L_TRANSPORT SK ON SK.TRANSPORT_ID = TGD.TRANSPORT_ID
    JOIN LMST_L_TRANS_OP_DETAIL TOD ON TOD.TRANSPORT_ID = SK.TRANSPORT_ID
    LEFT JOIN YUK_SAYISI YS ON YS.TRANSPORT_ID = SK.TRANSPORT_ID
    WHERE TGD.GOODS_ID > 0
    GROUP BY TGD.GOODS_ID, NVL(YS.YUK_ADEDI, 1)
),
SEVK_OZET AS (
    SELECT TGD.GOODS_ID,
           MIN(SK.TRANSPORT_NO) KEEP (DENSE_RANK FIRST ORDER BY SK.DOC_DATE) SEVK_NO,
           MIN(SK.DOC_DATE) KEEP (DENSE_RANK FIRST ORDER BY SK.DOC_DATE) SEVK_TARIHI,
           MIN(VH.LICENSE_PLATE) KEEP (DENSE_RANK FIRST ORDER BY SK.DOC_DATE) PLAKA,
           MIN(TTYPE.DESCRIPTION) KEEP (DENSE_RANK FIRST ORDER BY SK.DOC_DATE) ARAC_TIPI
    FROM LMST_L_TRANS_GOODS_DETAIL TGD
    JOIN LMST_L_TRANSPORT SK ON SK.TRANSPORT_ID = TGD.TRANSPORT_ID
    LEFT JOIN LMSW_VIEW_TRANSPORT_UNITS TU
        ON SK.TRACTOR_UNIT_ID = TU.TRANSPORT_UNIT_ID AND SK.TRACTOR_MAPID = TU.MAPID
    LEFT JOIN FLMD_VEHICLE VH ON TU.TRANSPORT_TYPE = 1 AND TU.UNIT_ID = VH.VEHICLE_ID
    LEFT JOIN FLMD_L_TRAILER_TYPE TTYPE ON TTYPE.TRAILER_TYPE_ID = VH.TRAILER_TYPE_ID
    WHERE TGD.GOODS_ID > 0
    GROUP BY TGD.GOODS_ID
),
CARI_OZET AS (
    SELECT GOD.GOODS_ID,
           MIN(FE.ENTITY_CODE) MUSTERI_KODU,
           MIN(FE.ENTITY_NAME) MUSTERI_ADI
    FROM LMST_L_GOODS_OP_DET GOD
    LEFT JOIN FIND_ENTITY FE ON FE.ENTITY_ID = GOD.CHARGED_L_ENTITY_ID
    WHERE GOD.PURCHASE_SALES_TYPE IN (2, 4)
    GROUP BY GOD.GOODS_ID
)
SELECT
    YK.REFERENCE_NO AS YUK_NO,
    YK.DOC_DATE AS YUK_TARIHI,
    SO.SEVK_NO,
    SO.SEVK_TARIHI,
    PJ.PROJECT_CODE AS PROJE_KODU,
    SO.PLAKA,
    SO.ARAC_TIPI,
    CARI.MUSTERI_KODU,
    CARI.MUSTERI_ADI,
    BR.BRANCH_CODE AS SUBE_KODU,
    NVL(BR.BRANCH_DESC, BR.BRANCH_CODE) AS SUBE,
    GPT.GOODS_PRICE_TYPE_CODE AS YUK_FIYAT_TIP_KODU,
    NVL(YFT.SATIS, 0) - NVL(YFT.SATIS_IADE, 0) AS SATIS_TUTAR,
    NVL(SFT.ALIS, 0) - NVL(SFT.ALIS_IADE, 0) AS ALIS_TUTAR,
    (NVL(YFT.SATIS, 0) - NVL(YFT.SATIS_IADE, 0))
        - (NVL(SFT.ALIS, 0) - NVL(SFT.ALIS_IADE, 0)) AS KAR_ZARAR,
    CASE
        WHEN NVL(YFT.SATIS, 0) - NVL(YFT.SATIS_IADE, 0) = 0 THEN 0
        ELSE ROUND(
            ((NVL(YFT.SATIS, 0) - NVL(YFT.SATIS_IADE, 0))
             - (NVL(SFT.ALIS, 0) - NVL(SFT.ALIS_IADE, 0)))
            * 100
            / (NVL(YFT.SATIS, 0) - NVL(YFT.SATIS_IADE, 0)),
            2
        )
    END AS MARJ_YUZDE,
    YK.GOODS_ID
FROM LMST_L_GOODS YK
LEFT JOIN LMSD_L_AGR_PROJ_TYPE PJ ON PJ.PROJECT_ID = YK.PROJECT_ID
LEFT JOIN LMSD_L_GOODSPRICE_TYPE GPT ON GPT.GOODS_PRICE_TYPE_ID = YK.GOODS_PRICE_TYPE_ID
LEFT JOIN CARI_OZET CARI ON CARI.GOODS_ID = YK.GOODS_ID
LEFT JOIN YFT ON YFT.GOODS_ID = YK.GOODS_ID
LEFT JOIN SFT ON SFT.GOODS_ID = YK.GOODS_ID
LEFT JOIN SEVK_OZET SO ON SO.GOODS_ID = YK.GOODS_ID
{yk_join}
WHERE YK.DOC_DATE BETWEEN TO_DATE(:bas, 'DD.MM.YYYY') AND TO_DATE(:bit, 'DD.MM.YYYY')
{yk_where}
ORDER BY YK.DOC_DATE, YK.REFERENCE_NO
"""


def _veri_satir_zenginlestir(satir: dict[str, Any]) -> dict[str, Any]:
    """Pivot şablonundaki alternatif kolon adları için takma alanlar."""
    satir["TOPLAM_SATIS"] = satir.get("SATIS_TUTAR")
    satir["TOPLAM_ALIS"] = satir.get("ALIS_TUTAR")
    satir["NET_KAR_ZARAR"] = satir.get("KAR_ZARAR")
    satir["MARJ_ORANI"] = satir.get("MARJ_YUZDE")
    return satir


def veri_satirlari_getir(cursor, bas: str, bit: str, bind: dict) -> list[dict[str, Any]]:
    ok, _ = veri_semasi_hazir()
    if not ok:
        return []
    cursor.execute(_veri_sql(), bind)
    sutunlar = [c[0] for c in cursor.description]
    return [_veri_satir_zenginlestir(dict(zip(sutunlar, satir))) for satir in cursor.fetchall()]


def hucre_degeri(deger: Any) -> Any:
    if deger is None:
        return None
    if isinstance(deger, Decimal):
        return float(deger)
    if isinstance(deger, (datetime, date)):
        return deger
    return deger

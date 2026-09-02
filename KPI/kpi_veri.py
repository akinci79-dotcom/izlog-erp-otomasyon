"""
Yük bazında detay — Uyumsoft VERİ sayfası kaynağı.

Şablondaki VERİ sayfasının 1. satır başlıkları ile Oracle kolon adları
eşleştirilir (büyük/küçük harf ve boşluk duyarsız).

Tam rapor için Uyumsoft'taki LojistikYükSevkKalemRaporu SQL'ini
referans/kpi_veri_rapor.sql dosyasına yapıştırın (ayarlar.py → KPI_VERI_SQL_DOSYASI).
"""
from __future__ import annotations

from datetime import date, datetime
from decimal import Decimal
from pathlib import Path
from typing import Any

import ayarlar
from oracle_baglanti import tablo_kolonlari, tablo_var_mi

_KPI_KOKU = Path(__file__).resolve().parent


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


def _veri_sql_dosya_yolu() -> Path | None:
    """Özel VERİ SQL dosyası — ayarlar veya referans/kpi_veri_rapor.sql."""
    ad = getattr(ayarlar, "KPI_VERI_SQL_DOSYASI", None)
    adaylar: list[Path] = []
    if ad:
        yol = Path(str(ad))
        if yol.is_absolute():
            adaylar.append(yol)
        else:
            adaylar.append(_KPI_KOKU / yol)
            adaylar.append(_KPI_KOKU / "referans" / yol)
    adaylar.append(_KPI_KOKU / "referans" / "kpi_veri_rapor.sql")
    for yol in adaylar:
        if yol.is_file() and yol.stat().st_size > 50:
            return yol
    return None


def _veri_sql_kaynak() -> str:
    dosya = _veri_sql_dosya_yolu()
    if dosya is not None:
        return dosya.read_text(encoding="utf-8-sig")
    if getattr(ayarlar, "KPI_VERI_GENISLETILMIS_SQL", False):
        return _veri_sql_varsayilan()
    return _veri_sql_basit()


def _kolon_var(tablo: str, kolon: str) -> bool:
    try:
        return kolon.upper() in {k.upper() for k in tablo_kolonlari_from_cache(tablo)}
    except Exception:
        return False


_tablo_kolon_onbellek: dict[str, list[str]] = {}


def tablo_kolonlari_from_cache(tablo: str) -> list[str]:
    if tablo not in _tablo_kolon_onbellek:
        with __import__("oracle_baglanti", fromlist=["baglanti_yonet"]).baglanti_yonet() as bag:
            _tablo_kolon_onbellek[tablo] = tablo_kolonlari(bag.cursor(), tablo)
    return _tablo_kolon_onbellek[tablo]


def _nokta_cte_parcasi() -> tuple[str, str, str, str]:
    """
    Yükleme/boşaltma noktaları — tablo adları kuruma göre değişebilir.
    LMST_L_GOODS_POINT + LMSD_L_POINT varsa CTE üretir.
    """
    if not tablo_var_mi("LMST_L_GOODS_POINT"):
        return "", "", "", ""
    point_tablo = "LMSD_L_POINT" if tablo_var_mi("LMSD_L_POINT") else ""
    if not point_tablo:
        return "", "", "", ""

    gp_kolonlar = {k.upper() for k in tablo_kolonlari_from_cache("LMST_L_GOODS_POINT")}
    tip_kolon = "POINT_TYPE" if "POINT_TYPE" in gp_kolonlar else (
        "OPERATION_TYPE" if "OPERATION_TYPE" in gp_kolonlar else None
    )
    sira_kolon = "LINE_NO" if "LINE_NO" in gp_kolonlar else (
        "ROW_NO" if "ROW_NO" in gp_kolonlar else "GOODS_POINT_ID"
    )
    yukleme_filtre = f"AND GP.{tip_kolon} IN (1, 10)" if tip_kolon else ""
    bosaltma_filtre = f"AND GP.{tip_kolon} IN (2, 20)" if tip_kolon else ""

    cte = f""",
NOKTA_YUKLEME AS (
    SELECT GP.GOODS_ID,
           MIN(P.POINT_CODE) KEEP (DENSE_RANK FIRST ORDER BY GP.{sira_kolon}) YUKLEME_YER_KODU,
           MIN(P.POINT_NAME) KEEP (DENSE_RANK FIRST ORDER BY GP.{sira_kolon}) YUKLEME_YER_ADI,
           MIN(CITY.CITY_NAME) KEEP (DENSE_RANK FIRST ORDER BY GP.{sira_kolon}) YUKLEME_SEHIR_ADI,
           MIN(TOWN.TOWN_NAME) KEEP (DENSE_RANK FIRST ORDER BY GP.{sira_kolon}) YUKLEME_ILCE_ADI,
           MIN(P.ADDRESS) KEEP (DENSE_RANK FIRST ORDER BY GP.{sira_kolon}) YUKLEME_ADRESI
    FROM LMST_L_GOODS_POINT GP
    LEFT JOIN {point_tablo} P ON P.POINT_ID = GP.POINT_ID
    LEFT JOIN GNLD_CITY CITY ON CITY.CITY_ID = P.CITY_ID
    LEFT JOIN GNLD_TOWN TOWN ON TOWN.TOWN_ID = P.TOWN_ID
    WHERE GP.GOODS_ID > 0 {yukleme_filtre}
    GROUP BY GP.GOODS_ID
),
NOKTA_BOSALTMA AS (
    SELECT GP.GOODS_ID,
           MIN(P.POINT_CODE) KEEP (DENSE_RANK FIRST ORDER BY GP.{sira_kolon}) BOSALTMA_YER_KODU,
           MIN(P.POINT_NAME) KEEP (DENSE_RANK FIRST ORDER BY GP.{sira_kolon}) BOSALTMA_YER_ADI,
           MIN(CITY.CITY_NAME) KEEP (DENSE_RANK FIRST ORDER BY GP.{sira_kolon}) BOSALTMA_SEHIR_ADI,
           MIN(TOWN.TOWN_NAME) KEEP (DENSE_RANK FIRST ORDER BY GP.{sira_kolon}) BOSALTMA_ILCE_ADI,
           MIN(P.ADDRESS) KEEP (DENSE_RANK FIRST ORDER BY GP.{sira_kolon}) BOSALTMA_ADRESI,
           COUNT(*) NOKTA_SAYISI
    FROM LMST_L_GOODS_POINT GP
    LEFT JOIN {point_tablo} P ON P.POINT_ID = GP.POINT_ID
    LEFT JOIN GNLD_CITY CITY ON CITY.CITY_ID = P.CITY_ID
    LEFT JOIN GNLD_TOWN TOWN ON TOWN.TOWN_ID = P.TOWN_ID
    WHERE GP.GOODS_ID > 0 {bosaltma_filtre}
    GROUP BY GP.GOODS_ID
)"""
    select = """
    NY.YUKLEME_YER_KODU,
    NY.YUKLEME_YER_ADI,
    NY.YUKLEME_SEHIR_ADI,
    NY.YUKLEME_ILCE_ADI,
    NY.YUKLEME_ADRESI,
    NB.BOSALTMA_YER_KODU,
    NB.BOSALTMA_YER_ADI,
    NB.BOSALTMA_SEHIR_ADI,
    NB.BOSALTMA_ILCE_ADI,
    NB.BOSALTMA_ADRESI,
    NB.NOKTA_SAYISI,"""
    join = """
LEFT JOIN NOKTA_YUKLEME NY ON NY.GOODS_ID = YK.GOODS_ID
LEFT JOIN NOKTA_BOSALTMA NB ON NB.GOODS_ID = YK.GOODS_ID"""
    return cte, select, join, ""


def _veri_sql_varsayilan() -> str:
    """Genişletilmiş yük-bazında sorgu — tam Uyumsoft raporu değil, kalem_detay ile uyumlu."""
    yk_join, yk_where = _yuk_filtre_parcasi()
    nokta_cte, nokta_select, nokta_join, _ = _nokta_cte_parcasi()

    agirlik = "YK.GROSS_WEIGHT" if _kolon_var("LMST_L_GOODS", "GROSS_WEIGHT") else (
        "YK.TOTAL_GROSS_WEIGHT" if _kolon_var("LMST_L_GOODS", "TOTAL_GROSS_WEIGHT") else "NULL"
    )
    arsiv = (
        "CASE NVL(YK.IS_ARCHIVED, 0) WHEN 1 THEN 'Arşiv kaydı var' END"
        if _kolon_var("LMST_L_GOODS", "IS_ARCHIVED")
        else "NULL"
    )
    mal_tipi_select = (
        "LGT.GOODS_TYPE_CODE AS MAL_TIPI,"
        if _kolon_var("LMST_L_GOODS", "GOODS_TYPE_ID") and tablo_var_mi("LMSD_L_GOODS_TYPE")
        else "GPT.GOODS_PRICE_TYPE_CODE AS MAL_TIPI,"
    )
    mal_tipi_join = (
        "LEFT JOIN LMSD_L_GOODS_TYPE LGT ON LGT.GOODS_TYPE_ID = YK.GOODS_TYPE_ID"
        if _kolon_var("LMST_L_GOODS", "GOODS_TYPE_ID") and tablo_var_mi("LMSD_L_GOODS_TYPE")
        else ""
    )
    surucu_select = """MIN(CAST(NULL AS VARCHAR2(200))) AS SURUCU_ADI,
           MIN(CAST(NULL AS VARCHAR2(50))) AS SURUCU_TEL"""
    surucu_join = ""
    if tablo_var_mi("HRMD_ADVANCE"):
        surucu_select = """MIN(SUR.ADVANCE_NAME || ' ' || SUR.ADVANCE_SURNAME)
               KEEP (DENSE_RANK FIRST ORDER BY SK.DOC_DATE) SURUCU_ADI,
           MIN(SUR.MOBILE_TEL) KEEP (DENSE_RANK FIRST ORDER BY SK.DOC_DATE) SURUCU_TEL"""
        surucu_join = "LEFT JOIN HRMD_ADVANCE SUR ON SUR.ADVANCE_ID = SK.DRIVER_ID"
    else:
        surucu_join = ""
    waybill_expr = (
        "TGD.TRAN_WAYBILL_DATE"
        if _kolon_var("LMST_L_TRANS_GOODS_DETAIL", "TRAN_WAYBILL_DATE")
        else "TGD.TRAN_WAYBILL_DOC_DATE"
        if _kolon_var("LMST_L_TRANS_GOODS_DETAIL", "TRAN_WAYBILL_DOC_DATE")
        else "CAST(NULL AS DATE)"
    )
    waybill_date = (
        f"MIN({waybill_expr}) KEEP (DENSE_RANK FIRST ORDER BY SK.DOC_DATE) SOZLESME_TARIHI,"
    )
    gonderici_join = ""
    gonderici_select = "NULL AS GONDERICI_CARI_ADI,"
    alici_select = "NULL AS ALICI_CARI_ADI,"
    if _kolon_var("LMST_L_GOODS", "SENDER_L_ENTITY_ID"):
        gonderici_join = "LEFT JOIN FIND_ENTITY GON ON GON.ENTITY_ID = YK.SENDER_L_ENTITY_ID"
        gonderici_select = "GON.ENTITY_NAME AS GONDERICI_CARI_ADI,"
    if _kolon_var("LMST_L_GOODS", "RECEIVER_L_ENTITY_ID"):
        gonderici_join += "\nLEFT JOIN FIND_ENTITY ALC ON ALC.ENTITY_ID = YK.RECEIVER_L_ENTITY_ID"
        alici_select = "ALC.ENTITY_NAME AS ALICI_CARI_ADI,"

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
           MIN(TTYPE.DESCRIPTION) KEEP (DENSE_RANK FIRST ORDER BY SK.DOC_DATE) ARAC_TIPI,
           MIN(CASE ED.OWNERSHIP_STATUS
                 WHEN 1 THEN 'Kurum Malı'
                 WHEN 2 THEN 'Kiralık'
                 WHEN 3 THEN 'Tedarikçi'
               END) KEEP (DENSE_RANK FIRST ORDER BY SK.DOC_DATE) MULKIYET,
           MIN(PLK_CARI.ENTITY_NAME) KEEP (DENSE_RANK FIRST ORDER BY SK.DOC_DATE) PLAKA_CARI_ADI,
           MIN(TGD.TRAN_WAYBILL_DOC_NO) KEEP (DENSE_RANK FIRST ORDER BY SK.DOC_DATE) SOZLESME_NO,
           {waybill_date}
           MIN(SK.START_KM) KEEP (DENSE_RANK FIRST ORDER BY SK.DOC_DATE) CIKIS_KM,
           MIN(SK.END_KM) KEEP (DENSE_RANK FIRST ORDER BY SK.DOC_DATE) VARIS_KM,
           MIN(NVL(SK.EMPTY_KM, NVL(SK.END_KM, 0) - NVL(SK.START_KM, 0)))
               KEEP (DENSE_RANK FIRST ORDER BY SK.DOC_DATE) BOS_KM,
           {surucu_select}
    FROM LMST_L_TRANS_GOODS_DETAIL TGD
    JOIN LMST_L_TRANSPORT SK ON SK.TRANSPORT_ID = TGD.TRANSPORT_ID
    LEFT JOIN LMSW_VIEW_TRANSPORT_UNITS TU
        ON SK.TRACTOR_UNIT_ID = TU.TRANSPORT_UNIT_ID AND SK.TRACTOR_MAPID = TU.MAPID
    LEFT JOIN FLMD_VEHICLE VH ON TU.TRANSPORT_TYPE = 1 AND TU.UNIT_ID = VH.VEHICLE_ID
    LEFT JOIN FLMD_L_TRAILER_TYPE TTYPE ON TTYPE.TRAILER_TYPE_ID = VH.TRAILER_TYPE_ID
    LEFT JOIN FLMD_VHC_ENTITY_DETAIL ED
        ON ED.VEHICLE_ID = VH.VEHICLE_ID
       AND ED.START_DATE <= SK.DOC_DATE AND ED.END_DATE >= SK.DOC_DATE
    LEFT JOIN FIND_ENTITY PLK_CARI ON PLK_CARI.ENTITY_ID = ED.L_ENTITY_ID
    {surucu_join}
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
{nokta_cte}
SELECT
    PJ.PROJECT_CODE AS PROJE_KODU,
    US.US_NAME || ' ' || US.US_SURNAME AS KULLANICI,
    {arsiv} AS ARSIV_DURUMU,
    NULL AS MUSTERI_HESAPLASMA_ACIKLAMA,
    NULL AS SEVK_DURUMU,
    SO.SEVK_TARIHI,
    SO.SEVK_NO,
    YK.DOC_DATE AS YUK_TARIHI,
    YK.REFERENCE_NO AS YUK_NO,
    GPT.GOODS_PRICE_TYPE_CODE AS YUK_FIYAT_TIP_KODU,
    {mal_tipi_select}
    NULL AS SEFER_TURU,
    NULL AS SIPARIS_TARIHI,
    NULL AS SIPARIS_NO,
    NULL AS SIPARIS_NOTLARI,
    CARI.MUSTERI_ADI,
    {gonderici_select}
    {alici_select}
    SO.SOZLESME_TARIHI,
    SO.SOZLESME_NO,
    NULL AS MUSTERI_EVRAK_NO,
    SO.PLAKA,
    SO.ARAC_TIPI,
    SO.PLAKA_CARI_ADI,
    SO.MULKIYET,
    SO.SURUCU_ADI,
    SO.SURUCU_TEL,
    NULL AS SEFER_TIPI,
    {nokta_select if nokta_select else "NULL AS NOKTA_SAYISI,"}
    {agirlik} AS BRUT_AGIRLIK,
    SO.CIKIS_KM,
    SO.VARIS_KM,
    SO.BOS_KM,
    NVL(SFT.ALIS, 0) - NVL(SFT.ALIS_IADE, 0) AS ALIS_TUTAR,
    NVL(SFT.ALIS_IADE, 0) AS ALIS_IADE_TUTAR,
    NVL(YFT.SATIS, 0) - NVL(YFT.SATIS_IADE, 0) AS SATIS_TUTAR,
    NVL(YFT.SATIS_IADE, 0) AS SATIS_IADE_TUTAR,
    NVL(SFT.ALIS, 0) - NVL(SFT.ALIS_IADE, 0) AS TOPLAM_ALIS,
    NVL(YFT.SATIS, 0) - NVL(YFT.SATIS_IADE, 0) AS TOPLAM_SATIS,
    (NVL(YFT.SATIS, 0) - NVL(YFT.SATIS_IADE, 0))
        - (NVL(SFT.ALIS, 0) - NVL(SFT.ALIS_IADE, 0)) AS KAR_ZARAR,
    CARI.MUSTERI_KODU,
    BR.BRANCH_CODE AS SUBE_KODU,
    NVL(BR.BRANCH_DESC, BR.BRANCH_CODE) AS SUBE,
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
{mal_tipi_join}
LEFT JOIN USERS US ON US.US_ID = YK.CREATE_USER_ID
LEFT JOIN CARI_OZET CARI ON CARI.GOODS_ID = YK.GOODS_ID
LEFT JOIN YFT ON YFT.GOODS_ID = YK.GOODS_ID
LEFT JOIN SFT ON SFT.GOODS_ID = YK.GOODS_ID
LEFT JOIN SEVK_OZET SO ON SO.GOODS_ID = YK.GOODS_ID
{gonderici_join}
{nokta_join}
{yk_join}
WHERE YK.DOC_DATE BETWEEN TO_DATE(:bas, 'DD.MM.YYYY') AND TO_DATE(:bit, 'DD.MM.YYYY')
{yk_where}
ORDER BY YK.DOC_DATE, YK.REFERENCE_NO
"""


def _veri_sql_basit() -> str:
    """Önceki çalışan minimal sorgu — genişletilmiş SQL hata verirse yedek."""
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
    PJ.PROJECT_CODE AS PROJE_KODU,
    YK.REFERENCE_NO AS YUK_NO,
    YK.DOC_DATE AS YUK_TARIHI,
    SO.SEVK_NO,
    SO.SEVK_TARIHI,
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


def _veri_sql() -> str:
    return _veri_sql_kaynak()


def veri_sql_kaynak_bilgisi() -> str:
    dosya = _veri_sql_dosya_yolu()
    if dosya:
        return f"özel SQL: {dosya.name}"
    if getattr(ayarlar, "KPI_VERI_GENISLETILMIS_SQL", False):
        return "genişletilmiş varsayılan SQL"
    return "basit varsayılan SQL (tutarlar + temel alanlar)"


def _veri_satir_zenginlestir(satir: dict[str, Any]) -> dict[str, Any]:
    """Pivot şablonundaki alternatif kolon adları için takma alanlar."""
    satis = satir.get("SATIS_TUTAR")
    alis = satir.get("ALIS_TUTAR")
    kar = satir.get("KAR_ZARAR")
    satir["TOPLAM_SATIS"] = satir.get("TOPLAM_SATIS", satis)
    satir["TOPLAM_ALIS"] = satir.get("TOPLAM_ALIS", alis)
    satir["NET_KAR_ZARAR"] = kar
    satir["TOPLAM_KAR_ZARAR"] = kar
    satir["MARJ_ORANI"] = satir.get("MARJ_YUZDE")
    satir["SATIS_TUTARI"] = satis
    satir["ALIS_TUTARI"] = alis
    satir["ALIS_IADE_TUTARI"] = satir.get("ALIS_IADE_TUTAR")
    satir["SATIS_IADE_TUTARI"] = satir.get("SATIS_IADE_TUTAR")
    return satir


def veri_satirlari_getir(cursor, bas: str, bit: str, bind: dict) -> list[dict[str, Any]]:
    sql = _veri_sql()
    try:
        cursor.execute(sql, bind)
    except Exception as exc:
        if _veri_sql_dosya_yolu() is not None:
            raise
        mesaj = str(exc)
        if any(k in mesaj for k in ("ORA-00923", "ORA-00936", "ORA-00904", "ORA-00942")):
            print(
                f"  Uyarı: Genişletilmiş VERİ SQL hatası ({mesaj.splitlines()[0]}) — "
                "basit sorguya düşülüyor.",
                flush=True,
            )
            cursor.execute(_veri_sql_basit(), bind)
        else:
            raise
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

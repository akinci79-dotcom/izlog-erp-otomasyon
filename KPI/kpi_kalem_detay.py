"""
Sevk / yük kalem bazında detay — Uyumsoft kalem detay raporu SQL'i ile uyumlu.

Bir sevkte birden fazla yük olduğunda sevk kalemleri TGD üzerinden yük kırılımıyla
tekrarlanır [DOĞRULANMIŞ — kullanıcı rapor SQL'i].
"""
from __future__ import annotations

from collections import defaultdict
from decimal import Decimal
from typing import Any

import ayarlar
from oracle_baglanti import satir_limit_sql, tablo_var_mi


def _decimal(deger) -> Decimal:
    if deger is None:
        return Decimal("0")
    return Decimal(str(deger))


def _sk_filtre_parcasi() -> tuple[str, str]:
    joins: list[str] = []
    wheres: list[str] = []
    if getattr(ayarlar, "CO_CODE", None):
        joins.append("JOIN GNLD_COMPANY CO ON CO.CO_ID = SK.CO_ID")
        wheres.append("CO.CO_CODE = :co_code")
    if getattr(ayarlar, "BRANCH_CODE", None):
        joins.append("JOIN GNLD_BRANCH BR ON BR.BRANCH_ID = SK.BRANCH_ID")
        wheres.append("BR.BRANCH_CODE = :branch_code")
    if getattr(ayarlar, "KPI_KAPI_KAPI_HARIC", True):
        wheres.append("NVL(YK.IS_DOOR_TO_DOOR, 0) = 0")
    join_sql = "\n        ".join(joins)
    where_sql = (" AND " + " AND ".join(wheres)) if wheres else ""
    return join_sql, where_sql


def _yk_filtre_parcasi() -> tuple[str, str]:
    joins: list[str] = []
    wheres: list[str] = []
    if getattr(ayarlar, "CO_CODE", None):
        joins.append("JOIN GNLD_COMPANY CO ON CO.CO_ID = YK.CO_ID")
        wheres.append("CO.CO_CODE = :co_code")
    if getattr(ayarlar, "BRANCH_CODE", None):
        joins.append("JOIN GNLD_BRANCH BR ON BR.BRANCH_ID = YK.BRANCH_ID")
        wheres.append("BR.BRANCH_CODE = :branch_code")
    if getattr(ayarlar, "KPI_KAPI_KAPI_HARIC", True):
        wheres.append("NVL(YK.IS_DOOR_TO_DOOR, 0) = 0")
    join_sql = "\n        ".join(joins)
    where_sql = (" AND " + " AND ".join(wheres)) if wheres else ""
    return join_sql, where_sql


def kalem_detay_semasi_hazir() -> tuple[bool, str]:
    if not tablo_var_mi("LMST_L_TRANS_OP_DETAIL"):
        return False, "LMST_L_TRANS_OP_DETAIL tablosu bulunamadı."
    if not tablo_var_mi("LMST_L_GOODS_OP_DET"):
        return False, "LMST_L_GOODS_OP_DET tablosu bulunamadı."
    if not tablo_var_mi("LMST_L_TRANS_GOODS_DETAIL"):
        return False, "LMST_L_TRANS_GOODS_DETAIL tablosu bulunamadı."
    return True, ""


def _detay_sql(limit: int) -> str:
    sk_join, sk_where = _sk_filtre_parcasi()
    yk_join, yk_where = _yk_filtre_parcasi()
    return satir_limit_sql(
        f"""
    SELECT 'Sevk Kalemi' TIP,
           US.US_NAME || ' ' || US.US_SURNAME KULLANICI,
           PJ.PROJECT_CODE PROJE_KODU,
           SK.TRANSPORT_NO SEVK_NO,
           SK.DOC_DATE SEVK_TARIHI,
           YK.REFERENCE_NO YUK_NO,
           YK.DOC_DATE YUK_TARIHI,
           TTYPE.DESCRIPTION ARAC_TIPI,
           VH.LICENSE_PLATE PLAKA,
           CASE ED.OWNERSHIP_STATUS
             WHEN 1 THEN 'Kurum Malı'
             WHEN 2 THEN 'Kiralık'
             WHEN 3 THEN 'Tedarikçi'
           END PLAKA_MULKIYET,
           CASE TD.PURCHASE_SALES_TYPE
             WHEN 1 THEN 'Alış'
             WHEN 2 THEN 'Satış'
             WHEN 3 THEN 'Satış İade'
             WHEN 4 THEN 'Alış İade'
           END ALIS_SATIS,
           TGD.TRAN_WAYBILL_DOC_NO SOZLESME_NO,
           FM.DOC_NO FATURA_NO,
           FM.DOC_DATE FATURA_TARIHI,
           HK.EXPENSE_CODE OPERASYON_KODU,
           TD.QTY MIKTAR,
           IU.UNIT_CODE BIRIM,
           TD.UNIT_PRICE_TRA BIRIM_FIYAT,
           TD.AMT_TRA TUTAR,
           GC.CUR_CODE PARA_BIRIMI,
           CK.ENTITY_CODE CARI_KODU,
           CK.ENTITY_NAME CARI_ADI,
           NVL(SPM.CASE_CODE, CPM.CASE_CODE) DOSYA_NO,
           NVL(YK.GOODS_ID, 0) YUK_ID,
           NVL(SK.TRANSPORT_ID, 0) SEVK_ID
    FROM LMST_L_TRANS_OP_DETAIL TD
    LEFT JOIN INVD_EXPENSE HK ON HK.EXPENSE_ID = TD.OPERATION_ID
    LEFT JOIN FIND_ENTITY CK ON CK.ENTITY_ID = TD.CHARGED_L_ENTITY_ID
    LEFT JOIN PSMT_INVOICE_M FM ON FM.INVOICE_M_ID = TD.INVOICE_M_ID
    LEFT JOIN LMST_CUST_PAYOFF_M CPM ON CPM.CUST_PAYOFF_ID = TD.SUP_PAYOFF_ID
    LEFT JOIN LMST_SUP_PAYOFF_M SPM ON SPM.SUP_PAYOFF_ID = TD.SUP_PAYOFF_ID
    LEFT JOIN INVD_UNIT IU ON IU.UNIT_ID = TD.UNIT_ID
    LEFT JOIN LMST_L_TRANSPORT SK ON SK.TRANSPORT_ID = TD.TRANSPORT_ID
    LEFT JOIN LMST_L_TRANS_GOODS_DETAIL TGD ON TGD.TRANSPORT_ID = TD.TRANSPORT_ID
    LEFT JOIN LMST_L_GOODS YK ON YK.GOODS_ID = TGD.GOODS_ID
    LEFT JOIN USERS US ON US.US_ID = TD.CREATE_USER_ID
    LEFT JOIN LMSW_VIEW_TRANSPORT_UNITS TU
        ON SK.TRACTOR_UNIT_ID = TU.TRANSPORT_UNIT_ID AND SK.TRACTOR_MAPID = TU.MAPID
    LEFT JOIN FLMD_VEHICLE VH ON TU.TRANSPORT_TYPE = 1 AND TU.UNIT_ID = VH.VEHICLE_ID
    LEFT JOIN FLMD_VHC_ENTITY_DETAIL ED
        ON ED.VEHICLE_ID = VH.VEHICLE_ID
       AND ED.START_DATE <= SK.DOC_DATE AND ED.END_DATE >= SK.DOC_DATE
    LEFT JOIN FLMD_L_TRAILER_TYPE TTYPE ON TTYPE.TRAILER_TYPE_ID = VH.TRAILER_TYPE_ID
    LEFT JOIN GNLD_CURRENCY GC ON GC.CUR_ID = TD.CUR_TRA_ID
    LEFT JOIN LMSD_L_AGR_PROJ_TYPE PJ ON PJ.PROJECT_ID = SK.PROJECT_ID
    {sk_join}
    WHERE SK.DOC_DATE BETWEEN TO_DATE(:bas, 'DD.MM.YYYY') AND TO_DATE(:bit, 'DD.MM.YYYY')
    {sk_where}
    UNION ALL
    SELECT 'Yük Kalemi' TIP,
           US.US_NAME || ' ' || US.US_SURNAME KULLANICI,
           PJ.PROJECT_CODE PROJE_KODU,
           SK.TRANSPORT_NO SEVK_NO,
           SK.DOC_DATE SEVK_TARIHI,
           YK.REFERENCE_NO YUK_NO,
           YK.DOC_DATE YUK_TARIHI,
           TTYPE.DESCRIPTION ARAC_TIPI,
           VH.LICENSE_PLATE PLAKA,
           CASE ED.OWNERSHIP_STATUS
             WHEN 1 THEN 'Kurum Malı'
             WHEN 2 THEN 'Kiralık'
             WHEN 3 THEN 'Tedarikçi'
           END PLAKA_MULKIYET,
           CASE TD.PURCHASE_SALES_TYPE
             WHEN 1 THEN 'Alış'
             WHEN 2 THEN 'Satış'
             WHEN 3 THEN 'Satış İade'
             WHEN 4 THEN 'Alış İade'
           END ALIS_SATIS,
           TGD.TRAN_WAYBILL_DOC_NO SOZLESME_NO,
           FM.DOC_NO FATURA_NO,
           FM.DOC_DATE FATURA_TARIHI,
           HK.EXPENSE_CODE OPERASYON_KODU,
           TD.QTY MIKTAR,
           IU.UNIT_CODE BIRIM,
           TD.UNIT_PRICE BIRIM_FIYAT,
           TD.AMT_TRA TUTAR,
           GC.CUR_CODE PARA_BIRIMI,
           CK.ENTITY_CODE CARI_KODU,
           CK.ENTITY_NAME CARI_ADI,
           CPM.CASE_CODE DOSYA_NO,
           NVL(YK.GOODS_ID, 0) YUK_ID,
           NVL(SK.TRANSPORT_ID, 0) SEVK_ID
    FROM LMST_L_GOODS_OP_DET TD
    LEFT JOIN INVD_EXPENSE HK ON HK.EXPENSE_ID = TD.OPERATION_ID
    LEFT JOIN FIND_ENTITY CK ON CK.ENTITY_ID = TD.CHARGED_L_ENTITY_ID
    LEFT JOIN PSMT_INVOICE_M FM ON FM.INVOICE_M_ID = TD.INVOICE_M_ID
    LEFT JOIN LMST_CUST_PAYOFF_M CPM ON CPM.CUST_PAYOFF_ID = TD.CUST_PAYOFF_ID
    LEFT JOIN INVD_UNIT IU ON IU.UNIT_ID = TD.UNIT_ID
    LEFT JOIN LMST_L_GOODS YK ON YK.GOODS_ID = TD.GOODS_ID
    LEFT JOIN LMST_L_TRANS_GOODS_DETAIL TGD ON TGD.GOODS_ID = TD.GOODS_ID
    LEFT JOIN LMST_L_TRANSPORT SK ON TGD.TRANSPORT_ID = SK.TRANSPORT_ID
    LEFT JOIN USERS US ON US.US_ID = TD.CREATE_USER_ID
    LEFT JOIN LMSW_VIEW_TRANSPORT_UNITS TU
        ON SK.TRACTOR_UNIT_ID = TU.TRANSPORT_UNIT_ID AND SK.TRACTOR_MAPID = TU.MAPID
    LEFT JOIN FLMD_VEHICLE VH ON TU.TRANSPORT_TYPE = 1 AND TU.UNIT_ID = VH.VEHICLE_ID
    LEFT JOIN FLMD_VHC_ENTITY_DETAIL ED
        ON ED.VEHICLE_ID = VH.VEHICLE_ID
       AND ED.START_DATE <= SK.DOC_DATE AND ED.END_DATE >= SK.DOC_DATE
    LEFT JOIN FLMD_L_TRAILER_TYPE TTYPE ON TTYPE.TRAILER_TYPE_ID = VH.TRAILER_TYPE_ID
    LEFT JOIN GNLD_CURRENCY GC ON GC.CUR_ID = TD.CUR_TRA_ID
    LEFT JOIN LMSD_L_AGR_PROJ_TYPE PJ ON PJ.PROJECT_ID = YK.PROJECT_ID
    {yk_join}
    WHERE YK.DOC_DATE BETWEEN TO_DATE(:bas, 'DD.MM.YYYY') AND TO_DATE(:bit, 'DD.MM.YYYY')
    {yk_where}
ORDER BY SEVK_NO, TIP, YUK_NO, OPERASYON_KODU
""",
        limit,
    )


def kalem_detay_getir(cursor, bas: str, bit: str, bind: dict) -> tuple[list[dict], bool]:
    """
    Sevk + yük kalem detay satırlarını döner.
    Returns: (satirlar, limit_asildi_mi)
    """
    ok, _ = kalem_detay_semasi_hazir()
    if not ok:
        return [], False

    limit = int(getattr(ayarlar, "KPI_KALEM_DETAY_LIMIT", 10000))
    cursor.execute(_detay_sql(limit), bind)
    sutunlar = [c[0] for c in cursor.description]
    satirlar = [dict(zip(sutunlar, satir)) for satir in cursor.fetchall()]

    limit_asildi = len(satirlar) >= limit
    return satirlar, limit_asildi


def kalem_detay_ozet_hesapla(detay: list[dict], mevcut: bool = True, limit_asildi: bool = False) -> dict[str, Any]:
    if not mevcut:
        return {"mevcut": False, "mesaj": "Kalem detay tabloları bulunamadı."}

    if not detay:
        return {
            "mevcut": True,
            "sevk_kalem_sayisi": 0,
            "yuk_kalem_sayisi": 0,
            "net_sevk_alis": Decimal("0"),
            "net_yuk_satis": Decimal("0"),
            "faturasiz_kalem_sayisi": 0,
            "limit_asildi": limit_asildi,
        }

    sevk_kalem = yuk_kalem = 0
    net_sevk_alis = Decimal("0")
    net_yuk_satis = Decimal("0")
    faturasiz = 0

    for satir in detay:
        tip = satir.get("TIP")
        tutar = _decimal(satir.get("TUTAR"))
        alis_satis = satir.get("ALIS_SATIS") or ""

        if not satir.get("FATURA_NO") and alis_satis in ("Satış", "Alış"):
            faturasiz += 1

        if tip == "Sevk Kalemi":
            sevk_kalem += 1
            if alis_satis in ("Alış", "Alış İade"):
                net_sevk_alis += tutar if alis_satis == "Alış" else -tutar
        elif tip == "Yük Kalemi":
            yuk_kalem += 1
            if alis_satis in ("Satış", "Satış İade"):
                net_yuk_satis += tutar if alis_satis == "Satış" else -tutar

    return {
        "mevcut": True,
        "sevk_kalem_sayisi": sevk_kalem,
        "yuk_kalem_sayisi": yuk_kalem,
        "net_sevk_alis": net_sevk_alis,
        "net_yuk_satis": net_yuk_satis,
        "faturasiz_kalem_sayisi": faturasiz,
        "limit_asildi": limit_asildi,
    }


def kalem_sevk_kirilim_hesapla(detay: list[dict]) -> list[dict]:
    """
    Sevk bazında yük kırılımı — birden fazla yük içeren sevkler için özet.
    Sevk kalemleri TGD join nedeniyle yük başına tekrarlanır; net tutarlar
    benzersiz (SEVK_NO, OPERASYON, ALIS_SATIS, TUTAR, YUK_NO) ile toplanır.
    """
    sevkler: dict[str, dict] = defaultdict(lambda: {
        "SEVK_NO": "",
        "SEVK_TARIHI": None,
        "PLAKA": "",
        "YUK_NOS": set(),
        "SEVK_ALIS": Decimal("0"),
        "YUK_SATIS": Decimal("0"),
    })
    gorulen_sevk_kalem: dict[str, set] = defaultdict(set)

    for satir in detay:
        sevk_no = satir.get("SEVK_NO") or "?"
        g = sevkler[sevk_no]
        g["SEVK_NO"] = sevk_no
        g["SEVK_TARIHI"] = satir.get("SEVK_TARIHI")
        g["PLAKA"] = satir.get("PLAKA") or g["PLAKA"]
        yuk_no = satir.get("YUK_NO")
        if yuk_no:
            g["YUK_NOS"].add(yuk_no)

        tip = satir.get("TIP")
        alis_satis = satir.get("ALIS_SATIS") or ""
        tutar = _decimal(satir.get("TUTAR"))
        op = satir.get("OPERASYON_KODU") or ""

        if tip == "Sevk Kalemi" and alis_satis in ("Alış", "Alış İade"):
            # TGD join aynı sevk kalemini yük sayısı kadar tekrarlar — yük_no hariç tekilleştir
            anahtar = (op, alis_satis, str(tutar))
            if anahtar not in gorulen_sevk_kalem[sevk_no]:
                gorulen_sevk_kalem[sevk_no].add(anahtar)
                g["SEVK_ALIS"] += tutar if alis_satis == "Alış" else -tutar
        elif tip == "Yük Kalemi" and alis_satis in ("Satış", "Satış İade"):
            g["YUK_SATIS"] += tutar if alis_satis == "Satış" else -tutar

    sonuc = []
    for g in sevkler.values():
        if len(g["YUK_NOS"]) < 2:
            continue
        yuk_listesi = ", ".join(sorted(g["YUK_NOS"]))
        sevk_alis = g["SEVK_ALIS"]
        yuk_satis = g["YUK_SATIS"]
        sonuc.append({
            "SEVK_NO": g["SEVK_NO"],
            "SEVK_TARIHI": g["SEVK_TARIHI"],
            "PLAKA": g["PLAKA"],
            "YUK_SAYISI": len(g["YUK_NOS"]),
            "YUK_LISTESI": yuk_listesi,
            "SEVK_ALIS_TOPLAM": sevk_alis,
            "YUK_SATIS_TOPLAM": yuk_satis,
            "NET_KAR_ZARAR": yuk_satis - sevk_alis,
        })

    sonuc.sort(key=lambda x: x["NET_KAR_ZARAR"])
    return sonuc


def kalem_detay_problemleri(ozet: dict, kirilim: list[dict]) -> list[dict]:
    problemler = []
    if not ozet.get("mevcut"):
        return problemler

    if ozet.get("faturasiz_kalem_sayisi", 0) > 50:
        problemler.append({
            "oncelik": "ORTA",
            "kategori": "Kalem Detay",
            "baslik": "Faturasız sevk/yük kalemleri yüksek",
            "detay": f"{ozet['faturasiz_kalem_sayisi']} kalemde fatura no yok.",
            "aksiyon": "Kalem detay sayfasından faturasız satırları proje bazında kapatın.",
        })

    zararli = [k for k in kirilim if _decimal(k.get("NET_KAR_ZARAR")) < 0]
    if zararli:
        ornek = ", ".join(
            f"{k['SEVK_NO']} ({_decimal(k['NET_KAR_ZARAR']):,.0f} TL)" for k in zararli[:3]
        )
        problemler.append({
            "oncelik": "YUKSEK",
            "kategori": "Kalem Detay",
            "baslik": f"Çoklu yüklü sevklerde zarar ({len(zararli)} sevk)",
            "detay": f"Birden fazla yük taşıyan sevklerde satış−alış negatif. Örnek: {ornek}",
            "aksiyon": "Sevk Yük Kırılım sayfasından maliyet paylaşımını ve fiyatlandırmayı inceleyin.",
        })

    if ozet.get("limit_asildi"):
        limit = getattr(ayarlar, "KPI_KALEM_DETAY_LIMIT", 10000)
        problemler.append({
            "oncelik": "ORTA",
            "kategori": "Kalem Detay",
            "baslik": "Kalem detay satır limiti aşıldı",
            "detay": f"Excel'e en fazla {limit} satır aktarıldı; tam liste için dönemi daraltın.",
            "aksiyon": f"ayarlar.py içinde KPI_KALEM_DETAY_LIMIT artırılabilir veya tarih aralığı kısaltılabilir.",
        })

    return problemler


def ornek_kalem_detay_verisi() -> tuple[dict, list[dict], list[dict]]:
    detay = [
        {
            "TIP": "Sevk Kalemi", "SEVK_NO": "S-1001", "SEVK_TARIHI": "10.01.2026",
            "YUK_NO": "Y-5001", "PLAKA": "34ABC123", "ALIS_SATIS": "Alış",
            "OPERASYON_KODU": "NAVLUN", "TUTAR": Decimal("45000"), "FATURA_NO": None,
            "CARI_ADI": "Tedarikçi A", "PROJE_KODU": "MAPFRE",
        },
        {
            "TIP": "Sevk Kalemi", "SEVK_NO": "S-1001", "SEVK_TARIHI": "10.01.2026",
            "YUK_NO": "Y-5002", "PLAKA": "34ABC123", "ALIS_SATIS": "Alış",
            "OPERASYON_KODU": "NAVLUN", "TUTAR": Decimal("45000"), "FATURA_NO": None,
            "CARI_ADI": "Tedarikçi A", "PROJE_KODU": "MAPFRE",
        },
        {
            "TIP": "Yük Kalemi", "SEVK_NO": "S-1001", "SEVK_TARIHI": "10.01.2026",
            "YUK_NO": "Y-5001", "PLAKA": "34ABC123", "ALIS_SATIS": "Satış",
            "OPERASYON_KODU": "NAVLUN", "TUTAR": Decimal("52000"), "FATURA_NO": "F-001",
            "CARI_ADI": "Müşteri X", "PROJE_KODU": "MAPFRE",
        },
        {
            "TIP": "Yük Kalemi", "SEVK_NO": "S-1001", "SEVK_TARIHI": "10.01.2026",
            "YUK_NO": "Y-5002", "PLAKA": "34ABC123", "ALIS_SATIS": "Satış",
            "OPERASYON_KODU": "NAVLUN", "TUTAR": Decimal("38000"), "FATURA_NO": "F-002",
            "CARI_ADI": "Müşteri Y", "PROJE_KODU": "MAPFRE",
        },
    ]
    ozet = kalem_detay_ozet_hesapla(detay)
    kirilim = kalem_sevk_kirilim_hesapla(detay)
    return ozet, detay, kirilim

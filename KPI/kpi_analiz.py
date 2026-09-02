"""
İzlog Lojistik — Uyumsoft ERP KPI analiz modülü.

Otomasyon projesinden tamamen bağımsız; yalnızca KPI/ klasöründeki ayarlar.py
ve oracle_baglanti.py kullanılır.

Sevk ↔ yük ve fiyat tabloları Uyumsoft "yük bazında detay raporu" SQL'i ile
uyumludur [DOĞRULANMIŞ — kullanıcı rapor SQL'i]:
  LMST_L_TRANS_GOODS_DETAIL  (TGD.GOODS_ID ↔ YK, TGD.TRANSPORT_ID ↔ SK)
  LMST_L_TRANS_OP_DETAIL     (sevk alış / alış iade)
  LMST_L_GOODS_OP_DET        (yük satış / satış iade)
"""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime, timedelta
from decimal import Decimal
from typing import Any

import ayarlar
from kpi_fatura_detay import (
    fatura_detay_getir,
    fatura_detay_ozet_hesapla,
    fatura_detay_problemleri,
    fatura_detay_semasi_hazir,
    faturasiz_kalemler,
    ornek_fatura_detay_verisi,
)
from kpi_kalem_detay import (
    kalem_detay_getir,
    kalem_detay_ozet_hesapla,
    kalem_detay_problemleri,
    kalem_detay_semasi_hazir,
    kalem_sevk_kirilim_hesapla,
    ornek_kalem_detay_verisi,
)
from kpi_problem_detay import PROBLEM_DETAY_SUTUNLARI, ornek_problem_detay, problem_detay_olustur
from kpi_kiralk_arac import (
    kiralk_arac_cari_ozet_hesapla,
    kiralk_arac_detay_getir,
    kiralk_arac_ozet_hesapla,
    kiralk_arac_problemleri,
    kiralk_arac_semasi_hazir,
    ornek_kiralk_arac_verisi,
)
from oracle_baglanti import baglanti_yonet, satir_limit_sql, tablo_var_mi


def _decimal(deger) -> Decimal:
    if deger is None:
        return Decimal("0")
    return Decimal(str(deger))


def _tarih_araligi():
    """ayarlar.py'den veya varsayılan olarak son 30 gün."""
    bugun = datetime.now().date()
    varsayilan_bas = (bugun - timedelta(days=30)).strftime("%d.%m.%Y")
    varsayilan_bit = bugun.strftime("%d.%m.%Y")
    bas = getattr(ayarlar, "KPI_BASLANGIC_TARIHI", varsayilan_bas)
    bit = getattr(ayarlar, "KPI_BITIS_TARIHI", varsayilan_bit)
    return bas, bit


def _bind_ortak(bas: str, bit: str) -> dict:
    bind = {"bas": bas, "bit": bit}
    co_code = getattr(ayarlar, "CO_CODE", None)
    branch_code = getattr(ayarlar, "BRANCH_CODE", None)
    if co_code:
        bind["co_code"] = co_code
    if branch_code:
        bind["branch_code"] = branch_code
    return bind


def _yuk_filtre_sql(bind: dict | None = None) -> tuple[str, str]:
    """
    Yük (LMST_L_GOODS) sorguları için firma/şube/kapıdan kapıya filtreleri.
    Returns: (extra_join_sql, extra_where_sql)
    """
    bind = bind if bind is not None else {}
    joins: list[str] = []
    wheres: list[str] = []

    if getattr(ayarlar, "CO_CODE", None):
        joins.append("JOIN GNLD_COMPANY CO ON CO.CO_ID = YK.CO_ID")
        wheres.append("CO.CO_CODE = :co_code")
    if getattr(ayarlar, "BRANCH_CODE", None):
        joins.append("JOIN GNLD_BRANCH BR ON BR.BRANCH_ID = YK.BRANCH_ID")
        wheres.append("BR.BRANCH_CODE = :branch_code")
    if getattr(ayarlar, "KPI_KAPI_KAPI_HARIC", True):
        wheres.append("YK.IS_DOOR_TO_DOOR = 0")

    join_sql = "\n        ".join(joins)
    where_sql = (" AND " + " AND ".join(wheres)) if wheres else ""
    return join_sql, where_sql


def _sevk_semasi_hazir() -> tuple[bool, str]:
    """Gerekli sevk tablolarının varlığını kontrol eder."""
    if not tablo_var_mi("LMST_L_TRANS_GOODS_DETAIL"):
        return False, "LMST_L_TRANS_GOODS_DETAIL tablosu bulunamadı."
    if not tablo_var_mi("LMST_L_TRANSPORT"):
        return False, "LMST_L_TRANSPORT tablosu bulunamadı."
    return True, ""


@dataclass
class KpiAnalizSonucu:
    baslangic: str
    bitis: str
    ozet: dict[str, Any] = field(default_factory=dict)
    aylik_trend: list[dict] = field(default_factory=list)
    proje_performans: list[dict] = field(default_factory=list)
    operasyon_dagilimi: list[dict] = field(default_factory=list)
    fatura_sagligi: dict[str, Any] = field(default_factory=dict)
    marj_analizi: dict[str, Any] = field(default_factory=dict)
    kiralk_arac_ozet: dict[str, Any] = field(default_factory=dict)
    kiralk_arac_detay: list[dict] = field(default_factory=list)
    kiralk_arac_cari: list[dict] = field(default_factory=list)
    kalem_detay_ozet: dict[str, Any] = field(default_factory=dict)
    kalem_detay: list[dict] = field(default_factory=list)
    kalem_sevk_kirilim: list[dict] = field(default_factory=list)
    fatura_detay_ozet: dict[str, Any] = field(default_factory=dict)
    fatura_detay: list[dict] = field(default_factory=list)
    faturasiz_kalemler: list[dict] = field(default_factory=list)
    sevksiz_yukler: list[dict] = field(default_factory=list)
    problem_detay: list[dict] = field(default_factory=list)
    problemler: list[dict] = field(default_factory=list)
    uyarilar: list[str] = field(default_factory=list)


def _gelir_payi_yuzde_ekle(satirlar: list[dict], tutar_kolonu: str) -> list[dict]:
    """Gelir payı yüzdesini Python tarafında hesaplar (Oracle 11g uyumluluğu)."""
    toplam = sum(_decimal(s.get(tutar_kolonu, 0)) for s in satirlar)
    for satir in satirlar:
        tutar = _decimal(satir.get(tutar_kolonu, 0))
        if toplam > 0:
            satir["GELIR_PAYI_YUZDE"] = round(float(tutar * 100 / toplam), 1)
        else:
            satir["GELIR_PAYI_YUZDE"] = 0.0
    return satirlar


def _satirlari_dict_yap(sutunlar, satirlar):
    return [dict(zip(sutunlar, satir)) for satir in satirlar]


def _temel_hacim_kpi(cursor, bas, bit):
    """Yük ve sevk hacmi KPI'ları (TGD üzerinden sevk eşlemesi)."""
    bind = _bind_ortak(bas, bit)
    yuk_join, yuk_where = _yuk_filtre_sql(bind)

    cursor.execute(
        f"""
        SELECT COUNT(*) AS YUK_SAYISI
        FROM LMST_L_GOODS YK
        {yuk_join}
        WHERE YK.DOC_DATE BETWEEN TO_DATE(:bas, 'DD.MM.YYYY') AND TO_DATE(:bit, 'DD.MM.YYYY')
        {yuk_where}
        """,
        bind,
    )
    yuk_sayisi = cursor.fetchone()[0]

    sevk_ok, _ = _sevk_semasi_hazir()
    sevk_sayisi = 0
    if sevk_ok:
        cursor.execute(
            f"""
            SELECT COUNT(DISTINCT SK.TRANSPORT_ID) AS SEVK_SAYISI
            FROM LMST_L_GOODS YK
            JOIN LMST_L_TRANS_GOODS_DETAIL TGD
              ON TGD.GOODS_ID = YK.GOODS_ID
             AND TGD.GOODS_ID > 0
            JOIN LMST_L_TRANSPORT SK ON SK.TRANSPORT_ID = TGD.TRANSPORT_ID
            {yuk_join}
            WHERE YK.DOC_DATE BETWEEN TO_DATE(:bas, 'DD.MM.YYYY') AND TO_DATE(:bit, 'DD.MM.YYYY')
            {yuk_where}
            """,
            bind,
        )
        sevk_sayisi = cursor.fetchone()[0]

    sevk_orani = round(sevk_sayisi / yuk_sayisi * 100, 1) if yuk_sayisi else 0

    return {
        "yuk_sayisi": yuk_sayisi,
        "sevk_sayisi": sevk_sayisi,
        "sevk_orani_yuzde": sevk_orani,
        "yuk_basina_ortalama_gelir": Decimal("0"),
    }


def _gelir_kpi(cursor, bas, bit):
    """
    Net satış geliri ve fatura sağlığı.
    Satış: PURCHASE_SALES_TYPE IN (2,4); iade: IN (1,3) — rapor YFT alt sorgusu ile uyumlu.
    """
    bind = _bind_ortak(bas, bit)
    yuk_join, yuk_where = _yuk_filtre_sql(bind)

    cursor.execute(
        f"""
        SELECT
            NVL(SUM(CASE WHEN OPDET.PURCHASE_SALES_TYPE IN (2, 4) THEN OPDET.AMT ELSE 0 END), 0)
            - NVL(SUM(CASE WHEN OPDET.PURCHASE_SALES_TYPE IN (1, 3) THEN OPDET.AMT ELSE 0 END), 0)
            AS TOPLAM_SATIS,
            COUNT(DISTINCT YK.GOODS_ID) AS GELIRLI_YUK_SAYISI,
            COUNT(*) AS SATIS_SATIR_SAYISI,
            NVL(SUM(CASE WHEN OPDET.INVOICE_M_ID IS NULL AND OPDET.PURCHASE_SALES_TYPE IN (2, 4)
                     THEN OPDET.AMT ELSE 0 END), 0) AS FATURASIZ_TUTAR,
            NVL(SUM(CASE WHEN OPDET.INVOICE_M_ID IS NOT NULL AND OPDET.PURCHASE_SALES_TYPE IN (2, 4)
                     THEN OPDET.AMT ELSE 0 END), 0) AS FATURALI_TUTAR,
            COUNT(CASE WHEN OPDET.INVOICE_M_ID IS NULL AND OPDET.PURCHASE_SALES_TYPE IN (2, 4)
                       THEN 1 END) AS FATURASIZ_SATIR,
            COUNT(CASE WHEN OPDET.INVOICE_M_ID IS NOT NULL AND OPDET.PURCHASE_SALES_TYPE IN (2, 4)
                       THEN 1 END) AS FATURALI_SATIR
        FROM LMST_L_GOODS_OP_DET OPDET
        JOIN LMST_L_GOODS YK ON YK.GOODS_ID = OPDET.GOODS_ID
        {yuk_join}
        WHERE YK.DOC_DATE BETWEEN TO_DATE(:bas, 'DD.MM.YYYY') AND TO_DATE(:bit, 'DD.MM.YYYY')
          AND OPDET.PURCHASE_SALES_TYPE IN (1, 2, 3, 4)
        {yuk_where}
        """,
        bind,
    )
    row = cursor.fetchone()
    sutunlar = [c[0] for c in cursor.description]
    veri = dict(zip(sutunlar, row))

    toplam_satis = _decimal(veri["TOPLAM_SATIS"])
    satir_sayisi = veri["SATIS_SATIR_SAYISI"] or 0
    faturasiz_satir = veri["FATURASIZ_SATIR"] or 0
    satis_satir = (veri["FATURALI_SATIR"] or 0) + faturasiz_satir

    return {
        "toplam_satis_geliri": toplam_satis,
        "gelirli_yuk_sayisi": veri["GELIRLI_YUK_SAYISI"] or 0,
        "satis_satir_sayisi": satir_sayisi,
        "faturasiz_tutar": _decimal(veri["FATURASIZ_TUTAR"]),
        "faturali_tutar": _decimal(veri["FATURALI_TUTAR"]),
        "fatura_baglama_orani_yuzde": round(
            (veri["FATURALI_SATIR"] or 0) / satis_satir * 100, 1
        )
        if satis_satir
        else 0,
    }


def _aylik_trend(cursor, bas, bit):
    bind = _bind_ortak(bas, bit)
    yuk_join, yuk_where = _yuk_filtre_sql(bind)

    cursor.execute(
        f"""
        SELECT
            TO_CHAR(YK.DOC_DATE, 'YYYY-MM') AS AY,
            COUNT(DISTINCT YK.GOODS_ID) AS YUK_SAYISI,
            NVL(SUM(CASE WHEN OPDET.PURCHASE_SALES_TYPE IN (2, 4) THEN OPDET.AMT ELSE 0 END), 0)
            - NVL(SUM(CASE WHEN OPDET.PURCHASE_SALES_TYPE IN (1, 3) THEN OPDET.AMT ELSE 0 END), 0)
            AS SATIS_GELIRI
        FROM LMST_L_GOODS YK
        LEFT JOIN LMST_L_GOODS_OP_DET OPDET
            ON OPDET.GOODS_ID = YK.GOODS_ID
           AND OPDET.PURCHASE_SALES_TYPE IN (1, 2, 3, 4)
        {yuk_join}
        WHERE YK.DOC_DATE BETWEEN TO_DATE(:bas, 'DD.MM.YYYY') AND TO_DATE(:bit, 'DD.MM.YYYY')
        {yuk_where}
        GROUP BY TO_CHAR(YK.DOC_DATE, 'YYYY-MM')
        ORDER BY 1
        """,
        bind,
    )
    sutunlar = [c[0] for c in cursor.description]
    return _satirlari_dict_yap(sutunlar, cursor.fetchall())


def _proje_performans(cursor, bas, bit):
    bind = _bind_ortak(bas, bit)
    yuk_join, yuk_where = _yuk_filtre_sql(bind)

    cursor.execute(
        f"""
        SELECT
            NVL(P.PROJECT_CODE, '(Proje Yok)') AS PROJE_KODU,
            COUNT(DISTINCT YK.GOODS_ID) AS YUK_SAYISI,
            NVL(SUM(CASE WHEN OPDET.PURCHASE_SALES_TYPE IN (2, 4) THEN OPDET.AMT ELSE 0 END), 0)
            - NVL(SUM(CASE WHEN OPDET.PURCHASE_SALES_TYPE IN (1, 3) THEN OPDET.AMT ELSE 0 END), 0)
            AS SATIS_GELIRI
        FROM LMST_L_GOODS YK
        LEFT JOIN LMSD_L_AGR_PROJ_TYPE P ON P.PROJECT_ID = YK.PROJECT_ID
        LEFT JOIN LMST_L_GOODS_OP_DET OPDET
            ON OPDET.GOODS_ID = YK.GOODS_ID
           AND OPDET.PURCHASE_SALES_TYPE IN (1, 2, 3, 4)
        {yuk_join}
        WHERE YK.DOC_DATE BETWEEN TO_DATE(:bas, 'DD.MM.YYYY') AND TO_DATE(:bit, 'DD.MM.YYYY')
        {yuk_where}
        GROUP BY NVL(P.PROJECT_CODE, '(Proje Yok)')
        ORDER BY 3 DESC
        """,
        bind,
    )
    sutunlar = [c[0] for c in cursor.description]
    satirlar = _gelir_payi_yuzde_ekle(_satirlari_dict_yap(sutunlar, cursor.fetchall()), "SATIS_GELIRI")
    return satirlar[:20]


def _operasyon_dagilimi(cursor, bas, bit):
    bind = _bind_ortak(bas, bit)
    yuk_join, yuk_where = _yuk_filtre_sql(bind)

    cursor.execute(
        f"""
        SELECT
            NVL(HK.EXPENSE_CODE, 'BILINMIYOR') AS OPERASYON_KODU,
            COUNT(*) AS SATIR_SAYISI,
            NVL(SUM(CASE WHEN OPDET.PURCHASE_SALES_TYPE IN (2, 4) THEN OPDET.AMT ELSE 0 END), 0)
            - NVL(SUM(CASE WHEN OPDET.PURCHASE_SALES_TYPE IN (1, 3) THEN OPDET.AMT ELSE 0 END), 0)
            AS TOPLAM_TUTAR
        FROM LMST_L_GOODS_OP_DET OPDET
        JOIN LMST_L_GOODS YK ON YK.GOODS_ID = OPDET.GOODS_ID
        LEFT JOIN INVD_EXPENSE HK ON HK.EXPENSE_ID = OPDET.OPERATION_ID
        {yuk_join}
        WHERE YK.DOC_DATE BETWEEN TO_DATE(:bas, 'DD.MM.YYYY') AND TO_DATE(:bit, 'DD.MM.YYYY')
          AND OPDET.PURCHASE_SALES_TYPE IN (1, 2, 3, 4)
        {yuk_where}
        GROUP BY NVL(HK.EXPENSE_CODE, 'BILINMIYOR')
        ORDER BY 3 DESC
        """,
        bind,
    )
    sutunlar = [c[0] for c in cursor.description]
    return _gelir_payi_yuzde_ekle(_satirlari_dict_yap(sutunlar, cursor.fetchall()), "TOPLAM_TUTAR")


def _fatura_gecikmesi(cursor, bas, bit):
    bind = _bind_ortak(bas, bit)
    yuk_join, yuk_where = _yuk_filtre_sql(bind)

    cursor.execute(
        f"""
        SELECT
            ROUND(AVG(INV.DOC_DATE - YK.DOC_DATE), 1) AS ORT_GECIKME_GUN,
            ROUND(MAX(INV.DOC_DATE - YK.DOC_DATE), 0) AS MAX_GECIKME_GUN,
            COUNT(*) AS FATURALI_SATIR
        FROM LMST_L_GOODS_OP_DET OPDET
        JOIN LMST_L_GOODS YK ON YK.GOODS_ID = OPDET.GOODS_ID
        JOIN PSMT_INVOICE_M INV ON INV.INVOICE_M_ID = OPDET.INVOICE_M_ID
        {yuk_join}
        WHERE YK.DOC_DATE BETWEEN TO_DATE(:bas, 'DD.MM.YYYY') AND TO_DATE(:bit, 'DD.MM.YYYY')
          AND OPDET.PURCHASE_SALES_TYPE IN (2, 4)
        {yuk_where}
        """,
        bind,
    )
    row = cursor.fetchone()
    if not row or row[2] == 0:
        return {"ort_gecikme_gun": None, "max_gecikme_gun": None, "faturali_satir": 0}

    return {
        "ort_gecikme_gun": float(row[0]) if row[0] is not None else None,
        "max_gecikme_gun": int(row[1]) if row[1] is not None else None,
        "faturali_satir": row[2],
    }


def _marj_analizi(cursor, bas, bit):
    """
    Brüt marj: net satış − sevk alış maliyeti (yük başına paylaştırılmış).
    Tablolar: LMST_L_GOODS_OP_DET (satış), LMST_L_TRANS_OP_DETAIL (alış),
    LMST_L_TRANS_GOODS_DETAIL (yük↔sevk), YUK_SAYISI CTE (rapor SQL'i ile aynı mantık).
    """
    if not tablo_var_mi("LMST_L_TRANS_OP_DETAIL"):
        return {
            "mevcut": False,
            "mesaj": "LMST_L_TRANS_OP_DETAIL tablosu bulunamadı — marj KPI atlandı.",
        }

    sevk_ok, sevk_mesaj = _sevk_semasi_hazir()
    if not sevk_ok:
        return {"mevcut": False, "mesaj": sevk_mesaj}

    bind = _bind_ortak(bas, bit)
    yuk_join, yuk_where = _yuk_filtre_sql(bind)

    cursor.execute(
        f"""
        WITH YUK_SAYISI AS (
            SELECT TGD.TRANSPORT_ID, COUNT(*) AS YUK_ADEDI
            FROM LMST_L_TRANS_GOODS_DETAIL TGD
            WHERE TGD.GOODS_ID > 0
            GROUP BY TGD.TRANSPORT_ID
        ),
        YFT AS (
            SELECT
                GOD.GOODS_ID,
                SUM(CASE WHEN GOD.PURCHASE_SALES_TYPE IN (2, 4) THEN GOD.AMT ELSE 0 END) AS SATIS,
                SUM(CASE WHEN GOD.PURCHASE_SALES_TYPE IN (1, 3) THEN GOD.AMT ELSE 0 END) AS SATIS_IADE
            FROM LMST_L_GOODS_OP_DET GOD
            GROUP BY GOD.GOODS_ID
        ),
        SFT AS (
            SELECT
                TGD.GOODS_ID,
                SUM(CASE WHEN TOD.PURCHASE_SALES_TYPE IN (1, 3) THEN TOD.AMT ELSE 0 END)
                    / NVL(YS.YUK_ADEDI, 1) AS ALIS,
                SUM(CASE WHEN TOD.PURCHASE_SALES_TYPE IN (2, 4) THEN TOD.AMT ELSE 0 END)
                    / NVL(YS.YUK_ADEDI, 1) AS ALIS_IADE
            FROM LMST_L_TRANS_GOODS_DETAIL TGD
            JOIN LMST_L_TRANSPORT SK ON SK.TRANSPORT_ID = TGD.TRANSPORT_ID
            JOIN LMST_L_TRANS_OP_DETAIL TOD ON TOD.TRANSPORT_ID = SK.TRANSPORT_ID
            LEFT JOIN YUK_SAYISI YS ON YS.TRANSPORT_ID = SK.TRANSPORT_ID
            WHERE TGD.GOODS_ID > 0
            GROUP BY TGD.GOODS_ID, NVL(YS.YUK_ADEDI, 1)
        )
        SELECT
            NVL(SUM(NVL(YFT.SATIS, 0) - NVL(YFT.SATIS_IADE, 0)), 0) AS TOPLAM_SATIS,
            NVL(SUM(NVL(SFT.ALIS, 0) - NVL(SFT.ALIS_IADE, 0)), 0) AS TOPLAM_ALIS
        FROM LMST_L_GOODS YK
        LEFT JOIN YFT ON YFT.GOODS_ID = YK.GOODS_ID
        LEFT JOIN SFT ON SFT.GOODS_ID = YK.GOODS_ID
        {yuk_join}
        WHERE YK.DOC_DATE BETWEEN TO_DATE(:bas, 'DD.MM.YYYY') AND TO_DATE(:bit, 'DD.MM.YYYY')
        {yuk_where}
        """,
        bind,
    )
    row = cursor.fetchone()
    toplam_satis = _decimal(row[0])
    toplam_alis = _decimal(row[1])
    brut_marj = toplam_satis - toplam_alis
    marj_orani = round(float(brut_marj / toplam_satis * 100), 1) if toplam_satis else 0

    return {
        "mevcut": True,
        "toplam_satis": toplam_satis,
        "toplam_alis": toplam_alis,
        "brut_marj": brut_marj,
        "brut_marj_orani_yuzde": marj_orani,
    }


def _sevksiz_yukleri_getir(cursor, bas, bit, limit: int = 500) -> list[dict]:
    sevk_ok, _ = _sevk_semasi_hazir()
    if not sevk_ok:
        return []

    bind = _bind_ortak(bas, bit)
    yuk_join, yuk_where = _yuk_filtre_sql(bind)
    cursor.execute(
        satir_limit_sql(
            f"""
            SELECT YK.REFERENCE_NO, YK.DOC_DATE, NVL(P.PROJECT_CODE, '-') AS PROJE
            FROM LMST_L_GOODS YK
            LEFT JOIN LMSD_L_AGR_PROJ_TYPE P ON P.PROJECT_ID = YK.PROJECT_ID
            LEFT JOIN LMST_L_TRANS_GOODS_DETAIL TGD
              ON TGD.GOODS_ID = YK.GOODS_ID
             AND TGD.GOODS_ID > 0
            {yuk_join}
            WHERE YK.DOC_DATE BETWEEN TO_DATE(:bas, 'DD.MM.YYYY') AND TO_DATE(:bit, 'DD.MM.YYYY')
              AND TGD.TRANSPORT_ID IS NULL
            {yuk_where}
            ORDER BY YK.DOC_DATE DESC
            """,
            limit,
        ),
        bind,
    )
    sutunlar = [c[0] for c in cursor.description]
    return _satirlari_dict_yap(sutunlar, cursor.fetchall())


def _problemleri_tespit_et(cursor, bas, bit, ozet, proje_listesi, marj, sevksiz_yukler: list[dict]):
    problemler = []

    if ozet.get("fatura_baglama_orani_yuzde", 100) < 95:
        problemler.append(
            {
                "oncelik": "YUKSEK",
                "kategori": "Fatura / Tahsilat",
                "baslik": "Fatura bağlama oranı hedefin altında",
                "detay": (
                    f"Satış satırlarının %{ozet['fatura_baglama_orani_yuzde']}'i faturalı. "
                    f"Faturasız tutar: {ozet['faturasiz_tutar']:,.2f} TL"
                ),
                "aksiyon": "Faturasız satır listesini operasyon ekibiyle paylaşın; fatura kesim sürecini hızlandırın.",
            }
        )

    if ozet.get("yuk_sayisi", 0) > 0 and ozet.get("sevk_orani_yuzde", 100) < 90:
        problemler.append(
            {
                "oncelik": "YUKSEK",
                "kategori": "Operasyon",
                "baslik": "Sevk tamamlama oranı düşük",
                "detay": (
                    f"{ozet['yuk_sayisi']} yükten yalnızca {ozet['sevk_sayisi']} sevk "
                    f"(%{ozet['sevk_orani_yuzde']})."
                ),
                "aksiyon": "Sevki oluşturulmamış yükleri listeleyin; nakliye planlama darboğazını inceleyin.",
            }
        )

    if proje_listesi:
        top3_pay = sum(_decimal(p.get("GELIR_PAYI_YUZDE", 0)) for p in proje_listesi[:3])
        if top3_pay > 60:
            projeler = ", ".join(p["PROJE_KODU"] for p in proje_listesi[:3])
            problemler.append(
                {
                    "oncelik": "ORTA",
                    "kategori": "Strateji / Risk",
                    "baslik": "Gelir proje konsantrasyonu yüksek",
                    "detay": f"İlk 3 proje toplam gelirin %{top3_pay}'ini oluşturuyor: {projeler}",
                    "aksiyon": "Müşteri çeşitlendirme planı ve yeni proje pipeline'ını gözden geçirin.",
                }
            )

    if marj.get("mevcut") and marj.get("brut_marj_orani_yuzde", 0) < 10:
        problemler.append(
            {
                "oncelik": "YUKSEK",
                "kategori": "Karlılık",
                "baslik": "Brüt marj oranı kritik seviyede",
                "detay": f"Brüt marj oranı: %{marj['brut_marj_orani_yuzde']} (net satış − sevk alış)",
                "aksiyon": "Düşük marjlı yükleri proje/operasyon bazında analiz edin; fiyatlandırma revizyonu değerlendirin.",
            }
        )

    sevk_ok, _ = _sevk_semasi_hazir()
    if sevk_ok and sevksiz_yukler:
        problemler.append(
            {
                "oncelik": "YUKSEK",
                "kategori": "Operasyon",
                "baslik": f"Sevki olmayan yükler ({len(sevksiz_yukler)}+ kayıt)",
                "detay": "Örnek: " + ", ".join(r["REFERENCE_NO"] for r in sevksiz_yukler[:5]),
                "aksiyon": "Sevki oluşturulmamış yük listesini günlük operasyon toplantısına taşıyın.",
            }
        )

    bind = _bind_ortak(bas, bit)
    yuk_join, yuk_where = _yuk_filtre_sql(bind)
    cursor.execute(
        f"""
        SELECT COUNT(*)
        FROM LMST_L_GOODS_OP_DET OPDET
        JOIN LMST_L_GOODS YK ON YK.GOODS_ID = OPDET.GOODS_ID
        {yuk_join}
        WHERE YK.DOC_DATE BETWEEN TO_DATE(:bas, 'DD.MM.YYYY') AND TO_DATE(:bit, 'DD.MM.YYYY')
          AND OPDET.PURCHASE_SALES_TYPE IN (2, 4)
          AND NVL(OPDET.AMT, 0) = 0
        {yuk_where}
        """,
        bind,
    )
    sifir_tutar = cursor.fetchone()[0]
    if sifir_tutar > 0:
        problemler.append(
            {
                "oncelik": "ORTA",
                "kategori": "Veri Kalitesi",
                "baslik": "Sıfır tutarlı satış satırları",
                "detay": f"{sifir_tutar} satırda tutar 0 TL.",
                "aksiyon": "ERP veri giriş kalite kontrolü; otomasyon öncesi satır doğrulama kuralı ekleyin.",
            }
        )

    oncelik_sira = {"YUKSEK": 0, "ORTA": 1, "DUSUK": 2}
    problemler.sort(key=lambda p: oncelik_sira.get(p["oncelik"], 9))
    return problemler


def kpi_analizi_yap(baslangic=None, bitis=None) -> KpiAnalizSonucu:
    bas, bit = baslangic or _tarih_araligi()[0], bitis or _tarih_araligi()[1]
    sonuc = KpiAnalizSonucu(baslangic=bas, bitis=bit)

    with baglanti_yonet() as baglanti:
        cursor = baglanti.cursor()

        hacim = _temel_hacim_kpi(cursor, bas, bit)
        gelir = _gelir_kpi(cursor, bas, bit)

        yuk_sayisi = hacim["yuk_sayisi"] or 0
        ort_gelir = gelir["toplam_satis_geliri"] / yuk_sayisi if yuk_sayisi else Decimal("0")

        sonuc.ozet = {
            **hacim,
            **gelir,
            "yuk_basina_ortalama_gelir": ort_gelir,
            "donem": f"{bas} — {bit}",
            "rapor_tarihi": datetime.now().strftime("%d.%m.%Y %H:%M"),
        }

        sonuc.aylik_trend = _aylik_trend(cursor, bas, bit)
        sonuc.proje_performans = _proje_performans(cursor, bas, bit)
        sonuc.operasyon_dagilimi = _operasyon_dagilimi(cursor, bas, bit)
        sonuc.fatura_sagligi = _fatura_gecikmesi(cursor, bas, bit)
        sonuc.marj_analizi = _marj_analizi(cursor, bas, bit)

        ka_ok, ka_mesaj = kiralk_arac_semasi_hazir()
        if ka_ok:
            bind = _bind_ortak(bas, bit)
            sonuc.kiralk_arac_detay = kiralk_arac_detay_getir(cursor, bas, bit, bind)
            sonuc.kiralk_arac_ozet = kiralk_arac_ozet_hesapla(sonuc.kiralk_arac_detay)
            sonuc.kiralk_arac_cari = kiralk_arac_cari_ozet_hesapla(sonuc.kiralk_arac_detay)
        else:
            sonuc.kiralk_arac_ozet = {"mevcut": False, "mesaj": ka_mesaj}

        kd_ok, kd_mesaj = kalem_detay_semasi_hazir()
        if kd_ok:
            bind = _bind_ortak(bas, bit)
            sonuc.kalem_detay, limit_asildi = kalem_detay_getir(cursor, bas, bit, bind)
            sonuc.kalem_detay_ozet = kalem_detay_ozet_hesapla(
                sonuc.kalem_detay, limit_asildi=limit_asildi
            )
            sonuc.kalem_sevk_kirilim = kalem_sevk_kirilim_hesapla(sonuc.kalem_detay)
        else:
            sonuc.kalem_detay_ozet = {"mevcut": False, "mesaj": kd_mesaj}

        fd_ok, fd_mesaj = fatura_detay_semasi_hazir()
        if fd_ok:
            bind = _bind_ortak(bas, bit)
            sonuc.fatura_detay, fd_limit = fatura_detay_getir(cursor, bas, bit, bind)
            sonuc.fatura_detay_ozet = fatura_detay_ozet_hesapla(
                sonuc.fatura_detay, limit_asildi=fd_limit
            )
            sonuc.faturasiz_kalemler = faturasiz_kalemler(sonuc.fatura_detay)
        else:
            sonuc.fatura_detay_ozet = {"mevcut": False, "mesaj": fd_mesaj}

        if getattr(ayarlar, "CO_CODE", None):
            sonuc.uyarilar.append(f"Firma filtresi: CO_CODE = {ayarlar.CO_CODE}")
        if getattr(ayarlar, "BRANCH_CODE", None):
            sonuc.uyarilar.append(f"Şube filtresi: BRANCH_CODE = {ayarlar.BRANCH_CODE}")
        if not sonuc.marj_analizi.get("mevcut"):
            sonuc.uyarilar.append(sonuc.marj_analizi.get("mesaj", "Marj analizi atlandı."))
        if not sonuc.kiralk_arac_ozet.get("mevcut"):
            mesaj = sonuc.kiralk_arac_ozet.get("mesaj")
            if mesaj and "bulunamadı" in mesaj.lower() and "tablo" in mesaj.lower():
                sonuc.uyarilar.append(mesaj)
        elif sonuc.kiralk_arac_ozet.get("dosya_sayisi", 0) == 0:
            mesaj = sonuc.kiralk_arac_ozet.get("mesaj")
            if mesaj:
                sonuc.uyarilar.append(mesaj)
        if not sonuc.kalem_detay_ozet.get("mevcut"):
            mesaj = sonuc.kalem_detay_ozet.get("mesaj")
            if mesaj and "tablo" in mesaj.lower():
                sonuc.uyarilar.append(mesaj)
        elif sonuc.kalem_detay_ozet.get("limit_asildi"):
            limit = getattr(ayarlar, "KPI_KALEM_DETAY_LIMIT", 10000)
            sonuc.uyarilar.append(
                f"Kalem detay {limit} satır limitine ulaştı — tam liste için dönemi daraltın."
            )
        if not sonuc.fatura_detay_ozet.get("mevcut"):
            mesaj = sonuc.fatura_detay_ozet.get("mesaj")
            if mesaj and "tablo" in mesaj.lower():
                sonuc.uyarilar.append(mesaj)
        elif sonuc.fatura_detay_ozet.get("limit_asildi"):
            limit = getattr(ayarlar, "KPI_FATURA_DETAY_LIMIT", 10000)
            sonuc.uyarilar.append(
                f"Fatura detay {limit} satır limitine ulaştı — tam liste için dönemi daraltın."
            )

        sonuc.sevksiz_yukler = _sevksiz_yukleri_getir(cursor, bas, bit)
        sonuc.problemler = _problemleri_tespit_et(
            cursor, bas, bit, sonuc.ozet, sonuc.proje_performans, sonuc.marj_analizi, sonuc.sevksiz_yukler
        )
        sonuc.problemler.extend(
            kiralk_arac_problemleri(sonuc.kiralk_arac_ozet, sonuc.kiralk_arac_detay)
        )
        sonuc.problemler.extend(
            kalem_detay_problemleri(sonuc.kalem_detay_ozet, sonuc.kalem_sevk_kirilim)
        )
        sonuc.problemler.extend(
            fatura_detay_problemleri(sonuc.fatura_detay_ozet, sonuc.fatura_detay)
        )
        sonuc.problem_detay = problem_detay_olustur(
            sonuc.fatura_detay,
            sonuc.kalem_detay,
            sonuc.kalem_sevk_kirilim,
            sonuc.sevksiz_yukler,
        )
        oncelik_sira = {"YUKSEK": 0, "ORTA": 1, "DUSUK": 2}
        sonuc.problemler.sort(key=lambda p: oncelik_sira.get(p["oncelik"], 9))

    return sonuc


def ornek_analiz_sonucu() -> KpiAnalizSonucu:
    sonuc = KpiAnalizSonucu(
        baslangic="01.01.2026",
        bitis="31.01.2026",
    )
    sonuc.ozet = {
        "donem": "01.01.2026 — 31.01.2026",
        "rapor_tarihi": datetime.now().strftime("%d.%m.%Y %H:%M"),
        "yuk_sayisi": 1247,
        "sevk_sayisi": 1189,
        "sevk_orani_yuzde": 95.3,
        "toplam_satis_geliri": Decimal("48752300.00"),
        "yuk_basina_ortalama_gelir": Decimal("39095.67"),
        "fatura_baglama_orani_yuzde": 92.1,
        "faturasiz_tutar": Decimal("3842100.00"),
        "faturali_tutar": Decimal("44910200.00"),
        "satis_satir_sayisi": 2891,
    }
    sonuc.aylik_trend = [
        {"AY": "2026-01", "YUK_SAYISI": 1247, "SATIS_GELIRI": Decimal("48752300")},
    ]
    sonuc.proje_performans = [
        {"PROJE_KODU": "MAPFRE", "YUK_SAYISI": 412, "SATIS_GELIRI": Decimal("18200000"), "GELIR_PAYI_YUZDE": 37.3},
        {"PROJE_KODU": "ARCELIK", "YUK_SAYISI": 298, "SATIS_GELIRI": Decimal("12100000"), "GELIR_PAYI_YUZDE": 24.8},
        {"PROJE_KODU": "TRENDYOL", "YUK_SAYISI": 187, "SATIS_GELIRI": Decimal("6800000"), "GELIR_PAYI_YUZDE": 13.9},
    ]
    sonuc.operasyon_dagilimi = [
        {"OPERASYON_KODU": "NAVLUN", "SATIR_SAYISI": 2100, "TOPLAM_TUTAR": Decimal("42000000"), "GELIR_PAYI_YUZDE": 86.2},
        {"OPERASYON_KODU": "UĞRAMA", "SATIR_SAYISI": 520, "TOPLAM_TUTAR": Decimal("4200000"), "GELIR_PAYI_YUZDE": 8.6},
        {"OPERASYON_KODU": "BEKLEME", "SATIR_SAYISI": 271, "TOPLAM_TUTAR": Decimal("2552300"), "GELIR_PAYI_YUZDE": 5.2},
    ]
    sonuc.fatura_sagligi = {"ort_gecikme_gun": 4.2, "max_gecikme_gun": 28, "faturali_satir": 2663}
    sonuc.marj_analizi = {
        "mevcut": True,
        "toplam_satis": Decimal("48752300"),
        "toplam_alis": Decimal("42100000"),
        "brut_marj": Decimal("6652300"),
        "brut_marj_orani_yuzde": 13.6,
    }
    ka_ozet, ka_detay, ka_cari = ornek_kiralk_arac_verisi()
    sonuc.kiralk_arac_ozet = ka_ozet
    sonuc.kiralk_arac_detay = ka_detay
    sonuc.kiralk_arac_cari = ka_cari
    kd_ozet, kd_detay, kd_kirilim = ornek_kalem_detay_verisi()
    sonuc.kalem_detay_ozet = kd_ozet
    sonuc.kalem_detay = kd_detay
    sonuc.kalem_sevk_kirilim = kd_kirilim
    fd_ozet, fd_detay, fd_faturasiz = ornek_fatura_detay_verisi()
    sonuc.fatura_detay_ozet = fd_ozet
    sonuc.fatura_detay = fd_detay
    sonuc.faturasiz_kalemler = fd_faturasiz
    sonuc.problemler = [
        {
            "oncelik": "YUKSEK",
            "kategori": "Fatura / Tahsilat",
            "baslik": "Fatura bağlama oranı hedefin altında",
            "detay": "Satış satırlarının %92.1'i faturalı. Faturasız tutar: 3.842.100,00 TL",
            "aksiyon": "Faturasız satır listesini operasyon ekibiyle paylaşın.",
        },
        {
            "oncelik": "ORTA",
            "kategori": "Strateji / Risk",
            "baslik": "Gelir proje konsantrasyonu yüksek",
            "detay": "İlk 3 proje toplam gelirin %76.0'ını oluşturuyor.",
            "aksiyon": "Müşteri çeşitlendirme planını gözden geçirin.",
        },
    ]
    sonuc.problemler.extend(kiralk_arac_problemleri(ka_ozet, ka_detay))
    sonuc.problemler.extend(kalem_detay_problemleri(kd_ozet, kd_kirilim))
    sonuc.problemler.extend(fatura_detay_problemleri(fd_ozet, fd_detay))
    sonuc.problem_detay = ornek_problem_detay()
    oncelik_sira = {"YUKSEK": 0, "ORTA": 1, "DUSUK": 2}
    sonuc.problemler.sort(key=lambda p: oncelik_sira.get(p["oncelik"], 9))
    sonuc.uyarilar = ["Bu rapor ORNEK veri ile üretilmiştir — gerçek analiz için Windows sunucusunda çalıştırın."]
    return sonuc


if __name__ == "__main__":
    import sys

    if "--ornek" in sys.argv:
        analiz = ornek_analiz_sonucu()
        print(f"[ORNEK] Dönem: {analiz.ozet['donem']}")
        print(f"  Yük: {analiz.ozet['yuk_sayisi']:,} | Sevk: {analiz.ozet['sevk_sayisi']:,}")
        print(f"  Satış geliri: {analiz.ozet['toplam_satis_geliri']:,.2f} TL")
        print(f"  Problem sayısı: {len(analiz.problemler)}")
    else:
        analiz = kpi_analizi_yap()
        print(f"Dönem: {analiz.ozet['donem']}")
        print(f"  Yük: {analiz.ozet['yuk_sayisi']:,} | Sevk: {analiz.ozet['sevk_sayisi']:,}")
        print(f"  Satış geliri: {analiz.ozet['toplam_satis_geliri']:,.2f} TL")
        print(f"  Fatura bağlama: %{analiz.ozet['fatura_baglama_orani_yuzde']}")
        print(f"  Tespit edilen problem: {len(analiz.problemler)}")
        for p in analiz.problemler:
            print(f"    [{p['oncelik']}] {p['baslik']}")

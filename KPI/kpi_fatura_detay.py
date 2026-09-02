"""
Sevk / yük fatura detay — Uyumsoft fatura bağlantı raporu SQL'i ile uyumlu.

Operasyon kodu INVD_EXPENSE.EXPENSE_CODE üzerinden okunur (rapor SQL'indeki
LMSD_L_OP_DEFINITION yerine — otomasyon/KPI ile aynı doğrulanmış kaynak).
"""
from __future__ import annotations

from datetime import datetime
from typing import Any

import ayarlar
from oracle_baglanti import satir_limit_sql, tablo_var_mi


def _sk_filtre_parcasi() -> tuple[str, str]:
    joins: list[str] = []
    wheres: list[str] = []
    if getattr(ayarlar, "CO_CODE", None):
        joins.append("JOIN GNLD_COMPANY CO ON CO.CO_ID = SE.CO_ID")
        wheres.append("CO.CO_CODE = :co_code")
    if getattr(ayarlar, "BRANCH_CODE", None):
        joins.append("JOIN GNLD_BRANCH BR ON BR.BRANCH_ID = SE.BRANCH_ID")
        wheres.append("BR.BRANCH_CODE = :branch_code")
    return "\n        ".join(joins), (" AND " + " AND ".join(wheres)) if wheres else ""


def _yk_filtre_parcasi() -> tuple[str, str]:
    joins: list[str] = []
    wheres: list[str] = []
    if getattr(ayarlar, "CO_CODE", None):
        joins.append("JOIN GNLD_COMPANY CO ON CO.CO_ID = Y.CO_ID")
        wheres.append("CO.CO_CODE = :co_code")
    if getattr(ayarlar, "BRANCH_CODE", None):
        joins.append("JOIN GNLD_BRANCH BR ON BR.BRANCH_ID = Y.BRANCH_ID")
        wheres.append("BR.BRANCH_CODE = :branch_code")
    if getattr(ayarlar, "KPI_KAPI_KAPI_HARIC", True):
        wheres.append("NVL(Y.IS_DOOR_TO_DOOR, 0) = 0")
    return "\n        ".join(joins), (" AND " + " AND ".join(wheres)) if wheres else ""


def fatura_detay_semasi_hazir() -> tuple[bool, str]:
    if not tablo_var_mi("LMST_L_TRANS_OP_DETAIL"):
        return False, "LMST_L_TRANS_OP_DETAIL tablosu bulunamadı."
    if not tablo_var_mi("LMST_L_GOODS_OP_DET"):
        return False, "LMST_L_GOODS_OP_DET tablosu bulunamadı."
    if not tablo_var_mi("PSMT_INVOICE_M"):
        return False, "PSMT_INVOICE_M tablosu bulunamadı."
    return True, ""


def _detay_sql(limit: int) -> str:
    sk_join, sk_where = _sk_filtre_parcasi()
    yk_join, yk_where = _yk_filtre_parcasi()
    return satir_limit_sql(
        f"""
    SELECT 'Sevk' TIP,
           SE.TRANSPORT_NO SEVK_YUK_NO,
           SE.DOC_DATE BELGE_TARIHI,
           F.DOC_NO FATURA_NO,
           F.DOC_DATE FATURA_TARIHI,
           HK.EXPENSE_CODE OPERASYON_KODU,
           F.INVOICE_M_ID FATURA_ID,
           SE.TRANSPORT_ID KAYIT_ID
    FROM LMST_L_TRANS_OP_DETAIL S
    LEFT JOIN PSMT_INVOICE_M F ON F.INVOICE_M_ID = S.INVOICE_M_ID
    LEFT JOIN INVD_EXPENSE HK ON HK.EXPENSE_ID = S.OPERATION_ID
    LEFT JOIN LMST_L_TRANSPORT SE ON SE.TRANSPORT_ID = S.TRANSPORT_ID
    {sk_join}
    WHERE SE.DOC_DATE BETWEEN TO_DATE(:bas, 'DD.MM.YYYY') AND TO_DATE(:bit, 'DD.MM.YYYY')
    {sk_where}
    UNION ALL
    SELECT 'Yük' TIP,
           Y.REFERENCE_NO SEVK_YUK_NO,
           Y.DOC_DATE BELGE_TARIHI,
           F.DOC_NO FATURA_NO,
           F.DOC_DATE FATURA_TARIHI,
           HK.EXPENSE_CODE OPERASYON_KODU,
           F.INVOICE_M_ID FATURA_ID,
           Y.GOODS_ID KAYIT_ID
    FROM LMST_L_GOODS_OP_DET G
    LEFT JOIN PSMT_INVOICE_M F ON F.INVOICE_M_ID = G.INVOICE_M_ID
    LEFT JOIN INVD_EXPENSE HK ON HK.EXPENSE_ID = G.OPERATION_ID
    LEFT JOIN LMST_L_GOODS Y ON Y.GOODS_ID = G.GOODS_ID
    {yk_join}
    WHERE Y.DOC_DATE BETWEEN TO_DATE(:bas, 'DD.MM.YYYY') AND TO_DATE(:bit, 'DD.MM.YYYY')
    {yk_where}
ORDER BY BELGE_TARIHI DESC, SEVK_YUK_NO, TIP
""",
        limit,
    )


def fatura_detay_getir(cursor, bas: str, bit: str, bind: dict) -> tuple[list[dict], bool]:
    ok, _ = fatura_detay_semasi_hazir()
    if not ok:
        return [], False

    limit = int(getattr(ayarlar, "KPI_FATURA_DETAY_LIMIT", 10000))
    cursor.execute(_detay_sql(limit), bind)
    sutunlar = [c[0] for c in cursor.description]
    satirlar = [dict(zip(sutunlar, satir)) for satir in cursor.fetchall()]
    return satirlar, len(satirlar) >= limit


def _tarih_gun_farki(belge_tarihi, fatura_tarihi) -> float | None:
    if belge_tarihi is None or fatura_tarihi is None:
        return None
    try:
        belge = belge_tarihi.date() if isinstance(belge_tarihi, datetime) else belge_tarihi
        fatura = fatura_tarihi.date() if isinstance(fatura_tarihi, datetime) else fatura_tarihi
        return (fatura - belge).days
    except (TypeError, AttributeError):
        return None


def fatura_detay_ozet_hesapla(
    detay: list[dict], mevcut: bool = True, limit_asildi: bool = False
) -> dict[str, Any]:
    if not mevcut:
        return {"mevcut": False, "mesaj": "Fatura detay tabloları bulunamadı."}

    if not detay:
        return {
            "mevcut": True,
            "toplam_satir": 0,
            "sevk_satir": 0,
            "yuk_satir": 0,
            "faturali_satir": 0,
            "faturasiz_satir": 0,
            "fatura_baglama_orani_yuzde": 0,
            "ort_fatura_gecikme_gun": None,
            "limit_asildi": limit_asildi,
        }

    sevk = yuk = faturali = faturasiz = 0
    gecikmeler: list[float] = []

    for satir in detay:
        if satir.get("TIP") == "Sevk":
            sevk += 1
        else:
            yuk += 1
        if satir.get("FATURA_NO"):
            faturali += 1
            gecikme = _tarih_gun_farki(satir.get("BELGE_TARIHI"), satir.get("FATURA_TARIHI"))
            if gecikme is not None:
                gecikmeler.append(gecikme)
        else:
            faturasiz += 1

    toplam = len(detay)
    return {
        "mevcut": True,
        "toplam_satir": toplam,
        "sevk_satir": sevk,
        "yuk_satir": yuk,
        "faturali_satir": faturali,
        "faturasiz_satir": faturasiz,
        "fatura_baglama_orani_yuzde": round(faturali / toplam * 100, 1) if toplam else 0,
        "ort_fatura_gecikme_gun": round(sum(gecikmeler) / len(gecikmeler), 1) if gecikmeler else None,
        "limit_asildi": limit_asildi,
    }


def faturasiz_kalemler(detay: list[dict]) -> list[dict]:
    return [s for s in detay if not s.get("FATURA_NO")]


def fatura_detay_problemleri(ozet: dict, detay: list[dict]) -> list[dict]:
    problemler = []
    if not ozet.get("mevcut"):
        return problemler

    oran = ozet.get("fatura_baglama_orani_yuzde", 100)
    if ozet.get("toplam_satir", 0) > 0 and oran < 95:
        faturasiz = faturasiz_kalemler(detay)
        ornek = ", ".join(
            f"{s.get('SEVK_YUK_NO', '?')} ({s.get('OPERASYON_KODU', '-')})"
            for s in faturasiz[:5]
        )
        problemler.append({
            "oncelik": "YUKSEK",
            "kategori": "Fatura Detay",
            "baslik": "Kalem bazında fatura bağlama oranı düşük",
            "detay": (
                f"Sevk/yük kalemlerinin %{oran}'i faturalı. "
                f"Faturasız: {ozet['faturasiz_satir']} satır. Örnek: {ornek}"
            ),
            "aksiyon": "Fatura Detay ve Faturasız Kalemler sayfalarından eksik faturaları kapatın.",
        })

    gecikme = ozet.get("ort_fatura_gecikme_gun")
    if gecikme is not None and gecikme > 7:
        problemler.append({
            "oncelik": "ORTA",
            "kategori": "Fatura Detay",
            "baslik": "Fatura kesim gecikmesi yüksek",
            "detay": f"Belge tarihi ile fatura tarihi arası ortalama {gecikme} gün.",
            "aksiyon": "Geciken projeleri fatura ekibiyle birlikte önceliklendirin.",
        })

    erken = [
        s for s in detay
        if s.get("FATURA_NO")
        and (_t := _tarih_gun_farki(s.get("BELGE_TARIHI"), s.get("FATURA_TARIHI"))) is not None
        and _t < 0
    ]
    if len(erken) > 10:
        problemler.append({
            "oncelik": "ORTA",
            "kategori": "Fatura Detay",
            "baslik": "Fatura tarihi belge tarihinden önce görünen kalemler",
            "detay": f"{len(erken)} satırda fatura tarihi sevk/yük tarihinden önce.",
            "aksiyon": "Veri giriş hatası veya ön fatura senaryosu olabilir — ERP kayıtlarını doğrulayın.",
        })

    if ozet.get("limit_asildi"):
        limit = getattr(ayarlar, "KPI_FATURA_DETAY_LIMIT", 10000)
        problemler.append({
            "oncelik": "ORTA",
            "kategori": "Fatura Detay",
            "baslik": "Fatura detay satır limiti aşıldı",
            "detay": f"Excel'e en fazla {limit} satır aktarıldı.",
            "aksiyon": "Dönemi daraltın veya ayarlar.py içinde KPI_FATURA_DETAY_LIMIT artırın.",
        })

    return problemler


def ornek_fatura_detay_verisi() -> tuple[dict, list[dict], list[dict]]:
    detay = [
        {
            "TIP": "Sevk", "SEVK_YUK_NO": "S-2001", "BELGE_TARIHI": "05.01.2026",
            "FATURA_NO": "ALF2026001", "FATURA_TARIHI": "10.01.2026",
            "OPERASYON_KODU": "NAVLUN", "FATURA_ID": 1001, "KAYIT_ID": 501,
        },
        {
            "TIP": "Sevk", "SEVK_YUK_NO": "S-2002", "BELGE_TARIHI": "08.01.2026",
            "FATURA_NO": None, "FATURA_TARIHI": None,
            "OPERASYON_KODU": "NAVLUN", "FATURA_ID": None, "KAYIT_ID": 502,
        },
        {
            "TIP": "Yük", "SEVK_YUK_NO": "Y-6001", "BELGE_TARIHI": "06.01.2026",
            "FATURA_NO": "SAT2026001", "FATURA_TARIHI": "12.01.2026",
            "OPERASYON_KODU": "NAVLUN", "FATURA_ID": 2001, "KAYIT_ID": 6001,
        },
        {
            "TIP": "Yük", "SEVK_YUK_NO": "Y-6002", "BELGE_TARIHI": "09.01.2026",
            "FATURA_NO": None, "FATURA_TARIHI": None,
            "OPERASYON_KODU": "UĞRAMA", "FATURA_ID": None, "KAYIT_ID": 6002,
        },
    ]
    ozet = fatura_detay_ozet_hesapla(detay)
    faturasiz = faturasiz_kalemler(detay)
    return ozet, detay, faturasiz

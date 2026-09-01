"""
Kiralık araç maliyet KPI — Uyumsoft tedarikçi hakediş raporu SQL'i ile uyumlu.

Tablolar [DOĞRULANMIŞ — kullanıcı rapor SQL'i]:
  LMST_SUP_PAYOFF_M / LMST_SUP_PAYOFF_OPDET / LMST_SUP_PAYOFF_T_D
  FLMD_VEHICLE, FIND_ENTITY, LMST_L_TRANSPORT, LMST_L_TRANS_GOODS_DETAIL
"""
from __future__ import annotations

from collections import defaultdict
from decimal import Decimal
from typing import Any

import ayarlar
from oracle_baglanti import tablo_var_mi


def _decimal(deger) -> Decimal:
    if deger is None:
        return Decimal("0")
    return Decimal(str(deger))


def _payoff_filtre_parcasi() -> tuple[str, str]:
    """PM alias'lı sorgular için firma/şube join ve WHERE parçası."""
    joins: list[str] = []
    wheres: list[str] = []
    if getattr(ayarlar, "CO_CODE", None):
        joins.append("LEFT JOIN GNLD_COMPANY CO ON CO.CO_ID = PM.CO_ID")
        wheres.append("CO.CO_CODE = :co_code")
    if getattr(ayarlar, "BRANCH_CODE", None):
        joins.append("LEFT JOIN GNLD_BRANCH BR ON BR.BRANCH_ID = PM.BRANCH_ID")
        wheres.append("BR.BRANCH_CODE = :branch_code")
    join_sql = "\n".join(joins)
    ek_filtre = (" AND " + " AND ".join(wheres)) if wheres else ""
    return join_sql, ek_filtre


def _sk_filtre_parcasi() -> tuple[str, str]:
    """Sevk (SK) alt sorguları için firma/şube filtresi."""
    joins: list[str] = []
    wheres: list[str] = []
    if getattr(ayarlar, "CO_CODE", None):
        joins.append("LEFT JOIN GNLD_COMPANY CO ON CO.CO_ID = SK.CO_ID")
        wheres.append("CO.CO_CODE = :co_code")
    if getattr(ayarlar, "BRANCH_CODE", None):
        joins.append("LEFT JOIN GNLD_BRANCH BR ON BR.BRANCH_ID = SK.BRANCH_ID")
        wheres.append("BR.BRANCH_CODE = :branch_code")
    join_sql = "\n".join(joins)
    ek_filtre = (" AND " + " AND ".join(wheres)) if wheres else ""
    return join_sql, ek_filtre


def kiralk_arac_semasi_hazir() -> tuple[bool, str]:
    if not tablo_var_mi("LMST_SUP_PAYOFF_M"):
        return False, "LMST_SUP_PAYOFF_M tablosu bulunamadı — kiralık araç KPI atlandı."
    if not tablo_var_mi("LMST_SUP_PAYOFF_OPDET"):
        return False, "LMST_SUP_PAYOFF_OPDET tablosu bulunamadı."
    return True, ""


def _detay_sql() -> str:
    pm_join, pm_filtre = _payoff_filtre_parcasi()
    sk_join, sk_filtre = _sk_filtre_parcasi()
    return f"""
WITH DOSYA_KALEM AS (
    SELECT POT.SUP_PAYOFF_ID,
           CASE POT.PURCHASE_SALES_TYPE
             WHEN 1 THEN 'Alış'
             WHEN 2 THEN 'Satış'
             WHEN 3 THEN 'Satış İade'
             WHEN 4 THEN 'Alış İade'
           END ALIS_SATIS,
           HK.EXPENSE_CODE OPERASYON_KODU,
           HK.DESCRIPTION OPERASYON_ADI,
           POT.AMT_TRA TUTAR
    FROM LMST_SUP_PAYOFF_OPDET POT
    LEFT JOIN LMST_SUP_PAYOFF_M PM ON PM.SUP_PAYOFF_ID = POT.SUP_PAYOFF_ID
    LEFT JOIN INVD_EXPENSE HK ON HK.EXPENSE_ID = POT.OPERATION_ID
    {pm_join}
    WHERE PM.DOC_DATE BETWEEN TO_DATE(:bas, 'DD.MM.YYYY') AND TO_DATE(:bit, 'DD.MM.YYYY')
    {pm_filtre}
),
DOSYA_KALEM_TOPLAM AS (
    SELECT POT.SUP_PAYOFF_ID,
           CASE POT.PURCHASE_SALES_TYPE
             WHEN 1 THEN 'Alış'
             WHEN 2 THEN 'Satış'
             WHEN 3 THEN 'Satış İade'
             WHEN 4 THEN 'Alış İade'
           END ALIS_SATIS,
           SUM(POT.AMT_TRA) TUTAR
    FROM LMST_SUP_PAYOFF_OPDET POT
    LEFT JOIN LMST_SUP_PAYOFF_M PM ON PM.SUP_PAYOFF_ID = POT.SUP_PAYOFF_ID
    LEFT JOIN INVD_EXPENSE HK ON HK.EXPENSE_ID = POT.OPERATION_ID
    {pm_join}
    WHERE HK.EXPENSE_CODE NOT IN (
            'ALDIĞI MAZOT TL', 'OTOBAN-KÖPRÜ-FERİBOT', 'NAVLUN', 'YAKIT', 'IZLOG OGS'
          )
      AND PM.DOC_DATE BETWEEN TO_DATE(:bas, 'DD.MM.YYYY') AND TO_DATE(:bit, 'DD.MM.YYYY')
    {pm_filtre}
    GROUP BY POT.SUP_PAYOFF_ID, POT.PURCHASE_SALES_TYPE
),
DOSYA_KALEM_YAKIT AS (
    SELECT POT.SUP_PAYOFF_ID,
           CASE POT.PURCHASE_SALES_TYPE
             WHEN 1 THEN 'Alış'
             WHEN 2 THEN 'Satış'
             WHEN 3 THEN 'Satış İade'
             WHEN 4 THEN 'Alış İade'
           END ALIS_SATIS,
           SUM(POT.AMT_TRA) TUTAR
    FROM LMST_SUP_PAYOFF_OPDET POT
    LEFT JOIN LMST_SUP_PAYOFF_M PM ON PM.SUP_PAYOFF_ID = POT.SUP_PAYOFF_ID
    LEFT JOIN INVD_EXPENSE HK ON HK.EXPENSE_ID = POT.OPERATION_ID
    {pm_join}
    WHERE HK.EXPENSE_CODE IN ('YAKIT')
      AND PM.DOC_DATE BETWEEN TO_DATE(:bas, 'DD.MM.YYYY') AND TO_DATE(:bit, 'DD.MM.YYYY')
    {pm_filtre}
    GROUP BY POT.SUP_PAYOFF_ID, POT.PURCHASE_SALES_TYPE
),
DOSYA_YUK_KALEMLERI AS (
    SELECT PM.SUP_PAYOFF_ID,
           SUM(NVL(CASE WHEN GOD.PURCHASE_SALES_TYPE IN (2, 4) THEN GOD.AMT_TRA
                        ELSE -1 * GOD.AMT_TRA END, 0)) YUK_SATIS_TOPLAMI
    FROM LMST_SUP_PAYOFF_T_D TD
    LEFT JOIN LMST_SUP_PAYOFF_M PM ON PM.SUP_PAYOFF_ID = TD.SUP_PAYOFF_ID
    LEFT JOIN LMST_L_TRANSPORT SK ON SK.TRANSPORT_ID = TD.TRANSPORT_ID
    LEFT JOIN LMST_L_TRANS_GOODS_DETAIL TGD ON TGD.TRANSPORT_ID = SK.TRANSPORT_ID
    LEFT JOIN LMST_L_GOODS YK ON YK.GOODS_ID = TGD.GOODS_ID
    LEFT JOIN LMST_L_GOODS_OP_DET GOD ON GOD.GOODS_ID = YK.GOODS_ID
    {pm_join}
    WHERE PM.DOC_DATE BETWEEN TO_DATE(:bas, 'DD.MM.YYYY') AND TO_DATE(:bit, 'DD.MM.YYYY')
    {pm_filtre}
    GROUP BY PM.SUP_PAYOFF_ID
),
DOSYA_YUK_KALEMLERI_ELDEN AS (
    SELECT PM.SUP_PAYOFF_ID,
           SUM(NVL(CASE WHEN GOD.PURCHASE_SALES_TYPE IN (2, 4) THEN GOD.AMT_TRA
                        ELSE -1 * GOD.AMT_TRA END, 0)) YUK_SATIS_TOPLAMI
    FROM LMST_SUP_PAYOFF_T_D TD
    LEFT JOIN LMST_SUP_PAYOFF_M PM ON PM.SUP_PAYOFF_ID = TD.SUP_PAYOFF_ID
    LEFT JOIN LMST_L_TRANSPORT SK ON SK.TRANSPORT_ID = TD.TRANSPORT_ID
    LEFT JOIN LMST_L_TRANS_GOODS_DETAIL TGD ON TGD.TRANSPORT_ID = SK.TRANSPORT_ID
    LEFT JOIN LMST_L_GOODS YK ON YK.GOODS_ID = TGD.GOODS_ID
    LEFT JOIN LMST_L_GOODS_OP_DET GOD ON GOD.GOODS_ID = YK.GOODS_ID
    {pm_join}
    WHERE GOD.OPERATION_ID = 867
      AND PM.DOC_DATE BETWEEN TO_DATE(:bas, 'DD.MM.YYYY') AND TO_DATE(:bit, 'DD.MM.YYYY')
    {pm_filtre}
    GROUP BY PM.SUP_PAYOFF_ID
)
SELECT VH.VEHICLE_CODE ARAC_KODU,
       TTYPE.DESCRIPTION ARAC_TIPI,
       LGT.GOODS_TYPE_CODE GOREV_YERI,
       CK.ENTITY_CODE CARI_KODU,
       CK.ENTITY_NAME CARI_ADI,
       PM.CASE_CODE DOSYA_NO,
       PM.DOC_DATE DOSYA_TARIHI,
       CASE PM.CASE_STATUS WHEN 1 THEN 'Açık' WHEN 2 THEN 'Kapalı' END DOSYA_DURUM,
       PM.PERIOD_START_DATE BASLANGIC_TARIHI,
       PM.PERIOD_END_DATE BITIS_TARIHI,
       NVL(TK.TOPLAM_KM, 0) TOPLAM_KM,
       NVL(VRD.RENT_AMT_TRA, 0) AYLIK_KIRA_TUTARI,
       NVL(VRD.FUEL_RATE, 0) AYLIK_YAKIT_ORANI,
       NVL(D3.TUTAR, 0) HAKEDIS_KIRA_TUTARI,
       NVL(D1.TUTAR, 0) ALDIGI_YAKIT_TUTARI,
       NVL(D2.TUTAR, 0) OTOBAN_KOPRU_VS,
       NVL(D4.TUTAR, 0) IZLOG_OGS,
       NVL(D5.YUK_SATIS_TOPLAMI, 0) ELDEN,
       NVL(DY1.TUTAR, 0) YAKIT_FARK_ARTI,
       NVL(DY2.TUTAR, 0) YAKIT_FARK_EKSI,
       NVL(DT1.TUTAR, 0) ALIS_DIGER,
       NVL(DT2.TUTAR, 0) ALIS_IADE_DIGER,
       (NVL(D3.TUTAR, 0) + NVL(D2.TUTAR, 0) + NVL(DY1.TUTAR, 0) + NVL(DT1.TUTAR, 0))
           TEDARIKCI_FATURA_TOPLAM,
       (NVL(DY2.TUTAR, 0) + NVL(DT2.TUTAR, 0)) IZLOG_IADE_TOPLAM,
       (NVL(D3.TUTAR, 0) + NVL(D2.TUTAR, 0) + NVL(DY1.TUTAR, 0) + NVL(DT1.TUTAR, 0))
           - (NVL(DY2.TUTAR, 0) + NVL(DT2.TUTAR, 0)) ODENECEK_TUTAR,
       ((NVL(D3.TUTAR, 0) + NVL(D2.TUTAR, 0) + NVL(DY1.TUTAR, 0) + NVL(DT1.TUTAR, 0))
           - (NVL(DY2.TUTAR, 0) + NVL(DT2.TUTAR, 0)))
           + NVL(D1.TUTAR, 0) + NVL(D4.TUTAR, 0) MALIYETLER_TOPLAMI,
       NVL(DYK.YUK_SATIS_TOPLAMI, 0) TOPLAM_SATIS,
       NVL(DYK.YUK_SATIS_TOPLAMI, 0)
           - (((NVL(D3.TUTAR, 0) + NVL(D2.TUTAR, 0) + NVL(DY1.TUTAR, 0) + NVL(DT1.TUTAR, 0))
               - (NVL(DY2.TUTAR, 0) + NVL(DT2.TUTAR, 0)))
               + NVL(D1.TUTAR, 0) + NVL(D4.TUTAR, 0)) KAR_ZARAR,
       PM.SUP_PAYOFF_ID DOSYA_ID
FROM LMST_SUP_PAYOFF_M PM
LEFT JOIN FIND_ENTITY CK ON CK.ENTITY_ID = PM.L_ENTITY_ID
LEFT JOIN FLMD_VEHICLE VH ON VH.VEHICLE_ID = PM.VEHICLE_ID
LEFT JOIN FLMD_VHC_GOODTYPE_DETAIL VGD ON VGD.VEHICLE_ID = VH.VEHICLE_ID
LEFT JOIN LMSD_L_GOODS_TYPE LGT ON LGT.GOODS_TYPE_ID = VGD.GOODS_TYPE_ID
LEFT JOIN FLMD_L_TRAILER_TYPE TTYPE ON TTYPE.TRAILER_TYPE_ID = VH.TRAILER_TYPE_ID
LEFT JOIN FLMD_VHC_RENT_DETAIL VRD
    ON VRD.VEHICLE_ID = VH.VEHICLE_ID
   AND VRD.START_DATE <= PM.PERIOD_END_DATE
   AND VRD.END_DATE >= PM.PERIOD_END_DATE
LEFT JOIN DOSYA_KALEM D1 ON D1.SUP_PAYOFF_ID = PM.SUP_PAYOFF_ID AND D1.OPERASYON_KODU = 'ALDIĞI MAZOT TL'
LEFT JOIN DOSYA_KALEM D2 ON D2.SUP_PAYOFF_ID = PM.SUP_PAYOFF_ID AND D2.OPERASYON_KODU = 'OTOBAN-KÖPRÜ-FERİBOT'
LEFT JOIN DOSYA_KALEM D3 ON D3.SUP_PAYOFF_ID = PM.SUP_PAYOFF_ID AND D3.OPERASYON_KODU = 'NAVLUN'
LEFT JOIN DOSYA_KALEM D4 ON D4.SUP_PAYOFF_ID = PM.SUP_PAYOFF_ID AND D4.OPERASYON_KODU = 'IZLOG OGS'
LEFT JOIN DOSYA_KALEM_YAKIT DY1 ON DY1.SUP_PAYOFF_ID = PM.SUP_PAYOFF_ID AND DY1.ALIS_SATIS = 'Alış'
LEFT JOIN DOSYA_KALEM_YAKIT DY2 ON DY2.SUP_PAYOFF_ID = PM.SUP_PAYOFF_ID AND DY2.ALIS_SATIS = 'Alış İade'
LEFT JOIN DOSYA_KALEM_TOPLAM DT1 ON DT1.SUP_PAYOFF_ID = PM.SUP_PAYOFF_ID AND DT1.ALIS_SATIS = 'Alış'
LEFT JOIN DOSYA_KALEM_TOPLAM DT2 ON DT2.SUP_PAYOFF_ID = PM.SUP_PAYOFF_ID AND DT2.ALIS_SATIS = 'Alış İade'
LEFT JOIN DOSYA_YUK_KALEMLERI DYK ON DYK.SUP_PAYOFF_ID = PM.SUP_PAYOFF_ID
LEFT JOIN DOSYA_YUK_KALEMLERI_ELDEN D5 ON D5.SUP_PAYOFF_ID = PM.SUP_PAYOFF_ID
LEFT JOIN (
    SELECT TD.SUP_PAYOFF_ID, SUM(NVL(SK.END_KM, 0)) TOPLAM_KM
    FROM LMST_SUP_PAYOFF_T_D TD
    LEFT JOIN LMST_L_TRANSPORT SK ON SK.TRANSPORT_ID = TD.TRANSPORT_ID
    {sk_join}
    WHERE SK.DOC_DATE BETWEEN TO_DATE(:bas, 'DD.MM.YYYY') AND TO_DATE(:bit, 'DD.MM.YYYY')
    {sk_filtre}
    GROUP BY TD.SUP_PAYOFF_ID
) TK ON TK.SUP_PAYOFF_ID = PM.SUP_PAYOFF_ID
{pm_join}
WHERE PM.DOC_DATE BETWEEN TO_DATE(:bas, 'DD.MM.YYYY') AND TO_DATE(:bit, 'DD.MM.YYYY')
{pm_filtre}
ORDER BY PM.DOC_DATE DESC, PM.CASE_CODE
"""


def kiralk_arac_detay_getir(cursor, bas: str, bit: str, bind: dict) -> list[dict]:
    """Dönem içindeki tüm kiralık araç hakediş dosyalarını döner."""
    ok, mesaj = kiralk_arac_semasi_hazir()
    if not ok:
        return []

    cursor.execute(_detay_sql(), bind)
    sutunlar = [c[0] for c in cursor.description]
    return [dict(zip(sutunlar, satir)) for satir in cursor.fetchall()]


def kiralk_arac_ozet_hesapla(detay: list[dict], mevcut: bool = True) -> dict[str, Any]:
    if not mevcut:
        return {"mevcut": False, "mesaj": "Kiralık araç tabloları bulunamadı."}

    if not detay:
        return {
            "mevcut": True,
            "dosya_sayisi": 0,
            "acik_dosya_sayisi": 0,
            "toplam_maliyet": Decimal("0"),
            "toplam_satis": Decimal("0"),
            "toplam_kar_zarar": Decimal("0"),
            "zararli_dosya_sayisi": 0,
            "kar_orani_yuzde": 0,
            "mesaj": "Dönemde kiralık araç hakediş dosyası bulunamadı.",
        }

    toplam_maliyet = sum(_decimal(r.get("MALIYETLER_TOPLAMI")) for r in detay)
    toplam_satis = sum(_decimal(r.get("TOPLAM_SATIS")) for r in detay)
    toplam_kz = sum(_decimal(r.get("KAR_ZARAR")) for r in detay)
    zararli = sum(1 for r in detay if _decimal(r.get("KAR_ZARAR")) < 0)
    acik = sum(1 for r in detay if r.get("DOSYA_DURUM") == "Açık")

    return {
        "mevcut": True,
        "dosya_sayisi": len(detay),
        "acik_dosya_sayisi": acik,
        "toplam_maliyet": toplam_maliyet,
        "toplam_satis": toplam_satis,
        "toplam_kar_zarar": toplam_kz,
        "zararli_dosya_sayisi": zararli,
        "kar_orani_yuzde": round(float(toplam_kz / toplam_satis * 100), 1) if toplam_satis else 0,
    }


def kiralk_arac_cari_ozet_hesapla(detay: list[dict]) -> list[dict]:
    """Tedarikçi (cari) bazında maliyet/kar özeti."""
    gruplar: dict[str, dict] = defaultdict(lambda: {
        "CARI_KODU": "",
        "CARI_ADI": "",
        "DOSYA_SAYISI": 0,
        "TOPLAM_MALIYET": Decimal("0"),
        "TOPLAM_SATIS": Decimal("0"),
        "KAR_ZARAR": Decimal("0"),
    })

    for satir in detay:
        anahtar = satir.get("CARI_KODU") or satir.get("CARI_ADI") or "?"
        g = gruplar[anahtar]
        g["CARI_KODU"] = satir.get("CARI_KODU") or ""
        g["CARI_ADI"] = satir.get("CARI_ADI") or ""
        g["DOSYA_SAYISI"] += 1
        g["TOPLAM_MALIYET"] += _decimal(satir.get("MALIYETLER_TOPLAMI"))
        g["TOPLAM_SATIS"] += _decimal(satir.get("TOPLAM_SATIS"))
        g["KAR_ZARAR"] += _decimal(satir.get("KAR_ZARAR"))

    sonuc = sorted(gruplar.values(), key=lambda x: x["KAR_ZARAR"])
    return sonuc


def kiralk_arac_problemleri(ozet: dict, detay: list[dict]) -> list[dict]:
    problemler = []
    if not ozet.get("mevcut"):
        return problemler

    if ozet.get("zararli_dosya_sayisi", 0) > 0:
        ornek = sorted(detay, key=lambda r: _decimal(r.get("KAR_ZARAR")))[:3]
        ornek_metin = ", ".join(
            f"{r.get('DOSYA_NO', '?')} ({_decimal(r.get('KAR_ZARAR')):,.0f} TL)"
            for r in ornek
        )
        problemler.append(
            {
                "oncelik": "YUKSEK",
                "kategori": "Kiralık Araç",
                "baslik": f"Zararlı hakediş dosyaları ({ozet['zararli_dosya_sayisi']} adet)",
                "detay": (
                    f"Toplam kar/zarar: {ozet['toplam_kar_zarar']:,.2f} TL. "
                    f"En düşük: {ornek_metin}"
                ),
                "aksiyon": "Zararlı dosyaları plaka/cari bazında inceleyin; kira ve yakıt oranlarını gözden geçirin.",
            }
        )

    if ozet.get("acik_dosya_sayisi", 0) > 5:
        problemler.append(
            {
                "oncelik": "ORTA",
                "kategori": "Kiralık Araç",
                "baslik": f"Açık hakediş dosyası yüksek ({ozet['acik_dosya_sayisi']} adet)",
                "detay": "Kapalı olmayan dosyalar nakit akışı ve mutabakat riski oluşturur.",
                "aksiyon": "Açık dosyaları kapatma takvimini operasyon ekibiyle netleştirin.",
            }
        )

    if ozet.get("toplam_satis") and ozet.get("kar_orani_yuzde", 0) < 5:
        problemler.append(
            {
                "oncelik": "YUKSEK",
                "kategori": "Kiralık Araç",
                "baslik": "Kiralık filo karlılığı düşük",
                "detay": (
                    f"Filo kar oranı: %{ozet['kar_orani_yuzde']} "
                    f"(satış {ozet['toplam_satis']:,.0f} TL − maliyet {ozet['toplam_maliyet']:,.0f} TL)"
                ),
                "aksiyon": "Düşük marjlı araç/cari sözleşmelerini yeniden değerlendirin.",
            }
        )

    return problemler


def ornek_kiralk_arac_verisi() -> tuple[dict, list[dict], list[dict]]:
    detay = [
        {
            "ARAC_KODU": "34ABC123",
            "ARAC_TIPI": "TIR",
            "GOREV_YERI": "SAKARYA",
            "CARI_KODU": "TED001",
            "CARI_ADI": "Örnek Taşımacılık A.Ş.",
            "DOSYA_NO": "HK-2026-001",
            "DOSYA_TARIHI": "15.01.2026",
            "DOSYA_DURUM": "Kapalı",
            "TOPLAM_KM": 4200,
            "HAKEDIS_KIRA_TUTARI": Decimal("85000"),
            "ALDIGI_YAKIT_TUTARI": Decimal("32000"),
            "OTOBAN_KOPRU_VS": Decimal("4500"),
            "MALIYETLER_TOPLAMI": Decimal("121500"),
            "TOPLAM_SATIS": Decimal("138000"),
            "KAR_ZARAR": Decimal("16500"),
        },
        {
            "ARAC_KODU": "06XYZ789",
            "ARAC_TIPI": "KAMYON",
            "GOREV_YERI": "KURUYÜK",
            "CARI_KODU": "TED002",
            "CARI_ADI": "Demo Lojistik Ltd.",
            "DOSYA_NO": "HK-2026-002",
            "DOSYA_TARIHI": "20.01.2026",
            "DOSYA_DURUM": "Açık",
            "TOPLAM_KM": 2800,
            "HAKEDIS_KIRA_TUTARI": Decimal("62000"),
            "ALDIGI_YAKIT_TUTARI": Decimal("41000"),
            "OTOBAN_KOPRU_VS": Decimal("2100"),
            "MALIYETLER_TOPLAMI": Decimal("105100"),
            "TOPLAM_SATIS": Decimal("98000"),
            "KAR_ZARAR": Decimal("-7100"),
        },
    ]
    ozet = kiralk_arac_ozet_hesapla(detay)
    cari = kiralk_arac_cari_ozet_hesapla(detay)
    return ozet, detay, cari

"""
İzlog Lojistik — Uyumsoft ERP KPI analiz modülü.

Otomasyon projesinden tamamen bağımsız; yalnızca KPI/ klasöründeki ayarlar.py
ve oracle_baglanti.py kullanılır.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime, timedelta
from decimal import Decimal
from typing import Any

import ayarlar
from oracle_baglanti import baglanti_yonet, tablo_kolonlari, tablo_var_mi


def _kolon_adayindan_bul(kolonlar: list[str], adaylar: list[str]) -> str | None:
    kolon_haritasi = {k.upper(): k for k in kolonlar}
    for aday in adaylar:
        if aday.upper() in kolon_haritasi:
            return kolon_haritasi[aday.upper()]
    return None


def _transport_semasini_coz(cursor) -> dict:
    """
    LMST_L_TRANSPORT tablosunun gerçek kolon adlarını keşfeder.
    İzlog ERP'de GOODS_ID değil L_GOODS_ID olabiliyor [DOĞRULANMIŞ — ORA-00904].
    """
    if not tablo_var_mi("LMST_L_TRANSPORT"):
        return {"mevcut": False}

    kolonlar = tablo_kolonlari(cursor, "LMST_L_TRANSPORT")

    elle = getattr(ayarlar, "KPI_SEVK_YUK_KOLONU", None)
    if elle and elle.upper() in {k.upper() for k in kolonlar}:
        yuk_kolon = next(k for k in kolonlar if k.upper() == elle.upper())
    else:
        yuk_kolon = _kolon_adayindan_bul(
            kolonlar,
            ["L_GOODS_ID", "GOODS_ID", "LG_GOODS_ID", "SOURCE_GOODS_ID", "SOURCE_L_GOODS_ID"],
        )
        if not yuk_kolon:
            for kolon in kolonlar:
                if "GOODS" in kolon.upper():
                    yuk_kolon = kolon
                    break

    tarih_kolon = _kolon_adayindan_bul(
        kolonlar,
        ["DOC_DATE", "DOCUMENT_DATE", "TRANSPORT_DATE", "TRANS_DATE"],
    )

    return {
        "mevcut": True,
        "yuk_kolon": yuk_kolon,
        "tarih_kolon": tarih_kolon,
        "kolonlar": kolonlar,
    }


def _sevk_sayisi_hesapla(cursor, bas, bit, transport_sema: dict) -> int:
    if not transport_sema.get("mevcut"):
        return 0

    yuk_kolon = transport_sema.get("yuk_kolon")
    if yuk_kolon:
        sql = f"""
            SELECT COUNT(DISTINCT T.TRANSPORT_ID) AS SEVK_SAYISI
            FROM LMST_L_TRANSPORT T
            JOIN LMST_L_GOODS YK ON YK.GOODS_ID = T.{yuk_kolon}
            WHERE YK.DOC_DATE BETWEEN TO_DATE(:bas, 'DD.MM.YYYY') AND TO_DATE(:bit, 'DD.MM.YYYY')
        """
        cursor.execute(sql, {"bas": bas, "bit": bit})
        return cursor.fetchone()[0]

    tarih_kolon = transport_sema.get("tarih_kolon")
    if tarih_kolon:
        sql = f"""
            SELECT COUNT(DISTINCT T.TRANSPORT_ID) AS SEVK_SAYISI
            FROM LMST_L_TRANSPORT T
            WHERE T.{tarih_kolon} BETWEEN TO_DATE(:bas, 'DD.MM.YYYY') AND TO_DATE(:bit, 'DD.MM.YYYY')
        """
        cursor.execute(sql, {"bas": bas, "bit": bit})
        return cursor.fetchone()[0]

    return 0


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
    problemler: list[dict] = field(default_factory=list)
    uyarilar: list[str] = field(default_factory=list)


def _satirlari_dict_yap(sutunlar, satirlar):
    return [dict(zip(sutunlar, satir)) for satir in satirlar]


def _temel_hacim_kpi(cursor, bas, bit, transport_sema: dict):
    """Yük ve sevk hacmi KPI'ları."""
    cursor.execute(
        """
        SELECT COUNT(*) AS YUK_SAYISI
        FROM LMST_L_GOODS
        WHERE DOC_DATE BETWEEN TO_DATE(:bas, 'DD.MM.YYYY') AND TO_DATE(:bit, 'DD.MM.YYYY')
        """,
        {"bas": bas, "bit": bit},
    )
    yuk_sayisi = cursor.fetchone()[0]

    sevk_sayisi = _sevk_sayisi_hesapla(cursor, bas, bit, transport_sema)

    sevk_orani = round(sevk_sayisi / yuk_sayisi * 100, 1) if yuk_sayisi else 0

    return {
        "yuk_sayisi": yuk_sayisi,
        "sevk_sayisi": sevk_sayisi,
        "sevk_orani_yuzde": sevk_orani,
        "yuk_basina_ortalama_gelir": Decimal("0"),  # gelir KPI'dan sonra güncellenir
    }


def _gelir_kpi(cursor, bas, bit):
    """Satış geliri ve operasyon/proje kırılımları."""
    cursor.execute(
        """
        SELECT
            NVL(SUM(OPDET.AMT), 0) AS TOPLAM_SATIS,
            COUNT(DISTINCT YK.GOODS_ID) AS GELIRLI_YUK_SAYISI,
            COUNT(*) AS SATIS_SATIR_SAYISI,
            NVL(SUM(CASE WHEN OPDET.INVOICE_M_ID IS NULL THEN OPDET.AMT ELSE 0 END), 0) AS FATURASIZ_TUTAR,
            NVL(SUM(CASE WHEN OPDET.INVOICE_M_ID IS NOT NULL THEN OPDET.AMT ELSE 0 END), 0) AS FATURALI_TUTAR,
            COUNT(CASE WHEN OPDET.INVOICE_M_ID IS NULL THEN 1 END) AS FATURASIZ_SATIR,
            COUNT(CASE WHEN OPDET.INVOICE_M_ID IS NOT NULL THEN 1 END) AS FATURALI_SATIR
        FROM LMST_L_GOODS_OP_DET OPDET
        JOIN LMST_L_GOODS YK ON YK.GOODS_ID = OPDET.GOODS_ID
        WHERE YK.DOC_DATE BETWEEN TO_DATE(:bas, 'DD.MM.YYYY') AND TO_DATE(:bit, 'DD.MM.YYYY')
          AND OPDET.PURCHASE_SALES_TYPE IN (2, 4)
        """,
        {"bas": bas, "bit": bit},
    )
    row = cursor.fetchone()
    sutunlar = [c[0] for c in cursor.description]
    veri = dict(zip(sutunlar, row))

    toplam_satis = _decimal(veri["TOPLAM_SATIS"])
    satir_sayisi = veri["SATIS_SATIR_SAYISI"] or 0
    faturasiz_satir = veri["FATURASIZ_SATIR"] or 0

    return {
        "toplam_satis_geliri": toplam_satis,
        "gelirli_yuk_sayisi": veri["GELIRLI_YUK_SAYISI"] or 0,
        "satis_satir_sayisi": satir_sayisi,
        "faturasiz_tutar": _decimal(veri["FATURASIZ_TUTAR"]),
        "faturali_tutar": _decimal(veri["FATURALI_TUTAR"]),
        "fatura_baglama_orani_yuzde": round(
            (satir_sayisi - faturasiz_satir) / satir_sayisi * 100, 1
        )
        if satir_sayisi
        else 0,
    }


def _aylik_trend(cursor, bas, bit):
    cursor.execute(
        """
        SELECT
            TO_CHAR(YK.DOC_DATE, 'YYYY-MM') AS AY,
            COUNT(DISTINCT YK.GOODS_ID) AS YUK_SAYISI,
            NVL(SUM(OPDET.AMT), 0) AS SATIS_GELIRI
        FROM LMST_L_GOODS YK
        LEFT JOIN LMST_L_GOODS_OP_DET OPDET
            ON OPDET.GOODS_ID = YK.GOODS_ID
           AND OPDET.PURCHASE_SALES_TYPE IN (2, 4)
        WHERE YK.DOC_DATE BETWEEN TO_DATE(:bas, 'DD.MM.YYYY') AND TO_DATE(:bit, 'DD.MM.YYYY')
        GROUP BY TO_CHAR(YK.DOC_DATE, 'YYYY-MM')
        ORDER BY 1
        """,
        {"bas": bas, "bit": bit},
    )
    sutunlar = [c[0] for c in cursor.description]
    return _satirlari_dict_yap(sutunlar, cursor.fetchall())


def _proje_performans(cursor, bas, bit):
    cursor.execute(
        """
        SELECT
            NVL(P.PROJECT_CODE, '(Proje Yok)') AS PROJE_KODU,
            COUNT(DISTINCT YK.GOODS_ID) AS YUK_SAYISI,
            NVL(SUM(OPDET.AMT), 0) AS SATIS_GELIRI,
            ROUND(
                NVL(SUM(OPDET.AMT), 0) * 100.0
                / NULLIF(SUM(SUM(OPDET.AMT)) OVER (), 0),
                1
            ) AS GELIR_PAYI_YUZDE
        FROM LMST_L_GOODS YK
        LEFT JOIN LMSD_L_AGR_PROJ_TYPE P ON P.PROJECT_ID = YK.PROJECT_ID
        LEFT JOIN LMST_L_GOODS_OP_DET OPDET
            ON OPDET.GOODS_ID = YK.GOODS_ID
           AND OPDET.PURCHASE_SALES_TYPE IN (2, 4)
        WHERE YK.DOC_DATE BETWEEN TO_DATE(:bas, 'DD.MM.YYYY') AND TO_DATE(:bit, 'DD.MM.YYYY')
        GROUP BY NVL(P.PROJECT_CODE, '(Proje Yok)')
        ORDER BY SATIS_GELIRI DESC
        FETCH FIRST 20 ROWS ONLY
        """,
        {"bas": bas, "bit": bit},
    )
    sutunlar = [c[0] for c in cursor.description]
    return _satirlari_dict_yap(sutunlar, cursor.fetchall())


def _operasyon_dagilimi(cursor, bas, bit):
    cursor.execute(
        """
        SELECT
            NVL(HK.EXPENSE_CODE, 'BILINMIYOR') AS OPERASYON_KODU,
            COUNT(*) AS SATIR_SAYISI,
            NVL(SUM(OPDET.AMT), 0) AS TOPLAM_TUTAR,
            ROUND(
                NVL(SUM(OPDET.AMT), 0) * 100.0
                / NULLIF(SUM(SUM(OPDET.AMT)) OVER (), 0),
                1
            ) AS GELIR_PAYI_YUZDE
        FROM LMST_L_GOODS_OP_DET OPDET
        JOIN LMST_L_GOODS YK ON YK.GOODS_ID = OPDET.GOODS_ID
        LEFT JOIN INVD_EXPENSE HK ON HK.EXPENSE_ID = OPDET.OPERATION_ID
        WHERE YK.DOC_DATE BETWEEN TO_DATE(:bas, 'DD.MM.YYYY') AND TO_DATE(:bit, 'DD.MM.YYYY')
          AND OPDET.PURCHASE_SALES_TYPE IN (2, 4)
        GROUP BY NVL(HK.EXPENSE_CODE, 'BILINMIYOR')
        ORDER BY TOPLAM_TUTAR DESC
        """,
        {"bas": bas, "bit": bit},
    )
    sutunlar = [c[0] for c in cursor.description]
    return _satirlari_dict_yap(sutunlar, cursor.fetchall())


def _fatura_gecikmesi(cursor, bas, bit):
    """Yük tarihi ile fatura tarihi arasındaki gün farkı."""
    cursor.execute(
        """
        SELECT
            ROUND(AVG(INV.DOC_DATE - YK.DOC_DATE), 1) AS ORT_GECIKME_GUN,
            ROUND(MAX(INV.DOC_DATE - YK.DOC_DATE), 0) AS MAX_GECIKME_GUN,
            COUNT(*) AS FATURALI_SATIR
        FROM LMST_L_GOODS_OP_DET OPDET
        JOIN LMST_L_GOODS YK ON YK.GOODS_ID = OPDET.GOODS_ID
        JOIN PSMT_INVOICE_M INV ON INV.INVOICE_M_ID = OPDET.INVOICE_M_ID
        WHERE YK.DOC_DATE BETWEEN TO_DATE(:bas, 'DD.MM.YYYY') AND TO_DATE(:bit, 'DD.MM.YYYY')
          AND OPDET.PURCHASE_SALES_TYPE IN (2, 4)
        """,
        {"bas": bas, "bit": bit},
    )
    row = cursor.fetchone()
    if not row or row[2] == 0:
        return {"ort_gecikme_gun": None, "max_gecikme_gun": None, "faturali_satir": 0}

    return {
        "ort_gecikme_gun": float(row[0]) if row[0] is not None else None,
        "max_gecikme_gun": int(row[1]) if row[1] is not None else None,
        "faturali_satir": row[2],
    }


def _marj_analizi(cursor, bas, bit, transport_sema: dict):
    """
    Brüt marj: satış geliri − sevk alış maliyeti.
    Sevk maliyeti LMST_L_TRANSPORT_OP_DET tablosundan okunur [VARSAYIM/TODO:
    tablo/kolon adı Uyumsoft lojistik modülüyle uyumludur].
    """
    if not tablo_var_mi("LMST_L_TRANSPORT_OP_DET"):
        return {
            "mevcut": False,
            "mesaj": "LMST_L_TRANSPORT_OP_DET tablosu bulunamadı — marj KPI atlandı.",
        }

    yuk_kolon = transport_sema.get("yuk_kolon")
    if not yuk_kolon:
        return {
            "mevcut": False,
            "mesaj": (
                "LMST_L_TRANSPORT ↔ yük eşleme kolonu bulunamadı — marj KPI atlandı. "
                f"Mevcut kolonlar: {', '.join(transport_sema.get('kolonlar', [])[:12])}..."
            ),
        }

    sql = f"""
        SELECT
            NVL(SUM(satis.AMT), 0) AS TOPLAM_SATIS,
            NVL(SUM(alış.AMT), 0) AS TOPLAM_ALIS
        FROM LMST_L_GOODS YK
        LEFT JOIN (
            SELECT GOODS_ID, SUM(AMT) AS AMT
            FROM LMST_L_GOODS_OP_DET
            WHERE PURCHASE_SALES_TYPE IN (2, 4)
            GROUP BY GOODS_ID
        ) satis ON satis.GOODS_ID = YK.GOODS_ID
        LEFT JOIN LMST_L_TRANSPORT T ON T.{yuk_kolon} = YK.GOODS_ID
        LEFT JOIN (
            SELECT TRANSPORT_ID, SUM(AMT) AS AMT
            FROM LMST_L_TRANSPORT_OP_DET
            GROUP BY TRANSPORT_ID
        ) alış ON alış.TRANSPORT_ID = T.TRANSPORT_ID
        WHERE YK.DOC_DATE BETWEEN TO_DATE(:bas, 'DD.MM.YYYY') AND TO_DATE(:bit, 'DD.MM.YYYY')
    """
    cursor.execute(sql, {"bas": bas, "bit": bit})
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


def _problemleri_tespit_et(cursor, bas, bit, ozet, proje_listesi, marj, transport_sema: dict):
    """Üst yönetimin odaklanması gereken operasyonel problemler."""
    problemler = []

    # 1) Faturasız satış satırları
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

    # 2) Sevk oranı düşük
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

    # 3) Proje konsantrasyon riski
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

    # 4) Negatif / düşük marj
    if marj.get("mevcut") and marj.get("brut_marj_orani_yuzde", 0) < 10:
        problemler.append(
            {
                "oncelik": "YUKSEK",
                "kategori": "Karlılık",
                "baslik": "Brüt marj oranı kritik seviyede",
                "detay": f"Brüt marj oranı: %{marj['brut_marj_orani_yuzde']} (satış − sevk alış)",
                "aksiyon": "Düşük marjlı yükleri proje/operasyon bazında analiz edin; fiyatlandırma revizyonu değerlendirin.",
            }
        )

    # 5) Sevki olmayan yükler
    yuk_kolon = transport_sema.get("yuk_kolon")
    if transport_sema.get("mevcut") and yuk_kolon:
        sql = f"""
            SELECT YK.REFERENCE_NO, YK.DOC_DATE, NVL(P.PROJECT_CODE, '-') AS PROJE
            FROM LMST_L_GOODS YK
            LEFT JOIN LMSD_L_AGR_PROJ_TYPE P ON P.PROJECT_ID = YK.PROJECT_ID
            LEFT JOIN LMST_L_TRANSPORT T ON T.{yuk_kolon} = YK.GOODS_ID
            WHERE YK.DOC_DATE BETWEEN TO_DATE(:bas, 'DD.MM.YYYY') AND TO_DATE(:bit, 'DD.MM.YYYY')
              AND T.TRANSPORT_ID IS NULL
            ORDER BY YK.DOC_DATE DESC
            FETCH FIRST 50 ROWS ONLY
        """
        cursor.execute(sql, {"bas": bas, "bit": bit})
        cursor.execute(sql, {"bas": bas, "bit": bit})
        sevksiz = cursor.fetchall()
        if sevksiz:
            problemler.append(
                {
                    "oncelik": "YUKSEK",
                    "kategori": "Operasyon",
                    "baslik": f"Sevki olmayan yükler ({len(sevksiz)}+ kayıt)",
                    "detay": "Örnek: " + ", ".join(r[0] for r in sevksiz[:5]),
                    "aksiyon": "Sevki oluşturulmamış yük listesini günlük operasyon toplantısına taşıyın.",
                }
            )

    # 6) Sıfır tutarlı satış satırları
    cursor.execute(
        """
        SELECT COUNT(*)
        FROM LMST_L_GOODS_OP_DET OPDET
        JOIN LMST_L_GOODS YK ON YK.GOODS_ID = OPDET.GOODS_ID
        WHERE YK.DOC_DATE BETWEEN TO_DATE(:bas, 'DD.MM.YYYY') AND TO_DATE(:bit, 'DD.MM.YYYY')
          AND OPDET.PURCHASE_SALES_TYPE IN (2, 4)
          AND NVL(OPDET.AMT, 0) = 0
        """,
        {"bas": bas, "bit": bit},
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
    """Oracle'dan KPI verilerini çeker ve analiz sonucunu döner."""
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

        if not tablo_var_mi("LMST_L_TRANSPORT"):
            sonuc.uyarilar.append(
                "LMST_L_TRANSPORT tablosu okunamadı — sevk KPI'ları kısıtlı."
            )
        if not sonuc.marj_analizi.get("mevcut"):
            sonuc.uyarilar.append(sonuc.marj_analizi.get("mesaj", "Marj analizi atlandı."))

        sonuc.problemler = _problemleri_tespit_et(
            cursor, bas, bit, sonuc.ozet, sonuc.proje_performans, sonuc.marj_analizi
        )

    return sonuc


def ornek_analiz_sonucu() -> KpiAnalizSonucu:
    """Oracle bağlantısı olmadan rapor şablonunu test etmek için örnek veri."""
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

"""
Oracle veri katmanı.

BU DOSYA BİR ŞABLONDUR (TODO içerir). `main.py` bu modülden şu iki fonksiyonu
bekler:

    kaynak_yuk_verilerini_getir(kaynak_yuk_no) -> dict
        {
            "PROJE_KODU": str,
            "YUK_TARIHI": str,           # "GG.AA.YYYY" formatında
            "SATIS_SATIRLARI": [
                {
                    "UCRET_TIPI": str,
                    "OPERASYON_KODU": str,
                    "SATIS_FIYATI": float | Decimal | str,
                    "FATURA_NO": str | None,
                    "FATURA_TARIHI": str | None,
                },
                ...
            ],
        }

    yeni_kayitlari_veritabaninda_guncelle(yeni_yuk_numaralari, yeni_sevk_numaralari) -> None
        Başarıyla oluşturulan Yük/Sevk numaralarını Oracle tarafında
        (örn. bir "işlendi" bayrağı/iz tablosu güncellenerek) temizler.

Gerçek SQL sorgularınızı ve tablo/şema adlarınızı bilmediğim için burada
sadece bağlantı iskeleti ve TODO işaretli yer tutucular bulunuyor. Kendi
şemanıza göre doldurmanız gerekir. Eğer elinizde bu dosyanın çalışan bir
versiyonu varsa, onu bu şablonun üzerine kopyalamanız yeterli olacaktır.
"""
import oracledb

import ayarlar


def _baglanti_ac():
    return oracledb.connect(
        user=ayarlar.ORACLE_KULLANICI,
        password=ayarlar.ORACLE_SIFRE,
        dsn=ayarlar.ORACLE_DSN,
    )


def kaynak_yuk_verilerini_getir(kaynak_yuk_no):
    """Kaynak yük numarasına göre proje/tarih/satış satırlarını Oracle'dan getirir."""
    with _baglanti_ac() as conn:
        cur = conn.cursor()

        # TODO: Gerçek tablo/kolon adlarınıza göre güncelleyin.
        cur.execute(
            """
            SELECT proje_kodu, TO_CHAR(yuk_tarihi, 'DD.MM.YYYY') AS yuk_tarihi
            FROM   yukler
            WHERE  yuk_no = :yuk_no
            """,
            {"yuk_no": kaynak_yuk_no},
        )
        satir = cur.fetchone()
        if not satir:
            raise ValueError(f"Oracle'da '{kaynak_yuk_no}' numaralı yük bulunamadı.")
        proje_kodu, yuk_tarihi = satir

        # TODO: Gerçek tablo/kolon adlarınıza göre güncelleyin.
        cur.execute(
            """
            SELECT ucret_tipi, operasyon_kodu, satis_fiyati, fatura_no,
                   TO_CHAR(fatura_tarihi, 'DD.MM.YYYY') AS fatura_tarihi
            FROM   yuk_satis_satirlari
            WHERE  yuk_no = :yuk_no
            ORDER BY satir_no
            """,
            {"yuk_no": kaynak_yuk_no},
        )
        satis_satirlari = [
            {
                "UCRET_TIPI": row[0],
                "OPERASYON_KODU": row[1],
                "SATIS_FIYATI": row[2],
                "FATURA_NO": row[3],
                "FATURA_TARIHI": row[4],
            }
            for row in cur.fetchall()
        ]

    return {
        "PROJE_KODU": proje_kodu,
        "YUK_TARIHI": yuk_tarihi,
        "SATIS_SATIRLARI": satis_satirlari,
    }


def yeni_kayitlari_veritabaninda_guncelle(yeni_yuk_numaralari, yeni_sevk_numaralari):
    """Başarıyla oluşturulan Yük/Sevk numaralarını Oracle tarafında işaretler."""
    if not yeni_yuk_numaralari and not yeni_sevk_numaralari:
        return

    with _baglanti_ac() as conn:
        cur = conn.cursor()

        # TODO: Gerçek güncelleme mantığınıza göre değiştirin.
        for yuk_no in yeni_yuk_numaralari:
            cur.execute(
                "UPDATE yukler SET aktarildi = 1 WHERE yuk_no = :yuk_no",
                {"yuk_no": yuk_no},
            )

        for sevk_no in yeni_sevk_numaralari:
            cur.execute(
                "UPDATE sevkler SET aktarildi = 1 WHERE sevk_no = :sevk_no",
                {"sevk_no": sevk_no},
            )

        conn.commit()

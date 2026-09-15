"""
"Zarar Detay" sayfası — sevk numarası bazında zarar eden işlerin (Tedarikçi/Kiralık
araç ayrımıyla) listesini VERİ satırlarından otomatik hesaplar.

Neden gerekli — kök neden analizi [bkz. .cursor/rules/erp-sistemi.mdc]:
"Zarar Detay" sayfasındaki ZararTedarikci/ZararKiralik tabloları eskiden elle
dolduruluyordu (birinin geçmiş ayın zarar eden sevklerini tek tek bulup Sevk No,
Müşteri, Rota gibi bilgileri yapıştırdığı statik bir liste). Bu tablolardaki
Alış/Satış/Kâr-Zarar sütunları ise VERİ sayfasındaki "Tablo5" tablosuna
SUMIF ile bakan FORMÜLLER. Otomasyon her ay Tablo5'i o ayın verisiyle sıfırdan
yazdığı için (eski ayın satırları silinir), önceki ayda elle girilmiş Sevk
No'lar artık Tablo5'te bulunmuyor ve bu formüller sessizce 0 dönüyor.

ÖNEMLİ — bu sadece "Zarar Detay" sayfasıyla sınırlı değil: "Özet" sayfasındaki
YÖNETİM ALARMLARI kutusu ("Toplam zarar büyüklüğü", "Kiralık araç zararı",
"Tedarikçi araç zararı", "Zarar Eden Sevkiyat Oranı") doğrudan
SUM(ZararKiralik[Kâr/Zarar]) / SUM(ZararTedarikci[Kâr/Zarar]) okuyor — yani bu
sayfa güncellenmezse o alarmlar da sessizce 0 / yanlış görünüyor.

Bu modül sadece SAF HESAPLAMA yapar (Oracle/Excel'e bağımlı değil, test edilebilir).
Excel'e yazma tarafı kpi_sablon_rapor.py içindedir (_com_zarar_detay_*).
"""
from __future__ import annotations

import unicodedata
from collections import defaultdict
from decimal import Decimal, InvalidOperation
from typing import Any

_BOS_ETIKET = "(boş)"

# Excel tablosundaki metin/sayaç sütunları — sırası ZararTedarikci/ZararKiralik
# tablolarının A-I sütunlarıyla birebir eşleşmeli. J-N (Alış/Satış/Kâr-Zarar/
# Zarar %/Zarar Payı) sütunlarına KASITLI OLARAK dokunulmuyor (bkz. kpi_sablon_rapor.py
# _com_zarar_detay_tablo_yaz) — oradaki orijinal formüller Sevk No güncellenince
# kendiliğinden doğru sonucu verir.
ZARAR_DETAY_METIN_SUTUNLARI = [
    "Tarih", "Şube", "Müşteri", "Rota", "Plaka", "Araç Tipi",
    "Sevk No", "Kullanıcı", "Kayıt Sayısı",
]


def _decimal(deger: Any) -> Decimal:
    if deger is None:
        return Decimal("0")
    if isinstance(deger, Decimal):
        return deger
    try:
        return Decimal(str(deger))
    except (InvalidOperation, ValueError):
        return Decimal("0")


def _ascii_kucult(metin: str) -> str:
    """Türkçe özel karakterleri (İ/I/ı/ş/ç/ğ/ö/ü) ASCII'ye indirger — Python'ın
    varsayılan .lower()'ı 'I' -> 'i' çevirdiği için 'KIRALIK' gibi büyük harfli
    değerlerde 'ı' (noktasız i) hiç eşleşmiyordu, bu yüzden basit .lower() YETERSİZ."""
    metin = metin.replace("İ", "i").replace("I", "i").replace("ı", "i")
    metin = unicodedata.normalize("NFKD", metin)
    return "".join(c for c in metin if not unicodedata.combining(c)).lower()


def _mulkiyet_grubu(deger: Any) -> str | None:
    """PLAKA_MULKIYET metnini Tedarikçi/Kiralık'a indirger (bilinen değerler:
    'Tedarikçi', 'Kiralık', 'Misafir Araç' — üçüncüsü [VARSAYIM/TODO] hiçbir
    tabloya net oturmuyor, çağıran taraf varsayılan bir listeye ekler)."""
    metin = _ascii_kucult(str(deger or "").strip())
    if not metin:
        return None
    if "kiral" in metin:
        return "Kiralık"
    if "tedarik" in metin:
        return "Tedarikçi"
    return None


def zarar_eden_sevkleri_hesapla(
    veri_satirlari: list[dict[str, Any]],
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], list[str]]:
    """VERİ satırlarını SEVK_NO'ya göre gruplar (aynı sevk no'daki satırlar
    birleştirilir — birden fazla yük taşıyan sevkler için), toplam
    TOPLAM_KAR_ZARAR'ı negatif olan (zarar eden) sevkleri PLAKA_MULKIYET'e göre
    Tedarikçi/Kiralık listelerine ayırır, en büyük zarardan küçüğe sıralar.

    Dönüş: (tedarikci_satirlari, kiralik_satirlari, uyarilar). Her satır dict'i
    ZARAR_DETAY_METIN_SUTUNLARI + "Alış"/"Satış"/"Kâr/Zarar"/"Zarar %"/
    "Zarar Payı" anahtarlarını içerir (formülle aynı hesap: Zarar %=Kâr-Zarar/Alış,
    Zarar Payı=-Kâr-Zarar/toplam zarar büyüklüğü — bkz. şablondaki
    '=IFERROR(-L/Özet!$K$7,0)' formülü).
    """
    gruplar: dict[str, dict[str, Any]] = {}

    for satir in veri_satirlari:
        sevk_no_ham = satir.get("SEVK_NO")
        anahtar = str(sevk_no_ham).strip() if sevk_no_ham not in (None, "") else _BOS_ETIKET

        grup = gruplar.get(anahtar)
        if grup is None:
            grup = {
                "sevk_no": anahtar,
                "kayit_sayisi": 0,
                "alis": Decimal("0"),
                "satis": Decimal("0"),
                "kar_zarar": Decimal("0"),
                "ilk_satir": satir,
                "mulkiyet_sayaci": defaultdict(int),
            }
            gruplar[anahtar] = grup

        grup["kayit_sayisi"] += 1
        grup["alis"] += _decimal(satir.get("TOPLAM_ALIS"))
        grup["satis"] += _decimal(satir.get("TOPLAM_SATIS"))
        grup["kar_zarar"] += _decimal(satir.get("TOPLAM_KAR_ZARAR"))
        mulkiyet = _mulkiyet_grubu(satir.get("PLAKA_MULKIYET"))
        if mulkiyet:
            grup["mulkiyet_sayaci"][mulkiyet] += 1

    zararli_gruplar = [g for g in gruplar.values() if g["kar_zarar"] < 0]
    toplam_zarar_buyuklugu = -sum((g["kar_zarar"] for g in zararli_gruplar), Decimal("0"))

    tedarikci: list[tuple[Decimal, dict]] = []
    kiralik: list[tuple[Decimal, dict]] = []
    siniflanmayan = 0

    for grup in zararli_gruplar:
        ilk = grup["ilk_satir"]
        if grup["mulkiyet_sayaci"]:
            baskin = max(grup["mulkiyet_sayaci"].items(), key=lambda kv: kv[1])[0]
        else:
            baskin = None
            siniflanmayan += 1

        yukleme = ilk.get("YUKLEME_SEHIR_ADI") or _BOS_ETIKET
        bosaltma = ilk.get("BOSALTMA_SEHIR_ADI") or _BOS_ETIKET
        alis = grup["alis"]
        kar_zarar = grup["kar_zarar"]
        zarar_yuzde = (kar_zarar / alis) if alis else Decimal("0")
        zarar_payi = (-kar_zarar / toplam_zarar_buyuklugu) if toplam_zarar_buyuklugu else Decimal("0")

        satir_out = {
            "Tarih": ilk.get("YUK_TARIHI") or ilk.get("SEVK_TARIHI"),
            "Şube": ilk.get("PROJE_KODU") or _BOS_ETIKET,
            "Müşteri": ilk.get("MUSTERI_ADI") or _BOS_ETIKET,
            "Rota": f"{yukleme} → {bosaltma}",
            "Plaka": ilk.get("PLAKA") or _BOS_ETIKET,
            "Araç Tipi": ilk.get("ARAC_TIPI") or _BOS_ETIKET,
            "Sevk No": grup["sevk_no"],
            "Kullanıcı": ilk.get("KULLANICI") or _BOS_ETIKET,
            "Kayıt Sayısı": grup["kayit_sayisi"],
            "Alış": float(alis),
            "Satış": float(grup["satis"]),
            "Kâr/Zarar": float(kar_zarar),
            "Zarar %": float(zarar_yuzde),
            "Zarar Payı": float(zarar_payi),
        }

        # Sınıflandırılamayan (örn. "Misafir Araç") varsayılan olarak Tedarikçi'ye eklenir.
        hedef = kiralik if baskin == "Kiralık" else tedarikci
        hedef.append((kar_zarar, satir_out))

    tedarikci.sort(key=lambda x: x[0])
    kiralik.sort(key=lambda x: x[0])

    uyarilar: list[str] = []
    if siniflanmayan:
        uyarilar.append(
            f"{siniflanmayan} zarar eden sevk MÜLKİYET bilgisiyle Tedarikçi/Kiralık "
            "olarak sınıflandırılamadı (varsayılan: Tedarikçi listesine eklendi)."
        )

    return [s for _, s in tedarikci], [s for _, s in kiralik], uyarilar

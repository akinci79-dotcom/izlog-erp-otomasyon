from openpyxl import Workbook

from yollar import islem_listesi_yolu, klasorleri_olustur


def excel_sablonu_olustur():
    klasorleri_olustur()
    dosya_adi = islem_listesi_yolu()

    kolonlar = [
        "KAYNAK_YUK_NO", "PLAKA", "SEVK_ALIS_FIYATI",
        "PROJE_KODU", "TARIH", "FATURA_NO", "FATURA_TARIHI",
        "YENI_YUK_NO", "YENI_SEVK_NO", "DURUM", "HATA_ACIKLAMASI"
    ]

    wb = Workbook()
    ws = wb.active
    ws.title = "Islemler"
    ws.append(kolonlar)
    wb.save(dosya_adi)
    print(f"BAŞARILI: '{dosya_adi}' şablonu oluşturuldu.")
    print("Lütfen Excel'i açıp ilk 3 kolona test verilerinizi (Örn: Y-575631, 31ATR52, 40000,50) girin ve kaydedin.")


if __name__ == "__main__":
    excel_sablonu_olustur()

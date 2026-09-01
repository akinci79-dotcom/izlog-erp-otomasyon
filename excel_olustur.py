from openpyxl import Workbook
import os

def excel_sablonu_olustur():
    dosya_adi = "islem_listesi.xlsx"

    # Şartnamemize uygun kolon listesi
    kolonlar = [
        "KAYNAK_YUK_NO", "PLAKA", "SEVK_ALIS_FIYATI",
        "PROJE_KODU", "TARIH", "FATURA_NO", "FATURA_TARIHI",
        "YENI_YUK_NO", "YENI_SEVK_NO", "DURUM", "HATA_ACIKLAMASI"
    ]

    # Yeni bir Excel kitabı oluştur ve aktif sayfayı al
    wb = Workbook()
    ws = wb.active
    ws.title = "Islemler"

    # Başlıkları ilk satıra yaz
    ws.append(kolonlar)

    # Dosyayı kaydet
    wb.save(dosya_adi)
    print(f"BAŞARILI: '{dosya_adi}' şablonu klasörünüze oluşturuldu.")
    print("Lütfen Excel'i açıp ilk 3 kolona test verilerinizi (Örn: Y-575631, 31ATR52, 40000,50) girin ve kaydedin.")

if __name__ == "__main__":
    excel_sablonu_olustur()

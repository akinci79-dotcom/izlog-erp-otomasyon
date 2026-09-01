"""
Proje alt klasörlerini oluşturur.

Windows sunucusunda pull sonrası bir kez çalıştırın:
  python klasorleri_olustur.py
"""
from yollar import klasorleri_olustur, islem_listesi_yolu, kpi_rapor_yolu, otomasyon_klasoru, raporlar_klasoru


def main():
    klasorleri_olustur()
    print("BAŞARILI: Klasör yapısı hazır.")
    print(f"  Otomasyon: {otomasyon_klasoru()}")
    print(f"  Raporlar:  {raporlar_klasoru()}")
    print(f"  Excel:     {islem_listesi_yolu()}")
    print(f"  KPI rapor: {kpi_rapor_yolu()}")


if __name__ == "__main__":
    main()

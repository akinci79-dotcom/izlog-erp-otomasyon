"""
Proje alt klasörlerini oluşturur.

Windows sunucusunda pull sonrası bir kez çalıştırın:
  python klasorleri_olustur.py
"""
from yollar import (
    islem_listesi_yolu,
    kpi_klasoru,
    kpi_rapor_yolu,
    klasorleri_olustur,
    otomasyon_klasoru,
)


def main():
    klasorleri_olustur()
    print("BAŞARILI: Klasör yapısı hazır.")
    print(f"  CANLI (otomasyon): {otomasyon_klasoru()}")
    print(f"  KPI (raporlar):    {kpi_klasoru()}")
    print(f"  Excel:             {islem_listesi_yolu()}")
    print(f"  KPI rapor:         {kpi_rapor_yolu()}")


if __name__ == "__main__":
    main()

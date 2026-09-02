"""
KPI Excel raporu — referans şablon modu (varsayılan).

Kullanım (KPI klasöründen):
  cd KPI
  python kpi_rapor_olustur.py              # Şablondan rapor (VERİ + Filo Detay)
  python kpi_rapor_olustur.py --analiz       # Eski çok sayfalı analiz raporu
  python kpi_rapor_olustur.py --ornek        # Analiz modunda örnek veri
"""
from __future__ import annotations

import sys

from kpi_sablon_rapor import sablon_rapor_olustur


def main():
    if "--analiz" in sys.argv or "--ornek" in sys.argv:
        from kpi_rapor_analiz import main as analiz_main

        analiz_main()
        return

    sablon_rapor_olustur()


if __name__ == "__main__":
    main()

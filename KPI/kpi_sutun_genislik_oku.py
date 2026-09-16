"""Bir Excel dosyasındaki sütun genişliklerini okuyup ayarlar.py için hazır satır üretir.

Elle sütun genişliklerini ayarladığınız kpi_rapor.xlsx'i bu scripte verin;
çıktıyı doğrudan ayarlar.py içindeki KPI_SABIT_SUTUN_GENISLIKLERI değişkenine
yapıştırın — o sayfadaki genişlikler bir daha hiç değişmez (AutoFit atlanır).

Kullanım (KPI klasöründen):
  python kpi_sutun_genislik_oku.py "raporlar\\kpi_rapor.xlsx" "Özet"

Sayfa adı verilmezse dosyadaki TÜM sayfalar için genişlikler yazdırılır.
"""
from __future__ import annotations

import sys
from pathlib import Path

from openpyxl import load_workbook


def main() -> None:
    if len(sys.argv) < 2:
        print('Kullanım: python kpi_sutun_genislik_oku.py "dosya.xlsx" ["Sayfa Adı"]')
        return

    dosya = Path(sys.argv[1])
    if not dosya.exists():
        print(f"HATA: Dosya yok → {dosya}")
        return

    sayfa_adi = sys.argv[2] if len(sys.argv) > 2 else None

    wb = load_workbook(dosya, data_only=True)
    if sayfa_adi:
        if sayfa_adi not in wb.sheetnames:
            print(f"HATA: Sayfa bulunamadı: {sayfa_adi}\nMevcut sayfalar: {wb.sheetnames}")
            wb.close()
            return
        sayfalar = [wb[sayfa_adi]]
    else:
        sayfalar = wb.worksheets

    for ws in sayfalar:
        genislikler: dict[str, float] = {}
        # column_dimensions sadece Excel'de gerçekten genişliği değiştirilmiş
        # sütunları içerir — max_column'a (dolu hücrelere) değil buna güvenilmeli.
        for harf, boyut in ws.column_dimensions.items():
            if boyut and boyut.width:
                genislikler[str(harf)] = round(float(boyut.width), 2)
        genislikler = dict(sorted(genislikler.items(), key=lambda kv: kv[0]))

        print(f"\n=== {ws.title} ===")
        if not genislikler:
            print("  (özel genişlik ayarlanmış sütun bulunamadı)")
            continue

        print("KPI_SABIT_SUTUN_GENISLIKLERI = {")
        print(f'    "{ws.title}": {{')
        for harf, w in genislikler.items():
            print(f'        "{harf}": {w},')
        print("    },")
        print("}")

    wb.close()


if __name__ == "__main__":
    main()

"""Excel COM ve AutoFit testi — KPI klasöründen: python kpi_excel_test.py"""
from __future__ import annotations

import sys
from pathlib import Path


def main() -> int:
    print("=== KPI Excel COM Testi ===\n")

    try:
        import win32com.client  # type: ignore
    except ImportError:
        print("HATA: pywin32 kurulu değil.")
        print("  pip install pywin32")
        return 1

    print("[1/3] Excel.Application başlatılıyor...")
    excel = None
    try:
        try:
            excel = win32com.client.gencache.EnsureDispatch("Excel.Application")
        except Exception:
            excel = win32com.client.Dispatch("Excel.Application")
        print(f"  OK — Excel sürümü: {excel.Version}")
    except Exception as exc:
        print(f"  HATA: Microsoft Excel bulunamadı veya açılamadı.")
        print(f"  Detay: {exc}")
        print("\n  pywin32 tek başına yetmez; masaüstü Excel (365/Office) kurulu olmalı.")
        return 1
    finally:
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass

    rapor = Path(__file__).resolve().parent / "raporlar" / "kpi_rapor.xlsx"
    if not rapor.exists():
        print(f"\n[2/3] Atlandı — rapor yok: {rapor}")
        print("  Önce: python kpi_rapor_olustur.py")
        return 0

    print(f"\n[2/3] Rapor dosyası: {rapor}")
    print("[3/3] Pivot yenile + AutoFit deneniyor...")
    print("  Not: Eski openpyxl kaydı bozuk dosya üretmiş olabilir — önce python kpi_rapor_olustur.py çalıştırın.")

    from kpi_sablon_rapor import _excel_islemleri

    ok, mesaj = _excel_islemleri(rapor, pivot_yenile=True)
    if ok:
        print("  OK — Excel işlemleri tamamlandı.")
        if mesaj:
            print(f"  Uyarı: {mesaj}")
        return 0

    print("  HATA — Excel işlemleri başarısız.")
    if mesaj:
        print(f"  Detay: {mesaj}")
    print("\n  Kontrol: Excel dosyayı açık tutmayın; Excel kurulu ve lisanslı olsun.")
    return 1


if __name__ == "__main__":
    sys.exit(main())

#!/usr/bin/env python3
"""referans/kpi_sablon.xlsx degistiginde sablon_surumu.txt dosyasini gunceller."""

from __future__ import annotations

import hashlib
from pathlib import Path

KPI_DIR = Path(__file__).resolve().parent
SABLON = KPI_DIR / "referans" / "kpi_sablon.xlsx"
SURUM = KPI_DIR / "referans" / "sablon_surumu.txt"


def main() -> None:
    if not SABLON.is_file():
        raise SystemExit(f"Sablon bulunamadi: {SABLON}")

    veri = SABLON.read_bytes()
    ozet = hashlib.sha256(veri).hexdigest()
    SURUM.write_text(
        "# Repodaki kpi_sablon.xlsx surumu — kpi_guncelle.ps1 karsilastirmasi icin\n"
        "# Sablon degistiginde bu dosyayi da guncelleyin (kpi_sablon_surumu_guncelle.py)\n"
        f"size={len(veri)}\n"
        f"sha256={ozet}\n",
        encoding="utf-8",
    )
    print(f"Guncellendi: {SURUM}")
    print(f"  size={len(veri)}")
    print(f"  sha256={ozet}")


if __name__ == "__main__":
    main()

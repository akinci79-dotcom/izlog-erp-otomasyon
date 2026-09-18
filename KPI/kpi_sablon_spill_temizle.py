#!/usr/bin/env python3
"""Özet sayfasındaki dinamik dizi spill alanlarına yazılmış eski değerleri temizler.

Excel 365 formül spill sonuçlarını bazen XML'e önbellekler. VERİ yenilenince bu
hücreler spill ile çakışır (#BAŞV! / #HESAPLA! hatası). Yalnızca kök formül
hücreleri kalır.
"""
from __future__ import annotations

import re
import shutil
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
Q = f"{{{NS}}}"

KPI_DIR = Path(__file__).resolve().parent
SABLON = KPI_DIR / "referans" / "kpi_sablon.xlsx"

# (anchor, spill_rows, spill_cols) — anchor dışındaki tüm hücreler silinir
_SPILL_BLOKLARI = [
    ("A7", 5, "ABCDEF"),
    ("A43", 5, "ABCDE"),
    ("A52", 5, "A"),
    ("B52", 5, "B"),
    ("C52", 5, "CDEF"),
    ("H52", 5, "HIJKLM"),
]

_B52_FORMUL = (
    "_xlfn.LET(_xlpm.r,_xlfn.ANCHORARRAY(A52),"
    "_xlfn.MAP(_xlpm.r,_xlfn.LAMBDA(_xlpm.x,"
    "_xlfn.LET(_xlpm.fr,Tablo5[YUKLEME_SEHIR_ADI]&\"\","
    "_xlpm.to,Tablo5[BOSALTMA_SEHIR_ADI]&\"\","
    "_xlpm.rk,IF(_xlpm.fr=\"\",\"(boş)\",_xlpm.fr)&\" → \"&IF(_xlpm.to=\"\",\"(boş)\",_xlpm.to),"
    "_xlpm.cust,Tablo5[MUSTERI_ADI]&\"\","
    "_xlpm.sales,Tablo5[TOPLAM_SATIS],"
    "_xlpm.cs,_xlfn.UNIQUE(_xlfn._xlws.FILTER(_xlpm.cust,_xlpm.rk=_xlpm.x)),"
    "_xlpm.ss,_xlfn.MAP(_xlpm.cs,_xlfn.LAMBDA(_xlpm.y,"
    "_xlfn.SUM(_xlfn._xlws.FILTER(_xlpm.sales,(_xlpm.rk=_xlpm.x)*(_xlpm.cust=_xlpm.y),0)))),"
    "IF(_xlfn.ROWS(_xlpm.cs)=0,\"\","
    "IFERROR(INDEX(_xlpm.cs,_xlfn.XMATCH(_xlfn.MAX(_xlpm.ss),_xlpm.ss)),\"-\"))))))"
)


def _col_num(col: str) -> int:
    n = 0
    for ch in col:
        n = n * 26 + ord(ch) - 64
    return n


def _num_col(n: int) -> str:
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _parse_ref(ref: str) -> tuple[str, int]:
    m = re.match(r"^([A-Z]+)(\d+)$", ref)
    if not m:
        raise ValueError(ref)
    return m.group(1), int(m.group(2))


def _spill_hucreleri() -> set[str]:
    silinecek: set[str] = set()
    for anchor, rows, cols in _SPILL_BLOKLARI:
        acol, arow = _parse_ref(anchor)
        for i in range(rows):
            for col in cols:
                ref = f"{col}{arow + i}"
                if ref != anchor:
                    silinecek.add(ref)
    return silinecek


def _ozet_sheet_path(z: zipfile.ZipFile) -> str:
    import xml.etree.ElementTree as ET

    wb = ET.fromstring(z.read("xl/workbook.xml"))
    rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    rid_to_target = {rel.get("Id"): rel.get("Target") for rel in rels}
    for sh in wb.findall(f"{Q}sheets")[0]:
        if sh.get("name") == "Özet":
            rid = sh.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
            return "xl/" + rid_to_target[rid].lstrip("/")
    raise RuntimeError("Özet sayfası bulunamadı")


def _hucre_guncelle(c_elem, ref: str) -> None:
    if ref != "B52":
        return
    for tag in ("f", "v", "is"):
        child = c_elem.find(f"{Q}{tag}")
        if child is not None:
            c_elem.remove(child)
    f = ET.Element(f"{Q}f")
    f.text = _B52_FORMUL
    c_elem.insert(0, f)


def main() -> None:
    if not SABLON.is_file():
        raise SystemExit(f"Şablon yok: {SABLON}")

    silinecek = _spill_hucreleri()
    yedek = SABLON.with_suffix(".xlsx.bak")
    shutil.copy2(SABLON, yedek)

    buf = SABLON.with_suffix(".xlsx.tmp")
    with zipfile.ZipFile(SABLON, "r") as zin, zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
        ozet_path = _ozet_sheet_path(zin)
        sheet_xml = zin.read(ozet_path)
        root = ET.fromstring(sheet_xml)
        silinen = 0

        for row in root.findall(f".//{Q}row"):
            for c in list(row.findall(f"{Q}c")):
                ref = c.get("r")
                if not ref:
                    continue
                if ref in silinecek:
                    row.remove(c)
                    silinen += 1
                elif ref == "B52":
                    _hucre_guncelle(c, ref)

        # boş kalan row elemanlarını bırak — Excel tolere eder
        new_xml = ET.tostring(root, encoding="utf-8", xml_declaration=True)
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == ozet_path:
                data = new_xml
            zout.writestr(item, data)

    buf.replace(SABLON)
    print(f"Spill temizligi tamam: {silinen} hucre silindi, B52 formulu guncellendi")
    print(f"Yedek: {yedek}")


if __name__ == "__main__":
    main()

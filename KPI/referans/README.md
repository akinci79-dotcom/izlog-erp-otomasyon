# KPI referans şablonu

Bu klasöre **mevcut KPI Excel dosyanızı** şu adla kopyalayın:

```
kpi_sablon.xlsx
```

Örnek: `7- Temmuz 2026 İzlog Lojistik Raporları.xlsx` → `kpi_sablon.xlsx`

## Şablonda olması gerekenler

| Sayfa | Açıklama |
|---|---|
| **VERİ** (gizli olabilir) | Pivot kaynağı — 1. satır kolon başlıkları |
| **Filo Detay** | Tedarikçi hesaplaşma yapıştırma alanı |

Pivotlu sayfalar (Özet, Filo Analizi, Şube KZ, …) şablonda kalır; script VERİ ve Filo Detay'ı doldurur, ardından pivotları yeniler.

## Kolon eşleşmesi

VERİ sayfasının 1. satırındaki başlıklar, **kpi_veri_rapor.sql** içindeki SELECT alias'larıyla eşleştirilir.

**Zorunlu:** Uyumsoft VERİ raporu SQL'inizi `referans/kpi_veri_rapor.sql` olarak kaydedin. Kod SQL'e dokunmaz; ne varsa aynen çalıştırır.

Tarih parametreleri: `:bas` ve `:bit` (ayarlar.py dönem tarihleri).

Eşleşmeyen başlık varsa `ayarlar.py` içinde:

```python
KPI_KOLON_ESLEME = {
    "Şablondaki Başlık": "ORACLE_KOLON_ADI",
}
```

## Çalıştırma

```powershell
cd KPI
python kpi_rapor_olustur.py
```

Çıktı: `raporlar\kpi_rapor.xlsx`

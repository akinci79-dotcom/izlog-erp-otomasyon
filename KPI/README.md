# İzlog Lojistik — KPI Analiz Modülü

Oracle ERP verilerinden üst yönetim KPI raporu üretir.

**Bu klasör otomasyon projesinden tamamen bağımsızdır.** Üst klasördeki
`izlog_yuk_otomasyon.py`, `ayarlar.py` veya `oracle_okuyucu.py` kullanılmaz.

## Konum

```
Cursor ERP Otomasyon\
└── KPI\                    ← sadece bu klasör
    ├── ayarlar.py          ← kendi ayarlarınız (git'e commit etmeyin)
    ├── ayarlar.example.py
    ├── oracle_baglanti.py
    ├── kpi_analiz.py
    ├── kpi_rapor_olustur.py
    ├── requirements.txt
    └── raporlar\
        └── kpi_rapor.xlsx  ← üretilen rapor
```

## Kurulum (Windows)

```powershell
cd "C:\Users\hakinci\Desktop\Kodlarım\Cursor ERP Otomasyon\KPI"
copy ayarlar.example.py ayarlar.py
```

`ayarlar.py` içinde `DB_SIFRE` alanını doldurun.

```powershell
pip install -r requirements.txt
```

Oracle Instant Client: `C:\instantclient\instantclient_19_32`

## Çalıştırma

```powershell
cd KPI
python kpi_rapor_olustur.py --ornek    # şablon testi (Oracle gerekmez)
python kpi_rapor_olustur.py            # gerçek veri
```

Rapor: `KPI\raporlar\kpi_rapor.xlsx`

## Dönem ayarı

`ayarlar.py`:

```python
KPI_BASLANGIC_TARIHI = "01.01.2026"
KPI_BITIS_TARIHI = "31.01.2026"
```

## Rapor içeriği

| Sayfa | İçerik |
|---|---|
| Yönetici Özeti | Yük/sevk, gelir, fatura oranı, marj, problem listesi |
| Aylık Trend | Aylık hacim ve gelir |
| Proje Performansı | Top 20 proje |
| Operasyon Dağılımı | NAVLUN, UĞRAMA vb. |

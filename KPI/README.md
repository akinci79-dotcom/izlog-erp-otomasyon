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

## Kurulum (Windows) — tek seferde

### Yöntem A: Tek blok kopyala-yapıştır (PowerShell)

PowerShell'i **Yönetici olarak açmanıza gerek yok.** Aşağıdaki bloğun **tamamını** seçip yapıştırın, Enter'a bir kez basın — hepsi sırayla çalışır:

```powershell
$Base = "C:\Users\hakinci\Desktop\Kodlarım"
$KpiDir = "$Base\Cursor ERP Otomasyon\KPI"
$Temp = "$Base\izlog-kpi-temp"
Set-Location $Base
if (Test-Path $Temp) { Remove-Item $Temp -Recurse -Force }
git clone -b cursor/kpi-analiz-rapor-0bd3 --depth 1 https://github.com/akinci79-dotcom/izlog-erp-otomasyon.git izlog-kpi-temp
New-Item -ItemType Directory -Force -Path $KpiDir | Out-Null
Copy-Item "$Temp\KPI\*" $KpiDir -Recurse -Force
Remove-Item $Temp -Recurse -Force
Set-Location $KpiDir
if (-not (Test-Path ayarlar.py)) { Copy-Item ayarlar.example.py ayarlar.py }
pip install -r requirements.txt
python kpi_rapor_olustur.py --ornek
Write-Host "Bitti! Rapor: $KpiDir\raporlar\kpi_rapor_ORNEK.xlsx" -ForegroundColor Green
```

Sonra `ayarlar.py` içinde `DB_SIFRE` doldurun.

### Yöntem B: Çift tıkla (dosyalar indikten sonra)

`KPI\kpi_kur.bat` dosyasına çift tıklayın — geri kalanını script yapar.

### Yöntem C: Manuel (adım adım)

```powershell
cd "C:\Users\hakinci\Desktop\Kodlarım\Cursor ERP Otomasyon\KPI"
copy ayarlar.example.py ayarlar.py
pip install -r requirements.txt
python kpi_rapor_olustur.py --ornek
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
| Kiralık Araç Detay | Tedarikçi hakediş dosyaları (maliyet/kar) |
| Kiralık Araç Cari | Tedarikçi bazında filo özeti |
| Kalem Detay | Sevk/yük kalem satırları (çoklu yük kırılımı) |
| Sevk Yük Kırılım | Birden fazla yük taşıyan sevklerin kar/zarar özeti |
| Fatura Detay | Sevk/yük kalemi ↔ fatura eşleşmesi |
| Faturasız Kalemler | Fatura no'su olmayan satırlar |

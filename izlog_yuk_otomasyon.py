import os
import shutil
import re
from datetime import datetime
from decimal import Decimal, InvalidOperation
from openpyxl import load_workbook
from playwright.sync_api import sync_playwright

import ayarlar
from oracle_okuyucu import kaynak_yuk_verilerini_getir, yeni_kayitlari_veritabaninda_guncelle

# --- YARDIMCI FONKSİYONLAR ---
def devexpress_tarih_yaz(sayfa, selector, tarih_metni):
    sayfa.wait_for_selector(selector, state="visible", timeout=15000)
    sayfa.click(selector)
    sayfa.keyboard.press("Home")
    sayfa.keyboard.press("Shift+End")
    sayfa.keyboard.press("Backspace")
    sayfa.type(selector, tarih_metni, delay=50)
    sayfa.press(selector, "Enter")
    sayfa.wait_for_timeout(500)
    sayfa.keyboard.press("Tab")
    sayfa.wait_for_timeout(500)


def _agsakinligini_bekle(sayfa, timeout=10000, yedek_bekleme=1500):
    """
    wait_for_load_state("networkidle") DevExpress'in arka planda sürekli
    keep-alive/heartbeat isteği attığı ekranlarda hiçbir zaman "idle" olmayıp
    timeout ile patlayabilir. Bu sarmalayıcı, timeout olursa akışı durdurmak
    yerine sabit bir yedek bekleme ile devam eder (eski davranışa güvenli düşüş).
    """
    try:
        sayfa.wait_for_load_state("networkidle", timeout=timeout)
    except Exception:
        sayfa.wait_for_timeout(yedek_bekleme)


def _tutar_adaylarini_ayikla(metin):
    """Bir metin içindeki Türkçe formatlı (1.234,56) tüm tutar adaylarını Decimal olarak döndürür."""
    adaylar = re.findall(r"-?\d{1,3}(?:\.\d{3})*,\d{2}", metin)
    sonuc = []
    for aday in adaylar:
        try:
            sonuc.append(Decimal(aday.replace(".", "").replace(",", ".")))
        except InvalidOperation:
            continue
    return sonuc


def _satirda_tutar_var_mi(satir_metni, hedef_tutar, tolerans=Decimal("0.01")):
    """
    Bir grid satırının metninde, hedef tutara (kuruş toleransıyla) eşit bir
    tutar olup olmadığını kontrol eder.

    NEDEN BÖYLE: Fatura Seç (LOV) penceresinde satırı tek bir CSS string'iyle
    ("78.279,00" gibi tam grid formatına `has_text` ile birebir eşleştirerek)
    bulmaya çalışmak kırılgan; ERP'nin gösterdiği format ile bizim ürettiğimiz
    string arasında ufak bir fark (boşluk, TL eki, farklı yuvarlama) olursa
    satır hiç bulunamaz ve akış patlar. Bunun yerine, fatura no ile filtrelenen
    satırların METNİ Python'da ayrıştırılıp gerçek tutarla (Decimal) tolerans
    dahilinde karşılaştırılıyor.
    """
    for deger in _tutar_adaylarini_ayikla(satir_metni):
        if abs(deger - hedef_tutar) <= tolerans:
            return True
    return False


# ==========================================
# UYUMSOFT ERP PLAYWRIGHT OPERASYONU V4.1 (FATURA TUTAR EŞLEŞTİRME DÜZELTMESİ)
# ==========================================
def uyumsoft_islemlerini_yap(page, kaynak_yuk_no, plaka, sevk_alis_fiyati, oracle_data, mevcut_durum, onceki_yeni_yuk_no, checkpoint_kaydet=None):

    atla_faz_123 = False
    islem_yuk_no = kaynak_yuk_no
    aktif_sayfa = page

    if mevcut_durum in ["YÜK OLUŞTU", "HATA_SEVK"] and onceki_yeni_yuk_no:
        atla_faz_123 = True
        islem_yuk_no = onceki_yeni_yuk_no

    # Güvenli dict.get() kullanımı
    satis_satirlari = oracle_data.get("SATIS_SATIRLARI", [])
    tum_faturalar = " + ".join([str(s.get("FATURA_NO", "")) for s in satis_satirlari if s.get("FATURA_NO") and str(s.get("FATURA_NO")).upper() != "NONE"])
    tum_fatura_tarihleri = " + ".join([str(s.get("FATURA_TARIHI", "")) for s in satis_satirlari if s.get("FATURA_TARIHI")])
    proje_kodu = oracle_data.get("PROJE_KODU", "")
    yuk_tarihi = oracle_data.get("YUK_TARIHI", "")
    saf_tarih = yuk_tarihi.replace(".", "") if yuk_tarihi else ""

    page.goto(ayarlar.ERP_YUK_LISTESI_URL)
    page.wait_for_selector("#myListPage_DXFREditorcol1_I", state="visible", timeout=15000)

    page.fill("#myListPage_DXFREditorcol1_I", islem_yuk_no)
    page.press("#myListPage_DXFREditorcol1_I", "Enter")

    # --- YÜK SEÇİMİ (ORİJİNAL SAĞLAM OMURGA) ---
    page.wait_for_timeout(2000)  # Arama yapıldıktan sonra listenin güncellenmesini bekle
    saglam_secici = f"tr.dxgvDataRow_Aqua:has-text('{islem_yuk_no}')"
    page.wait_for_selector(saglam_secici, state="visible", timeout=15000)
    page.click(saglam_secici)

    # DERİN TEST MODU: ayarlar.py'de DRY_RUN=True VE DERIN_TEST_MODU=True ise,
    # sadece arama/seçimle sınırlı kalınmaz; Kopyalama + tüm veri girişi +
    # fatura LOV eşleştirmesi (asıl düzeltilen hata burada) GERÇEKTEN yapılır,
    # ama ana "Kaydet" (#btnSave_CD) butonuna KESİNLİKLE basılmaz — pencere
    # kayıt yapılmadan kapatılır. Bu, "Kopya" tıklamanın ve satır bazlı
    # "Kaydet" tıklamalarının veritabanına yazmadığı (yalnızca ana Kaydet'in
    # yazdığı) varsayımına dayanır -- bunu ERP'nizde teyit etmeden production'a
    # karşı çalıştırmayın.
    derin_test = bool(ayarlar.DRY_RUN) and bool(getattr(ayarlar, "DERIN_TEST_MODU", False))

    if ayarlar.DRY_RUN and not derin_test:
        return {"durum": "DRY_RUN BAŞARILI", "yeni_yuk_no": None, "yeni_sevk_no": None, "proje": proje_kodu, "fatura_no": tum_faturalar, "fatura_tarihi": tum_fatura_tarihleri, "tarih": yuk_tarihi, "aktif_sayfa": aktif_sayfa}

    if derin_test:
        print(f"[{kaynak_yuk_no}] ⚠️ DERİN TEST MODU AKTİF: Kopyalama ve veri girişi GERÇEKTEN yapılacak, "
              f"ama ana 'Kaydet' butonuna kesinlikle basılmayacak.")

    yeni_yuk_no = onceki_yeni_yuk_no

    if not atla_faz_123:
        with page.context.expect_page() as yeni_pencere_beklentisi:
            page.click("#btnCopy_CD")

        aktif_sayfa = yeni_pencere_beklentisi.value
        aktif_sayfa.wait_for_load_state("networkidle")

        # Refactor: Yardımcı fonksiyon kullanımı
        devexpress_tarih_yaz(aktif_sayfa, "#TabControl_dte_DocDate_I", saf_tarih)

        aktif_sayfa.click("span.dx-vam:has-text('Yük Diğer Bilgiler')")
        aktif_sayfa.wait_for_selector("#TabControl_chk_IsGoodsInWhouse_S_D", state="visible", timeout=5000)
        aktif_sayfa.click("#TabControl_chk_IsGoodsInWhouse_S_D")

        aktif_sayfa.wait_for_timeout(500)
        aktif_sayfa.click("span.dx-vam:has-text('Yurtiçi Yük Tanımı')")
        aktif_sayfa.wait_for_selector("#TabControl_grd_LGoodsOpDetailCollection_EmptyRow_btnNew", state="visible")

        if derin_test:
            # ⚠️ BİLİNEN ERP HATASI: Fatura seçildikten sonra satır bazlı "Kaydet"e
            # basmak, ERP tarafında Yükü kilitliyor ve satır SQL'den boşaltılmadan
            # silinemiyor hale geliyor. Bu yüzden derin test modunda HİÇBİR satır
            # Kaydet'e basılmaz -- sadece TEK bir satır (tercihen faturalı olan,
            # asıl düzeltilen fatura eşleştirme mantığını doğrulamak için) test
            # edilir ve satır kaydedilmeden önce test sonlandırılır.
            faturali_satirlar = [
                s for s in satis_satirlari
                if s.get('FATURA_NO') and str(s.get('FATURA_NO')).strip().upper() not in ("", "NONE")
            ]
            test_edilecek_satirlar = (faturali_satirlar or satis_satirlari)[:1]
        else:
            test_edilecek_satirlar = satis_satirlari

        for satis in test_edilecek_satirlar:
            aktif_sayfa.click("#TabControl_grd_LGoodsOpDetailCollection_EmptyRow_btnNew")
            aktif_sayfa.wait_for_selector("#TabControl_grd_LGoodsOpDetailCollection_DXEditor4_I", state="visible", timeout=15000)

            op_kodu = satis.get('OPERASYON_KODU', 'NAVLUN')
            ucret_tipi = satis.get('UCRET_TIPI', 'NAVLUN')

            if op_kodu != 'NAVLUN':
                aktif_sayfa.fill("#TabControl_grd_LGoodsOpDetailCollection_DXEditor1_I", ucret_tipi)
                aktif_sayfa.press("#TabControl_grd_LGoodsOpDetailCollection_DXEditor1_I", "Tab")
                aktif_sayfa.fill("#TabControl_grd_LGoodsOpDetailCollection_DXEditor9_I", op_kodu)
                aktif_sayfa.press("#TabControl_grd_LGoodsOpDetailCollection_DXEditor9_I", "Tab")

            # Güvenli Decimal Çevirimi
            ham_fiyat = satis.get('SATIS_FIYATI')
            if ham_fiyat is None or str(ham_fiyat).strip() == "":
                raise ValueError(f"[{kaynak_yuk_no}] SATIS_FIYATI eksik veya geçersiz!")

            fiyat_decimal = Decimal(str(ham_fiyat))
            formatli_tutar = f"{fiyat_decimal:.2f}".replace(".", ",")

            aktif_sayfa.click("#TabControl_grd_LGoodsOpDetailCollection_DXEditor4_I", force=True)
            aktif_sayfa.fill("#TabControl_grd_LGoodsOpDetailCollection_DXEditor4_I", formatli_tutar)
            aktif_sayfa.press("#TabControl_grd_LGoodsOpDetailCollection_DXEditor4_I", "Tab")
            aktif_sayfa.wait_for_timeout(400)

            aktif_sayfa.keyboard.press("Tab")
            aktif_sayfa.wait_for_timeout(400)
            aktif_sayfa.keyboard.press("Tab")

            # Ağ trafiğinin dinlenmesi (Sabit 3000ms yerine) - timeout olursa yedek beklemeye düşer
            _agsakinligini_bekle(aktif_sayfa, timeout=10000, yedek_bekleme=1500)

            fatura_no_raw = satis.get('FATURA_NO')
            if not fatura_no_raw or str(fatura_no_raw).strip().upper() == "NONE" or str(fatura_no_raw).strip() == "":
                if derin_test:
                    print(f"[{kaynak_yuk_no}] DERİN TEST: Bu satırda fatura yok, test edilecek eşleştirme mantığı "
                          f"yok. Satır KAYDEDİLMEDEN (kilitleme riskine karşı) test sonlandırılıyor.")
                    break
                aktif_sayfa.locator("a[id*='editnew']:has-text('Kaydet')").first.click(force=True)
                aktif_sayfa.wait_for_selector("#TabControl_grd_LGoodsOpDetailCollection_EmptyRow_btnNew", state="visible", timeout=15000)
                continue

            fatura_no_str = str(fatura_no_raw).strip()

            aktif_kutu_tablosu = aktif_sayfa.locator("*:focus").locator("xpath=ancestor::table[1]")
            box = aktif_kutu_tablosu.bounding_box()

            if box:
                btn_x = box['x'] + box['width'] - 12
                btn_y = box['y'] + (box['height'] / 2)
                aktif_sayfa.mouse.move(btn_x, btn_y)
                aktif_sayfa.wait_for_timeout(300)
                aktif_sayfa.mouse.click(btn_x, btn_y)
            else:
                raise RuntimeError(f"[{kaynak_yuk_no}] HATA: Fatura kutusu ekranda odaklanamadı!")

            aktif_sayfa.wait_for_timeout(1500)
            lov_penceresi = aktif_sayfa.frames[-1]
            lov_penceresi.wait_for_selector("#myListPage_DXFREditorcol2_I", state="visible", timeout=20000)

            # --- ZIRH: Tutar kutusuna HİÇ dokunulmuyor (sadece arama/reset için boş bırakılıyor) ---
            # Tutar filtre kutusu (#myListPage_DXFREditorcol6_I) DevExpress'in maskeli/
            # formatlı sayısal editörü; buraya değer yazmak akışı bozuyor. Sadece
            # fatura no ile filtreleniyor, doğru satır aşağıda tutar metninden
            # ayrıştırılarak (Decimal ile) bulunuyor.
            lov_penceresi.fill("#myListPage_DXFREditorcol2_I", fatura_no_str)
            lov_penceresi.fill("#myListPage_DXFREditorcol6_I", "")
            lov_penceresi.press("#myListPage_DXFREditorcol2_I", "Enter")
            _agsakinligini_bekle(aktif_sayfa, timeout=10000, yedek_bekleme=1500)

            hedef_tutar = fiyat_decimal.quantize(Decimal("0.01"))

            # Fatura no'ya göre (kelime sınırı ile tam eşleşme) filtrelenmiş satırlar.
            fatura_satirlari = lov_penceresi.locator("tr.dxgvDataRow_Aqua").filter(
                has_text=re.compile(rf"\b{re.escape(fatura_no_str)}\b")
            )
            fatura_satirlari.first.wait_for(state="visible", timeout=10000)

            # NOT: Aynı fatura no + aynı tutar birden fazla satırda görünebilir
            # (örn. aynı faturanın farklı kalemleri aynı tutara sahip olabilir).
            # Bu durum önemli değil: fatura no zaten eşleşiyorsa ve tutar da
            # eşleşiyorsa, hangi satır olduğu fark etmez -- ilk eşleşen satır
            # bulunur bulunmaz seçilir, kalan satırlara bakılmaz.
            hedef_satir = None
            satir_sayisi = fatura_satirlari.count()
            incelenen_metinler = []
            for idx in range(satir_sayisi):
                aday = fatura_satirlari.nth(idx)
                try:
                    metin = aday.inner_text()
                except Exception:
                    continue
                incelenen_metinler.append(metin.replace("\n", " | "))
                if _satirda_tutar_var_mi(metin, hedef_tutar):
                    hedef_satir = aday
                    break  # İlk eşleşme yeterli; duplikasyon sorun değil.

            if hedef_satir is None:
                ornek = "\n".join(incelenen_metinler[:5])
                raise RuntimeError(
                    f"[{kaynak_yuk_no}] HATA: Fatura '{fatura_no_str}' için {hedef_tutar} TL tutarında "
                    f"eşleşen satır bulunamadı ({satir_sayisi} satır incelendi).\n"
                    f"İncelenen satır örnekleri:\n{ornek}"
                )

            hedef_satir.click()

            aktif_sayfa.wait_for_timeout(500)
            lov_penceresi.locator("#btnChoose_CD").click(force=True)
            aktif_sayfa.wait_for_timeout(1500)

            if derin_test:
                # Fatura arama + eşleştirme + seçme (asıl düzeltilen mantık) BAŞARILI.
                # ⚠️ Satır bazlı "Kaydet"e KESİNLİKLE basılmıyor -- bilinen ERP
                # kilitleme hatasına karşı. Test burada güvenle sonlandırılır.
                print(f"[{kaynak_yuk_no}] DERİN TEST: Fatura '{fatura_no_str}' başarıyla bulundu ve seçildi. "
                      f"Satır KAYDEDİLMEDEN (bilinen kilitleme riskine karşı) test sonlandırılıyor.")
                break

            aktif_sayfa.wait_for_selector("a[id*='editnew']:has-text('Kaydet')", state="visible", timeout=15000)
            aktif_sayfa.locator("a[id*='editnew']:has-text('Kaydet')").first.click(force=True)
            aktif_sayfa.wait_for_selector("#TabControl_grd_LGoodsOpDetailCollection_EmptyRow_btnNew", state="visible", timeout=15000)

        if derin_test:
            print(f"[{kaynak_yuk_no}] DERİN TEST TAMAMLANDI: Test edilen satır için veri girişi/fatura eşleştirmesi "
                  f"tamamlandı. Hiçbir satır 'Kaydet'e basılmadı, ana 'Kaydet' butonuna da BASILMADI. "
                  f"Pencere kayıt yapılmadan kapatılıyor.")
            try:
                aktif_sayfa.close()
            except Exception:
                pass
            return {
                "durum": "DERİN_TEST BAŞARILI",
                "yeni_yuk_no": None,
                "yeni_sevk_no": None,
                "proje": proje_kodu,
                "fatura_no": tum_faturalar,
                "fatura_tarihi": tum_fatura_tarihleri,
                "tarih": yuk_tarihi,
                "aktif_sayfa": page
            }

        aktif_sayfa.click("#btnSave_CD", force=True)

        # JS Güvenli Parametre Enjeksiyonu
        aktif_sayfa.wait_for_function('''
            (arg) => {
                let val = document.querySelector("#TabControl_txt_ReferenceNo_I").value.trim();
                return val !== "" && val.startsWith("Y-") && val !== arg;
            }
        ''', arg=kaynak_yuk_no, timeout=20000)

        yeni_yuk_no = aktif_sayfa.input_value("#TabControl_txt_ReferenceNo_I")

        if checkpoint_kaydet:
            checkpoint_kaydet(yeni_yuk_no)

    # --- FAZ 4: SEVK OLUŞTURMA ---
    aktif_sayfa.click("#TabControl_txt_ReferenceNo_I", button="right")
    aktif_sayfa.wait_for_selector("div.uyum-popup-menu span:text-is('Sevk Oluştur')", state="visible")
    aktif_sayfa.click("div.uyum-popup-menu span:text-is('Sevk Oluştur')")

    devexpress_tarih_yaz(aktif_sayfa, "#TabControl_dte_DocDate_I", saf_tarih)

    aktif_sayfa.fill("#TabControl_bte_TractorSerialNoPlateNo_I", plaka)
    aktif_sayfa.press("#TabControl_bte_TractorSerialNoPlateNo_I", "Enter")
    aktif_sayfa.wait_for_timeout(2000)

    aktif_sayfa.click("#TabControl_grd_LTransOpDetailCollection_EmptyRow_btnNew")
    aktif_sayfa.wait_for_selector("#TabControl_grd_LTransOpDetailCollection_DXEditor4_I", state="visible", timeout=15000)

    sevk_guvenli = Decimal(str(sevk_alis_fiyati)) if sevk_alis_fiyati else Decimal('0.00')
    formatli_sevk_fiyati = f"{sevk_guvenli:.2f}".replace(".", ",")

    aktif_sayfa.fill("#TabControl_grd_LTransOpDetailCollection_DXEditor4_I", formatli_sevk_fiyati)
    aktif_sayfa.press("#TabControl_grd_LTransOpDetailCollection_DXEditor4_I", "Tab")

    aktif_sayfa.click("a[id*='editnew']:has-text('Kaydet')", force=True)
    aktif_sayfa.wait_for_timeout(500)

    aktif_sayfa.click("#btnSave_CD", force=True)

    aktif_sayfa.wait_for_function('''
        () => {
            let val = document.querySelector("#TabControl_txt_TransportNo_I").value.trim();
            return val !== "" && val.startsWith("S-");
        }
    ''', timeout=20000)

    yeni_sevk_no = aktif_sayfa.input_value("#TabControl_txt_TransportNo_I")

    if not atla_faz_123:
        aktif_sayfa.close()

    return {
        "durum": "BAŞARILI",
        "yeni_yuk_no": yeni_yuk_no,
        "yeni_sevk_no": yeni_sevk_no,
        "proje": proje_kodu,
        "fatura_no": tum_faturalar,
        "fatura_tarihi": tum_fatura_tarihleri,
        "tarih": yuk_tarihi,
        "aktif_sayfa": page
    }

def main():
    excel_dosyasi = "islem_listesi.xlsx"

    if not os.path.exists(excel_dosyasi):
        print(f"HATA: '{excel_dosyasi}' dosyası bulunamadı. Lütfen dizini kontrol edin.")
        return

    yedek_isim = f"yedek_islem_listesi_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    try:
        shutil.copy(excel_dosyasi, yedek_isim)
        print(f"Orijinal veri güvene alındı. Yedek: {yedek_isim}")
    except PermissionError:
        print("HATA: Excel dosyası açık! Lütfen Excel'i kapatıp otomasyonu tekrar çalıştırın.")
        return

    wb = load_workbook(excel_dosyasi)
    ws = wb.active

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, channel="chrome")
        context = browser.new_context(viewport={'width': 1920, 'height': 1080}, locale='tr-TR')
        page = context.new_page()

        print("ERP sistemine giriş yapılıyor...")
        page.goto(ayarlar.ERP_LOGIN_URL)
        page.wait_for_selector("input#tbxUsername", state="visible")

        page.fill("input#tbxUsername", ayarlar.ERP_KULLANICI)
        page.fill("input#tbxPassword", ayarlar.ERP_SIFRE)
        page.press("input#tbxPassword", "Enter")

        page.wait_for_timeout(4000)

        basarili_yeni_yukler = []
        basarili_yeni_sevkler = []

        max_row = ws.max_row
        for i in range(2, max_row + 1):
            kaynak_yuk_no = ws.cell(row=i, column=1).value
            plaka = ws.cell(row=i, column=2).value
            sevk_alis_fiyati = ws.cell(row=i, column=3).value

            mevcut_durum = ws.cell(row=i, column=10).value
            onceki_yeni_yuk_no = ws.cell(row=i, column=8).value

            if not kaynak_yuk_no:
                continue

            if mevcut_durum in ["BAŞARILI", "DRY_RUN BAŞARILI"]:
                print(f"[{kaynak_yuk_no}] Durum: {mevcut_durum} -> İşlem Atlanıyor...")
                continue

            print(f"\n--- İŞLEM BAŞLIYOR: {kaynak_yuk_no} ---")

            hata_sayfasi = page
            try:
                oracle_data = kaynak_yuk_verilerini_getir(kaynak_yuk_no)

                def checkpoint_kaydet(yeni_yuk_no_degeri):
                    ws.cell(row=i, column=8).value = yeni_yuk_no_degeri
                    ws.cell(row=i, column=10).value = "YÜK OLUŞTU"
                    wb.save(excel_dosyasi)
                    print(f"[{kaynak_yuk_no}] CHECKPOINT ✅: Yük {yeni_yuk_no_degeri} Excel'e kaydedildi → Sevk aşamasına geçiliyor.")

                sonuc = uyumsoft_islemlerini_yap(page, kaynak_yuk_no, plaka, sevk_alis_fiyati, oracle_data, mevcut_durum, onceki_yeni_yuk_no, checkpoint_kaydet)

                hata_sayfasi = sonuc["aktif_sayfa"]

                ws.cell(row=i, column=4).value = sonuc["proje"]
                ws.cell(row=i, column=5).value = sonuc["tarih"]
                ws.cell(row=i, column=6).value = sonuc["fatura_no"]
                ws.cell(row=i, column=7).value = sonuc["fatura_tarihi"]
                if sonuc["yeni_yuk_no"]: ws.cell(row=i, column=8).value = sonuc["yeni_yuk_no"]
                if sonuc["yeni_sevk_no"]: ws.cell(row=i, column=9).value = sonuc["yeni_sevk_no"]
                ws.cell(row=i, column=10).value = sonuc["durum"]

                if sonuc["durum"] == "BAŞARILI":
                    if sonuc["yeni_yuk_no"]: basarili_yeni_yukler.append(sonuc["yeni_yuk_no"])
                    if sonuc["yeni_sevk_no"]: basarili_yeni_sevkler.append(sonuc["yeni_sevk_no"])

                wb.save(excel_dosyasi)

            except Exception as e:
                hata_mesaji = str(e)
                print(f"HATA OLUŞTU [{kaynak_yuk_no}]: {hata_mesaji}")

                zaman_damgasi = datetime.now().strftime("%H%M%S")
                hata_foto = f"hata_{kaynak_yuk_no}_{zaman_damgasi}.png"

                try:
                    if len(context.pages) > 1:
                        hata_sayfasi = context.pages[-1]

                    hata_sayfasi.screenshot(path=hata_foto)
                    hata_sayfasi.keyboard.press("Escape")

                    if mevcut_durum in ["YÜK OLUŞTU", "HATA_SEVK"]:
                        ws.cell(row=i, column=10).value = "HATA_SEVK"
                    else:
                        ws.cell(row=i, column=10).value = "HATA_YUK"
                        ws.cell(row=i, column=8).value = None

                    if len(context.pages) > 1:
                        hata_sayfasi.close()
                except Exception:
                    ws.cell(row=i, column=10).value = "HATA_BİLİNMEYEN"

                ws.cell(row=i, column=11).value = hata_mesaji
                wb.save(excel_dosyasi)
                print(f"Hata detayı kaydedildi. Ekran Görüntüsü: {hata_foto}")

                try:
                    page.goto(ayarlar.ERP_YUK_LISTESI_URL)
                except Exception:
                    print("Sistem başlangıç ekranına dönemedi! Güvenlik için otomasyon durduruluyor.")
                    break

        print("\nTüm Excel satırları tamamlandı.")

        if basarili_yeni_yukler or basarili_yeni_sevkler:
            print(f"\n>>> Veritabanı İz Temizliği (Toplu SQL İşlemi) Başlatılıyor... <<<")
            yeni_kayitlari_veritabaninda_guncelle(basarili_yeni_yukler, basarili_yeni_sevkler)

        print("\nERP'den güvenli çıkış yapılıyor...")
        try:
            page.click("div.logout-main")
            page.wait_for_timeout(2000)
        except Exception:
            pass

        browser.close()

if __name__ == "__main__":
    main()

import ssl
ssl._create_default_https_context = ssl._create_unverified_context
import streamlit as str_web
import os
import easyocr
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

# Web sitesinin başlığı ve tasarımı
str_web.set_page_config(page_title="Akıllı Çok Dilli Kitap Dönüştürücü", layout="centered")
str_web.title("📝 Akıllı Çok Dilli Arşiv ve Kitap Dönüştürücü")
str_web.write("Belgenizi yükleyin; yapay zeka satır satır dilleri kendi algılasın ve bilgisayar hattına çevirsin!")

klasor = "C:/Users/LENOVO/OneDrive/Desktop/proje"

if "okunan_sonuc" not in str_web.session_state:
    str_web.session_state.okunan_sonuc = ""

# Bilgisayardan dosya yükleme kutusu
yuklenen_dosya = str_web.file_uploader(
    "📌 Lütfen taratmak istediğiniz çok dilli belge resmini seçin (.jpg, .jpeg, .png, .jfif)", 
    type=["jpg", "jpeg", "png", "jfif"]
)

if yuklenen_dosya is not None:
    resim = Image.open(yuklenen_dosya)
    str_web.image(resim, caption="Yüklenen Belge", use_container_width=True)
    
    # 🔍 1. AŞAMA: YAPAY ZEKAYI TETİKLEME BUTONU
    if str_web.button("🔍 1. Adım: Akıllı Yapay Zeka Taramasını Başlat", type="primary"):
        with str_web.spinner("⏳ Yapay zeka tüm Doğu dillerini (Osmanlıca/Arapça/Farsça) aynı anda analiz ediyor..."):
            try:
                gecici_yol = f"{klasor}/gecici_tarama.jpg"
                resim.save(gecici_yol)
                
                # Yapay zekaya aynı anda Arapça, Farsça ve Osmanlıca harf gözlüklerini takıyoruz
                okuyucu = easyocr.Reader(['ar', 'fa', 'ug'], gpu=False)
                sonuc = okuyucu.readtext(gecici_yol, paragraph=True, canvas_size=3000, mag_ratio=2.0)
                
                if sonuc:
                    metinler = []
                    for item in sonuc:
                        metinler.append(item)
                    str_web.session_state.okunan_sonuc = "\n".join(metinler)
                    str_web.success("🎉 Akıllı tarama tamamlandı! Yapay zeka dilleri ayırt etmeye hazır.")
                else:
                    str_web.warning("🤖 Yapay zeka resimdeki harfleri seçemedi.")
                
                if os.path.exists(gecici_yol):
                    os.remove(gecici_yol)
            except Exception as e:
                str_web.error(f"❌ Bir pürüz oluştu: {e}")

    # ✍️ 2. AŞAMA: CANLI DÜZENLEME PANELİ
    if str_web.session_state.okunan_sonuc:
        str_web.markdown("---")
        str_web.subheader("✍️ 2. Adım: Metin Düzenleme ve Kontrol Paneli")
        
        duzenlenen_metin = str_web.text_area(
            "Karışık Metin İçeriği", 
            value=str_web.session_state.okunan_sonuc, 
            height=300
        )
        
        # 📄 3. AŞAMA: DİLLERİ AYIRT EDEREK WORD YAPMA BUTONU
        if str_web.button("💾 3. Adım: Dilleri Ayırt Et ve Word Kitabına Çevir", type="secondary"):
            try:
                doc = Document()
                
                # Ekrandaki her bir satırı tek tek inceliyoruz
                for satir in duzenlenen_metin.split('\n'):
                    if satir.strip():
                        paragraf = doc.add_paragraph(satir)
                        paragraf.paragraph_format.rtl = True
                        paragraf.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                        
                        # 🕵️‍♂️ AKILLI DİL DEDEKTİFİ:
                        # Satırın içindeki özel harflere bakarak hangi dil olduğunu tahmin ediyoruz
                        farsça_harfler = ["پ", "چ", "ژ", "گ"]
                        osmanlica_ekler = ["ڭ", "𐊠"] # Sağır kef gibi Osmanlıca harfler
                        
                        # Varsayılan olarak şık bir Osmanlıca/Arapça yazı tipi seçiyoruz
                        yazi_tipi = 'Traditional Arabic'
                        
                        # Eğer satırda Farsçaya özel harfler yoğunsa şık bir Nestalik/Farsça yazı tipi yapıyoruz
                        if any(harf in satir for harf in farsça_harfler):
                            yazi_tipi = 'IranNastaliq' # Bilgisayarda varsa muhteşem Fars hattı yapar
                        elif any(harf in satir for harf in osmanlica_ekler):
                            yazi_tipi = 'Traditional Arabic' # Osmanlıca matbu hattı
                        
                        for run in paragraf.runs:
                            run.font.name = yazi_tipi
                            run.font.size = 16
                            
                word_dosya_adi = f"{klasor}/dijital_rika_kitap.docx"
                doc.save(word_dosya_adi)
                
                str_web.success("🏆 BÜYÜK BAŞARI! Yapay zeka dilleri ayırt etti ve her satırı kendi yazı tipinde güncelledi.")
                str_web.balloons()
            except Exception as e:
                str_web.error(f"❌ Word üretilirken pürüz çıktı: {e}")

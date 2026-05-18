import ssl
ssl._create_default_https_context = ssl._create_unverified_context
import streamlit as str_web
import os
import easyocr
from PIL import Image
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

# Web sitesinin başlığı ve tasarımı
str_web.set_page_config(page_title="Evrensel Yapay Zeka Arşiv Dönüştürücü", layout="centered")
str_web.title("📝 Evrensel Yapay Zeka Arşiv ve Kitap Dönüştürücü")
str_web.write("Belgenizi yükleyin; yapay zeka (Osmanlıca/Arapça/Farsça/Türkçe/İngilizce) satırlarını otomatik ayırt etsin!")

klasor = "C:/Users/LENOVO/OneDrive/Desktop/proje"

if "okunan_sonuc" not in str_web.session_state:
    str_web.session_state.okunan_sonuc = ""

# Bilgisayardan çok dilli dosya yükleme kutusu
yuklenen_dosya = str_web.file_uploader(
    "📌 Lütfen taratmak istediğiniz karışık dilli Rika veya matbu belge resmini seçin (.jpg, .jpeg, .png, .jfif)", 
    type=["jpg", "jpeg", "png", "jfif"]
)

if yuklenen_dosya is not None:
    resim = Image.open(yuklenen_dosya)
    str_web.image(resim, caption="Yüklenen Karmaşık Belge", use_container_width=True)
    
    # 🔍 1. AŞAMA: YAPAY ZEKAYI TETİKLEME BUTONU
    if str_web.button("🔍 1. Adım: Evrensel Yapay Zeka Taramasını Başlat", type="primary"):
        with str_web.spinner("⏳ Yapay zeka tüm dünya dillerini (Arapça/Farsça/Osmanlıca/Türkçe/İngilizce) aynı anda tarıyor..."):
            try:
                gecici_yol = f"{klasor}/gecici_tarama.jpg"
                resim.save(gecici_yol)
                
                # 🔥 EVRENSAL ZEKÂ GÖZLÜĞÜ: Tüm alfabeleri tek seferde beyne yüklüyoruz
                okuyucu = easyocr.Reader(['ar', 'fa', 'ug', 'tr', 'en'], gpu=False)
                sonuc = okuyucu.readtext(gecici_yol, paragraph=True, canvas_size=3000, mag_ratio=2.0)
                
                if sonuc:
                    metinler = []
                    for item in sonuc:
                        metinler.append(item)
                    str_web.session_state.okunan_sonuc = "\n".join(metinler)
                    str_web.success("🎉 Evrensel tarama tamamlandı! Yapay zeka alfabeleri ayırt etmeye hazır.")
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
            "Karışık Alfabeli Metin İçeriği", 
            value=str_web.session_state.okunan_sonuc, 
            height=300
        )
        
        # 📄 3. AŞAMA: EVRENSEL DİL AYIRT EDİCİ WORD BUTONU
        if str_web.button("💾 3. Adım: Tüm Alfabeleri Ayırt Et ve Word Kitabına Çevir", type="secondary"):
            try:
                doc = Document()
                
                # Ekrandaki her bir satırı tek tek inceliyoruz
                for satir in duzenlenen_metin.split('\n'):
                    if satir.strip():
                        paragraf = doc.add_paragraph()
                        
                        # 🕵️‍♂️ EVRENSEL DİL DEDEKTİFİ:
                        # Satırın içindeki harf karakterlerine bakarak alfabeyi çözüyoruz
                        dogu_harfleri = ["ا", "ب", "ت", "ث", "ج", "ح", "خ", "د", "ذ", "ر", "ز", "س", "ش", "ص", "ض", "ط", "ظ", "ع", "غ", "ف", "ق", "ك", "ل", "م", "ن", "ه", "و", "ي", "پ", "چ", "ژ", "گ", "ڭ"]
                        farsça_harfler = ["پ", "چ", "ژ", "گ"]
                        
                        # Satırda Doğu alfabesi (Arap/Fars/Osmanlıca) harfleri var mı kontrolü
                        dogu_mu = any(harf in satir for harf in dogu_harfleri)
                        
                        if dogu_mu:
                            # 📝 Eğer satır Doğu alfabesiyse: Sağdan sola hizala
                            paragraf.paragraph_format.rtl = True
                            paragraf.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                            
                            # Farsça harfler varsa ona özel font ver
                            if any(harf in satir for harf in farsça_harfler):
                                yazi_tipi = 'IranNastaliq'
                            else:
                                yazi_tipi = 'Traditional Arabic'
                        else:
                            # 📝 Eğer satır Latin (Türkçe/İngilizce) alfabesiyse: Soldan sağa hizala
                            paragraf.paragraph_format.rtl = False
                            paragraf.alignment = WD_ALIGN_PARAGRAPH.LEFT
                            yazi_tipi = 'Calibri'
                        
                        # Metni paragrafa ekleyip yazı tipini basıyoruz
                        paragraf.add_run(satir)
                        for run in paragraf.runs:
                            run.font.name = yazi_tipi
                            run.font.size = 14 if not dogu_mu else 16
                            
                word_dosya_adi = f"{klasor}/dijital_rika_kitap.docx"
                doc.save(word_dosya_adi)
                
                str_web.success("🏆 EVRENSEL BAŞARI! Yapay zeka tüm alfabeleri/yönleri ayırt etti ve Word kitabınızı üretti.")
                str_web.balloons()
            except Exception as e:
                str_web.error(f"❌ Word üretilirken pürüz çıktı: {e}")

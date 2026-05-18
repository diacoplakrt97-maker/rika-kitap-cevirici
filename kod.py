import ssl
ssl._create_default_https_context = ssl._create_unverified_context
import streamlit as str_web
import os
import easyocr
from PIL import Image
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

# Web sitesinin başlığı ve tasarımı
str_web.set_page_config(page_title="Çok Dilli Yapay Zeka Dönüştürücü", layout="centered")
str_web.title("📝 Çok Dilli Arşiv ve Kitap Dönüştürücü")
str_web.write("Belgenizi yükleyin, dilinizi seçin, yapay zeka ile anında şık bir Word kitabına dönüştürün!")

# Sabit klasör yolumuz
klasor = "C:/Users/LENOVO/OneDrive/Desktop/proje"

# Sitenin hafızasını kontrol paneli için hazırlıyoruz
if "okunan_sonuc" not in str_web.session_state:
    str_web.session_state.okunan_sonuc = ""

# 🌍 YENİ ÖZELLİK: Sitenin en başına şık bir "Dil Seçim Menüsü" ekliyoruz
str_web.subheader("🌍 1. Dil Ayarını Seçin")
secilen_dil_adi = str_web.selectbox(
    "Taramak istediğiniz belgenin dilini seçiniz:",
    ["Osmanlıca / Arapça (Rika)", "Modern Türkçe", "İngilizce (English)"]
)

# Seçilen dile göre yapay zekanın anlayacağı resmi kodları belirliyoruz
if secilen_dil_adi == "Osmanlıca / Arapça (Rika)":
    dil_kodu = ['ar']
    rtl_ayari = True  # Sağdan sola yazı düzeni
elif secilen_dil_adi == "Modern Türkçe":
    dil_kodu = ['tr']
    rtl_ayari = False # Soldan sağa yazı düzeni
else:
    dil_kodu = ['en']
    rtl_ayari = False # Soldan sağa yazı düzeni

str_web.markdown("---")

# Bilgisayardan dosya yükleme kutusu
yuklenen_dosya = str_web.file_uploader(
    "📌 2. Lütfen taratmak istediğiniz belge resmini seçin (.jpg, .jpeg, .png, .jfif)", 
    type=["jpg", "jpeg", "png", "jfif"]
)

if yuklenen_dosya is not None:
    resim = Image.open(yuklenen_dosya)
    str_web.image(resim, caption="Yüklenen Belge", use_container_width=True)
    
    # 🔍 1. AŞAMA: YAPAY ZEKAYI TETİKLEME BUTONU
    if str_web.button("🔍 2. Adım: Yapay Zeka Taramasını Başlat", type="primary"):
        with str_web.spinner(f"⏳ Yapay zeka {secilen_dil_adi} harflerini çözüyor... Lütfen bekleyin..."):
            try:
                gecici_yol = f"{klasor}/gecici_tarama.jpg"
                resim.save(gecici_yol)
                
                # Seçtiğimiz dil koduyla yapay zekayı ayağa kaldırıyoruz
                okuyucu = easyocr.Reader(dil_kodu, gpu=False)
                sonuc = okuyucu.readtext(gecici_yol, paragraph=True, canvas_size=3000, mag_ratio=2.0)
                
                if sonuc:
                    metinler = []
                    for item in sonuc:
                        metinler.append(item)
                    str_web.session_state.okunan_sonuc = "\n".join(metinler)
                    str_web.success("🎉 Tarama Başarıyla Tamamlandı! Aşağıdaki panelden yazıları kontrol edebilirsiniz.")
                else:
                    str_web.warning("🤖 Yapay zeka resimdeki harfleri seçemedi.")
                
                if os.path.exists(gecici_yol):
                    os.remove(gecici_yol)
            except Exception as e:
                str_web.error(f"❌ Bir pürüz oluştu: {e}")

    # ✍️ 2. AŞAMA: CANLI DÜZENLEME PANELİ
    if str_web.session_state.okunan_sonuc:
        str_web.markdown("---")
        str_web.subheader("✍️ 3. Adım: Metin Düzenleme ve Kontrol Paneli")
        
        duzenlenen_metin = str_web.text_area(
            "Metin İçeriği (İhtiyaca göre düzenleyebilirsiniz)", 
            value=str_web.session_state.okunan_sonuc, 
            height=300
        )
        
        # 📄 3. AŞAMA: WORD DOSYASI ÜRETME BUTONU
        if str_web.button("💾 4. Adım: Son Hali Word Kitabına Çevir", type="secondary"):
            try:
                doc = Document()
                for satir in duzenlenen_metin.split('\n'):
                    if satir.strip():
                        paragraf = doc.add_paragraph(satir)
                        
                        # Seçilen dile göre yazının yönünü ayarlıyoruz (Sağdan sola veya Soldan sağa)
                        if rtl_ayari:
                            paragraf.paragraph_format.rtl = True
                            paragraf.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                            yazi_tipi = 'Traditional Arabic'
                        else:
                            paragraf.paragraph_format.rtl = False
                            paragraf.alignment = WD_ALIGN_PARAGRAPH.LEFT
                            yazi_tipi = 'Calibri'
                            
                        for run in paragraf.runs:
                            run.font.name = yazi_tipi
                            run.font.size = 14 if not rtl_ayari else 16
                            
                word_dosya_adi = f"{klasor}/dijital_rika_kitap.docx"
                doc.save(word_dosya_adi)
                
                str_web.success("🏆 KİTABINIZ HAZIR! Seçtiğiniz dil düzeninde Word belgeniz başarıyla güncellendi.")
                str_web.balloons()
            except Exception as e:
                str_web.error(f"❌ Word dosyası üretilirken pürüz çıktı: {e}")
else:
    str_web.session_state.okunan_sonuc = ""
    str_web.info("💡 Devam etmek için lütfen yukarıdaki kutuya bir belge resmi sürükleyin veya seçin.")

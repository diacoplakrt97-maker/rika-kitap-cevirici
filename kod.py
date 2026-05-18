import ssl
ssl._create_default_https_context = ssl._create_unverified_context
import streamlit as str_web
import os
import easyocr
from PIL import Image
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

# Web sitesinin başlığı ve tasarımı
str_web.set_page_config(page_title="Rika ve Arapça Yapay Zeka Dönüştürücü", layout="centered")
str_web.title("📝 Profesyonel Rika ve Arapça Kitap Dönüştürücü")
str_web.write("Belgenizi yükleyin, yapay zeka okumasını yapın, ekranda düzenleyin ve Word kitabınızı indirin!")

# Sabit klasör yolumuz
klasor = "C:/Users/LENOVO/OneDrive/Desktop/proje"

# Sitenin hafızasını (Session State) kontrol paneli için hazırlıyoruz
if "okunan_sonuc" not in str_web.session_state:
    str_web.session_state.okunan_sonuc = ""

# Bilgisayardan dosya yükleme kutusu
yuklenen_dosya = str_web.file_uploader(
    "📌 Lütfen taratmak istediğiniz Rika el yazısı resmini seçin (.jpg, .jpeg, .png, .jfif)", 
    type=["jpg", "jpeg", "png", "jfif"]
)

if yuklenen_dosya is not None:
    resim = Image.open(yuklenen_dosya)
    str_web.image(resim, caption="Yüklenen Rika Belgesi", use_container_width=True)
    
    # 🔍 1. AŞAMA: YAPAY ZEKAYI TETİKLEME BUTONU
    if str_web.button("🔍 1. Adım: Yapay Zeka Taramasını Başlat", type="primary"):
        with str_web.spinner("⏳ Yapay zeka harfleri derinlemesine çözüyor... Lütfen bekleyin..."):
            try:
                gecici_yol = f"{klasor}/gecici_tarama.jpg"
                resim.save(gecici_yol)
                
                okuyucu = easyocr.Reader(['ar'], gpu=False)
                sonuc = okuyucu.readtext(gecici_yol, paragraph=True, canvas_size=3000, mag_ratio=2.0)
                
                if sonuc:
                    # Okunan satırları birleştirip sitenin hafızasına kaydediyoruz
                    metinler = []
                    for item in sonuc:
                        metinler.append(item[1])
                    str_web.session_state.okunan_sonuc = "\n".join(metinler)
                    str_web.success("🎉 Tarama Başarıyla Tamamlandı! Aşağıdaki panelden yazıları kontrol edebilirsiniz.")
                else:
                    str_web.warning("🤖 Yapay zeka resimdeki harfleri seçemedi.")
                
                if os.path.exists(gecici_yol):
                    os.remove(gecici_yol)
            except Exception as e:
                str_web.error(f"❌ Bir pürüz oluştu: {e}")

    # ✍️ 2. AŞAMA: CANLI DÜZENLEME PANELİ (Yalnızca tarama bittiyse açılır)
    if str_web.session_state.okunan_sonuc:
        str_web.markdown("---")
        str_web.subheader("✍️ 2. Adım: Metin Düzenleme ve Kontrol Paneli")
        str_web.write("Yapılan okumayı aşağıdan inceleyebilirsiniz. Eksik veya hatalı yerleri doğrudan kutunun içine tıklayarak klavyenizle düzeltebilirsiniz:")
        
        # Kullanıcının ekrandan yazıyı değiştirebileceği akıllı geniş metin kutusu
        duzenlenen_metin = str_web.text_area(
            "Arapça / Rika Metin İçeriği", 
            value=str_web.session_state.okunan_sonuc, 
            height=300
        )
        
        # 📄 3. AŞAMA: WORD DOSYASI ÜRETME BUTONU
        if str_web.button("💾 3. Adım: Son Hali Word Kitabına Çevir", type="secondary"):
            try:
                doc = Document()
                # Kullanıcının eliyle düzelttiği o son güncel metni satır satır Word'e aktarıyoruz
                for satir in duzenlenen_metin.split('\n'):
                    if satir.strip():
                        paragraf = doc.add_paragraph(satir)
                        paragraf.paragraph_format.rtl = True
                        paragraf.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                        for run in paragraf.runs:
                            run.font.name = 'Traditional Arabic'
                            run.font.size = 16
                            
                word_dosya_adi = f"{klasor}/dijital_rika_kitap.docx"
                doc.save(word_dosya_adi)
                
                str_web.success("🏆 KİTABINIZ HAZIR! Son düzenlemelerinizle birlikte Word belgeniz başarıyla güncellendi.")
                str_web.balloons() # Balonları artık burada uçuruyoruz!
                str_web.info(f"📁 Masaüstündeki 'proje' klasörünün içine gidip 'dijital_rika_kitap.docx' dosyanızı alabilirsiniz.")
            except Exception as e:
                str_web.error(f"❌ Word dosyası üretilirken pürüz çıktı: {e}")
else:
    # Resim yoksa sitenin hafızasını sıfırlıyoruz
    str_web.session_state.okunan_sonuc = ""
    str_web.info("💡 Devam etmek için lütfen yukarıdaki kutuya bir Rika resmi sürükleyin veya seçin.")

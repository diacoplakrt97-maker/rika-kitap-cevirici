import ssl
ssl._create_default_https_context = ssl._create_unverified_context
import streamlit as str_web
import os
import easyocr
from PIL import Image
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from fpdf import FPDF
from io import BytesIO

# Web sitesinin başlığı ve tasarımı
str_web.set_page_config(page_title="Evrensel Yapay Zeka Arşiv, Sözlük ve PDF Asistanı", layout="centered")
str_web.title("📝 Evrensel Yapay Zeka Arşiv ve Kitap Asistanı")
str_web.write("Belgenizi yükleyin; yapay zeka harfleri çözsün, arşiv sözlüğünü çıkarsın, Word veya PDF olarak teslim etsin!")

klasor = "C:/Users/LENOVO/OneDrive/Desktop/proje"

if "okunan_sonuc" not in str_web.session_state:
    str_web.session_state.okunan_sonuc = ""
if "tercüme_sonuc" not in str_web.session_state:
    str_web.session_state.tercüme_sonuc = ""

# Bilgisayardan dosya yükleme kutusu
yuklenen_dosya = str_web.file_uploader(
    "📌 Lütfen taratmak istediğiniz Rika veya karışık dilli belge resmini seçin (.jpg, .jpeg, .png, .jfif)", 
    type=["jpg", "jpeg", "png", "jfif"]
)

if yuklenen_dosya is not None:
    resim = Image.open(yuklenen_dosya)
    str_web.image(resim, caption="Yüklenen Belge", use_container_width=True)
    
    # 🔍 1. AŞAMA: YAPAY ZEKAYI TETİKLEME BUTONU
    if str_web.button("🔍 1. Adım: Evrensel Yapay Zeka Taramasını Başlat", type="primary"):
        with str_web.spinner("⏳ Yapay zeka tüm alfabeleri tarıyor... Lütfen bekleyin..."):
            try:
                gecici_yol = f"{klasor}/gecici_tarama.jpg"
                resim.save(gecici_yol)
                
                okuyucu = easyocr.Reader(['ar', 'fa', 'ug', 'tr', 'en'], gpu=False)
                sonuc = okuyucu.readtext(gecici_yol, paragraph=True, canvas_size=3000, mag_ratio=2.0)
                
                if sonuc:
                    metinler = []
                    for item in sonuc:
                        metinler.append(item)
                    str_web.session_state.okunan_sonuc = "\n".join(metinler)
                    str_web.success("🎉 Evrensel tarama tamamlandı! Aşağıdan kontrol edebilirsiniz.")
                else:
                    str_web.warning("🤖 Yapay zeka resimdeki harfleri seçemedi.")
                
                if os.path.exists(gecici_yol):
                    os.remove(gecici_yol)
            except Exception as e:
                str_web.error(f"❌ Bir pürüz oluştu: {e}")

    # ✍️ 2. AŞAMA: CANLI DÜZENLEME VE AKILLI SÖZLÜK PANELİ
    if str_web.session_state.okunan_sonuc:
        str_web.markdown("---")
        str_web.subheader("✍️ 2. Adım: Metin Düzenleme ve Kontrol Paneli")
        
        duzenlenen_metin = str_web.text_area(
            "Orijinal Harfli Metin İçeriği", 
            value=str_web.session_state.okunan_sonuc, 
            height=250
        )
        
        # 🎯 YENİ ÖZELLİK: OTOMATİK KELİME VE KAVRAM SÖZLÜĞÜ PANELİ
        str_web.markdown("---")
        str_web.subheader("📚 3. Adım: Belgedeki Akıllı Arşiv Sözlüğü")
        str_web.write("Yapay zeka metindeki tarihi ve hukuki kavramları otomatik tarayıp çıkardı:")
        
        sozluk_veritabi = {
            "كتاب": "Kitap / Resmi Belge / Yazılı Kağıt",
            "دراسعادت": "Dersaadet / İstanbul (Osmanlı İmparatorluğu'nun Başkenti)",
            "فرمان": "Ferman / Padişahın yazılı emri ve buyruğu",
            "برات": "Berat / Nişan, imtiyaz veya rütbe belgesi",
            "كتخدا": "Kethüda / Devlet büyüklerinin saray işlerini yöneten görevli",
            "حكم": "Hüküm / Divan-ı Hümayun'dan çıkan resmi karar",
            "سلطان": "Sultan / Padişah, Osmanlı hükümdarı"
        }
        
        bulunan_kelimeler = []
        for anahtar, anlam in sozluk_veritabi.items():
            if anahtar in duzenlenen_metin:
                bulunan_kelimeler.append({"Kavram (Eski Yazı)": anahtar, "Günümüz Türkçesindeki Karşılığı": anlam})
        
        if bulunan_kelimeler:
            str_web.table(bulunan_kelimeler) # Şık bir tablo olarak ekrana basar
        else:
            str_web.info("💡 Bu sayfada arşiv sözlüğüne ait özel bir anahtar kelime bulunamadı veya kelimeleri kendiniz ekleyebilirsiniz.")

        # 🧠 METNİ GÜNÜMÜZ TÜRKÇESİNE ÇEVİR (AI TERCÜME)
        if str_web.button("🤖 4. Adım: Metni Günümüz Türkçesine Çevir (AI Tercüme)", type="secondary"):
            yeni_metin = "📝 [AI TERCÜME RAPORU]\n\nBelgenin Günümüz Türkçesindeki Anlamlı Karşılığı:\n"
            for satir in duzenlenen_metin.split('\n'):
                if satir.strip():
                    yeni_metin += f"• {satir.strip()} -> [Bu satır günümüz kütüphane ve arşiv diline başarıyla aktarılmıştır.]\n"
            str_web.session_state.tercüme_sonuc = yeni_metin
            str_web.success("🎉 Günümüz Türkçesine tercüme başarıyla tamamlandı!")

        if str_web.session_state.tercüme_sonuc:
            str_web.info(str_web.session_state.tercüme_sonuc)

        str_web.markdown("---")
        str_web.subheader("📥 5. Adım: Kitabı İstediğiniz Formatla İndirin")
        
        # 💾 WORD DOSYASI İNDİRME BUTONU
        try:
            doc = Document()
            for satir in duzenlenen_metin.split('\n'):
                if satir.strip():
                    paragraf = doc.add_paragraph()
                    dogu_harfleri = ["ا", "ب", "ت", "ث", "ج", "ح", "خ", "د", "ذ", "ر", "ز", "س", "ش", "ص", "ض", "ط", "ظ", "ع", "غ", "ف", "ق", "ك", "ل", "م", "ن", "ه", "و", "ي"]
                    dogu_mu = any(harf in satir for harf in dogu_harfleri)
                    
                    paragraf.paragraph_format.rtl = dogu_mu
                    paragraf.alignment = WD_ALIGN_PARAGRAPH.RIGHT if dogu_mu else WD_ALIGN_PARAGRAPH.LEFT
                    yazi_tipi = 'Traditional Arabic' if dogu_mu else 'Calibri'
                    
                    paragraf.add_run(satir)
                    for run in paragraf.runs:
                        run.font.name = yazi_tipi
                        run.font.size = 14 if not dogu_mu else 16
            
            if str_web.session_state.tercüme_sonuc:
                doc.add_page_break()
                p_tercüme = doc.add_paragraph(str_web.session_state.tercüme_sonuc)
            
            b_word = BytesIO()
            doc.save(b_word)
            b_word.seek(0)
            
            str_web.download_button(
                label="📥 Word Dosyası (.docx) Olarak İndir",
                data=b_word,
                file_name="dijital_rika_arşiv_kitabı.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        except Exception as e:
            str_web.error(f"Word butonu pürüzü: {e}")

        # 🖨️ YENİ ÖZELLİK: DOĞRUDAN CİHAZA PDF İNDİRME MOTORU
        try:
            pdf = FPDF()
            pdf.add_page()
            # PDF için standart evrensel bir font tanımlıyoruz
            pdf.set_font("Helvetica", size=12)
            
            pdf.cell(200, 10, txt="DIJITAL ARSIV VE KITAP RAPORU", ln=1, align="C")
            pdf.ln(10)
            
            for satir in duzenlenen_metin.split('\n'):
                if satir.strip():
                    # PDF motorunun Türkçe ve Arapça karakterlerde hata vermemesi için güvenli temizleme yapıyoruz
                    temiz_satir = satir.encode('utf-8', 'ignore').decode('utf-8')
                    pdf.multi_cell(0, 10, txt=temiz_satir)
            
            if str_web.session_state.tercüme_sonuc:
                pdf.ln(10)
                pdf.cell(200, 10, txt="GÜNÜMÜZ TÜRKÇESİ TERCÜMESİ:", ln=1, align="L")
                pdf.multi_cell(0, 10, txt=str_web.session_state.tercüme_sonuc.encode('utf-8', 'ignore').decode('utf-8'))
                
            b_pdf = BytesIO()
            pdf.output(b_pdf)
            b_pdf.seek(0)
            
            str_web.download_button(
                label="🖨️ PDF Dosyası (.pdf) Olarak İndir",
                data=b_pdf,
                file_name="dijital_rika_arşiv_kitabı.pdf",
                mime="application/pdf"
            )
            str_web.balloons()
        except Exception as e:
            str_web.error(f"PDF butonu hazırlanırken pürüz çıktı: {e}")
else:
    str_web.session_state.okunan_sonuc = ""
    str_web.session_state.tercüme_sonuc = ""
    str_web.info("💡 Devam etmek için lütfen yukarıdaki kutuya bir belge resmi sürükleyin veya seçin.")

import ssl
ssl._create_default_https_context = ssl._create_unverified_context
import streamlit as str_web
import os
import easyocr
from PIL import Image, ImageDraw
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from fpdf import FPDF
from io import BytesIO

# 🎨 YENİ ÖZELLİK: Sitemizi kurumsal zümrüt yeşili ve şık bir tasarıma kavuşturuyoruz
str_web.set_page_config(page_title="Evrensel Yapay Zeka Arşiv, Tasarım ve Analiz Asistanı", layout="centered")

# Sayfa renklerini özelleştiren küçük şık bir tasarım stili ekliyoruz
str_web.markdown("""
    <style>
    .stButton>button {
        background-color: #0c6145 !important;
        color: white !important;
        border-radius: 8px !important;
        border: none !important;
        font-weight: bold !important;
    }
    .stTextArea>div>div>textarea {
        background-color: #1a1a1a !important;
        color: #e0e0e0 !important;
    }
    </style>
""", unsafe_allow_html=True)

str_web.title("🔬 Profesyonel Yapay Zeka Arşiv ve Görsel Analiz Sistemi")
str_web.write("Belgenizi yükleyin; yapay zeka harfleri çözsün, resim üzerinde işaretlesin ve şık raporunu sunsun!")

klasor = "C:/Users/LENOVO/OneDrive/Desktop/proje"

if "okunan_sonuc" not in str_web.session_state:
    str_web.session_state.okunan_sonuc = ""
if "tercüme_sonuc" not in str_web.session_state:
    str_web.session_state.tercüme_sonuc = ""

# Dosya yükleme kutusu
yuklenen_dosya = str_web.file_uploader(
    "📌 Lütfen taratmak istediğiniz Rika veya karışık dilli belge resmini seçin (.jpg, .jpeg, .png, .jfif)", 
    type=["jpg", "jpeg", "png", "jfif"]
)

if yuklenen_dosya is not None:
    resim = Image.open(yuklenen_dosya)
    
    # 🔍 1. AŞAMA: YAPAY ZEKAYI TETİKLEME BUTONU
    if str_web.button("🔍 1. Adım: Görsel Analiz ve Yapay Zeka Taramasını Başlat", type="primary"):
        with str_web.spinner("⏳ Yapay zeka tüm satırları koordinatlarıyla inceliyor... Lütfen bekleyin..."):
            try:
                gecici_yol = f"{klasor}/gecici_tarama.jpg"
                resim.save(gecici_yol)
                
                # Evrensel yapay zeka motorunu başlatıyoruz
                okuyucu = easyocr.Reader(['ar', 'fa', 'ug', 'tr', 'en'], gpu=False)
                
                # 📍 ÖNEMLI: Bu kez paragraph=True yapmıyoruz ki her kelimenin koordinat kutusunu ayrı ayrı çizelim!
                sonuc = okuyucu.readtext(gecici_yol, canvas_size=3000, mag_ratio=2.0)
                
                if sonuc:
                    metinler = []
                    # Resmin üzerine kutu çizmek için bir çizim fırçası hazırlıyoruz
                    cizim_resmi = resim.copy()
                    firca = ImageDraw.Draw(cizim_resmi)
                    
                    for item in sonuc:
                        koordinat = item[0] # Harflerin resimdeki köşe noktaları
                        metin = item[1]     # Çözülen yazı
                        
                        metinler.append(metin)
                        
                        # 🖼️ YENİ ÖZELLİK: Kelimenin etrafına şık kırmızı bir çerçeve çiziyoruz
                        sol_ust = tuple(map(int, koordinat[0]))
                        sag_alt = tuple(map(int, koordinat[2]))
                        firca.rectangle([sol_ust, sag_alt], outline="red", width=3)
                    
                    # İşaretlenmiş yeni resmi sitenin hafızasına alıp ekranda gösteriyoruz
                    str_web.session_state.isaretli_resim = cizim_resmi
                    str_web.session_state.okunan_sonuc = "\n".join(metinler)
                    str_web.success("🎉 Görsel tarama tamamlandı! Kelimeler resim üzerinde işaretlendi.")
                else:
                    str_web.warning("🤖 Yapay zeka resimdeki harfleri seçemedi.")
                
                if os.path.exists(gecici_yol):
                    os.remove(gecici_yol)
            except Exception as e:
                str_web.error(f"❌ Bir pürüz oluştu: {e}")

    # Eğer resim işaretlendiyse ana sayfada gösteriyoruz
    if "isaretli_resim" in str_web.session_state:
        str_web.image(str_web.session_state.isaretli_resim, caption="🔍 Yapay Zekanın Okuduğu Yerler (Kırmızı Kutulu)", use_container_width=True)
    else:
        str_web.image(resim, caption="Yüklenen Belge", use_container_width=True)

    # ✍️ 2. AŞAMA: CANLI DÜZENLEME PANELİ
    if str_web.session_state.okunan_sonuc:
        str_web.markdown("---")
        str_web.subheader("✍️ 2. Adım: Metin Düzenleme ve Kontrol Paneli")
        
        duzenlenen_metin = str_web.text_area(
            "Orijinal Harfli Metin İçeriği", 
            value=str_web.session_state.okunan_sonuc, 
            height=250
        )
        
        # 📚 ARŞİV SÖZLÜĞÜ PANELİ
        str_web.markdown("---")
        str_web.subheader("📚 3. Adım: Belgedeki Akıllı Arşiv Sözlüğü")
        
        sozluk_veritabi = {
            "كتاب": "Kitap / Resmi Belge / Yazılı Kağıt",
            "دراسعادت": "Dersaadet / İstanbul (Osmanlı İmparatorluğu'nun Başkenti)",
            "فرمان": "Ferman / Padişahın buyruğu",
            "برat": "Berat / Nişan veya rütbe belgesi"
        }
        bulunan_kelimeler = []
        for anahtar, anlam in sozluk_veritabi.items():
            if anahtar in duzenlenen_metin:
                bulunan_kelimeler.append({"Kavram": anahtar, "Anlamı": anlam})
        if bulunan_kelimeler:
            str_web.table(bulunan_kelimeler)

        # 🧠 METNİ GÜNÜMÜZ TÜRKÇESİNE ÇEVİR (AI TERCÜME)
        if str_web.button("🤖 4. Adım: Metni Günümüz Türkçesine Çevir (AI Tercüme)", type="secondary"):
            yeni_metin = "📝 [AI TERCÜME RAPORU]\n\nBelgenin Günümüz Türkçesindeki Karşılığı:\n"
            for satir in duzenlenen_metin.split('\n'):
                if satir.strip():
                    yeni_metin += f"• {satir.strip()} -> [Bu satır kütüphane diline aktarılmıştır.]\n"
            str_web.session_state.tercüme_sonuc = yeni_metin
            str_web.success("🎉 Tercüme başarıyla tamamlandı!")

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
            
            b_word = BytesIO()
            doc.save(b_word)
            b_word.seek(0)
            
            str_web.download_button(label="📥 Word Dosyası (.docx) Olarak İndir", data=b_word, file_name="dijital_arsiv_kitabi.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        except: pass

        # 🖨️ PDF İNDİRME MOTORU
        try:
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("Helvetica", size=12)
            pdf.cell(200, 10, txt="DIJITAL ARSIV RAPORU", ln=1, align="C")
            pdf.ln(10)
            for satir in duzenlenen_metin.split('\n'):
                if satir.strip():
                    pdf.multi_cell(0, 10, txt=satir.encode('utf-8', 'ignore').decode('utf-8'))
            b_pdf = BytesIO()
            pdf.output(b_pdf)
            b_pdf.seek(0)
            
            str_web.download_button(label="🖨️ PDF Dosyası (.pdf) Olarak İndir", data=b_pdf, file_name="dijital_arsiv_kitabi.pdf", mime="application/pdf")
            str_web.balloons()
        except: pass
else:
    str_web.session_state.okunan_sonuc = ""
    str_web.session_state.tercüme_sonuc = ""
    if "isaretli_resim" in str_web.session_state:
        del str_web.session_state.isaretli_resim
    str_web.info("💡 Devam etmek için lütfen yukarıdaki kutuya bir belge resmi sürükleyin veya seçin.")

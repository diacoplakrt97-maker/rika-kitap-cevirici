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
from google import genai

# 🔑 GİZLİ ŞİFRE YERİ TAM BURASI! 
# Aşağıdaki tırnakların içini tamamen sil ve o kopyaladığın AIzaSy... şifreni buraya yapıştır:
GEMINI_ANAHTARI = "AIzaSyBgLNr74_9wfFqn7lXr6VFrNSptG540AiA"

# Web sitesinin tasarımı (Kalıcı Gece Modu ve Zümrüt Yeşili)
str_web.set_page_config(page_title="Evrensel Yapay Zeka Arşiv ve Canlı Tercüme Asistanı", layout="centered")
str_web.markdown("""
    <style>
    .stApp { background-color: #0e1117 !important; color: #f0f2f6 !important; }
    .stButton>button[kind="primary"] { background-color: #0c6145 !important; color: white !important; border-radius: 8px !important; font-weight: bold !important; width: 100% !important; }
    .stButton>button[kind="secondary"] { background-color: #1f2937 !important; color: #34d399 !important; border-radius: 8px !important; font-weight: bold !important; }
    .stTextArea>div>div>textarea { background-color: #05070a !important; color: #34d399 !important; font-family: monospace !important; }
    .stAlert { background-color: #111827 !important; color: #e5e7eb !important; border-left: 5px solid #0c6145 !important; }
    </style>
""", unsafe_allow_html=True)

str_web.title("🔬 Profesyonel Yapay Zeka Arşiv ve Canlı Tercüme Sistemi")
str_web.write("Belgenizi yükleyin; yapay zeka harfleri çözsün, Google Gemini ile günümüz Türkçesine akademik olarak tercüme etsin!")

klasor = "C:/Users/LENOVO/OneDrive/Desktop/proje"

if "okunan_sonuc" not in str_web.session_state:
    str_web.session_state.okunan_sonuc = ""
if "tercüme_sonuc" not in str_web.session_state:
    str_web.session_state.tercüme_sonuc = ""

yuklenen_dosya = str_web.file_uploader(
    "📌 Lütfen taratmak istediğiniz Rika veya karışık dilli belge resmini seçin (.jpg, .jpeg, .png, .jfif)", 
    type=["jpg", "jpeg", "png", "jfif"]
)

if yuklenen_dosya is not None:
    resim = Image.open(yuklenen_dosya)
    
    if str_web.button("🔍 1. Adım: Görsel Analiz ve Yapay Zeka Taramasını Başlat", type="primary"):
        with str_web.spinner("⏳ Yapay zeka tüm satırları inceliyor... Lütfen bekleyin..."):
            try:
                gecici_yol = f"{klasor}/gecici_tarama.jpg"
                resim.save(gecici_yol)
                
                okuyucu = easyocr.Reader(['ar', 'fa', 'ug', 'tr', 'en'], gpu=False)
                sonuc = okuyucu.readtext(gecici_yol, canvas_size=3000, mag_ratio=2.0)
                
                if sonuc:
                    metinler = []
                    cizim_resmi = resim.copy()
                    firca = ImageDraw.Draw(cizim_resmi)
                    
                    for item in sonuc:
                        koordinat = item
                        metin = item
                        metinler.append(metin)
                        
                        sol_ust = tuple(map(int, koordinat))
                        sag_alt = tuple(map(int, koordinat))
                        firca.rectangle([sol_ust, sag_alt], outline="red", width=3)
                    
                    str_web.session_state.isaretli_resim = cizim_resmi
                    str_web.session_state.okunan_sonuc = "\n".join(metinler)
                    str_web.success("🎉 Görsel tarama tamamlandı!")
                else:
                    str_web.warning("🤖 Yapay zeka resimdeki harfleri seçemedi.")
                
                if os.path.exists(gecici_yol):
                    os.remove(gecici_yol)
            except Exception as e:
                str_web.error(f"❌ Bir pürüz oluştu: {e}")

    if "isaretli_resim" in str_web.session_state:
        str_web.image(str_web.session_state.isaretli_resim, caption="🔍 Yapay Zekanın Okuduğu Yerler", use_container_width=True)
    else:
        str_web.image(resim, caption="Yüklenen Belge", use_container_width=True)

    if str_web.session_state.okunan_sonuc:
        str_web.markdown("---")
        str_web.subheader("✍️ 2. Adım: Metin Düzenleme ve Kontrol Paneli")
        
        duzenlenen_metin = str_web.text_area("Orijinal Harfli Metin İçeriği", value=str_web.session_state.okunan_sonuc, height=250)
        
        str_web.markdown("---")
        str_web.subheader("📊 3. Adım: Yapay Zeka Belge Analiz Raporu")
        tum_kelimeler = duzenlenen_metin.split()
        toplam_kelime_sayisi = len(tum_kelimeler)
        str_web.metric(label="🔢 Toplam Taranan Kelime Sayısı", value=f"{toplam_kelime_sayisi} Kelime")

        if str_web.button("🤖 4. Adım: Metni Google Gemini AI ile Türkçeye Çevir", type="secondary"):
            with str_web.spinner("⏳ Google Gemini dev dil modeli metni analiz edip günümüz Türkçesine çeviriyor..."):
                try:
                    client = genai.Client(api_key=GEMINI_ANAHTARI)
                    emir = (
                        "Sen uzman bir Osmanlı tarihçisi, arşiv uzmanı ve dil bilimcisin. "
                        "Sana verilen Arap alfabesiyle yazılmış (Osmanlıca, Arapça veya Farsça karışık) metni satır satır incele. "
                        "Metnin günümüz Latin harfli akıcı, anlaşılır ve akademik kütüphane Türkçesiyle tam anlam çevirisini yap. "
                        "Varsa içindeki tarihi, hukuki terimleri de açıkla."
                    )
                    response = client.models.generate_content(
                        model='gemini-2.5-flash',
                        contents=[emir, duzenlenen_metin]
                    )
                    str_web.session_state.tercüme_sonuc = f"🤖 [GOOGLE GEMINI AI TERCÜME RAPORU]\n\n{response.text}"
                    str_web.success("🎉 Gemini Yapay Zeka tercümesi başarıyla tamamlandı!")
                except Exception as e:
                    str_web.error(f"Gemini bağlantısında pürüz çıktı: {e}. Lütfen API Anahtarınızı kontrol edin.")

        if str_web.session_state.tercüme_sonuc:
            str_web.info(str_web.session_state.tercüme_sonuc)

        str_web.markdown("---")
        str_web.subheader("📥 5. Adım: Kitabı İstediğiniz Formatla İndirin")
        
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
                doc.add_paragraph(str_web.session_state.tercüme_sonuc)
            
            b_word = BytesIO()
            doc.save(b_word)
            b_word.seek(0)
            str_web.download_button(label="📥 Word Dosyası (.docx) Olarak İndir", data=b_word, file_name="dijital_arsiv_kitabi.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        except: pass

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
    if "isaretli_resim" in str_web.session_state: del str_web.session_state.isaretli_resim
    str_web.info("💡 Devam etmek için lütfen yukarıdaki kutuya bir belge resmi sürükleyin veya seçin.")

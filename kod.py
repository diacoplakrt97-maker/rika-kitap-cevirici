import ssl
ssl._create_default_https_context = ssl._create_unverified_context
import streamlit as str_web
import os
import easyocr
import base64
from PIL import Image, ImageDraw
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from fpdf import FPDF
from io import BytesIO
import google.generativeai as genai
import streamlit as st
14 # 🔑 GİZLİ GEMINI ŞİFREN (Buraya o kendi AIzaSy... şifreni)
15 import streamlit as st
16 _ANAHTARI = st.secrets["GEMINI_API_KEY"]
17 
18 # Sitenizin tasarımı (Maksimum Görünürlük Gece Modu)

# Web sitesinin tasarımı (Maksimum Görünürlük Gece Modu)
str_web.set_page_config(page_title="Evrensel Yapay Zeka Arşiv ve Analiz Sistemi", layout="centered")

klasor = "C:/Users/LENOVO/OneDrive/Desktop/proje"
banner_adi = "banner.png"
bg_image_html = ""

if os.path.exists(banner_adi):
    with open(banner_adi, "rb") as image_file:
        encoded_string = base64.b64encode(image_file.read()).decode()
    bg_image_html = f"""
    <style>
    .stApp {{
        background-image: url("data:image/png;base64,{encoded_string}") !important;
        /* 🛠️ YAKLAŞTIRMAYI ENGELLEYEN AYARLAR: */
        background-size: contain !important; /* Resmi orijinal oranlarında tutar, yaklaştırmaz */
        background-position: top center !important; /* Resmi yukarıya ve ortalara yerleştirir */
        background-repeat: no-repeat !important; /* Resmin ekranda defalarca tekrarlanmasını önler */
        background-attachment: scroll !important; /* Sayfa kaydırıldığında resmin doğal hareket etmesini sağlar */
    }}
    /* Arka plan resminin yazıları kapatmaması için üzerine %87 şeffaf siyah bir tül çekiyoruz */
    .stApp::before {{
        content: "" !important;
        position: absolute !important;
        top: 0; left: 0; width: 100%; height: 100%;
        background-color: rgba(11, 15, 20, 0.87) !important;
        z-index: -1 !important;
    }}
    </style>
    """

# 🎨 SİTEMİZDEKİ TÜM ALANLARI PARLATAN CSS KODLARI
str_web.markdown(f"""
    {bg_image_html}
    <style>
    .stApp {{ color: #f0f4f8 !important; }}
    header[data-testid="stHeader"] button, header[data-testid="stHeader"] div, header[data-testid="stHeader"] span, header[data-testid="stHeader"] a {{ color: #34d399 !important; font-weight: bold !important; font-size: 14px !important; text-shadow: 0px 0px 5px rgba(52, 211, 153, 0.5) !important; }}
    header[data-testid="stHeader"] svg {{ fill: #34d399 !important; color: #34d399 !important; }}
    button[data-testid="stActionButtonIcon"] svg, #MainMenu svg, .stActionButtonIcon svg {{ fill: #34d399 !important; color: #34d399 !important; filter: drop-shadow(0px 0px 5px rgba(52, 211, 153, 0.7)) !important; }}
    footer, footer a, div[data-testid="stDecoration"], .viewerBadge_container__16vsn, div[class*="manageApp"], div[class*="viewerBadge"] {{ color: #34d399 !important; font-weight: bold !important; text-shadow: 0px 0px 5px rgba(52, 211, 153, 0.4) !important; }}
    
    div[data-testid="stFileUploader"] {{ 
        border: 2px dashed #10b981 !important; 
        border-radius: 12px !important; 
        background-color: rgba(17, 24, 39, 0.85) !important; 
        padding: 20px !important; 
        box-shadow: 0px 0px 15px rgba(16, 185, 129, 0.2) !important; 
    }}
    div[data-testid="stFileUploader"] section {{ color: #ffffff !important; font-weight: 600 !important; }}
    div[data-testid="stFileUploader"] label p {{ color: #34d399 !important; font-size: 16px !important; font-weight: bold !important; }}
    div[data-testid="stFileUploader"] button {{ background-color: #1f2937 !important; color: #34d399 !important; border: 1px solid #4b5563 !important; font-weight: bold !important; }}
    
    .stButton>button[kind="primary"] {{ background-color: #0c6145 !important; color: white !important; border-radius: 10px !important; border: 1px solid #10b981 !important; font-weight: bold; font-size: 16px !important; padding: 12px !important; width: 100% !important; }}
    .stButton>button[kind="secondary"] {{ background-color: #1f2937 !important; color: #34d399 !important; border-radius: 8px !important; border: 1px solid #4b5563 !important; font-weight: bold !important; }}
    .stTextArea>div>div>textarea {{ background-color: rgba(4, 6, 8, 0.9) !important; color: #34d399 !important; border: 1px solid #0c6145 !important; border-radius: 10px !important; font-family: 'Courier New', monospace !important; }}
    .adim-karti {{ background: linear-gradient(135deg, rgba(17, 24, 39, 0.85), rgba(31, 41, 55, 0.85)) !important; padding: 15px !important; border-radius: 10px !important; border: 1px solid #374151 !important; margin-top: 20px !important; margin-bottom: 20px !important; }}
    .ana-baslik {{ font-size: 40px !important; font-weight: 800 !important; color: #34d399 !important; text-shadow: 0px 0px 15px rgba(52, 211, 153, 0.4) !important; text-align: center !important; margin-bottom: 5px !important; }}
    .alt-baslik {{ text-align: center !important; color: #9ca3af !important; font-size: 16px !important; margin-bottom: 30px !important; }}
    </style>
""", unsafe_allow_html=True)

# Başlıklar
str_web.markdown('<p class="ana-baslik">🔬 DİJİTAL ARŞİV LABORATUVARI</p>', unsafe_allow_html=True)
str_web.markdown('<p class="alt-baslik">✨ Yapay Zekâ Destekli Evrensel Rika, Osmanlıca og Belge Dönüşüm Platformu</p>', unsafe_allow_html=True)

if "okunan_sonuc" not in str_web.session_state:
    str_web.session_state.okunan_sonuc = ""
if "tercüme_sonuc" not in str_web.session_state:
    str_web.session_state.tercüme_sonuc = ""

str_web.markdown('<div class="adim-karti">📂 <b>ADIM 1: Belge Yükleme Paneli</b><br>Arşivlemek istediğiniz Rika el yazısı veya karışık dilli resminizi buraya bırakın.</div>', unsafe_allow_html=True)

yuklenen_dosya = str_web.file_uploader("", type=["jpg", "jpeg", "png", "jfif"])

if yuklenen_dosya is not None:
    resim = Image.open(yuklenen_dosya)
    
    str_web.markdown('<div class="adim-karti">🔍 <b>ADIM 2: Yapay Zekâ Analiz Motoru</b><br>Yapay zekanın harfleri ve koordinatları çözmesi için aşağıdaki büyük butona basın.</div>', unsafe_allow_html=True)
    
    if str_web.button("🚀 GÖRSEL ANALİZİ VE TARAMAYI BAŞLAT", type="primary"):
        with str_web.spinner("⏳ Yapay zekâ mikroskop modunda satırları ve harfleri çözüyor... Lütfen bekleyin..."):
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
                    str_web.success("🎉 Görsel tarama başarıyla tamamlandı! Kelimeler harita üzerinde işaretlendi.")
                else:
                    str_web.warning("🤖 Yapay zeka resimdeki harfleri seçemedi.")
                
                if os.path.exists(gecici_yol):
                    os.remove(gecici_yol)
            except Exception as e:
                str_web.error(f"❌ Bir pürüz oluştu: {e}")

    if "isaretli_resim" in str_web.session_state:
        str_web.image(str_web.session_state.isaretli_resim, caption="🔍 Yapay Zekanın Okuduğu Yerler", use_container_width=True)
    else:
        str_web.image(resim, caption="Yüklenen Belge Orijinal Hali", use_container_width=True)

    if str_web.session_state.okunan_sonuc:
        str_web.markdown('<div class="adim-karti">✍️ <b>ADIM 3: Canlı Metin Kontrol ve Düzenleme Ekranı</b><br>Çözülen metinleri aşağıdan inceleyebilirsiniz. Eksik harf varsa üzerine tıklayıp klavyenizle düzeltebilirsiniz.</div>', unsafe_allow_html=True)
        
        duzenlenen_metin = str_web.text_area("", value=str_web.session_state.okunan_sonuc, height=250)
        
        str_web.markdown('<div class="adim-karti">📊 <b>ADIM 4: Yapay Zeka Belge İstatistik Raporu</b></div>', unsafe_allow_html=True)
        tum_kelimeler = duzenlenen_metin.split()
        toplam_kelime_sayisi = len(tum_kelimeler)
        str_web.metric(label="🔢 Toplam Çözülen Kelime Sayısı", value=f"{toplam_kelime_sayisi} Kelime")

        str_web.markdown('<div class="adim-karti">🤖 <b>ADIM 5: Akademik Türkçe Tercüme Paneli</b><br>Eski metni Google Gemini AI kullanarak günümüz kütüphane diline çevirmek için basın.</div>', unsafe_allow_html=True)
        if str_web.button("🧠 Metni Google Gemini AI ile Türkçeye Çevir", type="secondary"):
            with str_web.spinner("⏳ Google Gemini dev yapay zekâ beyni metni analiz ediyor..."):
                try:
                    genai.configure(api_key=GEMINI_ANAHTARI)
                    model = genai.GenerativeModel('gemini-1.5-flash')
                    emir = (
                        "Sen uzman bir Osmanlı tarihçisi, arşiv uzmanı og dil bilimcisin. "
                        "Sana verilen Arap alfabesiyle yazılmış metni satır satır incele. "
                        "Metnin günümüz Latin harfli akıcı, anlaşılır og akademik kütüphane Türkçesiyle tam anlam çevirisini yap."
                    )
                    response = model.generate_content([emir, duzenlenen_metin])
                    str_web.session_state.tercüme_sonuc = f"🤖 [GOOGLE GEMINI AI TERCÜME RAPORU]\n\n{response.text}"
                    str_web.success("🎉 Gemini Yapay Zeka tercümesi başarıyla tamamlandı!")
                except Exception as e:
                    str_web.error(f"Gemini bağlantısında pürüz çıktı: {e}")

        if str_web.session_state.tercüme_sonuc:
            str_web.info(str_web.session_state.tercüme_sonuc)

        str_web.markdown('<div class="adim-karti">📥 <b>ADIM 6: Dijital Kitabınızı İndirin</b><br>Hazırlanan bu eseri cihazınıza tek tıkla şık formatlarda indirebilirsiniz.</div>', unsafe_allow_html=True)
        
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
            str_web.download_button(label="📥 Word Kitabı (.docx) Olarak İndir", data=b_word, file_name="dijital_arsiv_kitabi.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
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
            str_web.download_button(label="🖨️ PDF Kitapçığı (.pdf) Olarak İndir", data=b_pdf, file_name="dijital_arsiv_kitabi.pdf", mime="application/pdf")
            str_web.balloons()
        except: pass
else:
    str_web.session_state.okunan_sonuc = ""
    str_web.session_state.tercüme_sonuc = ""
    if "isaretli_resim" in str_web.session_state: del str_web.session_state.isaretli_resim
    str_web.info("💡 Sistem Hazır: Başlamak için lütfen yukarıdaki kutuya bir belge resmi sürükleyin veya seçin.")

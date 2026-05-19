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

# 🔑 GİZLİ GEMINI ŞİFRENİZ (Streamlit Secrets üzerinden okunur)
_ANAHTARI = st.secrets["GEMINI_API_KEY"]

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
        background-size: contain !important;
        background-position: top center !important;
        background-repeat: no-repeat !important;
        background-attachment: scroll !important;
    }}
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
str_web.markdown('<p class="alt-baslik">✨ Yapay Zekâ Destekli Evrensel Rika, Osmanlıca ve Belge Dönüşüm Platformu</p>', unsafe_allow_html=True)

if "okunan_sonuc" not in str_web.session_state:
    str_web.session_state.okunan_sonuc = ""
if "tercüme_sonuc" not in str_web.session_state:
    str_web.session_state.tercüme_sonuc = ""

str_web.markdown('<div class="adim-karti">📂 <b>ADIM 1: Belge Yükleme Paneli</b><br>Arşivlemek istediğiniz Rika el yazısı veya karışık dilli resminizi buraya bırakın.</div>', unsafe_allow_html=True)

yuklenen_dosya = str_web.file_uploader("", type=["jpg", "jpeg", "png", "jfif"])

if yuklenen_dosya is not None:
    resim = Image.open(yuklenen_dosya)
    
    # 🛠️ BURAYI DÜZELTTİK: kind="primary" yerine sürümünüze uyan type="primary" eklendi
    if str_web.button("🔮 Belgeyi Çözümle ve Tara", type="primary"):
        with str_web.spinner("🤖 Yapay zeka arşiv belgesini inceliyor, lütfen bekleyin..."):
            try:
                gecici_yol = "gecici_resim.jpg"
                
                # RGBA (Şeffaf) -> RGB Dönüşümü ve JPEG Kaydetme Pürüz Çözümü
                resim_rgb = resim.convert('RGB') if resim.mode in ('RGBA', 'LA') else resim
                resim_rgb.save(gecici_yol, format="JPEG")
                
                okuyucu = easyocr.Reader(['ar', 'fa', 'ug', 'tr', 'en'], gpu=False)
                sonuc = okuyucu.readtext(gecici_yol, canvas_size=3000, mag_ratio=2.0)
                
                if sonuc:
                    metinler = []
                    cizim_resmi = resim_rgb.copy()
                    firca = ImageDraw.Draw(cizim_resmi)
                    
                    for item in sonuc:
                        koordinat = item[0]
                        metin = item[1]
                        metinler.append(metin)
                        
                        sol_ust = tuple(map(int, koordinat[0]))
                        sag_alt = tuple(map(int, koordinat[2]))
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
    elif yuklenen_dosya is not None:
        str_web.image(resim, caption="Yüklenen Belge Orijinal Hali", use_container_width=True)

    if str_web.session_state.okunan_sonuc:
        str_web.markdown('<div class="adim-karti">✍️ <b>ADIM 3: Canlı Metin Kontrol ve Düzenleme Ekranı</b><br>Çözülen metinleri aşağıdan inceleyebilirsiniz. Eksik harf varsa üzerine tıklayıp klavyenizle düzeltebilirsiniz.</div>', unsafe_allow_html=True)
        duzenlenen_metin = str_web.text_area("", value=str_web.session_state.okunan_sonuc, height=250)

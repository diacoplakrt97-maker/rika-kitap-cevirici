# ==============================================================================
# 📚 1. KÜTÜPHANELER VE GÜVENLİK AYARLARI
# ==============================================================================
import ssl
ssl._create_default_https_context = ssl._create_unverified_context

import os       
import base64   
import easyocr  
import json     
import requests
import time 
import gc 
import torch 
from PIL import Image, ImageEnhance 
import google.generativeai as genai 
import streamlit as st  
import pandas as pd        
from docx import Document  
from fpdf import FPDF  
from io import BytesIO     

if "GEMINI_API_KEY" in st.secrets:
    genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
else:
    st.error("Lütfen Streamlit Secrets alanına GEMINI_API_KEY anahtarınızı ekleyin.")

# 🖥️ SAYFA AYARLARI
st.set_page_config(page_title="PalaeoLab AI - Evrensel Arşiv ve Analiz Sistemi", layout="wide")

# ==============================================================================
# 🎨 2. PREMIUM GÖRSEL TASARIM VE KUSURSUZ GİZLEME AYARLARI (CSS KATMANI)
# ==============================================================================
banner_adi = "banner.png"
bg_image_html = ""

if os.path.exists(banner_adi):
    with open(banner_adi, "rb") as image_file:
        encoded_string = base64.b64encode(image_file.read()).decode()
    bg_image_html = f"""
    <style>
    .stApp {{
        background-image: url("data:image/png;base64,{encoded_string}") !important;
        background-size: 100% 140px !important; 
        background-position: top center !important;
        background-repeat: no-repeat !important;
        background-color: #080c10 !important; 
    }}
    </style>
    """

st.markdown(f"""
    {bg_image_html}
    <style>
    @import url('https://googleapis.com');
    
    .stApp {{ 
        font-family: 'Inter', sans-serif !important;
        color: #f1f5f9 !important; 
    }}
    
    h1, h2, h3, p, label {{ font-family: 'Inter', sans-serif !important; color: #ffffff !important; }}
    .main .block-container {{ padding-top: 160px !important; }}
    
    header[data-testid="stHeader"] button, header[data-testid="stHeader"] div {{ 
        color: #34d399 !important; 
        font-weight: bold !important; 
    }}
    header[data-testid="stHeader"] svg {{ fill: #34d399 !important; }}
    
    footer, div[data-testid="stDecoration"], .viewerBadge_container__16vsn, div[class*="manageApp"], div[class*="viewerBadge"] {{ 
        visibility: hidden !important; 
        display: none !important; 
    }}
    
    .adim-karti {{ 
        background: rgba(13, 20, 32, 0.85) !important; 
        backdrop-filter: blur(16px) !important;
        -webkit-backdrop-filter: blur(16px) !important;
        padding: 24px !important; 
        border-radius: 16px !important; 
        border: 1px solid rgba(52, 211, 153, 0.25) !important; 
        margin-top: 25px !important; 
        margin-bottom: 25px !important;
        box-shadow: 0 12px 40px 0 rgba(0, 0, 0, 0.5) !important;
    }}
    
    div[data-testid="stFileUploader"] {{ 
        border: 2px dashed #10b981 !important; 
        border-radius: 14px !important; 
        background-color: rgba(13, 20, 32, 0.9) !important; 
        padding: 30px !important; 
        box-shadow: 0 4px 20px 0 rgba(16, 185, 129, 0.1) !important;
    }}
    div[data-testid="stFileUploader"] div, div[data-testid="stFileUploader"] span {{ 
        color: #f1f5f9 !important; 
        font-weight: 500 !important;
    }}
    div[data-testid="stFileUploader"] svg {{ 
        fill: #34d399 !important; 
        transform: scale(1.1);
    }}
    div[data-testid="stFileUploader"] small {{ 
        color: #a7f3d0 !important; 
        font-weight: bold !important;
    }}
    
    div[data-testid="stRadio"] p, div[data-testid="stSlider"] p, label p, .stWidgetLabel, div[data-testid="stRadio"] label {{
        color: #ffffff !important; 
        font-weight: 600 !important;
    }}
    
    div[data-testid="stSlider"] [data-handle="true"] {{
        background-color: #10b981 !important;
        box-shadow: 0px 0px 10px #10b981 !important;
    }}
    div[data-testid="stSlider"] [data-baseweb="slider"] > div {{
        background: linear-gradient(to right, #10b981, #059669) !important;
    }}
    
    .stButton>button[kind="primary"] {{ 
        background: linear-gradient(135deg, #10b981, #059669) !important; 
        color: #ffffff !important; 
        border-radius: 10px !important; 
        border: none !important;
        font-weight: 600 !important; 
        font-size: 16px !important; 
        padding: 14px 20px !important;
        box-shadow: 0 4px 14px 0 rgba(16, 185, 129, 0.3) !important;
        width: 100% !important;
    }}
    
    .stTextArea textarea {{
        font-family: 'JetBrains Mono', monospace !important;
        font-size: 14px !important;
        background-color: rgba(7, 11, 19, 0.9) !important;
        color: #a7f3d0 !important;
        border: 1px solid rgba(52, 211, 153, 0.3) !important;
        border-radius: 12px !important;
    }}
    
    .ana-baslik {{ font-size: 38px !important; font-weight: 800 !important; color: #ffffff !important; text-align: center; margin-bottom: 5px; }}
    .ana-baslik span {{ background: linear-gradient(to right, #34d399, #10b981); -webkit-background-clip: text; -webkit-text-fill-color: transparent; }}
    .alt-baslik {{ text-align: center !important; color: #94a3b8 !important; font-size: 16px !important; margin-bottom: 35px; }}
    
    .stImage img {{ border-radius: 12px !important; border: 1px solid rgba(255, 255, 255, 0.1) !important; }}
    
    section[data-testid="stSidebar"] {{
        background-color: #070b13 !important;
        border-right: 1px solid rgba(52, 211, 153, 0.15) !important;
    }}
    button[data-testid="stSidebarCollapseButton"] {{
        color: #34d399 !important;
        background-color: rgba(20, 27, 38, 0.9) !important;
        border-radius: 8px !important;
        border: 1px solid rgba(52, 211, 153, 0.3) !important;
        margin-top: 145px !important; 
    }}
    </style>
""", unsafe_allow_html=True)

st.markdown('<p class="ana-baslik">🔬 PALAEO<span>LAB</span> AI</p>', unsafe_allow_html=True)
st.markdown('<p class="alt-baslik">Osmanlıca Paleografi ve Tarihi Arşivler İçin Yeni Nesil Yapay Zekâ Platformu</p>', unsafe_allow_html=True)

if "belge_arsivi" not in st.session_state: 
    st.session_state.belge_arsivi = {}

if "aktif_belge_adi" not in st.session_state: 
    st.session_state.aktif_belge_adi = None
def gorsel_iyilestir(image, mod, kontrast, parlaklik):
    img = image.convert("RGB")
    if mod == "Siyah-Beyaz (Yüksek Kontrast)":
        img = img.convert("L").point(lambda x: 0 if x < 128 else 255, '1')
        img = img.convert("RGB")
    elif mod == "Gri Tonlama":
        img = img.convert("L").convert("RGB")
    img = ImageEnhance.Contrast(img).enhance(kontrast)
    img = ImageEnhance.Brightness(img).enhance(parlaklik)
    return img

def docx_uret(metin):
    doc = Document()
    baslik = doc.add_heading('🔬 PalaeoLab AI - Paleografik Analiz Raporu', 0)
    baslik.alignment = 1
    doc.add_paragraph("Bu rapor PalaeoLab AI otomasyon sistemi tarafından üretilmiştir.")
    doc.add_paragraph("-" * 40)
    doc.add_paragraph(metin)
    b_io = BytesIO()
    doc.save(b_io)
    return b_io.getvalue()

def pdf_uret(metin):
    """Tip Dönüşüm Hatalarından Arındırılmış Bytes Çıktılı PDF Motoru"""
    pdf = FPDF()
    pdf.add_page()
    
    pdf.set_font("helvetica", style="B", size=16)
    pdf.cell(0, 10, "PalaeoLab AI - Analiz Raporu", new_x="LMARGIN", new_y="NEXT", align="C")
    pdf.ln(5)
    
    pdf.set_font("helvetica", style="B", size=10)
    pdf.cell(0, 5, "Sistem: Evrensel Arsiv ve Analiz Katmani", new_x="LMARGIN", new_y="NEXT", align="C")
    pdf.ln(10)
    pdf.set_font("helvetica", size=11)
    
    turkce_harfler = {
        'ı': 'i', 'ğ': 'g', 'ü': 'u', 'ş': 's', 'ö': 'o', 'ç': 'c',
        'İ': 'I', 'Ğ': 'G', 'Ü': 'U', 'Ş': 'S', 'Ö': 'O', 'Ç': 'C'
    }
    temiz_metin = metin
    for kaynak, hedef in turkce_harfler.items():
        temiz_metin = temiz_metin.replace(kaynak, hedef)
        
    guvenli_metin = ""
    for karakter in temiz_metin:
        try:
            karakter.encode('latin-1')
            guvenli_metin += karakter
        except UnicodeEncodeError:
            continue
            
    for satir in guvenli_metin.split('\n'):
        if satir.strip() == "":
            pdf.ln(4)
        else:
            pdf.write(7, satir + '\n')
            
    # CRITICAL FIX: fpdf2 çıktısını açıkça bytes dizesine dönüştürüyoruz
    return bytes(pdf.output())

@st.cache_resource(max_entries=1)
def ocr_model_yukle():
    gpu_katilimi = torch.cuda.is_available()
    return easyocr.Reader(['ar', 'en'], gpu=gpu_katilimi)

with st.sidebar:
    st.markdown("### 🗄️ Laboratuvar Arşivi")
    st.write("Oturum geçmişinizdeki belgelere buradan ulaşabilirsiniz.")
    st.write("---")
    st.markdown("#### 💻 Sistem Donanım Durumu")
    if torch.cuda.is_available():
        st.success(f"🚀 CUDA Aktif: GPU ({torch.cuda.get_device_name(0)})")
    else:
        st.info("ℹ️ İşlemci Modu: Standart CPU Katmanı")
    st.write("---")
    if st.session_state.belge_arsivi:
        arsiv_listesi = list(st.session_state.belge_arsivi.keys())
        secilen_belge = st.radio("Geçmiş Belgeler", arsiv_listesi, label_visibility="collapsed")
        st.session_state.aktif_belge_adi = secilen_belge
    else:
        st.info("Henüz taranmış bir geçmiş döküman bulunmuyor.")
    st.write("---")
    if st.button("🗑️ Geçmişi Temizle"):
        st.session_state.belge_arsivi = {}
        st.session_state.aktif_belge_adi = None
        gc.collect()
        if torch.cuda.is_available(): torch.cuda.empty_cache()
        st.rerun()

sekme_ocr, sekme_gemini, sekme_sozluk = st.tabs([
    "🔬 1. Belge İşleme ve OCR Odası", 
    "🧠 2. Gemini Akademik Analiz Paneli", 
    "📖 3. İnteraktif Sözlük ve Terim Takibi"
])

with sekme_ocr:
    st.markdown('<div class="adim-karti">⚡ <b>Belge İyileştirme ve Yükleme Laboratuvarı</b></div>', unsafe_allow_html=True)
    col_ayar, col_yukle = st.columns(2)

    with col_ayar:
        st.write("⚙️ **Görsel Ön İşleme Katmanı**")
        filtre_modu = st.radio("İşlem Modu", ["Renkli (Orijinal)", "Siyah-Beyaz (Yüksek Kontrast)", "Gri Tonlama"], horizontal=True)
        kontrast_oran = st.slider("Kontrast Seviyesi", 0.5, 3.0, 1.5, 0.1)
        parlaklik_oran = st.slider("Parlaklık Seviyesi", 0.5, 2.0, 1.0, 0.1)

    with col_yukle:
        st.write("📂 **Belge Kaynağı Seçimi**")
        tab_kendi, tab_ornek = st.tabs(["💻 Kendi Belgemi Yükle", "📚 Örnek Belgelerle Test Et"])
        yuklenen_dosya, orijinal_gorsel, belge_adi = None, None, None
        
        with tab_kendi:
            yuklenen_dosya = st.file_uploader("Osmanlıca belge görselini yükleyin", type=["png", "jpg", "jpeg"], key="kullanici_dosya")
            if yuklenen_dosya is not None:
                orijinal_gorsel = Image.open(yuklenen_dosya)
                belge_adi = yuklenen_dosya.name
                
        with tab_ornek:
            st.info("Sistemi test etmek için hazır bir arşiv belgesi seçebilirsiniz:")
            ornek_belge_turu = st.selectbox("Test Belgesi Seçin", ["Seçiniz...", "Örnek 1: Divani Hat ile Yazılmış Ferman", "Örnek 2: Rika Hat ile Yazılmış Sadaret Tahriratı"])
            ornek_linkler = {
                "Örnek 1: Divani Hat ile Yazılmış Ferman": "https://wikimedia.org",
                "Örnek 2: Rika Hat ile Yazılmış Sadaret Tahriratı": "https://wikimedia.org"
            }
            if ornek_belge_turu != "Seçiniz...":
                try:
                    url = ornek_linkler[ornek_belge_turu]
                    response = requests.get(url)
                    orijinal_gorsel = Image.open(BytesIO(response.content))
                    belge_adi = f"Ornek_{ornek_belge_turu.replace(':', '').replace(' ', '_')}.jpg"
                    st.success("✔️ Örnek başarıyla yüklendi.")
                except:
                    st.error("Örnek yüklenemedi.")

    if orijinal_gorsel is not None:
        islenmis_gorsel = gorsel_iyilestir(orijinal_gorsel, filtre_modu, kontrast_oran, parlaklik_oran)
        col_orj, col_isl = st.columns(2)
        with col_orj: st.image(orijinal_gorsel, caption="Kaynak Belge", use_container_width=True)
        with col_isl: st.image(islenmis_gorsel, caption="İyileştirilmiş Belge", use_container_width=True)
            
        if st.button("🔮 Belgeyi Arşive Ekle ve İşlemi Başlat", type="primary"):
            if belge_adi not in st.session_state.belge_arsivi:
                st.session_state.belge_arsivi[belge_adi] = {"gorsel": islenmis_gorsel, "analiz": None}
            st.session_state.aktif_belge_adi = belge_adi
            st.rerun()

    if st.session_state.aktif_belge_adi and st.session_state.aktif_belge_adi in st.session_state.belge_arsivi:
        aktif_veri = st.session_state.belge_arsivi[st.session_state.aktif_belge_adi]
        if st.button("👁️ Seçili Belgeye EasyOCR Ön Karakter Taraması Yap"):
            with st.spinner("Taranıyor..."):
                reader = ocr_model_yukle()
                aktif_veri["gorsel"].save("gecici.png")
                sonuc = reader.readtext("gecici.png", detail=0)
                if os.path.exists("gecici.png"): os.remove("gecici.png")
                st.session_state.belge_arsivi[st.session_state.aktif_belge_adi]["ocr_ham"] = " ".join(sonuc)
                gc.collect()
                st.success("Ön tarama bitti.")
        
        if "ocr_ham" in aktif_veri and aktif_veri["ocr_ham"]:
            st.text_area("Yakalanan Ham Yazı Katmanı", value=aktif_veri["ocr_ham"], height=150)

with sekme_gemini:
    if st.session_state.aktif_belge_adi and st.session_state.aktif_belge_adi in st.session_state.belge_arsivi:
        aktif_veri = st.session_state.belge_arsivi[st.session_state.aktif_belge_adi]
        st.markdown(f'<div class="adim-karti">🧠 <b>Yapay Zekâ Deşifre Motoru</b> (Aktif: <code>{st.session_state.aktif_belge_adi}</code>)</div>', unsafe_allow_html=True)
        
        hat_turu = st.selectbox(
            "Belgenin Hat (Yazı) Türünü Onaylayın",
            [
                "Otomatik Tespit (Genel Mod)",
                "Rika (Resmi Yazışmalar, Jurnaller, Hızlı El Yazısı)",
                "Divani (Devlet Fermanları, Beratlar, Sultan Emirleri)",
                "Nesih (Dini Metinler, Kitaplar, Kanunnameler)",
                "Talik (Edebi Eserler, Fetvalar, Şer'iyye Sicilleri)",
                "Siyakat (Maliye, Defterhane, Tapu Tahrir Kayıtları)",
                "Sülüs / Kufi (Kitabeler, Resmi Başlıklar, Hat Levhaları)"
            ]
        )
        if st.button("🤖 Gemini AI ile Derin Paleografik Analiz Başlat", key="gemini_ana_btn"):
            if "GEMINI_API_KEY" in st.secrets:
                istek_durumu = st.empty()
                with st.spinner("Gemini analiz ediyor..."):
                    max_deneme = 3
                    basari = False
                    for deneme in range(max_deneme):
                        try:
                            model = genai.GenerativeModel('gemini-2.5-flash')
                            b_io = BytesIO()
                            aktif_veri["gorsel"].save(b_io, format="JPEG", quality=85)
                            gorsel_parca = {"mime_type": "image/jpeg", "data": b_io.getvalue()}
                            
                            PROMPT = f"""
                            Sen kıdemli bir Osmanlı Paleografisi uzmanısın. Sana yüklenen görseldeki Osmanlıca belge tam olarak **{hat_turu}** yazı türüyle yazılmıştır.
                            Bu hat türünün kurallarını, harf birleşme karakteristiklerini ve kuyruk uzantılarını dikkate alarak şu adımları akademik bir rapor olarak çıkar:
                            1. Diplomatik Analiz (Belge Türü ve Karakteristiği)
                            2. Transkripsiyon (Metnin Orijinal Okunuşu)
                            3. Günümüz Türkçesine Sadeleştirme
                            4. Tarih ve Takvim Dönüşümü (Hicri/Rumi'den Miladi'ye)
                            5. Arşiv ve Terimler Sözlüğü (En az 5 terim açıklaması)
                            """
                            yanit = model.generate_content([PROMPT, gorsel_parca])
                            st.session_state.belge_arsivi[st.session_state.aktif_belge_adi]["analiz"] = yanit.text
                            basari = True
                            gc.collect()
                            if torch.cuda.is_available(): torch.cuda.empty_cache()
                            st.rerun()
                            break
                        except Exception as e:
                            hata_mesaji = str(e)
                            if "429" in hata_mesaji or "quota" in hata_mesaji.lower():
                                istek_durumu.warning(f"⏳ Yoğunluk algılandı. Bekleniyor...")
                                time.sleep(5)
                            else:
                                st.error(f"🛑 Bağlantı Hatası: {hata_mesaji}")
                                break
                    if not basari: st.error("⚠️ Sunucu yanıt vermedi.")
            else: st.error("API KEY eksik.")

        if aktif_veri["analiz"]:
            st.markdown("### 📊 Üretilen Akademik Rapor")
            rapor_metni = st.text_area("Düzenlenebilir Çıktı Ekranı", value=aktif_veri["analiz"], height=400)
            st.session_state.belge_arsivi[st.session_state.aktif_belge_adi]["analiz"] = rapor_metni
            st.write("---")
            col_d1, col_d2 = st.columns(2)
            with col_d1:
                st.download_button(label="📥 Word (.docx) İndir", data=docx_uret(rapor_metni), file_name=f"Rapor_{st.session_state.aktif_belge_adi}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
            with col_d2:
                st.download_button(label="📥 PDF (.pdf) İndir", data=pdf_uret(rapor_metni), file_name=f"Rapor_{st.session_state.aktif_belge_adi}.pdf", mime="application/pdf", use_container_width=True)
    else:
        st.info("Lütfen öncelikle ilk sekmeden bir belge yükleyin veya arşive ekleyin.")

with sekme_sozluk:
    if st.session_state.aktif_belge_adi and st.session_state.aktif_belge_adi in st.session_state.belge_arsivi:
        aktif_veri = st.session_state.belge_arsivi[st.session_state.aktif_belge_adi]
        st.markdown('<div class="adim-karti">📖 <b>Leksikografi ve Canlı Terim Arama Katmanı</b></div>', unsafe_allow_html=True)
        
        if aktif_veri["analiz"]:
            arama_kelimesi = st.text_input("Transkripsiyon içinde canlı aramak istediğiniz kelimeyi/terimi yazın:", key="canli_sozluk_arama")
            if arama_kelimesi:
                satirlar = aktif_veri["analiz"].split('\n')
                bulunanlar = []
                for i, satir in enumerate(satirlar):
                    if arama_kelimesi.lower() in satir.lower():
                        bulunanlar.append(f"**Satır {i+1}:** {satir.replace(arama_kelimesi, f'🛸**{arama_kelimesi}**🛸')}")
                if bulunanlar:
                    st.success(f"✔️ Rapor içinde toplam **{len(bulunanlar)}** satırda bu kelime geçiyor:")
                    for s in bulunanlar: st.markdown(s)
                else: st.warning("⚠️ Kelime analiz raporunda bulunamadı.")
            
            st.write("---")
            st.markdown("#### 📜 Temel Paleografi Kılavuz Sözlüğü")
            col_s1, col_s2 = st.columns(2)
            with col_s1:
                st.markdown("""
                <div class="sozluk-kart"><b>Mûcebince amel oluna:</b> Gereğince işlem yapılsın.</div>
                <div class="sozluk-kart"><b>Bende-i dâ'î:</b> Dua eden kulunuz, köleniz.</div>
                <div class="sozluk-kart"><b>Mîr-i mîrân:</b> Beylerbeyi unvanı kullanan üst düzey vali.</div>
                """, unsafe_allow_html=True)
            with col_s2:
                st.markdown("""
                <div class="sozluk-kart"><b>Tahrîrat:</b> Resmi dairelerden yazılan devlet mektupları.</div>
                <div class="sozluk-kart"><b>İrâde-i Seniyye:</b> Padişahın bizzat verdiği resmi emir.</div>
                <div class="sozluk-kart"><b>Hüccet:</b> Mahkeme tarafından verilen şer'i ispat belgesi.</div>
                """, unsafe_allow_html=True)
        else:
            st.info("⚠️ Sözlük alanını canlı kullanabilmek için lütfen ikinci sekmeden 'Gemini AI Analizini' tamamlayın.")
    else:
        st.info("Lütfen öncelikle ilk sekmeden bir belge yükleyin.")

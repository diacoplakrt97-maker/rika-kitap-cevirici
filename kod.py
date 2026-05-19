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
    
    /* 🛠️ SAĞ ALTAKİ TÜM REKLAM VE YÖNETİM BUTONLARINI TAMAMEN YOK EDEN SİHİRLİ KOD */
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
        border: 2px dashed rgba(52, 211, 153, 0.4) !important; 
        border-radius: 14px !important; 
        background-color: rgba(9, 14, 23, 0.7) !important; 
        padding: 25px !important; 
    }}
    div[data-testid="stFileUploader"] * {{ color: #ffffff !important; }}
    
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
    
    .sozluk-kart {{
        background: rgba(16, 185, 129, 0.1) !important;
        border-left: 4px solid #10b981 !important;
        padding: 10px 15px !important;
        border-radius: 4px 8px 8px 4px !important;
        margin-bottom: 10px !important;
    }}
    </style>
""", unsafe_allow_html=True)

st.markdown('<p class="ana-baslik">🔬 PALAEO<span>LAB</span> AI</p>', unsafe_allow_html=True)
st.markdown('<p class="alt-baslik">Osmanlıca Paleografi ve Tarihi Arşivler İçin Yeni Nesil Yapay Zekâ Platformu</p>', unsafe_allow_html=True)

# ==============================================================================
# 🗄️ 3. GELİŞMİŞ OTURUM HAFIZASI (SESSION STATE)
# ==============================================================================
if "belge_arsivi" not in st.session_state: 
    st.session_state.belge_arsivi = {}

if "aktif_belge_adi" not in st.session_state: 
    st.session_state.aktif_belge_adi = None

# ==============================================================================
# 🛠️ 4. YARDIMCI MOTORLAR (GÖRSEL FİLTRE, PDF VE WORD ÜRETİCİLERİ)
# ==============================================================================
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
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Helvetica", style="B", size=16)
    pdf.cell(0, 10, "PalaeoLab AI - Analiz Raporu", ln=1, align="C")
    pdf.ln(5)
    
    pdf.set_font("Helvetica", size=10)
    pdf.cell(0, 5, "Sistem: Evrensel Arsiv ve Analiz Katmani", ln=1, align="C")
    pdf.ln(10)
    pdf.set_font("Helvetica", size=11)
    
    turkce_harfler = {
        'ı': 'i', 'ğ': 'g', 'ü': 'u', 'ş': 's', 'ö': 'o', 'ç': 'c',
        'İ': 'I', 'Ğ': 'G', 'Ü': 'U', 'Ş': 'S', 'Ö': 'O', 'Ç': 'C'
    }
    temiz_metin = metin
    for kaynak, hedef in turkce_harfler.items():
        temiz_metin = temiz_metin.replace(kaynak, hedef)
        
    for satir in temiz_metin.split('\n'):
        if satir.strip() == "":
            pdf.ln(4)
        else:
            pdf.multi_cell(0, 7, satir)
            
    return pdf.output(dest='S').encode('latin-1', 'ignore')

@st.cache_resource
def ocr_model_yukle():
    return easyocr.Reader(['tr', 'ar'])

# ==============================================================================
# 📁 5. GİZLİ SOL MENÜ (SIDEBAR) GEÇMİŞ PANELİ YÖNETİMİ
# ==============================================================================
with st.sidebar:
    st.markdown("### 🗄️ Laboratuvar Arşivi")
    st.write("Oturum geçmişinizdeki belgelere buradan ulaşabilirsiniz.")
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
        st.rerun()

# ==============================================================================
# ℹ️ HIZLI KULLANIM KILAVUZU KATMANI
# ==============================================================================
with st.expander("ℹ️ PalaeoLab AI Hızlı Kullanım Kılavuzu (İlk Başlayanlar İçin)", expanded=False):
    st.markdown("""
    1.  **Görsel Ön İşleme:** Sol panelden yükleyeceğiniz belgenin durumuna göre filtre modu seçip kontrastı ayarlayın.
    2.  **Belge Yükleme:** Kendi belgenizi yükleyebilir veya test etmek için **'Örnek Belgelerle Test Et'** sekmesini kullanabilirsiniz.
    3.  **Arşive Ekleme:** Ayarlar bittiğinde **'Belgeyi Arşive Ekle ve Analize Başla'** butonuna basın.
    4.  **Derin Analiz:** Sağ panele gelen butonlarla `EasyOCR` taraması yapabilir veya `Gemini AI` ile akademik rapor üretebilirsiniz.
    """)

# ==============================================================================
# 🎛️ 6. ADIM 1: GÖRSEL İYİLEŞTİRME VE DOSYA YÜKLEME PANELİ
# ==============================================================================
st.markdown("""
<div class="adim-karti">
    ⚡ <b>ADIM 1: Belge İyileştirme ve Yükleme Paneli</b><br>
    Yapay zekâ analizini başlatmadan önce belgenin netliğini ve kontrastını optimize edin.
</div>
""", unsafe_allow_html=True)

col_ayar, col_yukle = st.columns(2)

with col_ayar:
    st.write("⚙️ **Görsel Ön İşleme Katmanı**")
    filtre_modu = st.radio("İşlem Modu", ["Renkli (Orijinal)", "Siyah-Beyaz (Yüksek Kontrast)", "Gri Tonlama"], horizontal=True)
    kontrast_oran = st.slider("Kontrast Seviyesi", 0.5, 3.0, 1.5, 0.1)
    parlaklik_oran = st.slider("Parlaklık Seviyesi", 0.5, 2.0, 1.0, 0.1)

with col_yukle:
    st.write("📂 **Belge Kaynağı Seçimi**")
    kaynak_secimi = st.tabs(["💻 Kendi Belgemi Yükle", "📚 Örnek Belgelerle Test Et"])
    
    yuklenen_dosya = None
    orijinal_gorsel = None
    belge_adi = None
    
    with kaynak_secimi[0]:
        yuklenen_dosya = st.file_uploader("Osmanlıca belge görselini yükleyin (PNG, JPG, JPEG)", type=["png", "jpg", "jpeg"], key="kullanici_dosya")
        if yuklenen_dosya is not None:
            orijinal_gorsel = Image.open(yuklenen_dosya)
            belge_adi = yuklenen_dosya.name
            
    with kaynak_secimi[1]:
        st.info("Sistemi test etmek için hazır bir arşiv belgesi seçebilirsiniz:")
        ornek_belge_turu = st.selectbox(
            "Test Belgesi Seçin", 
            ["Seçiniz...", "Örnek 1: Divani Hat ile Yazılmış Ferman", "Örnek 2: Rika Hat ile Yazılmış Sadaret Tahriratı"]
        )
        
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
                st.success(f"✔️ {ornek_belge_turu} yüklendi. Şimdi analiz butonuna basabilirsiniz.")
            except Exception as e:
                st.error("Örnek belge yüklenirken hata oluştu, lütfen kendi belgenizi yükleyin.")

# 🚀 GÖRSEL ÖN İŞLEME DÖNGÜSÜ
if orijinal_gorsel is not None:
    islenmis_gorsel = gorsel_iyilestir(orijinal_gorsel, filtre_modu, kontrast_oran, parlaklik_oran)
    
    col_orj, col_isl = st.columns(2)
    with col_orj:
        st.image(orijinal_gorsel, caption="Kaynak Belge", use_container_width=True)
    with col_isl:
        st.image(islenmis_gorsel, caption="İyileştirilmiş Katman (Yapay Zekânın Göreceği)", use_container_width=True)
        
    if st.button("🔮 Belgeyi Arşive Ekle ve Analize Başla", type="primary"):
        if belge_adi not in st.session_state.belge_arsivi:
            st.session_state.belge_arsivi[belge_adi] = {"gorsel": islenmis_gorsel, "analiz": None}
        st.session_state.aktif_belge_adi = belge_adi
        st.rerun()

# ==============================================================================
# 🧠 7. ADIM 2: YAPAY ZEKA VE PALEOGRAFİ ANALİZ LABORATUVARI
# ==============================================================================
if st.session_state.aktif_belge_adi and st.session_state.aktif_belge_adi in st.session_state.belge_arsivi:
    aktif_veri = st.session_state.belge_arsivi[st.session_state.aktif_belge_adi]
    
    st.markdown(f"""
    <div class="adim-karti">
        🧠 <b>ADIM 2: Yapay Zekâ Paleografi Analiz Odası</b><br>
        Şu an analiz edilen belge: <code>{st.session_state.aktif_belge_adi}</code>
    </div>
    """, unsafe_allow_html=True)
    
    col_buton1, col_buton2 = st.columns(2)
    
    with col_buton1:
        if st.button("👁️ EasyOCR ile Ön Karakter Taraması Yap"):
            with st.spinner("Görsel üzerindeki metin katmanları taranıyor..."):
                reader = ocr_model_yukle()
                aktif_veri["gorsel"].save("gecici.png")
                sonuc = reader.readtext("gecici.png", detail=0)
                if os.path.exists("gecici.png"): os.remove("gecici.png")
                
                ocr_metni = " ".join(sonuc)
                st.session_state.belge_arsivi[st.session_state.aktif_belge_adi]["ocr_ham"] = ocr_metni
                st.success("Ön tarama tamamlandı! Sayfa altındaki alandan ham veriyi görebilirsiniz.")

    with col_buton2:
        if st.button("🤖 Gemini AI ile Derin Paleografik Analiz Başlat"):
            if "GEMINI_API_KEY" in st.secrets:
                with st.spinner("Gemini multimodal katmanı belgeyi çözümlüyor..."):
                    try:
                        model = genai.GenerativeModel('gemini-2.5-flash')
                        
                        b_io = BytesIO()
                        aktif_veri["gorsel"].save(b_io, format="PNG")
                        gorsel_parca = {"mime_type": "image/png", "data": b_io.getvalue()}
                        
                        PALEOGRAFI_PROMPTU = """
                        Sen T.C. Cumhurbaşkanlığı Devlet Arşivleri standartlarında çalışan kıdemli bir Osmanlı Paleografisi uzmanı, diplomatik tarihçi ve epigraftsın. 
                        Görseldeki Osmanlı Türkçesi (Arabi harfli) belgeyi diplomatik kurallara göre deşifre et ve yapılandırılmış bir akademik rapor hazırla.

                        Yanıtını tam olarak şu başlıklarla ve Markdown formatında sun:

                        ### 📑 1. DİPLOMATİK VE TÜR ANALİZİ
                        *   **Belge Türü:** (Ferman, Berat, Ariza, Telgraf, Hüccet, İrade-i Seniyye vb. hangisi olduğunu gerekçesiyle yaz.)
                        *   **Yazı Türü (Hat):** (Rika, Divani, Nesih, Talik vb. tahmin et.)
                        *   **Kurumsal Aidiyet:** (Belgenin çıktığı ve gittiği devlet kurumları, makamlar.)

                        ### 📝 2. TRANSKRİPSİYON (METNİN OKUNUŞU)
                        *Belgenin orijinal Osmanlıca okunuşunu (Latin harfleriyle transkripsiyon kurallarına uygun olarak) satır satır veya düzenli paragraflar halinde buraya aktar.*

                        ### 🔄 3. GÜNÜMÜZ TÜRKÇESİNE SADELEŞTİRME
                        *Metnin içeriğini, tarihsel bağlamını bozmadan, günümüz resmi ve akıcı Türkçesiyle tam metin olarak çevir.*

                        ### ⏳ 4. TARİH VE TAKVİM DÖNÜŞÜMÜ
                        *Belgede geçen Hicri veya Rumi tarihleri tespit et ve Miladi takvime dönüştür.*

                        ### 📖 5. ARŞİV AND TERİMLER SÖZLÜĞÜ
                        *Belgede geçen unvanlar, devlet görevleri veya ağır Arapça/Farsça tamlamalardan en az 5 tanesini seçerek anlamlarını açıkla.*
                        """
                        
                        yanit = model.generate_content([PALEOGRAFI_PROMPTU, gorsel_parca])
                        st.session_state.belge_arsivi[st.session_state.aktif_belge_adi]["analiz"] = yanit.text
                        st.rerun()
                    except Exception as e:
                        st.error(f"Gemini API hatası oluştu: {str(e)}")
            else:
                st.error("Lütfen önce Streamlit Secrets alanına geçerli bir GEMINI_API_KEY tanımlayın.")

    # 📊 SONUÇLARIN GÖSTERİLMESİ
    if "ocr_ham" in aktif_veri and aktif_veri["ocr_ham"]:
        with st.expander("📝 EasyOCR Tarafından Yakalanan Ham Metin Katmanı"):
            st.code(aktif_veri["ocr_ham"], language="text")

    if aktif_veri["analiz"]:
        st.markdown("### 📊 Yapay Zekâ Analiz ve Deşifre Raporu")
        
        rapor_metni = st.text_area("Düzenlenebilir Rapor Çıktısı", value=aktif_veri["analiz"], height=450)
        st.session_state.belge_arsivi[st.session_state.aktif_belge_adi]["analiz"] = rapor_metni
        
        st.write("---")
        st.markdown("### 📥 Raporu Dışa Aktar")
        col_down1, col_down2 = st.columns(2)
        
        with col_down1:
            docx_data = docx_uret(rapor_metni)
            st.download_button(
                label="📥 Word Dosyası (.docx) Olarak İndir",
                data=docx_data,
                file_name=f"PalaeoLab_{st.session_state.aktif_belge_adi}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
            
        with col_down2:
            pdf_data = pdf_uret(rapor_metni)
            st.download_button(
                label="📥 PDF Dosyası (.pdf) Olarak İndir",
                data=pdf_data,
                file_name=f"PalaeoLab_{st.session_state.aktif_belge_adi}.pdf",
                mime="application/pdf",
                use_container_width=True
            )

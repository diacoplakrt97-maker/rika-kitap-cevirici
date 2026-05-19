# ==============================================================================
# 📚 1. KÜTÜPHANELER VE GÜVENLİK AYARLARI
# ==============================================================================
import ssl
ssl._create_default_https_context = ssl._create_unverified_context

import os       
import base64   
import easyocr  
import json     
from PIL import Image, ImageEnhance 
import google.generativeai as genai 
import streamlit as st  
from docx import Document  
from fpdf import FPDF  
from io import BytesIO     

if "GEMINI_API_KEY" in st.secrets:
    genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
else:
    st.error("Lütfen Streamlit Secrets alanına GEMINI_API_KEY anahtarınızı ekleyin.")

# 🖥️ SAYFA AYARLARI (Kurumsal UI düzeni için Centered mod)
st.set_page_config(page_title="PalaeoLab AI - Evrensel Arşiv ve Analiz Sistemi", layout="centered")

# ==============================================================================
# 🎨 2. DÜNYA STANDARTLARINDA PREMIUM GÖRSEL TASARIM (GLOBAL UI/UX CSS)
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
    /* 🌐 Google Fonts Entegrasyonu (Modern ve Profesyonel Fontlar) */
    @import url('https://googleapis.com');
    
    .stApp {{ 
        font-family: 'Inter', sans-serif !important;
        color: #f1f5f9 !important; 
    }}
    
    h1, h2, h3, p, label {{ font-family: 'Inter', sans-serif !important; }}
    .main .block-container {{ padding-top: 160px !important; }}
    
    /* 🧪 Buzlu Cam (Glassmorphism) Efektli Kart Yapısı */
    .adim-karti {{ 
        background: rgba(15, 23, 42, 0.65) !important; 
        backdrop-filter: blur(12px) !important;
        -webkit-backdrop-filter: blur(12px) !important;
        padding: 24px !important; 
        border-radius: 16px !important; 
        border: 1px solid rgba(52, 211, 153, 0.2) !important; 
        margin-top: 25px !important; 
        margin-bottom: 25px !important;
        box-shadow: 0 8px 32px 0 rgba(0, 0, 0, 0.37) !important;
    }}
    
    /* 📂 Sürükle Bırak Kutusu (Premium Canva Tarzı) */
    div[data-testid="stFileUploader"] {{ 
        border: 2px dashed rgba(52, 211, 153, 0.4) !important; 
        border-radius: 14px !important; 
        background-color: rgba(15, 23, 42, 0.4) !important; 
        padding: 25px !important; 
        transition: all 0.3s ease-in-out !important;
    }}
    div[data-testid="stFileUploader"]:hover {{
        border-color: #34d399 !important;
        box-shadow: 0px 0px 20px rgba(52, 211, 153, 0.15) !important;
    }}
    
    /* 🔮 Premium İnteraktif Butonlar */
    .stButton>button[kind="primary"] {{ 
        background: linear-gradient(135deg, #10b981, #059669) !important; 
        color: #ffffff !important; 
        border-radius: 10px !important; 
        border: none !important;
        font-weight: 600 !important; 
        font-size: 16px !important; 
        padding: 14px 20px !important;
        letter-spacing: 0.5px !important;
        box-shadow: 0 4px 14px 0 rgba(16, 185, 129, 0.3) !important;
        transition: all 0.3s ease !important;
        width: 100% !important;
    }}
    .stButton>button[kind="primary"]:hover {{
        transform: translateY(-2px) !important;
        box-shadow: 0 6px 20px 0 rgba(16, 185, 129, 0.5) !important;
        background: linear-gradient(135deg, #34d399, #10b981) !important;
    }}
    
    /* 📝 Kod ve Metin Düzenleme Alanı */
    .stTextArea textarea {{
        font-family: 'JetBrains Mono', monospace !important;
        font-size: 14px !important;
        background-color: rgba(10, 15, 26, 0.8) !important;
        color: #a7f3d0 !important;
        border: 1px solid rgba(52, 211, 153, 0.3) !important;
        border-radius: 12px !important;
    }}
    
    /* 👑 Başlık Tasarımları */
    .ana-baslik {{ font-size: 36px !important; font-weight: 800 !important; color: #ffffff !important; text-align: center; margin-bottom: 5px; letter-spacing: -0.5px; }}
    .ana-baslik span {{ background: linear-gradient(to right, #34d399, #059669); -webkit-background-clip: text; -webkit-text-fill-color: transparent; }}
    .alt-baslik {{ text-align: center !important; color: #94a3b8 !important; font-size: 16px !important; margin-bottom: 30px; font-weight: 400; }}
    
    /* 🖼️ Resim Çerçeveleri */
    .stImage img {{
        border-radius: 12px !important;
        border: 1px solid rgba(255, 255, 255, 0.1) !important;
        box-shadow: 0 10px 25px rgba(0,0,0,0.5) !important;
    }}
    </style>
""", unsafe_allow_html=True)

st.markdown('<p class="ana-baslik">🔬 PALAEO<span>LAB</span> AI</p>', unsafe_allow_html=True)
st.markdown('<p class="alt-baslik">Osmanlıca Paleografi ve Tarihi Arşivler İçin Yeni Nesil Yapay Zekâ Platformu</p>', unsafe_allow_html=True)

# ==============================================================================
# 🗄️ 3. HAFIZA YÖNETİMİ (SESSION STATE)
# ==============================================================================
if "okunan_sonuc" not in st.session_state: st.session_state.okunan_sonuc = ""
if "analiz_sonuc" not in st.session_state: st.session_state.analiz_sonuc = None
if "islenmis_resim" not in st.session_state: st.session_state.islenmis_resim = None

# ==============================================================================
# 🎛️ 4. ADIM 1: GÖRSEL İYİLEŞTİRME VE DOSYA YÜKLEME PANELİ
# ==============================================================================
st.markdown("""
<div class="adim-karti">
    ⚡ <b>ADIM 1: Belge İyileştirme ve Yükleme Paneli</b><br>
    Yapay zekâ analizini başlatmadan önce belgenin netliğini ve kontrastını optimize edin.
</div>
""", unsafe_allow_html=True)

with st.container():
    st.write("⚙️ **Görsel Ön İşleme Katmanı**")
    filtre_modu = st.radio("Renk Modu", ["Orijinal Renk Spektrumu", "Siyah-Beyaz Stüdyo (Yapay Zeka İçin Önerilen)"], horizontal=True)
    kontrast_seviyesi = st.slider("Metin Belirginleştirme (Kontrast)", 1.0, 3.0, 1.6, step=0.1)
    
    st.write("---")
    yuklenen_dosya = st.file_uploader("Tarihi belgenizi buraya sürükleyip bırakın...", type=["jpg", "jpeg", "png", "jfif"])

# ==============================================================================
# 🔮 5. HİBRİT OKUMA VEYA ANALİZ MOTORU
# ==============================================================================
if yuklenen_dosya is not None:
    orijinal_resim = Image.open(yuklenen_dosya)
    resim_islem = orijinal_resim.convert('RGB') if orijinal_resim.mode in ('RGBA', 'LA') else orijinal_resim
    
    if filtre_modu == "Siyah-Beyaz Stüdyo (Yapay Zeka İçin Önerilen)":
        resim_islem = resim_islem.convert('L').convert('RGB')
        
    gelistirici = ImageEnhance.Contrast(resim_islem)
    resim_islem = gelistirici.enhance(kontrast_seviyesi)
    st.session_state.islenmis_resim = resim_islem

    st.write("")
    if st.button("🔮 Akıllı Çözümleme ve Analizi Başlat", type="primary"):
        with st.spinner("🧠 Belge katmanları inceleniyor ve paleografik yazılar çözülüyor..."):
            try:
                gecici_yol = "gecici_resim.jpg"
                st.session_state.islenmis_resim.save(gecici_yol, format="JPEG")
                
                okuyucu = easyocr.Reader(['ar', 'tr'], gpu=False)
                sonuc = okuyucu.readtext(gecici_yol)
                ocr_metni = " ".join([item for item in sonuc]) if sonuc else ""
                
                model = genai.GenerativeModel('gemini-1.5-flash')
                
                prompt = f"""Sen Osmanlı dönemi arşivleri ve paleografi alanında uzman kıdemli bir bilgi bilimcisin. 
                Sana sunulan bu tarihi el yazması veya matbu belgeyi analiz et ve şu protokolleri yerine getir:
                
                1. Metnin orijinal matbu/Arap harfli transkriptini eksiksiz çıkar.
                2. Metnin temiz, Latin harfli transkripsiyonlu (okunuş) halini yaz.
                3. Belge istihbaratını ve analiz verilerini tamamen Türkçe olarak ve KESİNLİKLE sadece şu JSON şablonu formatında döndür. Markdown etiketleri ekleme, doğrudan ham JSON olsun:
                   {{
                     "belge_turu": "Belge tipi (Örn: Ferman, Berat, Hüküm, Mektup, Arzuhal)",
                     "tarih_hicri": "Metinde geçen Hicri veya Rumi tarih",
                     "tarih_miladi": "Dönüştürülmüş Miladi takvim karşılığı",
                     "sahislar": ["Metinde adı veya unvanı geçen önemli tarihi kişiler"],
                     "yerler": ["Metinde adı geçen coğrafi konumlar veya bölgeler"],
                     "ozet": "Belgenin ana konusunu anlatan tek cümlelik Türkçe özet"
                   }}
                OCR motorundan gelen kaba kelime ipuçları: '{ocr_metni}'. Maksimum doğruluğa ulaşmak için görsel bağlam ile bu ipuçlarını harmanla."""
                
                response = model.generate_content([prompt, st.session_state.islenmis_resim])
                tam_yanit = response.text
                
                if "{" in tam_yanit and "}" in tam_yanit:
                    baslangic = tam_yanit.find("{")
                    bitis = tam_yanit.rfind("}") + 1
                    json_kismi = tam_yanit[baslangic:bitis]
                    metin_kismi = tam_yanit[:baslangic].strip()
                    
                    st.session_state.okunan_sonuc = metin_kismi
                    try:
                        st.session_state.analiz_sonuc = json.loads(json_kismi)
                    except:
                        st.session_state.analiz_sonuc = None
                else:
                    st.session_state.okunan_sonuc = tam_yanit
                    st.session_state.analiz_sonuc = None
                    
                st.success("🎉 Belge analizi ve veri tablosu başarıyla derlendi!")
                if os.path.exists(gecici_yol): os.remove(gecici_yol)
            except Exception as e:
                st.error(f"❌ Sistem Hatası: {e}")

    # ==============================================================================
    # 👁️ 6. ADIM 2: İNCELEME VE CANLI METİN DÜZENLEME EKRANI
    # ==============================================================================
    st.markdown('<div class="adim-karti">👁️ <b>ADIM 2: Çalışma Alanı ve İnteraktif Düzenleme Terminali</b></div>', unsafe_allow_html=True)
    
    if st.session_state.islenmis_resim:
        st.subheader("🔍 İyileştirilmiş Kaynak Belge")
        st.image(st.session_state.islenmis_resim, use_container_width=True)
            
    if st.session_state.okunan_sonuc:
        st.subheader("✍️ Yapay Zekâ Transkript Çıktısı")
        st.session_state.okunan_sonuc = st.text_area("Canlı Düzenleme Modu", value=st.session_state.okunan_sonuc, height=300)

    # ==============================================================================
    # 📊 7. ADIM 3: TARİHSEL AKILLI ANALİZ PANELİ (METADATA)
    # ==============================================================================
    if st.session_state.analiz_sonuc:
        st.markdown('<div class="adim-karti">📊 <b>ADIM 3: Tarihsel Veri ve Arşiv Kataloğu</b></div>', unsafe_allow_html=True)
        
        veri = st.session_state.analiz_sonuc
        
        c1, c2, c3 = st.columns(3)
        c1.metric("📜 Belge Sınıflandırması", veri.get("belge_turu", "Bilinmiyor"))
        c2.metric("📅 Hicri Takvim", veri.get("tarih_hicri", "Belirtilmemiş"))
        c3.metric("🌍 Miladi Takvim Karşılığı", veri.get("tarih_miladi", "Hesaplanamadı"))
        
        st.write("---")
        st.markdown(f"💡 **Yönetici Özeti:** *{veri.get('ozet', 'Özet mevcut değil.')}*")
        st.write("")
        
        col_sahis, col_yer = st.columns(2)
        with col_sahis:
            st.write("👥 **Belirlenen Tarihi Kişiler / Unvanlar:**")
            st.write(", ".join(veri.get("sahislar", [])) if veri.get("sahislar") else "Kişi adı ayıklanamadı.")
        with col_yer:
            st.write("📍 **Ayıklanan Coğrafi Konumlar:**")
            st.write(", ".join(veri.get("yerler", [])) if veri.get("yerler") else "Konum bilgisi ayıklanamadı.")

        # ==============================================================================
        # 💾 8. ADIM 4: GELİŞMİŞ RAPORLAMA VE DOSYA İNDİRME MODÜLÜ
        # ==============================================================================
        st.markdown('<div class="adim-karti">💾 <b>ADIM 4: Dışa Aktarım ve Sertifikalı Çıktı Motoru</b><br>Resmi dijital arşiv analiz raporlarını bilgisayarınıza indirin.</div>', unsafe_allow_html=True)
        
        col_word, col_pdf = st.columns(2)
        
        # Word Raporu Oluşturma
        doc = Document()
        doc.add_heading("PalaeoLab AI - Arşiv Analiz Raporu", 0)
        doc.add_heading("1. Çözümlenen Metin Çıktısı", level=1)
        doc.add_paragraph(st.session_state.okunan_sonuc)
        doc.add_heading("2. Katalog ve Metadata Bilgileri", level=1)
        doc.add_paragraph(f"Belge Türü: {veri.get('belge_turu')}")
        doc.add_paragraph(f"Hicri Tarih: {veri.get('tarih_hicri')} | Miladi Karşılığı: {veri.get('tarih_miladi')}")
        doc.add_paragraph(f"Ana Özet: {veri.get('ozet')}")
        
        word_akisi = BytesIO()
        doc.save(word_akisi)
        word_akisi.seek(0)
        
        with col_word:
            st.download_button(
                label="📥 Raporu Word Olarak İndir (.docx)",
                data=word_akisi,
                file_name="palaeolab_arsiv_raporu.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            
        # PDF Raporu Oluşturma
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=12)
        pdf.cell(200, 10, txt="PalaeoLab AI - Onayli Veri Raporu", ln=1, align="C")
        pdf.ln(10)
        
        temiz_metin = st.session_state.okunan_sonuc.encode('latin-1', 'ignore').decode('latin-1')
        pdf.multi_cell(0, 10, txt=temiz_metin)
        pdf_akisi = pdf.output(dest='S').encode('latin-1')
        
        with col_pdf:
            st.download_button(
                label="📥 Raporu PDF Olarak İndir (.pdf)",
                data=pdf_akisi,
                file_name="palaeolab_arsiv_raporu.pdf",
                mime="application/pdf"
            )

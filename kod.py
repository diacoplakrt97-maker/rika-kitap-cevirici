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

if "GEMINI_API_KEY" in st.secrets:
    genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
else:
    st.error("Lütfen Streamlit Secrets alanına GEMINI_API_KEY anahtarınızı ekleyin.")

# 🖥️ SAYFA AYARLARI (Biçimsiz uzamayı engellemek için tekrar CENTERED moduna alındı)
st.set_page_config(page_title="Evrensel Yapay Zeka Arşiv ve Analiz Sistemi", layout="centered")

# ==============================================================================
# 🎨 2. GELİŞMİŞ VE OKUNABİLİR GÖRSEL TASARIM (YENİ CSS)
# ==============================================================================
banner_adi = "banner.png"
bg_image_html = ""

if os.path.exists(banner_adi):
    with open(banner_adi, "rb") as image_file:
        encoded_string = base64.b64encode(image_file.read()).decode()
    # Banner sadece en üstte şık bir şerit olarak kalacak şekilde sınırlandı
    bg_image_html = f"""
    <style>
    .stApp {{
        background-image: url("data:image/png;base64,{encoded_string}") !important;
        background-size: 100% 120px !important; /* Yüksekliği 120px ile sınırladık */
        background-position: top center !important;
        background-repeat: no-repeat !important;
        background-color: #0b0f14 !important; /* Arka planı derin koyu gri yaptık */
    }}
    </style>
    """

# Tüm arayüzü okunaklı kılan özel panel tasarımları
st.markdown(f"""
    {bg_image_html}
    <style>
    /* Genel Yazı Renkleri */
    .stApp {{ color: #e2e8f0 !important; }}
    h1, h2, h3, p, label {{ color: #ffffff !important; }}
    
    /* Üst Boşluk Ayarı (Banner altında kalmaması için) */
    .main .block-container {{ padding-top: 140px !important; }}
    
    /* 📦 Modern Okunaklı Kart Yapısı */
    .adim-karti {{ 
        background: rgba(20, 27, 38, 0.95) !important; 
        padding: 20px !important; 
        border-radius: 12px !important; 
        border: 1px solid #10b981 !important; 
        margin-top: 20px !important; 
        margin-bottom: 20px !important;
        box-shadow: 0px 4px 20px rgba(0, 0, 0, 0.5) !important;
    }}
    
    /* 📂 Dosya Yükleme Kutusu Düzenlemesi */
    div[data-testid="stFileUploader"] {{ 
        border: 2px dashed #34d399 !important; 
        border-radius: 10px !important; 
        background-color: rgba(13, 18, 26, 0.8) !important; 
        padding: 15px !important; 
    }}
    
    /* 🔮 Ana Çalıştırma Butonu */
    .stButton>button[kind="primary"] {{ 
        background: linear-gradient(135deg, #059669, #047857) !important; 
        color: white !important; 
        border-radius: 8px !important; 
        border: none !important;
        font-weight: bold !important; 
        font-size: 16px !important; 
        padding: 12px !important;
        transition: 0.3s !important;
    }}
    .stButton>button[kind="primary"]:hover {{
        background: linear-gradient(135deg, #10b981, #059669) !important;
        box-shadow: 0px 0px 15px rgba(52, 211, 153, 0.4) !important;
    }}
    
    /* Ana Başlık Tasarımı */
    .ana-baslik {{ font-size: 34px !important; font-weight: 800 !important; color: #34d399 !important; text-shadow: 0px 2px 10px rgba(0,0,0,0.8) !important; text-align: center; margin-bottom: 5px; }}
    .alt-baslik {{ text-align: center !important; color: #9ca3af !important; font-size: 15px !important; margin-bottom: 20px; }}
    </style>
""", unsafe_allow_html=True)

st.markdown('<p class="ana-baslik">🔬 DİJİTAL ARŞİV LABORATUVARI</p>', unsafe_allow_html=True)
st.markdown('<p class="alt-baslik">✨ Gelişmiş Filtreleme, Rika Paleografi Analizi ve Otomatik Takvim Dönüştürücü</p>', unsafe_allow_html=True)

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
    📂 <b>ADIM 1: Laboratuvar İyileştirme ve Belge Yükleme</b><br>
    Eski arşiv belgesindeki lekeleri temizlemek ve yazıyı parlatmak için filtreleri kullanın.
</div>
""", unsafe_allow_html=True)

# Ayarlar ve yükleme alanını tek bir şık panel içinde alt alta topladık
with st.container():
    st.write("🔧 **Görsel Ön İşleme Filtreleri**")
    filtre_modu = st.radio("Renk İyileştirme", ["Orijinal Tonlar", "Siyah-Beyaz Modu (Yapay Zeka İçin En İyisi)"], horizontal=True)
    kontrast_seviyesi = st.slider("Harf Keskinliği (Kontrast)", 1.0, 3.0, 1.6, step=0.1)
    
    st.write("---")
    yuklenen_dosya = st.file_uploader("Arşiv Belgesini Seçin", type=["jpg", "jpeg", "png", "jfif"])

# ==============================================================================
# 🔮 5. HİBRİT OKUMA VE AKILLI TARAMA ALGORİTMASI
# ==============================================================================
if yuklenen_dosya is not None:
    orijinal_resim = Image.open(yuklenen_dosya)
    resim_islem = orijinal_resim.convert('RGB') if orijinal_resim.mode in ('RGBA', 'LA') else orijinal_resim
    
    if filtre_modu == "Siyah-Beyaz Modu (Yapay Zeka İçin En İyisi)":
        resim_islem = resim_islem.convert('L').convert('RGB')
        
    gelistirici = ImageEnhance.Contrast(resim_islem)
    resim_islem = gelistirici.enhance(kontrast_seviyesi)
    st.session_state.islenmis_resim = resim_islem

    st.write("")
    if st.button("🔮 Belgeyi Laboratuvarda Çözümle ve Analiz Et", type="primary"):
        with st.spinner("🤖 Yapay zeka ve OCR el yazısını inceliyor, lütfen bekleyin..."):
            try:
                gecici_yol = "gecici_resim.jpg"
                st.session_state.islenmis_resim.save(gecici_yol, format="JPEG")
                
                okuyucu = easyocr.Reader(['ar', 'tr'], gpu=False)
                sonuc = okuyucu.readtext(gecici_yol)
                ocr_metni = " ".join([item for item in sonuc]) if sonuc else ""
                
                model = genai.GenerativeModel('gemini-1.5-flash')
                
                prompt = f"""Sen Osmanlı dönemine ait belgeleri inceleyen kıdemli bir arşiv uzmanı ve paleografsın. 
                Sana sunulan bu tarihi belgedeki metni çöz ve şu kurallara göre analiz et:
                
                1. Metnin orijinal matbu/Arap harfli transkriptini eksiksiz çıkar.
                2. Metnin günümüz Latin harfleriyle transkripsiyonlu halini yaz.
                3. Belge hakkındaki tarihsel çıkarımlarını tam olarak şu JSON formatında ek bir bilgi olarak ver. Başka hiçbir açıklama yazma, sadece JSON formatı olsun:
                   {{
                     "belge_turu": "Ferman mı, berat mı, mektup mu, nüfus defteri mi?",
                     "tarih_hicri": "Belgede geçen Hicri/Rumi tarih",
                     "tarih_miladi": "Hicri tarihin günümüz Miladi takvimindeki karşılığı",
                     "sahislar": ["Metinde adı geçen padişah, devlet adamı veya kişiler"],
                     "yerler": ["Metinde adı geçen şehir, kasaba veya bölgeler"],
                     "ozet": "Belgenin kısaca ne anlattığına dair tek cümlelik özet"
                   }}
                OCR motorumuz resimden şu kaba kelimeleri yakaladı: '{ocr_metni}'. Bu veriyi ve görseli harmanlayarak en doğru transkripti oluştur."""
                
                response = model.generate_content([prompt, st.session_state.islenmis_resim])
                tam_yanit = response.text
                
                if "{" in tam_yanit and "}" in tam_yanit:
                    metin_kismi = tam_yanit.split("{")[0].strip()
                    json_kismi = "{" + tam_yanit.split("{")[1].split("}")[0] + "}"
                    
                    st.session_state.okunan_sonuc = metin_kismi
                    try:
                        st.session_state.analiz_sonuc = json.loads(json_kismi)
                    except:
                        st.session_state.analiz_sonuc = None
                else:
                    st.session_state.okunan_sonuc = tam_yanit
                    st.session_state.analiz_sonuc = None
                    
                st.success("🎉 Çözümleme ve Tarihsel Analiz Başarıyla Tamamlandı!")
                if os.path.exists(gecici_yol): os.remove(gecici_yol)
            except Exception as e:
                st.error(f"❌ Laboratuvar Hatası: {e}")

    # ==============================================================================
    # 👁️ 6. ADIM 2: YAN YANA İNCELEME VEYA TEKLİ GÖRÜNÜM
    # ==============================================================================
    st.markdown('<div class="adim-karti">👁️ <b>ADIM 2: Laboratuvar İncelemesi ve Canlı Metin Düzenleme</b></div>', unsafe_allow_html=True)
    
    if st.session_state.islenmis_resim:
        st.subheader("🔍 Filtrelenmiş Arşiv Belgesi")
        st.image(st.session_state.islenmis_resim, use_container_width=True)
            
    if st.session_state.okunan_sonuc:
        st.subheader("✍️ Yapay Zekâ Transkript Çıktısı")
        st.session_state.okunan_sonuc = st.text_area("Metin Düzenleme Alanı (Eksikleri Düzeltebilirsiniz)", value=st.session_state.okunan_sonuc, height=300)

    # ==============================================================================
    # 📊 7. ADIM 3: TARİHSEL AKILLI ANALİZ PANELİ (METADATA)
    # ==============================================================================
    if st.session_state.analiz_sonuc:
        st.markdown('<div class="adim-karti">📊 <b>ADIM 3: Yapay Zekâ Tarihsel Analiz ve Katalog Paneli</b></div>', unsafe_allow_html=True)
        
        veri = st.session_state.analiz_sonuc
        
        c1, c2, c3 = st.columns(3)
        c1.metric("📜 Belge Türü", veri.get("belge_turu", "Tespit Edilemedi"))
        c2.metric("📅 Hicri / Rumi Tarih", veri.get("tarih_hicri", "Bilinmiyor"))
        c3.metric("🌍 Miladi Takvim Karşılığı", veri.get("tarih_miladi", "Dönüştürülemedi"))
        
        st.write("---")
        st.write(f"ℹ️ **Belge Özeti:** {veri.get('ozet', 'Özet çıkarılamadı.')}")
        
        col_sahis, col_yer = st.columns(2)
        with col_sahis:
            st.write("👥 **Belgede Geçen Kişiler / Unvanlar:**")
            st.write(", ".join(veri.get("sahislar", [])) if veri.get("sahislar") else "Kişi adı bulunamadı.")
        with col_yer:
            st.write("📍 **Belgede Geçen Yer İsimleri:**")
            st.write(", ".join(veri.get("yerler", [])) if veri.get("yerler") else "Yer adı bulunamadı.")

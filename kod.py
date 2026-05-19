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

# 🖥️ SAYFA AYARLARI (Sol menü kullanımı için Centered yerine geniş yerleşim daha uygundur)
st.set_page_config(page_title="PalaeoLab AI - Evrensel Arşiv ve Analiz Sistemi", layout="wide")

# ==============================================================================
# 🎨 2. PREMIUM GÖRSEL TASARIM VE CAM EFEKTLERİ (CSS KATMANI)
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
    
    h1, h2, h3, p, label {{ font-family: 'Inter', sans-serif !important; }}
    .main .block-container {{ padding-top: 160px !important; }}
    
    /* 🧪 Buzlu Cam Efektli Kart Yapısı */
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
    
    /* 📂 Sürükle Bırak Kutusu */
    div[data-testid="stFileUploader"] {{ 
        border: 2px dashed rgba(52, 211, 153, 0.4) !important; 
        border-radius: 14px !important; 
        background-color: rgba(15, 23, 42, 0.4) !important; 
        padding: 25px !important; 
    }}
    
    /* 🔮 Premium Butonlar */
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
    
    /* 📝 Terminal Düzenleme Alanı */
    .stTextArea textarea {{
        font-family: 'JetBrains Mono', monospace !important;
        font-size: 14px !important;
        background-color: rgba(10, 15, 26, 0.8) !important;
        color: #a7f3d0 !important;
        border: 1px solid rgba(52, 211, 153, 0.3) !important;
        border-radius: 12px !important;
    }}
    
    .ana-baslik {{ font-size: 36px !important; font-weight: 800 !important; color: #ffffff !important; text-align: center; margin-bottom: 5px; }}
    .ana-baslik span {{ background: linear-gradient(to right, #34d399, #059669); -webkit-background-clip: text; -webkit-text-fill-color: transparent; }}
    .alt-baslik {{ text-align: center !important; color: #94a3b8 !important; font-size: 16px !important; margin-bottom: 30px; }}
    
    .stImage img {{ border-radius: 12px !important; border: 1px solid rgba(255, 255, 255, 0.1) !important; }}
    
    /* 🗄️ Sol Menü (Sidebar) Özelleştirmeleri */
    section[data-testid="stSidebar"] {{
        background-color: #0b0f17 !important;
        border-right: 1px solid rgba(52, 211, 153, 0.2) !important;
    }}
    </style>
""", unsafe_allow_html=True)

st.markdown('<p class="ana-baslik">🔬 PALAEO<span>LAB</span> AI</p>', unsafe_allow_html=True)
st.markdown('<p class="alt-baslik">Osmanlıca Paleografi ve Tarihi Arşivler İçin Yeni Nesil Yapay Zekâ Platformu</p>', unsafe_allow_html=True)

# ==============================================================================
# 🗄️ 3. GELİŞMİŞ OTURUM HAFIZASI (SESSION STATE)
# ==============================================================================
# Tüm yüklenen belgeleri saklayacağımız ana sözlük yapısı
if "belge_arsivi" not in st.session_state: 
    st.session_state.belge_arsivi = {}

# Şu an ekranda aktif olarak seçili olan belgenin adı
if "aktif_belge_adi" not in st.session_state: 
    st.session_state.aktif_belge_adi = None

# ==============================================================================
# 📁 4. SOL MENÜ (SIDEBAR) GEÇMİŞ PANELİ YÖNETİMİ
# ==============================================================================
with st.sidebar:
    st.markdown("### 🗄️ Laboratuvar Arşivi")
    st.write("Oturum boyunca analiz ettiğiniz belgeler burada listelenir:")
    
    if st.session_state.belge_arsivi:
        # Arşivdeki tüm belgeleri bir liste halinde seçime sunuyoruz
        arsiv_listesi = list(st.session_state.belge_arsivi.keys())
        
        # Kullanıcı sol menüden bir belge seçtiğinde tetiklenir
        secilen_belge = st.radio("Geçmiş Belgeler", arsiv_listesi, label_visibility="collapsed")
        st.session_state.aktif_belge_adi = secilen_belge
    else:
        st.info("Henüz taranmış bir belge bulunmuyor.")
        
    st.write("---")
    if st.button("🗑️ Arşivi Temizle"):
        st.session_state.belge_arsivi = {}
        st.session_state.aktif_belge_adi = None
        st.rerun()

# ==============================================================================
# 🎛️ 5. ADIM 1: GÖRSEL İYİLEŞTİRME VEYA DOSYA YÜKLEME PANELİ
# ==============================================================================
st.markdown("""
<div class="adim-karti">
    ⚡ <b>ADIM 1: Belge İyileştirme ve Yükleme Paneli</b><br>
    Yapay zekâ analizini başlatmadan önce belgenin netliğini ve kontrastını optimize edin.
</div>
""", unsafe_allow_html=True)

# Geniş ekranda yan yana daha şık durması için kolon yapısı
col_ayar, col_yukle = st.columns([1, 2])

with col_ayar:
    st.write("⚙️ **Görsel Ön İşleme Katmanı**")
    filtre_modu = st.radio("Renk Modu", ["Orijinal Renk Spektrumu", "Siyah-Beyaz Stüdyo (Önerilen)"])
    kontrast_seviyesi = st.slider("Metin Belirginleştirme (Kontrast)", 1.0, 3.0, 1.6, step=0.1)

with col_yukle:
    st.write("📂 **Yeni Döküman Girişi**")
    yuklenen_dosya = st.file_uploader("Tarihi belgenizi buraya sürükleyip bırakın...", type=["jpg", "jpeg", "png", "jfif"])

# ==============================================================================
# 🔮 6. HİBRİT OKUMA VE ANALİZ MOTORU
# ==============================================================================
if yuklenen_dosya is not None:
    orijinal_resim = Image.open(yuklenen_dosya)
    resim_islem = orijinal_resim.convert('RGB') if orijinal_resim.mode in ('RGBA', 'LA') else orijinal_resim
    
    if filtre_modu == "Siyah-Beyaz Stüdyo (Önerilen)":
        resim_islem = resim_islem.convert('L').convert('RGB')
        
    gelistirici = ImageEnhance.Contrast(resim_islem)
    resim_islem = gelistirici.enhance(kontrast_seviyesi)

    st.write("")
    if st.button("🔮 Akıllı Çözümleme ve Analizi Başlat", type="primary"):
        with st.spinner("🧠 Belge katmanları inceleniyor ve paleografik yazılar çözülüyor..."):
            try:
                gecici_yol = "gecici_resim.jpg"
                resim_islem.save(gecici_yol, format="JPEG")
                
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
                
                response = model.generate_content([prompt, resim_islem])
                tam_yanit = response.text
                
                # JSON Ayıklama Mantığı
                if "{" in tam_yanit and "}" in tam_yanit:
                    baslangic = tam_yanit.find("{")
                    bitis = tam_yanit.rfind("}") + 1
                    json_kismi = tam_yanit[baslangic:bitis]
                    metin_kismi = tam_yanit[:baslangic].strip()
                    
                    try:
                        analiz_verisi = json.loads(json_kismi)
                    except:
                        analiz_verisi = None
                else:
                    metin_kismi = tam_yanit
                    analiz_verisi = None
                
                # Yeni yüklenen belgeyi dinamik bir isimle hafıza sözlüğüne kaydediyoruz
                sayac = len(st.session_state.belge_arsivi) + 1
                tip = analiz_verisi.get("belge_turu", "Belge") if analiz_verisi else "Belge"
                yeni_isim = f"#{sayac} - {tip}"
                
                st.session_state.belge_arsivi[yeni_isim] = {
                    "resim": resim_islem,
                    "metin": metin_kismi,
                    "analiz": analiz_verisi
                }
                st.session_state.aktif_belge_adi = yeni_isim
                
                st.success("🎉 Belge analizi başarıyla tamamlandı ve sol menü arşivine eklendi!")
                if os.path.exists(gecici_yol): os.remove(gecici_yol)
                st.rerun() # Sol menünün hemen güncellenmesi için sayfayı tetikler
            except Exception as e:
                st.error(f"❌ Sistem Hatası: {e}")

# ==============================================================================
# 👁️ 7. AKTİF SEÇİLİ BELGENİN EKRANA YANSITILMASI
# ==============================================================================
if st.session_state.aktif_belge_adi and st.session_state.aktif_belge_adi in st.session_state.belge_arsivi:
    
    # Hafızadaki verileri çekiyoruz
    gecerli_belge = st.session_state.belge_arsivi[st.session_state.aktif_belge_adi]
    aktif_resim = gecerli_belge["resim"]
    aktif_metin = gecerli_belge["metin"]
    aktif_analiz = gecerli_belge["analiz"]

    # ADIM 2: ÇALIŞMA ALANI
    st.markdown(f'<div class="adim-karti">👁️ <b>ADIM 2: Çalışma Alanı — {st.session_state.aktif_belge_adi} İncelemesi</b></div>', unsafe_allow_html=True)
    
    col_img, col_txt = st.columns(2)
    
    with col_img:
        st.subheader("🔍 İyileştirilmiş Kaynak Belge")
        st.image(aktif_resim, use_container_width=True)
            
    with col_txt:
        st.subheader("✍️ Yapay Zekâ Transkript Çıktısı")
        # Kullanıcı düzenleme yaparsa anlık olarak hafızadaki veriyi günceller
        yeni_duzenleme = st.text_area("Canlı Düzenleme Modu", value=aktif_metin, height=380, key=f"txt_{st.session_state.aktif_belge_adi}")
        st.session_state.belge_arsivi[st.session_state.aktif_belge_adi]["metin"] = yeni_duzenleme

    # ADIM 3: METADATA PANELİ
    if aktif_analiz:
        st.markdown('<div class="adim-karti">📊 <b>ADIM 3: Tarihsel Veri ve Arşiv Kataloğu</b></div>', unsafe_allow_html=True)
        
        c1, c2, c3 = st.columns(3)
        c1.metric("📜 Belge Sınıflandırması", aktif_analiz.get("belge_turu", "Bilinmiyor"))
        c2.metric("📅 Hicri Takvim", aktif_analiz.get("tarih_hicri", "Belirtilmemiş"))
        c3.metric("🌍 Miladi Takvim Karşılığı", aktif_analiz.get("tarih_miladi", "Hesaplanamadı"))
        
        st.write("---")
        st.markdown(f"💡 **Yönetici Özeti:** *{aktif_analiz.get('ozet', 'Özet mevcut değil.')}*")
        st.write("")
        
        col_sahis, col_yer = st.columns(2)
        with col_sahis:
            st.write("👥 **Belirlenen Tarihi Kişiler / Unvanlar:**")
            st.write(", ".join(aktif_analiz.get("sahislar", [])) if aktif_analiz.get("sahislar") else "Kişi adı ayıklanamadı.")
        with col_yer:
            st.write("📍 **Ayıklanan Coğrafi Konumlar:**")
            st.write(", ".join(aktif_analiz.get("yerler", [])) if aktif_analiz.get("yerler") else "Konum bilgisi ayıklanamadı.")

        # ADIM 4: RAPORLAMA VE İNDİRME
        st.markdown('<div class="adim-karti">💾 <b>ADIM 4: Dışa Aktarım ve Sertifikalı Çıktı Motoru</b><br>Seçili belgeye ait resmi raporu bilgisayarınıza indirin.</div>', unsafe_allow_html=True)
        
        col_word, col_pdf = st.columns(2)
        
        # Word Raporu Motoru
        doc = Document()
        doc.add_heading("PalaeoLab AI - Arşiv Analiz Raporu", 0)
        doc.add_heading("1. Çözümlenen Metin Çıktısı", level=1)
        doc.add_paragraph(yeni_duzenleme)
        doc.add_heading("2. Katalog ve Metadata Bilgileri", level=1)
        doc.add_paragraph(f"Belge Türü: {aktif_analiz.get('belge_turu')}")
        doc.add_paragraph(f"Hicri Tarih: {aktif_analiz.get('tarih_hicri')} | Miladi Karşılığı: {aktif_analiz.get('tarih_miladi')}")
        doc.add_paragraph(f"Ana Özet: {aktif_analiz.get('ozet')}")
        
        word_akisi = BytesIO()
        doc.save(word_akisi)
        word_akisi.seek(0)
        
        with col_word:
            st.download_button(
                label=f"📥 {st.session_state.aktif_belge_adi} Raporunu Word Yap (.docx)",
                data=word_akisi,
                file_name=f"palaeolab_{st.session_state.aktif_belge_adi}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key=f"w_{st.session_state.aktif_belge_adi}"
            )
            
        # PDF Raporu Motoru
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=12)
        pdf.cell(200, 10, txt="PalaeoLab AI - Onayli Veri Raporu", ln=1, align="C")
        pdf.ln(10)
        
        temiz_metin = yeni_duzenleme.encode('latin-1', 'ignore').decode('latin-1')
        pdf.multi_cell(0, 10, txt=temiz_metin)
        pdf_akisi = pdf.output(dest='S').encode('latin-1')
        
        with col_pdf:
            st.download_button(
                label=f"📥 {st.session_state.aktif_belge_adi} Raporunu PDF Yap (.pdf)",
                data=pdf_akisi,
                file_name=f"palaeolab_{st.session_state.aktif_belge_adi}.pdf",
                mime="application/pdf",
                key=f"p_{st.session_state.aktif_belge_adi}"
            )

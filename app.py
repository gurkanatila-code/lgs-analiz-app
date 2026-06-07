import streamlit as st
from pdf2image import convert_from_bytes
from PIL import Image
import io
import json
import pandas as pd
import google.generativeai as genai
import docx

# Sayfa Ayarları
st.set_page_config(page_title="LGS Kazanım Analiz", page_icon="📚", layout="wide")

st.title("📚 LGS Sınav Kitapçığı Analiz Aracı")
st.markdown("A ve B kitapçıklarını yükleyin, yapay zeka sizin için cevap anahtarını ve kazanım analizini çıkarsın.")

# Yardımcı Fonksiyonlar
def extract_images_from_pdf(uploaded_file):
    pdf_bytes = uploaded_file.read()
    # pdf2image kullanarak görsele çevir
    images = convert_from_bytes(pdf_bytes, dpi=150)
    return images
    
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        img_data = pix.tobytes("jpeg")
        img = Image.open(io.BytesIO(img_data))
        images.append(img)
    doc.close()
    return images

def read_docx(uploaded_file):
    """Yüklenen Word dosyasındaki kazanım metinlerini okur."""
    doc = docx.Document(uploaded_file)
    fullText = []
    for para in doc.paragraphs:
        if para.text.strip():
            fullText.append(para.text)
    return '\n'.join(fullText)

def analyze_with_gemini(api_key, images_a, images_b, kazanim_text):
    """Gemini 1.5 Pro API'sine istek atar."""
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel('gemini-1.5-pro-latest')
    
    system_prompt = f"""
    Sen uzman bir öğretmensin. Görevin LGS deneme sınavlarını analiz etmektir.
    Sana A ve B kitapçıklarının sayfaları görsel olarak verilecek.
    Kullanacağın kazanım listesi:
    {kazanim_text}
    
    A kitapçığındaki her bir soruyu incele:
    1. Soruyu çöz ve doğru şıkkı (A, B, C, D) bul.
    2. B kitapçığı görsellerine bakarak, bu sorunun B kitapçığında kaçıncı soru olduğunu tespit et.
    3. Sorunun içeriğine bakarak, yukarıdaki listeden uygun olan 6 haneli 'Kazanım Kodunu' belirle.
    
    SADECE aşağıdaki formatta geçerli bir JSON döndür. Başka metin yazma:
    [
      {{"SORU": 1, "A CEVAP": "A", "DÖNÜŞÜM (B)": 4, "KAZANIM KODU": "961281"}},
      {{"SORU": 2, "A CEVAP": "C", "DÖNÜŞÜM (B)": 1, "KAZANIM KODU": "961282"}}
    ]
    """
    
    contents = [system_prompt]
    contents.append("A KİTAPÇIĞI SAYFALARI:")
    contents.extend(images_a)
    contents.append("B KİTAPÇIĞI SAYFALARI:")
    contents.extend(images_b)
    
    response = model.generate_content(
        contents,
        generation_config=genai.types.GenerationConfig(
            response_mime_type="application/json",
            temperature=0.1
        )
    )
    return json.loads(response.text)

# --- YAN MENÜ (SIDEBAR) AYARLARI ---
with st.sidebar:
    st.header("⚙️ Ayarlar")
    api_key = st.text_input("Gemini API Anahtarı", type="password", help="Google AI Studio'dan aldığınız anahtarı buraya girin.")
    st.markdown("---")
    st.info("Bu araç Google Gemini 1.5 Pro Vision modelini kullanmaktadır. Lütfen analiz edilecek PDF boyutlarının çok yüksek olmamasına dikkat edin.")

# --- ANA EKRAN YÜKLEME ALANLARI ---
col1, col2 = st.columns(2)

with col1:
    st.subheader("📄 Sınav Kitapçıkları")
    pdf_a = st.file_uploader("A Kitapçığı Yükle (PDF)", type=["pdf"])
    pdf_b = st.file_uploader("B Kitapçığı Yükle (PDF)", type=["pdf"])

with col2:
    st.subheader("🎯 Kazanım Listesi")
    kazanim_doc = st.file_uploader("İlgili Dersin Kazanım Listesi (Word)", type=["docx"])

# --- ANALİZİ BAŞLATMA VE SONUÇ ---
if st.button("🚀 Analizi Başlat", use_container_width=True):
    if not api_key:
        st.error("Lütfen sol menüden API Anahtarınızı girin.")
    elif not (pdf_a and pdf_b and kazanim_doc):
        st.error("Lütfen A Kitapçığı, B Kitapçığı ve Kazanım Word dosyasını yüklediğinizden emin olun.")
    else:
        try:
            with st.status("Analiz yürütülüyor, lütfen bekleyin... (Bu işlem 1-2 dakika sürebilir)", expanded=True) as status:
                st.write("1. Belgeler okunuyor...")
                kazanim_metni = read_docx(kazanim_doc)
                
                st.write("2. A Kitapçığı görsellere çevriliyor...")
                images_a = extract_images_from_pdf(pdf_a)
                
                st.write("3. B Kitapçığı görsellere çevriliyor...")
                images_b = extract_images_from_pdf(pdf_b)
                
                st.write("4. Yapay Zeka analizi yapılıyor (Gemini 1.5 Pro devrede)...")
                json_result = analyze_with_gemini(api_key, images_a, images_b, kazanim_metni)
                
                status.update(label="Analiz Başarıyla Tamamlandı!", state="complete", expanded=False)
            
            st.success("İşlem Başarılı! Aşağıdan sonuçları inceleyebilir ve Excel olarak indirebilirsiniz.")
            
            # Veriyi Tabloya Çevirme ve Gösterme
            df = pd.DataFrame(json_result)
            
            # Excel Çıktısı Hazırlama
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Analiz')
            processed_data = output.getvalue()
            
            # Ekranda Göster
            st.dataframe(df, use_container_width=True)
            
            # İndirme Butonu
            st.download_button(
                label="📥 Excel Olarak İndir",
                data=processed_data,
                file_name="LGS_Analiz_Sonucu.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            st.error(f"Bir hata oluştu: {str(e)}")

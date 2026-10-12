import os
import shutil
import streamlit as st
from docx import Document
import google.generativeai as genai

st.set_page_config(page_title="Oto Form Doldurucu", page_icon="📄", layout="centered")

st.title("Yapay Zeka Destekli Otonom Form Doldurucu")
st.write("Word belgesini yükleyin; sistem boş alanları tespit edip kişisel bilgilerinizle otomatik doldursun.")

if "GEMINI_API_KEY" in st.secrets:
    genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
else:
    st.error("Lütfen Streamlit ayarlarından GEMINI_API_KEY tanımlayın.")

yuklenen_dosya = st.file_uploader("Doldurulacak Word Belgesini Seçin (.docx)", type=["docx"])

if yuklenen_dosya is not None:
    input_path = "temp_input.docx"
    output_path = "temp_output.docx"

    with open(input_path, "wb") as f:
        f.write(yuklenen_dosya.getbuffer())

    doc = Document(input_path)

    bos_alanlar = []
    for tablo_idx, table in enumerate(doc.tables):
        for row_idx, row in enumerate(table.rows):
            for col_idx, cell in enumerate(row.cells):
                metin = cell.text.strip()
                if metin and len(metin) < 60:
                    if col_idx + 1 < len(row.cells):
                        sag_hucre = row.cells[col_idx + 1]
                        if not sag_hucre.text.strip():
                            bos_alanlar.append({
                                "tablo": tablo_idx,
                                "satir": row_idx,
                                "sutun": col_idx + 1,
                                "etiket": metin
                            })

    if bos_alanlar:
        st.success(f"Belgede {len(bos_alanlar)} adet doldurulacak alan tespit edildi.")
        
        if st.button("Formu Bilgilerinizle Otomatik Doldur"):
            with st.spinner("Yapay zeka kişisel verilerinizi kullanarak formu işliyor..."):
                try:
                    model = genai.GenerativeModel("gemini-3.8-flash")
                    
                    # Kullanıcı bağlamı
                    kullanici_baglami = (
                        "Kişi Bilgileri: Prof. Ufuk Tolga Savaş, "
                        "Kurum: Hacettepe Üniversitesi Güzel Sanatlar Fakültesi Seramik Bölümü, "
                        "Unvan: Profesör / Öğretim Üyesi."
                    )
                    
                    for alan in bos_alanlar:
                        etiket = alan["etiket"]
                        prompt = f"{kullanici_baglami}\n\nBu bağlamı kullanarak şu form alanı için uygun bilgiyi kısa ve net bir şekilde üret: {etiket}"
                        cevap = model.generate_content(prompt)
                        deger = cevap.text.strip()
                        
                        hedef_hucre = doc.tables[alan["tablo"]].cell(alan["satir"], alan["sutun"])
                        hedef_hucre.text = deger

                    doc.save(output_path)
                    st.success("Form başarıyla ve kişisel bilgilerinizle dolduruldu.")

                    with open(output_path, "rb") as f:
                        st.download_button(
                            label="Doldurulmuş Formu İndir",
                            data=f,
                            file_name="otomatik_doldurulmus_form.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
                except Exception as e:
                    st.error(f"İşlem sırasında bir hata oluştu: {e}")
    else:
        st.warning("Belgede otomatik tespit edilebilir boş alan bulunamadı.")
        

import os
import shutil
import streamlit as st
from docx import Document

st.set_page_config(page_title="Bulut Form Doldurucu", page_icon="📄", layout="centered")

st.title("Bulut Tabanlı Form Yönetim Sistemi")
st.write("Bu sistemi telefonunuzdan veya herhangi bir cihazdan tarayıcı ile kullanarak formlarınızı bulut üzerinden doldurabilirsiniz.")

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
        
        girilen_veriler = {}
        for alan in bos_alanlar:
            anahtar_adi = alan["etiket"]
            key_id = f"t_{alan['tablo']}_s_{alan['satir']}_c_{alan['sutun']}"
            girilen_veriler[key_id] = {
                "tablo": alan["tablo"],
                "satir": alan["satir"],
                "sutun": alan["sutun"],
                "deger": st.text_input(f"'{anahtar_adi}' alanı için değer:")
            }

        if st.button("Formu Yapıyı Bozmadan Doldur"):
            for k, veri in girilen_veriler.items():
                if veri["deger"]:
                    hedef_hucre = doc.tables[veri["tablo"]].cell(veri["satir"], veri["sutun"])
                    hedef_hucre.text = str(veri["deger"])

            doc.save(output_path)

            with open(output_path, "rb") as f:
                st.download_button(
                    label="Güncellenmiş Formu İndir",
                    data=f,
                    file_name="tamamlanmis_form.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            st.success("Form başarıyla hazırlandı.")
    else:
        st.warning("Belgede otomatik tespit edilebilir boş alan bulunamadı.")
      

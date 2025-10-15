import streamlit as st
import fitz  # PyMuPDF
from PIL import Image
import io
import pandas as pd
import os
import shutil

st.set_page_config(page_title="Extraction des photos et des plans Archipad", layout="wide")
st.title("📄 Extraction Photos PDF depuis Archipad")

st.markdown("""
Cette application permet d'extraire :
- Les photos des désordres
- Les plans depuis Archipad
""")

# --- INITIALISATION session_state ---
for key in ['uploaded_excel', 'uploaded_pdf', 'col_values', 'nb_unique', 'extracted', 'zip_path']:
    if key not in st.session_state:
        st.session_state[key] = None

# --- Upload Excel ---
col1, col2 = st.columns(2)

with col1:
    uploaded_excel = st.file_uploader("📂 Choisis ton fichier Excel Archipad (.xlsx)", type="xlsx", key="excel_uploader")
    if uploaded_excel is not None:
        st.session_state.uploaded_excel = uploaded_excel
        df = pd.read_excel(uploaded_excel, sheet_name="Observations")
        st.session_state.col_values = df["Plan"].dropna().tolist()
        st.session_state.nb_unique = len(set(st.session_state.col_values))
        st.success(f"✅ Rapport Excel Archipad importé avec succès !")

# --- Upload PDF ---
with col2:
    uploaded_pdf = st.file_uploader("📂 Choisis ton fichier PDF Archipad", type="pdf", key="pdf_uploader")
    if uploaded_pdf is not None:
        st.session_state.uploaded_pdf = uploaded_pdf
        st.success(f"✅ Rapport PDF Archipad importé avec succès !")

# --- Extraction si fichiers chargés et non déjà extraits ---
if (st.session_state.uploaded_excel and st.session_state.uploaded_pdf 
        and st.session_state.nb_unique is not None 
        and not st.session_state.extracted):

    output_folder = "Extraction_temp"
    if os.path.exists(output_folder):
        shutil.rmtree(output_folder)
    os.makedirs(output_folder, exist_ok=True)

    doc = fitz.open(stream=st.session_state.uploaded_pdf.read(), filetype="pdf")
    count = 0
    pages_to_extract = len(doc) - st.session_state.nb_unique

    # --- Extraction photos de désordres ---
    extraction_photos_msg = st.info("⏳ Extraction des photos de désordres …")
    progress_bar = st.progress(0)
    for page_num in range(pages_to_extract):
        page = doc[page_num]
        images = page.get_images(full=True)
        for img_index, img in enumerate(images[1:], start=2):
            xref = img[0]
            base_image = doc.extract_image(xref)
            image_bytes = base_image["image"]
            image_ext = base_image["ext"]
            image = Image.open(io.BytesIO(image_bytes))
            count += 1
            image_filename = f"img{count}.{image_ext}"
            image.save(os.path.join(output_folder, image_filename))
        progress_bar.progress((page_num+1)/pages_to_extract)
    extraction_photos_msg.empty()
    progress_bar.empty()
    st.success(f"✅ Photos de désordres extraites")

    # --- Extraction des plans ---
    extraction_plans_msg = st.info("⏳ Extraction des plans …")
    last_pages = range(len(doc) - st.session_state.nb_unique, len(doc))
    for idx, page_num in enumerate(last_pages, start=1):
        page = doc[page_num]
        pix = page.get_pixmap(dpi=200)
        page_filename = f"P{idx}.png"
        pix.save(os.path.join(output_folder, page_filename))
    extraction_plans_msg.empty()
    st.success(f"✅ Plans extraits")

    # --- Création ZIP ---
    st.session_state.zip_path = "Extraction_finale.zip"
    shutil.make_archive(st.session_state.zip_path.replace(".zip", ""), 'zip', output_folder)
    st.session_state.extracted = True

# --- Bouton téléchargement ---
if st.session_state.extracted and st.session_state.zip_path is not None:
    with open(st.session_state.zip_path, "rb") as f:
        st.download_button(
            label="⬇️ Télécharger le dossier ZIP",
            data=f,
            file_name="Extraction_finale.zip",
            mime="application/zip"
        )



import streamlit as st
import fitz  # PyMuPDF
from PIL import Image
import io
import pandas as pd
import os
import shutil
import re

st.set_page_config(page_title="Extraction Photos PDF", layout="wide")
st.title("📄 Extraction Photos PDF depuis Archipad")

st.markdown("""
Cette application permet d'extraire :
- Les photos des désordres et les plans depuis Archipad
- Vérification automatique de cohérence entre le nombre de désordres et le nombre de photos par désordre
""")

# --- Bouton Nouvelle extraction ---
if st.button("🔄 Nouvelle extraction"):
    for key in st.session_state.keys():
        del st.session_state[key]
    st.experimental_rerun()

# --- Upload Excel ---
col1, col2 = st.columns(2)

with col1:
    uploaded_excel = st.file_uploader("📂 Choisis ton fichier Excel Archipad (.xlsx)", type="xlsx", key="excel")
    if uploaded_excel and "df" not in st.session_state:
        st.session_state.df = pd.read_excel(uploaded_excel, sheet_name="Observations")
        st.session_state.col_values = st.session_state.df["Plan"].dropna().tolist()
        st.session_state.nb_unique = len(set(st.session_state.col_values))
        st.session_state.excel_imported = True

    if st.session_state.get("excel_imported"):
        st.success("✅ Rapport Excel Archipad importé avec succès !")

# --- Upload PDF ---
with col2:
    uploaded_pdf = st.file_uploader("📂 Choisis ton fichier PDF Archipad", type="pdf", key="pdf")
    if uploaded_pdf and "pdf_data" not in st.session_state:
        st.session_state.pdf_data = uploaded_pdf.read()
        st.session_state.pdf_imported = True

    if st.session_state.get("pdf_imported"):
        st.success("✅ Rapport PDF Archipad importé avec succès !")

# --- Extraction si les deux fichiers sont chargés ---
if st.session_state.get("excel_imported") and st.session_state.get("pdf_imported"):

    # Préparer dossier temporaire
    output_folder = "Extraction_temp"
    if os.path.exists(output_folder):
        shutil.rmtree(output_folder)
    os.makedirs(output_folder, exist_ok=True)

    doc = fitz.open(stream=st.session_state.pdf_data, filetype="pdf")
    count = 0
    pages_to_extract = len(doc) - st.session_state.nb_unique

    # --- Extraction des photos de désordres ---
    if "photos_extracted" not in st.session_state:
        extraction_photos_msg = st.info("⏳ Extraction des photos de désordres …")
        progress_bar = st.progress(0)
        for page_num in range(pages_to_extract):
            page = doc[page_num]
            images = page.get_images(full=True)
            for img_index, img in enumerate(images, start=1):
                xref = img[0]
                base_image = doc.extract_image(xref)
                image_bytes = base_image["image"]
                image_ext = base_image["ext"]
                image = Image.open(io.BytesIO(image_bytes))
                count += 1
                image_filename = f"img{count}.{image_ext}"
                image.save(os.path.join(output_folder, image_filename))
            progress_bar.progress((page_num+1)/pages_to_extract)
        st.session_state.photos_extracted = count
        progress_bar.empty()
        extraction_photos_msg.empty()

    st.success(f"✅ {st.session_state.photos_extracted} photos de désordres extraites")

    # --- Extraction des plans ---
    if "plans_extracted" not in st.session_state:
        extraction_plans_msg = st.info("⏳ Extraction des plans …")
        last_pages = range(len(doc) - st.session_state.nb_unique, len(doc))
        for idx, page_num in enumerate(last_pages, start=1):
            page = doc[page_num]
            pix = page.get_pixmap(dpi=200)
            page_filename = f"P{idx}.png"
            pix.save(os.path.join(output_folder, page_filename))
        st.session_state.plans_extracted = st.session_state.nb_unique
        extraction_plans_msg.empty()

    st.success(f"✅ {st.session_state.plans_extracted} plans extraits")

    # --- Supprimer img1, img8, img15, … ---
    for file in os.listdir(output_folder):
        if file.startswith("img"):
            match = re.match(r"img(\d+)", file)
            if match:
                num = int(match.group(1))
                if (num - 1) % 7 == 0:
                    os.remove(os.path.join(output_folder, file))

    # --- Vérification cohérence ---
    nb_img_restantes = len([f for f in os.listdir(output_folder) if f.startswith("img")])
    nb_lignes_plan = len(st.session_state.col_values)

    if not (nb_img_restantes == nb_lignes_plan or nb_img_restantes // 2 == nb_lignes_plan):
        st.error("❌ Incohérence détectée : vérifie le nombre de photos par désordre sur Archipad.")
        shutil.rmtree(output_folder)
        st.stop()
    else:
        st.success("✅ Vérification OK : nombre de photos par désordre respecté")

    # --- Création ZIP ---
    zip_path = "Extraction_finale.zip"
    shutil.make_archive(zip_path.replace(".zip", ""), 'zip', output_folder)

    # --- Bouton téléchargement (ne relance pas l’extraction) ---
    with open(zip_path, "rb") as f:
        st.download_button(
            label="⬇️ Télécharger le dossier ZIP",
            data=f,
            file_name="Extraction_finale.zip",
            mime="application/zip"
        )

    # --- Nettoyage ---
    shutil.rmtree(output_folder)
    os.remove(zip_path)

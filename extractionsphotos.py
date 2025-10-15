import streamlit as st
import fitz
from PIL import Image
import io
import pandas as pd
import os
import shutil
import re
import tempfile

st.set_page_config(page_title="Extraction Photos PDF", layout="wide")
st.title("📄 Extraction Photos PDF depuis Archipad")

st.markdown("""
Cette application permet d'extraire :
- Les photos des désordres et les plans depuis Archipad
- Vérification automatique de cohérence entre le nombre de désordres et le nombre de photos par désordre
""")

# --- Nouvelle extraction ---
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
        st.success("✅ Rapport Excel Archipad importé avec succès !")

with col2:
    uploaded_pdf = st.file_uploader("📂 Choisis ton fichier PDF Archipad", type="pdf", key="pdf")
    if uploaded_pdf and "pdf_data" not in st.session_state:
        st.session_state.pdf_data = uploaded_pdf.read()
        st.success("✅ Rapport PDF Archipad importé avec succès !")

# --- Extraction ---
if "df" in st.session_state and "pdf_data" in st.session_state and "nb_unique" in st.session_state:
    if "zip_path" not in st.session_state:
        # Préparer dossier temporaire
        output_folder = "Extraction_temp"
        if os.path.exists(output_folder):
            shutil.rmtree(output_folder)
        os.makedirs(output_folder, exist_ok=True)

        doc = fitz.open(stream=st.session_state.pdf_data, filetype="pdf")
        count = 0
        pages_to_extract = len(doc) - st.session_state.nb_unique

        # Photos désordres
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
        extraction_photos_msg.empty()
        progress_bar.empty()
        st.success(f"✅ {count} photos de désordres extraites")

        # Plans
        extraction_plans_msg = st.info("⏳ Extraction des plans …")
        last_pages = range(len(doc) - st.session_state.nb_unique, len(doc))
        for idx, page_num in enumerate(last_pages, start=1):
            page = doc[page_num]
            pix = page.get_pixmap(dpi=200)
            page_filename = f"P{idx}.png"
            pix.save(os.path.join(output_folder, page_filename))
        extraction_plans_msg.empty()
        st.success(f"✅ {st.session_state.nb_unique} plans extraits")

        # Supprimer img1, img8, …
        for file in os.listdir(output_folder):
            if file.startswith("img"):
                match = re.match(r"img(\d+)", file)
                if match:
                    num = int(match.group(1))
                    if (num - 1) % 7 == 0:
                        os.remove(os.path.join(output_folder, file))

        # Vérification cohérence
        nb_img_restantes = len([f for f in os.listdir(output_folder) if f.startswith("img")])
        nb_lignes_plan = len(st.session_state.col_values)
        if not (nb_img_restantes == nb_lignes_plan or nb_img_restantes // 2 == nb_lignes_plan):
            st.error("❌ Incohérence détectée !")
            shutil.rmtree(output_folder)
            st.stop()
        else:
            st.success("✅ Vérification OK : nombre de photos par désordre respecté")

        # Création ZIP
        temp_zip = tempfile.NamedTemporaryFile(delete=False, suffix=".zip")
        shutil.make_archive(temp_zip.name.replace(".zip",""), 'zip', output_folder)
        st.session_state.zip_path = temp_zip.name

        # Nettoyage dossier
        shutil.rmtree(output_folder)

    # --- Bouton téléchargement ---
    with open(st.session_state.zip_path, "rb") as f:
        st.download_button(
            label="⬇️ Télécharger le dossier ZIP",
            data=f,
            file_name="Extraction_finale.zip",
            mime="application/zip"
        )

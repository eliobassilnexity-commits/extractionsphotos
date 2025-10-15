import streamlit as st
import fitz
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

# --- Bouton pour recommencer ---
if st.button("🔄 Nouvelle extraction"):
    for key in list(st.session_state.keys()):
        del st.session_state[key]
    st.experimental_rerun()

# --- Initialisation de la timeline ---
#if "timeline" not in st.session_state:
   # st.session_state.timeline = {
   #    "excel_uploaded": "⏳ Excel non chargé",
   #     "pdf_uploaded": "⏳ PDF non chargé",
   #     "photos_extracted": "⏳ Photos de désordres non extraites",
   #     "plans_extracted": "⏳ Plans non extraits",
   #     "zip_ready": "⏳ ZIP non prêt",
   #      "coherence_checked": "⏳ Vérification non effectuée"
 #   }

def display_timeline():
    for val in st.session_state.timeline.values():
        st.markdown(f"- {val}")

st.subheader("🕒 État du processus")
display_timeline()

# --- Upload Excel ---
col1, col2 = st.columns(2)
nb_unique = None

with col1:
    uploaded_excel = st.file_uploader("📂 Choisis ton fichier Excel Archipad (.xlsx)", type="xlsx", key="excel_uploader")
    if uploaded_excel:
        df = pd.read_excel(uploaded_excel, sheet_name="Observations")
        col_values = df["Plan"].dropna().tolist()
        nb_unique = len(set(col_values))
        st.session_state.timeline["excel_uploaded"] = "✅ Excel chargé"
        st.success(f"✅ Rapport Excel Archipad importé avec succès !")
        display_timeline()

with col2:
    uploaded_pdf = st.file_uploader("📂 Choisis ton fichier PDF Archipad", type="pdf", key="pdf_uploader")
    if uploaded_pdf:
        st.session_state.timeline["pdf_uploaded"] = "✅ PDF chargé"
        st.success(f"✅ Rapport PDF Archipad importé avec succès !")
        display_timeline()

# --- Extraction si les deux fichiers sont chargés ---
if uploaded_excel and uploaded_pdf and nb_unique is not None:

    if "zip_ready" not in st.session_state:

        output_folder = "Extraction_temp"
        if os.path.exists(output_folder):
            shutil.rmtree(output_folder)
        os.makedirs(output_folder, exist_ok=True)

        doc = fitz.open(stream=uploaded_pdf.read(), filetype="pdf")
        count = 0
        pages_to_extract = len(doc) - nb_unique

        # --- Extraction des photos de désordres ---
        extraction_photos_msg = st.info("⏳ Extraction des photos de désordres …")
        progress_bar_photos = st.progress(0)
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
            progress_bar_photos.progress((page_num + 1) / pages_to_extract)
        extraction_photos_msg.empty()
        progress_bar_photos.empty()
        st.session_state.timeline["photos_extracted"] = f"✅ {count} photos extraites"
        st.success(f"✅ {count} photos de désordres extraites")
        display_timeline()

        # --- Extraction des plans ---
        extraction_plans_msg = st.info("⏳ Extraction des plans …")
        last_pages = range(len(doc) - nb_unique, len(doc))
        progress_bar_plans = st.progress(0)
        for idx, page_num in enumerate(last_pages, start=1):
            page = doc[page_num]
            pix = page.get_pixmap(dpi=200)
            page_filename = f"P{idx}.png"
            pix.save(os.path.join(output_folder, page_filename))
            progress_bar_plans.progress(idx / nb_unique)
        extraction_plans_msg.empty()
        progress_bar_plans.empty()
        st.session_state.timeline["plans_extracted"] = f"✅ {nb_unique} plans extraits"
        st.success(f"✅ {nb_unique} plans extraits")
        display_timeline()

        # --- Supprimer img1, img8, img15, …
        for file in os.listdir(output_folder):
            if file.startswith("img"):
                match = re.match(r"img(\d+)", file)
                if match:
                    num = int(match.group(1))
                    if (num - 1) % 7 == 0:
                        os.remove(os.path.join(output_folder, file))

        # --- Vérification cohérence ---
        nb_img_restantes = len([f for f in os.listdir(output_folder) if f.startswith("img")])
        nb_lignes_plan = len(col_values)
        if not (nb_img_restantes == nb_lignes_plan or nb_img_restantes // 2 == nb_lignes_plan):
            st.error("❌ Incohérence détectée : vérifie le nombre de photos par désordre sur Archipad.")
            shutil.rmtree(output_folder)
            st.stop()
        else:
            st.session_state.timeline["coherence_checked"] = "✅ Vérification OK"
            st.success("✅ Vérification OK : nombre de photos par désordre respecté")
            display_timeline()

        # --- Création ZIP ---
        zip_path = "Extraction_finale.zip"
        shutil.make_archive(zip_path.replace(".zip", ""), 'zip', output_folder)
        st.session_state.timeline["zip_ready"] = "✅ ZIP prêt"
        st.session_state.zip_ready = zip_path
        st.session_state.output_folder = output_folder
        display_timeline()

    # --- Bouton téléchargement ---
    with open(st.session_state.zip_ready, "rb") as f:
        if st.download_button(
            label="⬇️ Télécharger le dossier ZIP",
            data=f,
            file_name="Extraction_finale.zip",
            mime="application/zip"
        ):
            # Nettoyage après téléchargement
            shutil.rmtree(st.session_state.output_folder)
            os.remove(st.session_state.zip_ready)
            st.session_state.pop("zip_ready")
            st.session_state.pop("output_folder")
            st.success("🧹 Nettoyage terminé après téléchargement !")


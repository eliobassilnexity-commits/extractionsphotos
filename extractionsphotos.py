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
- Vérification automatique de cohérence
""")

# --- Upload Excel ---
col1, col2 = st.columns(2)

with col1:
    uploaded_excel = st.file_uploader("📂 Choisis ton fichier Excel Archipad (.xlsx)", type="xlsx")
    nb_unique = None
    if uploaded_excel:
        df = pd.read_excel(uploaded_excel, sheet_name="Observations")
        col_values = df["Plan"].dropna().tolist()
        nb_unique = len(set(col_values))
        st.success(f"📊 Nombre de valeurs distinctes dans 'Plan' : {nb_unique}")
        st.info(f"Nombre total de lignes non vides dans 'Plan' : {len(col_values)}")

# --- Upload PDF ---
with col2:
    uploaded_pdf = st.file_uploader("📂 Choisis ton fichier PDF Archipad", type="pdf")

if uploaded_excel and uploaded_pdf and nb_unique is not None:
    output_folder = "Extraction_temp"
    if os.path.exists(output_folder):
        shutil.rmtree(output_folder)
    os.makedirs(output_folder, exist_ok=True)

    doc = fitz.open(stream=uploaded_pdf.read(), filetype="pdf")
    st.info(f"📄 PDF chargé : {len(doc)} pages")
    count = 0
    pages_to_extract = len(doc) - nb_unique

    st.info("⏳ Extraction des images internes…")
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

    st.success(f"✅ {count} images internes extraites (img1, img2, …)")

    st.info("⏳ Extraction des dernières pages en images…")
    last_pages = range(len(doc) - nb_unique, len(doc))
    for idx, page_num in enumerate(last_pages, start=1):
        page = doc[page_num]
        pix = page.get_pixmap(dpi=200)
        page_filename = f"P{idx}.png"
        pix.save(os.path.join(output_folder, page_filename))
    st.success(f"✅ Dernières pages extraites : {nb_unique} pages (P1, P2, …)")

    # Supprimer img1, img8, img15, …
    for file in os.listdir(output_folder):
        if file.startswith("img"):
            match = re.match(r"img(\d+)", file)
            if match:
                num = int(match.group(1))
                if (num - 1) % 7 == 0:
                    os.remove(os.path.join(output_folder, file))

    # Vérification cohérence
    nb_img_restantes = len([f for f in os.listdir(output_folder) if f.startswith("img")])
    nb_lignes_plan = len(col_values)

    st.info(f"🔍 Images restantes : {nb_img_restantes}, Lignes Excel 'Plan' : {nb_lignes_plan}")

    if not (nb_img_restantes == nb_lignes_plan or nb_img_restantes // 2 == nb_lignes_plan):
        st.error("❌ Incohérence détectée ! Vérifie le nombre de photos par désordre sur Archipad.")
        shutil.rmtree(output_folder)
        st.stop()
    else:
        st.success("✅ Vérification OK : cohérence respectée")

    # Création ZIP
    zip_path = "Extraction_finale.zip"
    shutil.make_archive(zip_path.replace(".zip", ""), 'zip', output_folder)

    # Bouton téléchargement
    with open(zip_path, "rb") as f:
        st.download_button(
            label="⬇️ Télécharger le dossier ZIP",
            data=f,
            file_name="Extraction_finale.zip",
            mime="application/zip"
        )

    # Nettoyage
    shutil.rmtree(output_folder)
    os.remove(zip_path)
    st.success("🧹 Nettoyage terminé")



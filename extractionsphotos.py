import streamlit as st
import fitz
from PIL import Image
import io
import pandas as pd
import os
import shutil

st.set_page_config(page_title="Extraction des photos et des plans Archipad", layout="wide")
st.title("📄 Extraction des photos et des plans Archipad")

st.markdown("""
Cette application permet d'extraire depuis les rapports d'Archipad :
- Les photos des désordres
- Les plans
- Un fichier Excel "repère" indiquant les dimensions (en points) des pages de plans
""")

# --- INITIALISATION session_state ---
for key in ['uploaded_excel', 'uploaded_pdf', 'col_values', 'nb_unique', 
            'extracted', 'zip_path', 'progress_photos', 'progress_plans', 'tailles_pages']:
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
        st.success("✅ Rapport Excel Archipad importé avec succès !")

# --- Upload PDF ---
with col2:
    uploaded_pdf = st.file_uploader("📂 Choisis ton fichier PDF Archipad", type="pdf", key="pdf_uploader")
    if uploaded_pdf is not None:
        st.session_state.uploaded_pdf = uploaded_pdf
        st.success("✅ Rapport PDF Archipad importé avec succès !")

# --- Extraction ---
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

    # --- Photos ---
    if st.session_state.progress_photos is None:
        st.session_state.progress_photos = 0
    if st.session_state.progress_photos == 0:
        extraction_photos_msg = st.info("⏳ Extraction des photos de désordres …")
        progress_bar_photos = st.progress(0)
        for page_num in range(pages_to_extract):
            if page_num == 0:
                continue
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
            progress_bar_photos.progress((page_num+1)/pages_to_extract)
        extraction_photos_msg.empty()
        progress_bar_photos.empty()
        st.success("✅ Photos extraites")
        st.session_state.progress_photos = 1
    else:
        st.success("✅ Photos déjà extraites")

    # --- Plans ---
    if st.session_state.progress_plans is None:
        st.session_state.progress_plans = 0
    if st.session_state.progress_plans == 0:
        extraction_plans_msg = st.info("⏳ Extraction des plans …")
        last_pages = range(len(doc) - st.session_state.nb_unique, len(doc))
        st.session_state.tailles_pages = []
        for idx, page_num in enumerate(last_pages, start=1):
            page = doc[page_num]
            rect = page.rect
            st.session_state.tailles_pages.append({
                "Plan": f"P{idx}",
                "Largeur (pt)": rect.width,
                "Hauteur (pt)": rect.height
            })
            pix = page.get_pixmap(dpi=200)
            page_filename = f"P{idx}.png"
            pix.save(os.path.join(output_folder, page_filename))
        extraction_plans_msg.empty()
        st.success("✅ Plans extraits")
        st.session_state.progress_plans = 1
    else:
        st.success("✅ Plans déjà extraits")

    # --- Excel repère ---
    df_tailles = pd.DataFrame(st.session_state.tailles_pages)
    excel_repere_path = os.path.join(output_folder, "excel_repere.xlsx")
    df_tailles.to_excel(excel_repere_path, index=False)
    # st.success("📊 Fichier 'excel_repere.xlsx' généré")

    # --- Copier Excel original ---
    excel_orig_copy_path = os.path.join(output_folder, "excelarchipad.xlsx")
    with open(excel_orig_copy_path, "wb") as f:
        f.write(st.session_state.uploaded_excel.getbuffer())
    # st.success("📊 Copie de l'Excel original ajoutée")

    # --- Vérification photos ---
    nb_img_restantes = len([f for f in os.listdir(output_folder) if f.startswith("img")])
    nb_lignes_plan = len(st.session_state.col_values)

    # Première vérification : ignorer la dernière page
    nb_images_derniere_page = len(doc[pages_to_extract - 1].get_images(full=True)) - 1
    nb_photos_sans_derniere_page = nb_img_restantes - nb_images_derniere_page

    if nb_photos_sans_derniere_page == nb_lignes_plan:
        st.success("✅ Vérification OK : 1 photo par désordre (hors dernière page)")
    elif nb_photos_sans_derniere_page == nb_lignes_plan * 2:
        st.success("✅ Vérification OK : 2 photos par désordre (hors dernière page)")
    else:
        # Vérification finale : inclure toutes les photos
        if nb_img_restantes == nb_lignes_plan:
            st.success("✅ Vérification finale OK : 1 photo par désordre (toutes pages)")
        elif nb_img_restantes == nb_lignes_plan * 2:
            st.success("✅ Vérification finale OK : 2 photos par désordre (toutes pages)")
        else:
            st.error(f"❌ Incohérence détectée : nombre total de photos = {nb_img_restantes}")
            shutil.rmtree(output_folder)
            st.stop()

    # --- Création ZIP ---
    zip_path = "Extraction_finale.zip"
    shutil.make_archive(zip_path.replace(".zip", ""), 'zip', output_folder)
    st.session_state.zip_path = zip_path
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


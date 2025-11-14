import datetime

print("App visit detected at:", datetime.datetime.now())



import streamlit as st
import fitz
from PIL import Image
import io
import pandas as pd
import os
import shutil
import re



st.set_page_config(page_title="Extraction des photos et des plans Archipad", layout="wide")
st.title("📄 Extraction des photos et des plans Archipad")

st.markdown("""
Cette application permet d'extraire depuis les rapports d'Archipad :
- Les photos des désordres (avant de lancer cette application, s'assurer sur archipad que chaque désordre possède 2 photos)
- Les plans
- Un fichier Excel "repère" indiquant les dimensions (en points) des pages de plans
- Une copie du fichier excel archipad importé

(Pour pouvoir réutiliser l’application, il faudra actualiser la page.)
""")

# --- INITIALISATION session_state ---
for key in ['uploaded_excel', 'uploaded_pdf', 'col_values', 'nb_unique', 
            'extracted', 'zip_path', 'progress_photos', 'progress_plans', 'tailles_pages', 'plan_names']:
    if key not in st.session_state:
        st.session_state[key] = None

# --- Upload Excel ---
col1, col2 = st.columns(2)
with col1:
    uploaded_excel = st.file_uploader("📂 Choisis ton fichier Excel Archipad (.xlsx)", type="xlsx", key="excel_uploader")
    if uploaded_excel is not None:
        st.session_state.uploaded_excel = uploaded_excel
        df = pd.read_excel(uploaded_excel, sheet_name="Observations")

        # On prend les noms des plans dans la colonne H
        col_values = df["Plan"].dropna().astype(str).tolist()

        # On nettoie les extensions (".pdf", ".png", ".jpg", etc.)
        cleaned_plan_names = []
        for name in col_values:
            clean_name = re.sub(r'\.[a-zA-Z0-9]+$', '', name.strip())  # supprime extension éventuelle
            cleaned_plan_names.append(clean_name)

        # Plans uniques en conservant l’ordre d’apparition
        unique_plan_names = list(dict.fromkeys(cleaned_plan_names))

        st.session_state.col_values = cleaned_plan_names
        st.session_state.nb_unique = len(unique_plan_names)
        st.session_state.plan_names = unique_plan_names

        st.success(f"✅ Rapport Excel archipad importé avec succès !")

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
    pages_to_extract = len(doc) - st.session_state.nb_unique  # pages avec photos
    photo_last_page_idx = pages_to_extract - 1                 # dernière page "photos" (exclue du contrôle 3/6)

    # --- Photos ---
    if st.session_state.progress_photos is None:
        st.session_state.progress_photos = 0
    if st.session_state.progress_photos == 0:
        extraction_photos_msg = st.info("⏳ Extraction des photos de désordres …")
        progress_bar_photos = st.progress(0)
        for page_num in range(pages_to_extract):
            if page_num == 0:
                # On continue à ignorer la première page comme dans ton code d'origine
                progress_bar_photos.progress((page_num+1)/pages_to_extract)
                continue

            page = doc[page_num]
            images = page.get_images(full=True)

            # --- 🔎 Vérification de cohérence page par page (3 ou 6 photos) ---
            # On contrôle toutes les pages "photos" SAUF la dernière (photo_last_page_idx)
            # Le nombre de photos extraites par page correspond à len(images) - 1,
            # car on saute systématiquement la première image (images[1:], arrière-plan).
           # if page_num != photo_last_page_idx:
            #    nb_photos_page = max(len(images) - 1, 0)
             #   if nb_photos_page not in (3, 6):
                    # Nettoyage visuel et suppression du dossier temporaire avant arrêt
              #      extraction_photos_msg.empty()
               #     progress_bar_photos.empty()
                #    st.error(
                 #       f"❌ Incohérence détectée dans le nombre de photos par désordre à la page {page_num + 1}. Corrige celà directement sur le projet Archipad, exporte à nouveau les rapports PDF et Excel, puis reviens ici pour le traitement."
                  #  )
                  #  if os.path.exists(output_folder):
                   #     shutil.rmtree(output_folder)
                  #  st.stop()

            # --- Extraction effective des images de la page (en conservant le comportement existant) ---
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
        st.success("✅ Photos extraites")

    # --- Plans ---
    if st.session_state.progress_plans is None:
        st.session_state.progress_plans = 0
    if st.session_state.progress_plans == 0:
        extraction_plans_msg = st.info("⏳ Extraction des plans …")
        last_pages = range(len(doc) - st.session_state.nb_unique, len(doc))
        st.session_state.tailles_pages = []

        # Utiliser les noms uniques issus d'Excel
        plan_names = st.session_state.plan_names

        for idx, page_num in enumerate(last_pages):
            plan_name = plan_names[idx] if idx < len(plan_names) else f"Plan_{idx+1}"
            safe_name = re.sub(r'[\\/*?:"<>|]', "_", plan_name)  # retire caractères illégaux
            page = doc[page_num]
            rect = page.rect
            st.session_state.tailles_pages.append({
                "Plan": safe_name,
                "Largeur (pt)": rect.width,
                "Hauteur (pt)": rect.height
            })
            pix = page.get_pixmap(dpi=200)
            page_filename = f"{safe_name}.png"
            pix.save(os.path.join(output_folder, page_filename))
        extraction_plans_msg.empty()
        st.success("✅ Plans extraits")
        st.session_state.progress_plans = 1
    else:
        st.success("✅ Plans extraits")

    # --- Excel repère ---
    df_tailles = pd.DataFrame(st.session_state.tailles_pages)
    excel_repere_path = os.path.join(output_folder, "excel_repere.xlsx")
    df_tailles.to_excel(excel_repere_path, index=False)
    st.success("✅ Fichier excel repère crée")

    # --- Copier Excel original ---
    excel_orig_copy_path = os.path.join(output_folder, "excelarchipad.xlsx")
    with open(excel_orig_copy_path, "wb") as f:
        f.write(st.session_state.uploaded_excel.getbuffer())
    st.success("✅ Copie du fichier excel archipad créée")
                    
    # --- Vérification photos (globale, conservée) ---
    nb_img_restantes = len([f for f in os.listdir(output_folder) if f.startswith("img")])
    nb_lignes_plan = len(st.session_state.col_values)

    nb_images_derniere_page = len(doc[pages_to_extract - 1].get_images(full=True)) - 1
    nb_photos_sans_derniere_page = nb_img_restantes - nb_images_derniere_page

    if nb_photos_sans_derniere_page == nb_lignes_plan:
        st.success("✅ Vérification OK : 1 photo par désordre (hors dernière page)")
    elif nb_photos_sans_derniere_page == nb_lignes_plan * 2:
        st.success("✅ Vérification OK : 2 photos par désordre (hors dernière page)")
    else:
        if nb_img_restantes == nb_lignes_plan:
            st.success("✅ Vérification finale OK : 1 photo par désordre (toutes pages)")
        elif nb_img_restantes == nb_lignes_plan * 2:
            st.success("✅ Vérification finale OK : 2 photos par désordre (toutes pages)")
        else:
            st.error(f"❌ Incohérence détectée dans le nombre de photos par désordre. Corrige celà directement sur le projet Archipad, exporte à nouveau les rapports PDF et Excel, puis reviens ici pour le traitement.")
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













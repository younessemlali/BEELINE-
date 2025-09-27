import streamlit as st
import pandas as pd
import pdfplumber
from rapidfuzz import fuzz
import unidecode
import io

def normalize(s):
    if pd.isna(s) or s is None:
        return ""
    return unidecode.unidecode(str(s)).replace(" ", "").lower()

def extract_pdf_text(pdf_file):
    try:
        with pdfplumber.open(pdf_file) as pdf:
            text = ""
            for page in pdf.pages:
                t = page.extract_text()
                if t:
                    text += t
            return text
    except Exception as e:
        return ""

def fuzzy_match(val, text, threshold=90):
    # Utilise RapidFuzz pour trouver des correspondances floues
    if not val or not text:
        return False
    return fuzz.partial_ratio(val, text) >= threshold

st.title("Rapprochement multi-fichiers PDF ↔ Excel – Version robuste")

pdf_files = st.file_uploader("Importer plusieurs PDF", type="pdf", accept_multiple_files=True)
excel_files = st.file_uploader("Importer plusieurs Excel", type=["xlsx", "xls"], accept_multiple_files=True)
fuzzy = st.checkbox("Activer le rapprochement flou (tolérance d'écriture)", value=True)
fuzzy_threshold = st.slider("Niveau de tolérance du fuzzy matching (plus élevé = plus strict)", 70, 100, 90) if fuzzy else 100

if pdf_files and excel_files:
    st.info("Traitement en cours, merci de patienter…")
    results = []
    progress = st.progress(0)
    total = len(pdf_files) * len(excel_files)
    count = 0

    # Extraction et rapprochement
    for i, excel_file in enumerate(excel_files):
        try:
            df = pd.read_excel(excel_file)
        except Exception as e:
            st.error(f"Erreur de lecture du fichier Excel {excel_file.name} : {e}")
            continue

        for j, pdf_file in enumerate(pdf_files):
            pdf_text = extract_pdf_text(pdf_file)
            if not pdf_text:
                st.warning(f"PDF {pdf_file.name} : texte non extrait ou vide.")
            # Afficher un extrait du PDF pour contrôle
            st.subheader(f"Extrait du PDF : {pdf_file.name}")
            st.write(pdf_text[:1000] + "..." if len(pdf_text) > 1000 else pdf_text)

            matched = []
            norm_pdf_text = normalize(pdf_text)

            for col in df.columns:
                for val in df[col]:
                    norm_val = normalize(val)
                    if not norm_val:
                        continue
                    found = False
                    # Recherche exacte
                    if norm_val in norm_pdf_text:
                        found = True
                    # Recherche floue si activée
                    elif fuzzy and fuzzy_match(norm_val, norm_pdf_text, threshold=fuzzy_threshold):
                        found = True
                    if found:
                        matched.append(val)
            matched = list(set(matched))
            results.append({
                'excel': excel_file.name,
                'pdf': pdf_file.name,
                'matched_values': matched
            })
            count += 1
            progress.progress(min(count / total, 1.0))

    # Affichage des résultats
    st.success("Traitement terminé !")
    export_rows = []
    for res in results:
        st.write(f"**Fichier Excel :** {res['excel']}  —  **Fichier PDF :** {res['pdf']}")
        if res['matched_values']:
            st.write("Valeurs trouvées :", res['matched_values'])
            for v in res['matched_values']:
                export_rows.append({'Excel': res['excel'], 'PDF': res['pdf'], 'Valeur trouvée': v})
        else:
            st.write(":x: Aucune valeur trouvée.")
        st.write("---")
    # Export CSV
    if export_rows:
        export_df = pd.DataFrame(export_rows)
        csv = export_df.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="Télécharger les résultats au format CSV",
            data=csv,
            file_name='resultats_rapprochement.csv',
            mime='text/csv'
        )

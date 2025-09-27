import streamlit as st
import pandas as pd
import pdfplumber

st.title("Rapprochement multi-fichiers PDF ↔ Excel")

pdf_files = st.file_uploader("Importer plusieurs PDF", type="pdf", accept_multiple_files=True)
excel_files = st.file_uploader("Importer plusieurs Excel", type=["xlsx", "xls"], accept_multiple_files=True)

if pdf_files and excel_files:
    # Extraction des textes des PDFs
    pdf_texts = []
    for pdf_file in pdf_files:
        with pdfplumber.open(pdf_file) as pdf:
            text = ""
            for page in pdf.pages:
                text += page.extract_text() or ""
            pdf_texts.append({'filename': pdf_file.name, 'text': text})

    # Extraction des données des Excels
    excel_dfs = []
    for excel_file in excel_files:
        df = pd.read_excel(excel_file)
        excel_dfs.append({'filename': excel_file.name, 'df': df})

    # Rapprochement : pour chaque Excel, chercher les valeurs dans chaque PDF
    results = []
    for excel in excel_dfs:
        for pdf in pdf_texts:
            found = []
            for col in excel['df'].columns:
                for val in excel['df'][col].astype(str):
                    if pd.notna(val) and val in pdf['text']:
                        found.append(val)
            found = list(set(found))
            results.append({
                'excel': excel['filename'],
                'pdf': pdf['filename'],
                'matched_values': found
            })

    # Affichage des résultats
    for res in results:
        st.write(f"**Fichier Excel :** {res['excel']}  —  **Fichier PDF :** {res['pdf']}")
        if res['matched_values']:
            st.write("Valeurs trouvées :", res['matched_values'])
        else:
            st.write(":x: Aucune valeur trouvée.")
        st.write("---")

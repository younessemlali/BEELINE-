"""
Application Streamlit simple pour rapprochement PDF/Excel
"""

import streamlit as st
import pandas as pd
import pdfplumber
import io

def extract_pdf_text(pdf_file):
    """Extrait le texte d'un fichier PDF"""
    text = ""
    try:
        with pdfplumber.open(pdf_file) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    text += page_text + "\n"
    except Exception as e:
        st.error(f"Erreur lors de l'extraction du PDF: {e}")
    return text

def load_excel_file(excel_file):
    """Charge un fichier Excel et retourne un DataFrame"""
    try:
        # Essayer de lire comme Excel
        if excel_file.name.endswith(('.xlsx', '.xls')):
            df = pd.read_excel(excel_file)
        else:
            # Essayer de lire comme CSV
            df = pd.read_csv(excel_file)
        
        return df
    except Exception as e:
        st.error(f"Erreur lors de la lecture du fichier Excel: {e}")
        return None

def simple_reconciliation(pdf_text, excel_df):
    """Effectue un rapprochement simple entre PDF et Excel"""
    matches = []
    
    if excel_df is None or pdf_text == "":
        return matches
    
    # Pour chaque cellule dans le DataFrame Excel
    for idx, row in excel_df.iterrows():
        for col_name, cell_value in row.items():
            if pd.notna(cell_value):
                cell_str = str(cell_value).strip()
                # Vérifier si la valeur de la cellule apparaît dans le texte PDF
                if cell_str and len(cell_str) > 2 and cell_str in pdf_text:
                    matches.append({
                        'Ligne Excel': idx + 1,
                        'Colonne': col_name,
                        'Valeur trouvée': cell_str,
                        'Type': type(cell_value).__name__
                    })
    
    return matches

def main():
    st.title("📄📊 Rapprochement PDF/Excel")
    st.markdown("Application simple pour rapprocher les données entre un fichier PDF et un fichier Excel")
    
    # Section upload des fichiers
    st.header("📤 Upload des fichiers")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📄 Fichier PDF")
        pdf_file = st.file_uploader(
            "Choisissez un fichier PDF", 
            type=['pdf'],
            help="Sélectionnez un fichier PDF pour extraire le texte"
        )
    
    with col2:
        st.subheader("📊 Fichier Excel")
        excel_file = st.file_uploader(
            "Choisissez un fichier Excel", 
            type=['xlsx', 'xls', 'csv'],
            help="Sélectionnez un fichier Excel ou CSV"
        )
    
    # Traitement si les deux fichiers sont uploadés
    if pdf_file is not None and excel_file is not None:
        
        st.header("🔍 Extraction et analyse")
        
        # Extraction du texte PDF
        with st.spinner("Extraction du texte PDF..."):
            pdf_text = extract_pdf_text(pdf_file)
        
        # Chargement du fichier Excel
        with st.spinner("Chargement du fichier Excel..."):
            excel_df = load_excel_file(excel_file)
        
        # Affichage des résultats d'extraction
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("📄 Contenu du PDF")
            if pdf_text:
                st.text_area(
                    "Texte extrait:", 
                    pdf_text[:1000] + ("..." if len(pdf_text) > 1000 else ""),
                    height=300
                )
                st.info(f"Texte extrait: {len(pdf_text)} caractères")
            else:
                st.warning("Aucun texte extrait du PDF")
        
        with col2:
            st.subheader("📊 Aperçu Excel")
            if excel_df is not None:
                st.dataframe(excel_df.head(10))
                st.info(f"Fichier Excel: {len(excel_df)} lignes, {len(excel_df.columns)} colonnes")
            else:
                st.warning("Impossible de charger le fichier Excel")
        
        # Rapprochement
        if pdf_text and excel_df is not None:
            st.header("⚖️ Résultat du rapprochement")
            
            with st.spinner("Analyse en cours..."):
                matches = simple_reconciliation(pdf_text, excel_df)
            
            if matches:
                st.success(f"🎯 {len(matches)} correspondances trouvées!")
                
                # Afficher les résultats dans un tableau
                matches_df = pd.DataFrame(matches)
                st.dataframe(matches_df, use_container_width=True)
                
                # Statistiques
                st.subheader("📊 Statistiques")
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    st.metric("Correspondances", len(matches))
                
                with col2:
                    unique_values = len(set(match['Valeur trouvée'] for match in matches))
                    st.metric("Valeurs uniques", unique_values)
                
                with col3:
                    coverage = len(matches) / len(excel_df) * 100 if len(excel_df) > 0 else 0
                    st.metric("Couverture", f"{coverage:.1f}%")
                
                # Permettre le téléchargement des résultats
                csv_results = matches_df.to_csv(index=False)
                st.download_button(
                    label="📥 Télécharger les résultats (CSV)",
                    data=csv_results,
                    file_name="rapprochement_resultats.csv",
                    mime="text/csv"
                )
                
            else:
                st.warning("❌ Aucune correspondance trouvée entre le PDF et l'Excel")
                st.info("💡 Vérifiez que les données sont dans le bon format et que les valeurs correspondent")
    
    # Instructions d'utilisation
    with st.sidebar:
        st.header("ℹ️ Instructions")
        st.markdown("""
        **Comment utiliser cette application:**
        
        1. **Uploadez un fichier PDF** contenant du texte
        2. **Uploadez un fichier Excel** (.xlsx, .xls ou .csv)
        3. L'application va:
           - Extraire le texte du PDF
           - Analyser le contenu Excel
           - Chercher les valeurs Excel dans le texte PDF
           - Afficher les correspondances trouvées
        
        **Formats supportés:**
        - PDF: Fichiers avec texte extractible
        - Excel: .xlsx, .xls, .csv
        
        **Note:** Cette application effectue un rapprochement simple basé sur la recherche de chaînes de caractères.
        """)

if __name__ == "__main__":
    main()
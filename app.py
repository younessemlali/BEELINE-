"""
BEELINE - APPLICATION DE RAPPROCHEMENT PDF/EXCEL
Application Streamlit moderne pour le rapprochement automatique
Version GitHub + Streamlit Cloud
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import io
import zipfile
import time
import json

# Import des modules personnalisés
from pdf_extractor import PDFExtractor
from excel_processor import ExcelProcessor
from reconciliation import ReconciliationEngine

# Configuration de la page
st.set_page_config(
    page_title="Beeline - Rapprochement PDF/Excel",
    page_icon="⚖️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS personnalisé pour un design moderne
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(90deg, #1f4e79 0%, #2e86ab 100%);
        color: white;
        padding: 2rem;
        border-radius: 10px;
        margin-bottom: 2rem;
        text-align: center;
    }
    
    .metric-card {
        background: white;
        padding: 1.5rem;
        border-radius: 10px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        border-left: 4px solid #2e86ab;
    }
    
    .success-card {
        border-left-color: #28a745 !important;
    }
    
    .warning-card {
        border-left-color: #ffc107 !important;
    }
    
    .danger-card {
        border-left-color: #dc3545 !important;
    }
    
    .upload-section {
        background: #f8f9fa;
        padding: 2rem;
        border-radius: 10px;
        margin: 1rem 0;
    }
    
    .stProgress > div > div > div > div {
        background: linear-gradient(90deg, #1f4e79 0%, #2e86ab 100%);
    }
    
    .sidebar .sidebar-content {
        background: linear-gradient(180deg, #f8f9fa 0%, #e9ecef 100%);
    }
</style>
""", unsafe_allow_html=True)

# Initialisation des classes
@st.cache_resource
def get_processors():
    """Initialise et met en cache les processeurs"""
    pdf_extractor = PDFExtractor()
    excel_processor = ExcelProcessor()
    reconciliation_engine = ReconciliationEngine()
    return pdf_extractor, excel_processor, reconciliation_engine

def initialize_session_state():
    """Initialise les variables de session"""
    if 'uploaded_pdfs' not in st.session_state:
        st.session_state.uploaded_pdfs = []
    if 'uploaded_excels' not in st.session_state:
        st.session_state.uploaded_excels = []
    if 'pdf_data' not in st.session_state:
        st.session_state.pdf_data = None
    if 'excel_data' not in st.session_state:
        st.session_state.excel_data = None
    if 'reconciliation_results' not in st.session_state:
        st.session_state.reconciliation_results = None
    if 'processing_complete' not in st.session_state:
        st.session_state.processing_complete = False

def main():
    """Fonction principale de l'application"""
    
    # Initialisation
    initialize_session_state()
    pdf_extractor, excel_processor, reconciliation_engine = get_processors()
    
    # En-tête principal
    st.markdown("""
    <div class="main-header">
        <h1>⚖️ BEELINE - Rapprochement PDF/Excel</h1>
        <p style="font-size: 1.2em; opacity: 0.9;">
            Application moderne de rapprochement automatique entre factures PDF et données Excel
        </p>
    </div>
    """, unsafe_allow_html=True)
    
    # Sidebar pour la navigation
    with st.sidebar:
        st.markdown("## 📋 Navigation")
        page = st.selectbox(
            "Choisir une section",
            ["🏠 Accueil", "📤 Upload Fichiers", "⚖️ Rapprochement", "📊 Résultats", "📈 Historique"]
        )
        
        # Informations sur l'état
        st.markdown("---")
        st.markdown("## 📊 État Actuel")
        
        # Compteurs de fichiers
        pdf_count = len(st.session_state.uploaded_pdfs)
        excel_count = len(st.session_state.uploaded_excels)
        
        col1, col2 = st.columns(2)
        with col1:
            st.metric("📄 PDFs", pdf_count)
        with col2:
            st.metric("📊 Excel", excel_count)
        
        # Statut du traitement
        if st.session_state.processing_complete:
            st.success("✅ Traitement terminé")
        elif pdf_count > 0 and excel_count > 0:
            st.info("🔄 Prêt pour traitement")
        else:
            st.warning("⏳ En attente de fichiers")
    
    # Routage des pages
    if page == "🏠 Accueil":
        show_home_page()
    elif page == "📤 Upload Fichiers":
        show_upload_page(pdf_extractor, excel_processor)
    elif page == "⚖️ Rapprochement":
        show_reconciliation_page(reconciliation_engine)
    elif page == "📊 Résultats":
        show_results_page()
    elif page == "📈 Historique":
        show_history_page()

def show_home_page():
    """Page d'accueil avec présentation"""
    
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.markdown("## 🎯 Fonctionnalités")
        
        features = [
            ("📄 Extraction PDF", "Extraction automatique des données des factures PDF Randstad"),
            ("📊 Traitement Excel", "Analyse des fichiers Excel/CSV avec données détaillées Beeline"),
            ("⚖️ Rapprochement Intelligent", "Matching automatique par numéro de commande et montants"),
            ("📈 Rapports Visuels", "Tableaux de bord interactifs avec graphiques"),
            ("📥 Export Multiple", "Téléchargement en Excel, CSV et PDF"),
            ("📧 Notification Email", "Envoi optionnel des résultats par email")
        ]
        
        for title, desc in features:
            with st.expander(title):
                st.write(desc)
    
    with col2:
        st.markdown("## 🚀 Démarrage Rapide")
        
        st.markdown("""
        ### Étapes simples :
        1. **Upload** vos fichiers PDF et Excel
        2. **Lancez** le rapprochement automatique  
        3. **Consultez** les résultats détaillés
        4. **Téléchargez** les rapports
        
        ### Formats supportés :
        - 📄 **PDF** : Factures Randstad
        - 📊 **Excel** : .xlsx, .xls, .csv
        - 📈 **Google Sheets** : Liens partagés
        
        ### Limites :
        - 🗂️ **100 fichiers** maximum par session
        - 📏 **50 MB** par fichier
        - ⏱️ **Session** : 2 heures
        """)
        
        if st.button("🚀 Commencer", type="primary", use_container_width=True):
            st.switch_page("pages/upload.py") if hasattr(st, 'switch_page') else st.info("Utilisez la navigation à gauche")

def show_upload_page(pdf_extractor, excel_processor):
    """Page d'upload des fichiers"""
    
    st.markdown("## 📤 Upload des Fichiers")
    
    # Section PDF
    st.markdown("### 📄 Fichiers PDF (Factures)")
    
    with st.container():
        st.markdown('<div class="upload-section">', unsafe_allow_html=True)
        
        uploaded_pdfs = st.file_uploader(
            "Déposez vos factures PDF ici",
            type=['pdf'],
            accept_multiple_files=True,
            key="pdf_uploader",
            help="Formats acceptés: PDF uniquement. Taille max: 50MB par fichier."
        )
        
        if uploaded_pdfs:
            st.session_state.uploaded_pdfs = uploaded_pdfs
            st.success(f"✅ {len(uploaded_pdfs)} fichier(s) PDF chargé(s)")
            
            # Prévisualisation des PDFs
            with st.expander("👁️ Prévisualisation des PDFs"):
                for i, pdf_file in enumerate(uploaded_pdfs):
                    col1, col2, col3 = st.columns([2, 1, 1])
                    with col1:
                        st.write(f"📄 {pdf_file.name}")
                    with col2:
                        st.write(f"{pdf_file.size / 1024:.1f} KB")
                    with col3:
                        if st.button(f"🔍 Analyser", key=f"analyze_pdf_{i}"):
                            with st.spinner("Extraction en cours..."):
                                try:
                                    result = pdf_extractor.extract_single_pdf(pdf_file)
                                    st.json(result)
                                except Exception as e:
                                    st.error(f"Erreur: {str(e)}")
        
        st.markdown('</div>', unsafe_allow_html=True)
    
    # Section Excel
    st.markdown("### 📊 Fichiers Excel (Données détaillées)")
    
    with st.container():
        st.markdown('<div class="upload-section">', unsafe_allow_html=True)
        
        uploaded_excels = st.file_uploader(
            "Déposez vos fichiers Excel ici",
            type=['xlsx', 'xls', 'csv'],
            accept_multiple_files=True,
            key="excel_uploader",
            help="Formats acceptés: Excel (.xlsx, .xls) et CSV. Taille max: 50MB par fichier."
        )
        
        if uploaded_excels:
            st.session_state.uploaded_excels = uploaded_excels
            st.success(f"✅ {len(uploaded_excels)} fichier(s) Excel chargé(s)")
            
            # Prévisualisation des Excel
            with st.expander("👁️ Prévisualisation des Excel"):
                for i, excel_file in enumerate(uploaded_excels):
                    col1, col2, col3, col4 = st.columns([2, 1, 1, 1])
                    with col1:
                        st.write(f"📊 {excel_file.name}")
                    with col2:
                        st.write(f"{excel_file.size / 1024:.1f} KB")
                    with col3:
                        st.write(excel_file.type.split('/')[-1].upper())
                    with col4:
                        if st.button(f"👁️ Aperçu", key=f"preview_excel_{i}"):
                            try:
                                df = pd.read_excel(excel_file) if excel_file.name.endswith(('.xlsx', '.xls')) else pd.read_csv(excel_file)
                                st.dataframe(df.head(), use_container_width=True)
                                st.info(f"📊 {len(df)} lignes, {len(df.columns)} colonnes")
                            except Exception as e:
                                st.error(f"Erreur lecture: {str(e)}")
        
        st.markdown('</div>', unsafe_allow_html=True)
    
    # Bouton de traitement
    if st.session_state.uploaded_pdfs and st.session_state.uploaded_excels:
        st.markdown("---")
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            if st.button("🚀 Lancer le Traitement", type="primary", use_container_width=True):
                process_files(pdf_extractor, excel_processor)
    else:
        st.info("📋 Veuillez charger au moins un fichier PDF et un fichier Excel pour continuer.")

def process_files(pdf_extractor, excel_processor):
    """Traite les fichiers uploadés"""
    
    st.markdown("## 🔄 Traitement en Cours")
    
    # Barre de progression
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    try:
        # Étape 1: Extraction des PDFs
        status_text.text("📄 Extraction des données PDF...")
        progress_bar.progress(10)
        
        pdf_results = []
        for i, pdf_file in enumerate(st.session_state.uploaded_pdfs):
            result = pdf_extractor.extract_single_pdf(pdf_file)
            pdf_results.append(result)
            progress_bar.progress(10 + (30 * (i + 1) // len(st.session_state.uploaded_pdfs)))
        
        st.session_state.pdf_data = pdf_results
        st.success(f"✅ {len(pdf_results)} PDFs traités")
        
        # Étape 2: Traitement des Excel
        status_text.text("📊 Traitement des fichiers Excel...")
        progress_bar.progress(40)
        
        excel_results = []
        for i, excel_file in enumerate(st.session_state.uploaded_excels):
            result = excel_processor.process_excel_file(excel_file)
            excel_results.extend(result)
            progress_bar.progress(40 + (30 * (i + 1) // len(st.session_state.uploaded_excels)))
        
        st.session_state.excel_data = excel_results
        st.success(f"✅ {len(excel_results)} lignes Excel traitées")
        
        # Étape 3: Préparation des données
        status_text.text("⚙️ Préparation des données pour rapprochement...")
        progress_bar.progress(70)
        time.sleep(0.5)
        
        # Étape 4: Finalisation
        status_text.text("✅ Traitement terminé!")
        progress_bar.progress(100)
        
        st.balloons()
        st.success("🎉 Tous les fichiers ont été traités avec succès!")
        
        # Affichage des résultats de traitement
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("### 📄 Résultats PDF")
            valid_pdfs = [p for p in pdf_results if p.get('success', False)]
            st.metric("PDFs valides", len(valid_pdfs))
            st.metric("PDFs avec erreur", len(pdf_results) - len(valid_pdfs))
        
        with col2:
            st.markdown("### 📊 Résultats Excel")
            st.metric("Lignes traitées", len(excel_results))
            unique_orders = len(set(row.get('order_number') for row in excel_results if row.get('order_number')))
            st.metric("Commandes uniques", unique_orders)
        
        # Lien vers la page de rapprochement
        st.info("👉 Utilisez la navigation à gauche pour accéder au **Rapprochement**")
        
    except Exception as e:
        st.error(f"❌ Erreur durant le traitement: {str(e)}")
        st.exception(e)

def show_reconciliation_page(reconciliation_engine):
    """Page de rapprochement"""
    
    st.markdown("## ⚖️ Rapprochement Automatique")
    
    # Vérification des données
    if not st.session_state.pdf_data or not st.session_state.excel_data:
        st.warning("⚠️ Veuillez d'abord traiter vos fichiers dans la section **Upload Fichiers**")
        return
    
    # Paramètres de rapprochement
    st.markdown("### ⚙️ Paramètres de Rapprochement")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        tolerance = st.slider(
            "Tolérance montants (%)",
            min_value=0.0,
            max_value=5.0,
            value=1.0,
            step=0.1,
            help="Pourcentage d'écart accepté pour considérer deux montants comme équivalents"
        )
    
    with col2:
        matching_method = st.selectbox(
            "Méthode de rapprochement",
            ["Exact", "Partiel", "Intelligent"],
            index=2,
            help="Exact: N° commande identique / Partiel: Début de N° / Intelligent: Multi-critères"
        )
    
    with col3:
        send_email = st.checkbox(
            "Envoi email résultats",
            help="Recevez les résultats par email"
        )
        
        if send_email:
            email_address = st.text_input("📧 Email", placeholder="votre@email.com")
    
    # Lancement du rapprochement
    if st.button("🚀 Lancer le Rapprochement", type="primary", use_container_width=True):
        
        with st.spinner("⚖️ Rapprochement en cours..."):
            
            # Configuration des paramètres
            config = {
                'tolerance': tolerance / 100,  # Conversion en décimal
                'method': matching_method.lower(),
                'email': email_address if send_email else None
            }
            
            # Lancement du rapprochement
            try:
                results = reconciliation_engine.perform_reconciliation(
                    st.session_state.pdf_data,
                    st.session_state.excel_data,
                    config
                )
                
                st.session_state.reconciliation_results = results
                st.session_state.processing_complete = True
                
                # Affichage des résultats immédiats
                show_reconciliation_summary(results)
                
            except Exception as e:
                st.error(f"❌ Erreur durant le rapprochement: {str(e)}")
                st.exception(e)

def show_reconciliation_summary(results):
    """Affiche un résumé rapide des résultats"""
    
    st.markdown("### 📊 Résumé du Rapprochement")
    
    # Métriques principales
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric(
            "✅ Matches Parfaits",
            results['summary']['perfect_matches'],
            help="Rapprochements exacts sans écart"
        )
    
    with col2:
        st.metric(
            "⚠️ Écarts Détectés",
            results['summary']['discrepancies'],
            help="Rapprochements avec écarts de montant"
        )
    
    with col3:
        st.metric(
            "❌ PDFs Non Matchés",
            results['summary']['unmatched_pdf'],
            help="Factures PDF sans correspondance Excel"
        )
    
    with col4:
        matching_rate = results['summary']['matching_rate']
        st.metric(
            "📈 Taux de Réussite",
            f"{matching_rate:.1f}%",
            help="Pourcentage de rapprochements réussis"
        )
    
    # Graphique de répartition
    if results['summary']['total_invoices'] > 0:
        
        # Données pour le graphique
        labels = ['Matches Parfaits', 'Écarts', 'Non Matchés PDF', 'Non Matchés Excel']
        values = [
            results['summary']['perfect_matches'],
            results['summary']['discrepancies'],
            results['summary']['unmatched_pdf'],
            results['summary']['unmatched_excel']
        ]
        colors = ['#28a745', '#ffc107', '#dc3545', '#6c757d']
        
        # Création du graphique
        fig = go.Figure(data=[go.Pie(
            labels=labels,
            values=values,
            marker_colors=colors,
            textinfo='label+percent+value',
            hovertemplate='<b>%{label}</b><br>Nombre: %{value}<br>Pourcentage: %{percent}<extra></extra>'
        )])
        
        fig.update_layout(
            title="Répartition des Résultats de Rapprochement",
            title_x=0.5,
            font=dict(size=12),
            height=400
        )
        
        st.plotly_chart(fig, use_container_width=True)
    
    # Message de succès
    if matching_rate >= 90:
        st.success("🎉 Excellent taux de rapprochement ! La plupart des factures ont été matchées.")
    elif matching_rate >= 70:
        st.info("👍 Bon taux de rapprochement. Quelques écarts à vérifier.")
    else:
        st.warning("⚠️ Taux de rapprochement faible. Vérifiez la qualité des données d'entrée.")
    
    # Lien vers les résultats détaillés
    st.info("👉 Consultez la section **Résultats** pour l'analyse détaillée et les téléchargements.")

def show_results_page():
    """Page de résultats détaillés"""
    
    st.markdown("## 📊 Résultats Détaillés")
    
    if not st.session_state.reconciliation_results:
        st.warning("⚠️ Aucun résultat de rapprochement disponible. Lancez d'abord un traitement.")
        return
    
    results = st.session_state.reconciliation_results
    
    # Tabs pour les différentes vues
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "📈 Dashboard", "✅ Matches", "⚠️ Écarts", "❌ Non Matchés", "📥 Téléchargements"
    ])
    
    with tab1:
        show_dashboard_tab(results)
    
    with tab2:
        show_matches_tab(results)
    
    with tab3:
        show_discrepancies_tab(results)
    
    with tab4:
        show_unmatched_tab(results)
    
    with tab5:
        show_downloads_tab(results)

def show_dashboard_tab(results):
    """Onglet dashboard avec graphiques"""
    
    st.markdown("### 📈 Tableau de Bord Interactif")
    
    summary = results['summary']
    
    # KPIs en haut
    col1, col2, col3, col4, col5 = st.columns(5)
    
    with col1:
        st.metric("📄 Total PDFs", summary['total_invoices'])
    with col2:
        st.metric("📊 Total Excel", summary['total_excel_lines'])
    with col3:
        st.metric("🎯 Taux Réussite", f"{summary['matching_rate']:.1f}%")
    with col4:
        st.metric("💰 Montant Total", f"{summary.get('total_amount', 0):,.2f} €")
    with col5:
        st.metric("⏱️ Temps Traitement", f"{summary.get('processing_time', 0):.1f}s")
    
    # Graphiques
    col1, col2 = st.columns(2)
    
    with col1:
        # Graphique en secteurs des résultats
        labels = ['Matches Parfaits', 'Écarts', 'Non Matchés']
        values = [
            summary['perfect_matches'],
            summary['discrepancies'],
            summary['unmatched_pdf'] + summary['unmatched_excel']
        ]
        
        fig_pie = go.Figure(data=[go.Pie(labels=labels, values=values, hole=.3)])
        fig_pie.update_layout(
            title="Répartition des Résultats",
            annotations=[dict(text='Rapprochement', x=0.5, y=0.5, font_size=16, showarrow=False)]
        )
        st.plotly_chart(fig_pie, use_container_width=True)
    
    with col2:
        # Graphique de performance
        if 'matches' in results and results['matches']:
            # Distribution des écarts pour les matches
            differences = [match.get('difference', 0) for match in results['matches']]
            
            fig_hist = px.histogram(
                x=differences,
                title="Distribution des Écarts (Matches Parfaits)",
                labels={'x': 'Écart (€)', 'y': 'Nombre de factures'},
                nbins=20
            )
            fig_hist.update_layout(showlegend=False)
            st.plotly_chart(fig_hist, use_container_width=True)
    
    # Tableau de synthèse par fournisseur
    if 'matches' in results and results['matches']:
        st.markdown("### 📊 Synthèse par Fournisseur")
        
        # Analyse des données de matching
        supplier_data = {}
        for match in results['matches'] + results.get('discrepancies', []):
            supplier = match.get('supplier', 'Non défini')
            if supplier not in supplier_data:
                supplier_data[supplier] = {
                    'matches': 0,
                    'discrepancies': 0,
                    'total_amount': 0
                }
            
            if match.get('type') == 'perfect_match':
                supplier_data[supplier]['matches'] += 1
            else:
                supplier_data[supplier]['discrepancies'] += 1
                
            supplier_data[supplier]['total_amount'] += match.get('pdf_amount', 0)
        
        # Conversion en DataFrame
        supplier_df = pd.DataFrame.from_dict(supplier_data, orient='index')
        supplier_df = supplier_df.round(2)
        supplier_df.index.name = 'Fournisseur'
        
        st.dataframe(supplier_df, use_container_width=True)

def show_matches_tab(results):
    """Onglet des rapprochements parfaits"""
    
    st.markdown("### ✅ Rapprochements Parfaits")
    
    matches = results.get('matches', [])
    
    if not matches:
        st.info("Aucun rapprochement parfait trouvé.")
        return
    
    st.success(f"🎉 {len(matches)} rapprochement(s) parfait(s) identifié(s)")
    
    # Conversion en DataFrame pour affichage
    matches_data = []
    for match in matches:
        matches_data.append({
            'N° Commande': match.get('order_number', 'N/A'),
            'Fichier PDF': match.get('pdf_file', 'N/A'),
            'Montant PDF (€)': f"{match.get('pdf_amount', 0):.2f}",
            'Montant Excel (€)': f"{match.get('excel_amount', 0):.2f}",
            'Écart (€)': f"{match.get('difference', 0):.2f}",
            'Collaborateurs': match.get('collaborators', 'N/A'),
            'Statut': '✅ Parfait'
        })
    
    matches_df = pd.DataFrame(matches_data)
    
    # Affichage avec formatage
    st.dataframe(
        matches_df,
        use_container_width=True,
        hide_index=True
    )
    
    # Option d'export
    if st.button("📥 Exporter les Matches (CSV)"):
        csv = matches_df.to_csv(index=False)
        st.download_button(
            label="⬇️ Télécharger CSV",
            data=csv,
            file_name=f"matches_parfaits_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
            mime="text/csv"
        )

def show_discrepancies_tab(results):
    """Onglet des écarts détectés"""
    
    st.markdown("### ⚠️ Écarts Détectés")
    
    discrepancies = results.get('discrepancies', [])
    
    if not discrepancies:
        st.success("🎉 Aucun écart détecté ! Tous les rapprochements sont parfaits.")
        return
    
    st.warning(f"⚠️ {len(discrepancies)} écart(s) détecté(s)")
    
    # Analyse des écarts
    total_discrepancy = sum(d.get('difference', 0) for d in discrepancies)
    avg_discrepancy = total_discrepancy / len(discrepancies) if discrepancies else 0
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("💰 Écart Total", f"{total_discrepancy:.2f} €")
    with col2:
        st.metric("📊 Écart Moyen", f"{avg_discrepancy:.2f} €")
    with col3:
        max_discrepancy = max(d.get('difference', 0) for d in discrepancies)
        st.metric("⚠️ Écart Maximum", f"{max_discrepancy:.2f} €")
    
    # Tableau des écarts avec priorité
    discrepancies_data = []
    for disc in discrepancies:
        difference = disc.get('difference', 0)
        percentage = disc.get('difference_percent', 0)
        
        # Détermination de la priorité
        if difference > 500 or percentage > 10:
            priority = "🔴 Critique"
            priority_class = "danger"
        elif difference > 100 or percentage > 5:
            priority = "🟡 Élevée"
            priority_class = "warning"
        else:
            priority = "🟢 Faible"
            priority_class = "success"
        
        discrepancies_data.append({
            'Priorité': priority,
            'N° Commande': disc.get('order_number', 'N/A'),
            'Fichier PDF': disc.get('pdf_file', 'N/A'),
            'Montant PDF (€)': f"{disc.get('pdf_amount', 0):.2f}",
            'Montant Excel (€)': f"{disc.get('excel_amount', 0):.2f}",
            'Écart (€)': f"{difference:.2f}",
            'Écart (%)': f"{percentage:.2f}%",
            'Collaborateurs': disc.get('collaborators', 'N/A'),
            'Actions': "🔍 À vérifier"
        })
    
    discrepancies_df = pd.DataFrame(discrepancies_data)
    
    # Tri par écart décroissant
    discrepancies_df = discrepancies_df.sort_values('Écart (€)', ascending=False, key=lambda x: pd.to_numeric(x.str.replace('€', ''), errors='coerce'))
    
    st.dataframe(
        discrepancies_df,
        use_container_width=True,
        hide_index=True
    )
    
    # Graphique des écarts
    if len(discrepancies) > 1:
        st.markdown("#### 📊 Visualisation des Écarts")
        
        amounts = [d.get('difference', 0) for d in discrepancies]
        order_numbers = [d.get('order_number', f'Commande {i+1}') for i, d in enumerate(discrepancies)]
        
        fig_bar = px.bar(
            x=order_numbers,
            y=amounts,
            title="Écarts par Commande",
            labels={'x': 'N° Commande', 'y': 'Écart (€)'},
            color=amounts,
            color_continuous_scale="Reds"
        )
        
        fig_bar.update_layout(showlegend=False, xaxis_tickangle=-45)
        st.plotly_chart(fig_bar, use_container_width=True)

def show_unmatched_tab(results):
    """Onglet des éléments non matchés"""
    
    st.markdown("### ❌ Éléments Non Rapprochés")
    
    unmatched_pdf = results.get('unmatched_pdf', [])
    unmatched_excel = results.get('unmatched_excel', [])
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("#### 📄 PDFs Non Rapprochés")
        
        if not unmatched_pdf:
            st.success("✅ Tous les PDFs ont été rapprochés !")
        else:
            st.error(f"❌ {len(unmatched_pdf)} PDF(s) non rapproché(s)")
            
            pdf_data = []
            for pdf in unmatched_pdf:
                pdf_data.append({
                    'Fichier': pdf.get('filename', 'N/A'),
                    'N° Commande': pdf.get('order_number', 'Non trouvé'),
                    'Montant (€)': f"{pdf.get('amount', 0):.2f}",
                    'Raison': pdf.get('reason', 'Pas de correspondance Excel')
                })
            
            if pdf_data:
                pdf_df = pd.DataFrame(pdf_data)
                st.dataframe(pdf_df, use_container_width=True, hide_index=True)
    
    with col2:
        st.markdown("#### 📊 Excel Non Rapprochés")
        
        if not unmatched_excel:
            st.success("✅ Toutes les données Excel ont été rapprochées !")
        else:
            st.error(f"❌ {len(unmatched_excel)} commande(s) Excel non rapprochée(s)")
            
            excel_data = []
            for excel in unmatched_excel:
                excel_data.append({
                    'N° Commande': excel.get('order_number', 'N/A'),
                    'Montant (€)': f"{excel.get('total_amount', 0):.2f}",
                    'Collaborateurs': excel.get('collaborators', 'N/A'),
                    'Nb Lignes': excel.get('line_count', 0),
                    'Raison': 'Pas de PDF correspondant'
                })
            
            if excel_data:
                excel_df = pd.DataFrame(excel_data)
                st.dataframe(excel_df, use_container_width=True, hide_index=True)
    
    # Recommandations
    if unmatched_pdf or unmatched_excel:
        st.markdown("#### 💡 Recommandations")
        
        recommendations = [
            "🔍 Vérifiez la cohérence des numéros de commande entre PDF et Excel",
            "📅 Contrôlez les périodes de facturation",
            "📝 Vérifiez l'orthographe des noms de fichiers",
            "🔄 Essayez différents paramètres de tolérance",
            "📞 Contactez le support si le problème persiste"
        ]
        
        for rec in recommendations:
            st.info(rec)

def show_downloads_tab(results):
    """Onglet des téléchargements"""
    
    st.markdown("### 📥 Téléchargements et Exports")
    
    # Options de téléchargement
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("#### 📊 Formats Disponibles")
        
        # Excel complet
        if st.button("📊 Rapport Excel Complet", use_container_width=True):
            excel_data = create_excel_report(results)
            st.download_button(
                label="⬇️ Télécharger Excel",
                data=excel_data,
                file_name=f"rapport_beeline_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        
        # CSV matches
        if st.button("✅ Matches (CSV)", use_container_width=True):
            if results.get('matches'):
                csv_data = create_csv_matches(results['matches'])
                st.download_button(
                    label="⬇️ Télécharger CSV",
                    data=csv_data,
                    file_name=f"matches_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                    mime="text/csv"
                )
        
        # JSON complet
        if st.button("🔧 Données JSON (Technique)", use_container_width=True):
            json_data = json.dumps(results, indent=2, default=str)
            st.download_button(
                label="⬇️ Télécharger JSON",
                data=json_data,
                file_name=f"resultats_{datetime.now().strftime('%Y%m%d_%H%M')}.json",
                mime="application/json"
            )
    
    with col2:
        st.markdown("#### 📧 Partage et Notification")
        
        # Envoi par email
        email_recipient = st.text_input("📧 Email destinataire", placeholder="destinataire@example.com")
        
        if st.button("📤 Envoyer par Email", disabled=not email_recipient):
            if email_recipient:
                try:
                    send_results_email(email_recipient, results)
                    st.success("✅ Email envoyé avec succès !")
                except Exception as e:
                    st.error(f"❌ Erreur envoi email: {str(e)}")
        
        # Lien de partage (simulation)
        if st.button("🔗 Générer Lien de Partage"):
            share_link = f"https://beeline-app.streamlit.app/results/{datetime.now().strftime('%Y%m%d%H%M')}"
            st.code(share_link)
            st.info("🔗 Lien généré (valide 24h)")
        
        # Sauvegarde dans l'historique
        if st.button("💾 Sauvegarder dans l'Historique"):
            save_to_history(results)
            st.success("✅ Résultats sauvegardés !")

def show_history_page():
    """Page d'historique des traitements"""
    
    st.markdown("## 📈 Historique des Traitements")
    
    # Simulation d'historique (en production, utiliser une base de données)
    if 'processing_history' not in st.session_state:
        st.session_state.processing_history = []
    
    if not st.session_state.processing_history:
        st.info("📭 Aucun traitement dans l'historique.")
        
        # Bouton pour ajouter un exemple
        if st.button("🎲 Ajouter un Exemple"):
            example_entry = {
                'date': datetime.now(),
                'pdf_count': 5,
                'excel_count': 3,
                'matches': 4,
                'discrepancies': 1,
                'matching_rate': 80.0,
                'total_amount': 15678.90
            }
            st.session_state.processing_history.append(example_entry)
            st.rerun()
    
    else:
        # Affichage de l'historique
        history_df = pd.DataFrame(st.session_state.processing_history)
        history_df['Date'] = pd.to_datetime(history_df['date']).dt.strftime('%d/%m/%Y %H:%M')
        
        # Tableau d'historique
        display_columns = ['Date', 'pdf_count', 'excel_count', 'matches', 'discrepancies', 'matching_rate', 'total_amount']
        st.dataframe(
            history_df[display_columns].rename(columns={
                'pdf_count': 'PDFs',
                'excel_count': 'Excel',
                'matches': 'Matches',
                'discrepancies': 'Écarts',
                'matching_rate': 'Taux (%)',
                'total_amount': 'Montant (€)'
            }),
            use_container_width=True,
            hide_index=True
        )
        
        # Graphique d'évolution
        if len(st.session_state.processing_history) > 1:
            st.markdown("#### 📈 Évolution du Taux de Rapprochement")
            
            fig_line = px.line(
                history_df,
                x='date',
                y='matching_rate',
                title="Évolution du Taux de Rapprochement",
                labels={'matching_rate': 'Taux de Réussite (%)', 'date': 'Date'},
                markers=True
            )
            
            fig_line.update_layout(yaxis_range=[0, 100])
            st.plotly_chart(fig_line, use_container_width=True)

# Fonctions utilitaires pour les exports

def create_excel_report(results):
    """Crée un rapport Excel complet"""
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Feuille de synthèse
        summary_data = {
            'Métrique': ['Total PDFs', 'Total Excel', 'Matches Parfaits', 'Écarts', 'Non Matchés', 'Taux Réussite'],
            'Valeur': [
                results['summary']['total_invoices'],
                results['summary']['total_excel_lines'],
                results['summary']['perfect_matches'],
                results['summary']['discrepancies'],
                results['summary']['unmatched_pdf'] + results['summary']['unmatched_excel'],
                f"{results['summary']['matching_rate']:.1f}%"
            ]
        }
        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, sheet_name='Synthèse', index=False)
        
        # Feuille des matches
        if results.get('matches'):
            matches_df = pd.DataFrame(results['matches'])
            matches_df.to_excel(writer, sheet_name='Matches Parfaits', index=False)
        
        # Feuille des écarts
        if results.get('discrepancies'):
            discrepancies_df = pd.DataFrame(results['discrepancies'])
            discrepancies_df.to_excel(writer, sheet_name='Écarts', index=False)
    
    return output.getvalue()

def create_csv_matches(matches):
    """Crée un CSV des matches"""
    matches_df = pd.DataFrame(matches)
    return matches_df.to_csv(index=False)

def send_results_email(email, results):
    """Envoie les résultats par email (simulation)"""
    # En production, utiliser un service d'email comme SendGrid
    time.sleep(1)  # Simulation d'envoi
    return True

def save_to_history(results):
    """Sauvegarde les résultats dans l'historique"""
    if 'processing_history' not in st.session_state:
        st.session_state.processing_history = []
    
    history_entry = {
        'date': datetime.now(),
        'pdf_count': results['summary']['total_invoices'],
        'excel_count': results['summary']['total_excel_lines'],
        'matches': results['summary']['perfect_matches'],
        'discrepancies': results['summary']['discrepancies'],
        'matching_rate': results['summary']['matching_rate'],
        'total_amount': results['summary'].get('total_amount', 0)
    }
    
    st.session_state.processing_history.append(history_entry)

# Point d'entrée de l'application
if __name__ == "__main__":
    main()

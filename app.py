"""
BEELINE - APPLICATION DE RAPPROCHEMENT PDF/EXCEL
Application Streamlit moderne pour le rapprochement automatique
Version finale optimisée pour les données Select T.T et Excel Beeline
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import io
import json
import time

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
        border: 2px dashed #dee2e6;
    }
    
    .upload-section:hover {
        border-color: #2e86ab;
        background: #f0f8ff;
    }
    
    .stProgress > div > div > div > div {
        background: linear-gradient(90deg, #1f4e79 0%, #2e86ab 100%);
    }
    
    .highlight-info {
        background: #e8f4fd;
        padding: 1rem;
        border-radius: 5px;
        border-left: 3px solid #2e86ab;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

# Initialisation des classes avec cache
@st.cache_resource
def get_processors():
    """Initialise et met en cache les processeurs"""
    pdf_extractor = PDFExtractor()
    excel_processor = ExcelProcessor()
    reconciliation_engine = ReconciliationEngine()
    return pdf_extractor, excel_processor, reconciliation_engine

def initialize_session_state():
    """Initialise les variables de session"""
    defaults = {
        'uploaded_pdfs': [],
        'uploaded_excels': [],
        'pdf_data': None,
        'excel_data': None,
        'reconciliation_results': None,
        'processing_complete': False,
        'processing_history': [],
        'current_step': 1
    }
    
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value

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
            Application optimisée pour le rapprochement entre factures Select T.T et données Excel Beeline
        </p>
        <p style="font-size: 0.9em; opacity: 0.8;">
            Version 2.1.0 - Extraction native PDF • Matching intelligent par références • Export avancé
        </p>
    </div>
    """, unsafe_allow_html=True)
    
    # Sidebar pour la navigation et le suivi
    with st.sidebar:
        st.markdown("## 📋 Navigation")
        page = st.selectbox(
            "Choisir une section",
            ["🏠 Accueil", "📤 Upload Fichiers", "⚖️ Rapprochement", "📊 Résultats", "📈 Historique"],
            index=0
        )
        
        # Indicateur de progression
        st.markdown("---")
        st.markdown("## 🔄 Progression")
        
        # Étapes du processus
        steps = [
            ("1. Upload", len(st.session_state.uploaded_pdfs) > 0 and len(st.session_state.uploaded_excels) > 0),
            ("2. Traitement", st.session_state.pdf_data is not None and st.session_state.excel_data is not None),
            ("3. Rapprochement", st.session_state.reconciliation_results is not None),
            ("4. Résultats", st.session_state.processing_complete)
        ]
        
        for step_name, completed in steps:
            if completed:
                st.success(f"✅ {step_name}")
            else:
                st.info(f"⏳ {step_name}")
        
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
            if st.session_state.reconciliation_results:
                summary = st.session_state.reconciliation_results.get('summary', {})
                st.metric("🎯 Taux réussite", f"{summary.get('matching_rate', 0):.1f}%")
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
    """Page d'accueil avec présentation optimisée"""
    
    # Section de bienvenue
    col1, col2 = st.columns([3, 2])
    
    with col1:
        st.markdown("## 🎯 Application Optimisée pour Beeline")
        
        st.markdown("""
        <div class="highlight-info">
        <strong>🔥 Nouveauté Version 2.1.0</strong><br>
        • Extraction native des factures <strong>Select T.T</strong><br>
        • Rapprochement intelligent par <strong>références et centres de coût</strong><br>
        • Adaptation complète aux colonnes <strong>Excel Beeline</strong><br>
        • Matching multi-niveaux avec scoring de confiance
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown("### ⚡ Fonctionnalités Principales")
        
        features = [
            ("📄 Extraction PDF Native", "Extraction optimisée des factures Select T.T avec reconnaissance des références 4949_65744"),
            ("📊 Traitement Excel Beeline", "Mapping automatique des colonnes N° commande, Centre de coût, Collaborateur, Supplier"),
            ("🧠 Rapprochement Intelligent", "Matching par N° commande + validation croisée par références et centres de coût"),
            ("📈 Analytics Avancés", "Tableaux de bord interactifs avec métriques de performance et qualité"),
            ("📥 Export Professionnel", "Rapports Excel multi-onglets, CSV détaillés, JSON technique"),
            ("🔍 Diagnostic Complet", "Analyse des échecs avec recommandations d'amélioration automatiques")
        ]
        
        for title, desc in features:
            with st.expander(title):
                st.write(desc)
    
    with col2:
        st.markdown("## 🚀 Démarrage Express")
        
        # Guide rapide
        st.markdown("""
        ### ⚡ Guide Rapide (3 min)
        
        **1. 📤 Upload vos fichiers**
        - Factures PDF Select T.T
        - Fichiers Excel Beeline
        
        **2. 🔄 Traitement automatique**
        - Extraction native des données
        - Validation et nettoyage
        
        **3. ⚖️ Rapprochement intelligent**  
        - Par numéro de commande
        - Validation par références
        
        **4. 📊 Résultats détaillés**
        - Dashboard interactif
        - Export professionnel
        """)
        
        # Bouton de démarrage
        if st.button("🚀 Commencer Maintenant", type="primary", use_container_width=True):
            st.session_state.current_step = 2
            st.rerun()
        
        st.markdown("---")
        
        # Informations techniques
        st.markdown("""
        ### 📋 Formats Supportés
        
        **PDFs acceptés :**
        - ✅ Factures Select T.T natives
        - ✅ Format MSP Contingent Worker
        - ✅ Extraction références 4949_65744
        
        **Excel acceptés :**
        - ✅ Beeline Payment Register  
        - ✅ Formats .xlsx, .xls, .csv
        - ✅ Colonnes auto-détectées
        
        **Limites système :**
        - 📁 100 fichiers max/session
        - 📁 50 MB max/fichier
        - ⏱ Session : 2 heures
        """)

def show_upload_page(pdf_extractor, excel_processor):
    """Page d'upload optimisée avec preview"""
    
    st.markdown("## 📤 Upload des Fichiers")
    
    # Instructions spécifiques
    st.markdown("""
    <div class="highlight-info">
    <strong>💡 Instructions spécifiques Beeline :</strong><br>
    • <strong>PDFs</strong> : Factures Select T.T avec références 4949_65744, 4950_65744, etc.<br>
    • <strong>Excel</strong> : Fichiers Beeline Payment Register avec colonnes N° commande, Centre de coût<br>
    • <strong>Correspondance</strong> : Le système cherchera les numéros de commande identiques entre PDF et Excel
    </div>
    """, unsafe_allow_html=True)
    
    # Section PDF avec preview amélioré
    st.markdown("### 📄 Factures PDF Select T.T")
    
    with st.container():
        st.markdown('<div class="upload-section">', unsafe_allow_html=True)
        
        uploaded_pdfs = st.file_uploader(
            "Déposez vos factures PDF Select T.T ici",
            type=['pdf'],
            accept_multiple_files=True,
            key="pdf_uploader",
            help="Factures Select T.T avec format MSP Contingent Worker Invoice. Extraction automatique des références et numéros de commande."
        )
        
        if uploaded_pdfs:
            st.session_state.uploaded_pdfs = uploaded_pdfs
            st.success(f"✅ {len(uploaded_pdfs)} fichier(s) PDF chargé(s)")
            
            # Prévisualisation avancée
            with st.expander(f"👁️ Prévisualisation des {len(uploaded_pdfs)} PDFs"):
                for i, pdf_file in enumerate(uploaded_pdfs[:10]):  # Limiter à 10 pour performance
                    col1, col2, col3, col4 = st.columns([3, 1, 1, 1])
                    with col1:
                        st.write(f"📄 {pdf_file.name}")
                    with col2:
                        st.write(f"{pdf_file.size / 1024:.1f} KB")
                    with col3:
                        st.write(f"Page {i+1}")
                    with col4:
                        if st.button(f"🔍 Test", key=f"analyze_pdf_{i}"):
                            with st.spinner("Extraction test..."):
                                try:
                                    result = pdf_extractor.extract_single_pdf(pdf_file)
                                    
                                    # Affichage optimisé des résultats
                                    if result.get('success'):
                                        st.success("✅ Extraction réussie")
                                        col_a, col_b = st.columns(2)
                                        with col_a:
                                            st.write(f"**N° commande:** {result.get('purchase_order', 'Non trouvé')}")
                                            st.write(f"**Montant:** {result.get('total_net', 0):.2f} €")
                                        with col_b:
                                            st.write(f"**ID facture:** {result.get('invoice_id', 'N/A')}")
                                            st.write(f"**Fournisseur:** {result.get('supplier', 'N/A')}")
                                        
                                        # Références extraites
                                        refs = result.get('invoice_references', [])
                                        if refs:
                                            st.write(f"**Références:** {len(refs)} trouvée(s)")
                                            for ref in refs[:3]:  # Max 3
                                                st.code(ref.get('full_reference', 'N/A'), language=None)
                                    else:
                                        st.error(f"❌ Erreur: {result.get('error', 'Inconnue')}")
                                        
                                except Exception as e:
                                    st.error(f"❌ Erreur: {str(e)}")
                
                if len(uploaded_pdfs) > 10:
                    st.info(f"... et {len(uploaded_pdfs) - 10} autres fichiers")
        
        st.markdown('</div>', unsafe_allow_html=True)
    
    # Section Excel avec détection automatique
    st.markdown("### 📊 Fichiers Excel Beeline")
    
    with st.container():
        st.markdown('<div class="upload-section">', unsafe_allow_html=True)
        
        uploaded_excels = st.file_uploader(
            "Déposez vos fichiers Excel Beeline ici",
            type=['xlsx', 'xls', 'csv'],
            accept_multiple_files=True,
            key="excel_uploader",
            help="Fichiers Beeline Payment Register avec colonnes N° commande, Centre de coût, Collaborateur."
        )
        
        if uploaded_excels:
            st.session_state.uploaded_excels = uploaded_excels
            st.success(f"✅ {len(uploaded_excels)} fichier(s) Excel chargé(s)")
            
            # Prévisualisation avec mapping des colonnes
            with st.expander(f"👁️ Prévisualisation des {len(uploaded_excels)} Excel"):
                for i, excel_file in enumerate(uploaded_excels):
                    col1, col2, col3, col4 = st.columns([3, 1, 1, 1])
                    with col1:
                        st.write(f"📊 {excel_file.name}")
                    with col2:
                        st.write(f"{excel_file.size / 1024:.1f} KB")
                    with col3:
                        st.write(excel_file.type.split('/')[-1].upper())
                    with col4:
                        if st.button(f"👁️ Aperçu", key=f"preview_excel_{i}"):
                            try:
                                # Lecture intelligente
                                if excel_file.name.endswith(('.xlsx', '.xls')):
                                    df = pd.read_excel(excel_file)
                                else:
                                    df = pd.read_csv(excel_file, encoding='utf-8', sep=None, engine='python')
                                
                                st.success(f"✅ Lecture réussie: {len(df)} lignes, {len(df.columns)} colonnes")
                                
                                # Détection des colonnes clés
                                key_columns = {}
                                for col in df.columns:
                                    col_lower = str(col).lower()
                                    if 'commande' in col_lower:
                                        key_columns['N° Commande'] = col
                                    elif 'coût' in col_lower or 'cout' in col_lower:
                                        key_columns['Centre de Coût'] = col
                                    elif 'collaborateur' in col_lower:
                                        key_columns['Collaborateur'] = col
                                    elif 'supplier' in col_lower or 'fournisseur' in col_lower:
                                        key_columns['Fournisseur'] = col
                                
                                if key_columns:
                                    st.write("**🔍 Colonnes clés détectées:**")
                                    for key, col in key_columns.items():
                                        st.write(f"• {key}: `{col}`")
                                
                                # Aperçu des données
                                st.write("**📋 Aperçu des données:**")
                                st.dataframe(df.head(3), use_container_width=True)
                                
                                # Statistiques
                                if 'N° Commande' in key_columns:
                                    unique_orders = df[key_columns['N° Commande']].nunique()
                                    st.write(f"**📊 {unique_orders} commandes uniques détectées**")
                                        
                            except Exception as e:
                                st.error(f"❌ Erreur lecture: {str(e)}")
                                st.info("💡 Vérifiez le format du fichier (encodage, délimiteurs)")
        
        st.markdown('</div>', unsafe_allow_html=True)
    
    # Bouton de traitement avec vérifications
    if st.session_state.uploaded_pdfs and st.session_state.uploaded_excels:
        st.markdown("---")
        
        # Résumé avant traitement
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("📄 PDFs à traiter", len(st.session_state.uploaded_pdfs))
        with col2:
            st.metric("📊 Excel à traiter", len(st.session_state.uploaded_excels))
        with col3:
            total_size = sum(f.size for f in st.session_state.uploaded_pdfs + st.session_state.uploaded_excels)
            st.metric("📁 Taille totale", f"{total_size / (1024*1024):.1f} MB")
        
        # Bouton de lancement
        col_btn1, col_btn2, col_btn3 = st.columns([1, 2, 1])
        with col_btn2:
            if st.button("🚀 Lancer le Traitement", type="primary", use_container_width=True):
                process_files(pdf_extractor, excel_processor)
    else:
        st.markdown("---")
        st.info("📋 **Prochaine étape :** Chargez au moins un fichier PDF et un fichier Excel pour continuer.")
        
        # Aide contextuelle
        if not st.session_state.uploaded_pdfs:
            st.warning("📄 **PDFs manquants** - Chargez vos factures Select T.T")
        if not st.session_state.uploaded_excels:
            st.warning("📊 **Excel manquants** - Chargez vos fichiers Beeline Payment Register")

def process_files(pdf_extractor, excel_processor):
    """Traite les fichiers avec feedback détaillé"""
    
    st.markdown("## 🔄 Traitement en Cours")
    
    # Barre de progression principale
    progress_container = st.container()
    with progress_container:
        progress_bar = st.progress(0)
        status_text = st.empty()
        current_file = st.empty()
    
    # Container pour les résultats en temps réel
    results_container = st.container()
    
    try:
        # Étape 1: Extraction des PDFs
        with results_container:
            st.markdown("### 📄 Extraction des PDFs")
        
        status_text.text("🔍 Extraction des données PDF en cours...")
        
        pdf_results = []
        pdf_success = 0
        pdf_errors = 0
        
        for i, pdf_file in enumerate(st.session_state.uploaded_pdfs):
            current_file.text(f"📄 Traitement: {pdf_file.name}")
            
            result = pdf_extractor.extract_single_pdf(pdf_file)
            pdf_results.append(result)
            
            if result.get('success', False):
                pdf_success += 1
                with results_container:
                    st.success(f"✅ {pdf_file.name}: N° {result.get('purchase_order', 'N/A')}, {result.get('total_net', 0):.2f}€")
            else:
                pdf_errors += 1
                with results_container:
                    st.error(f"❌ {pdf_file.name}: {result.get('error', 'Erreur inconnue')}")
            
            # Mise à jour de la progression
            progress = 10 + (40 * (i + 1) // len(st.session_state.uploaded_pdfs))
            progress_bar.progress(progress)
        
        st.session_state.pdf_data = pdf_results
        current_file.text(f"✅ PDFs traités: {pdf_success} succès, {pdf_errors} erreurs")
        
        # Étape 2: Traitement des Excel
        with results_container:
            st.markdown("### 📊 Traitement des Excel")
        
        status_text.text("📊 Traitement des fichiers Excel en cours...")
        
        excel_results = []
        excel_lines = 0
        excel_valid = 0
        
        for i, excel_file in enumerate(st.session_state.uploaded_excels):
            current_file.text(f"📊 Traitement: {excel_file.name}")
            
            result = excel_processor.process_excel_file(excel_file)
            excel_results.extend(result)
            
            # Statistiques du fichier
            file_lines = len(result)
            file_valid = len([r for r in result if r.get('is_valid', False)])
            excel_lines += file_lines
            excel_valid += file_valid
            
            with results_container:
                st.info(f"📊 {excel_file.name}: {file_lines} lignes, {file_valid} valides")
                
                # Affichage des commandes uniques détectées
                unique_orders = set()
                for r in result:
                    if r.get('order_number'):
                        unique_orders.add(str(r['order_number']))
                
                if unique_orders:
                    st.write(f"   📢 Commandes détectées: {', '.join(list(unique_orders)[:5])}" + 
                            (f" ... (+{len(unique_orders)-5})" if len(unique_orders) > 5 else ""))
            
            # Mise à jour de la progression
            progress = 50 + (30 * (i + 1) // len(st.session_state.uploaded_excels))
            progress_bar.progress(progress)
        
        st.session_state.excel_data = excel_results
        current_file.text(f"✅ Excel traités: {excel_lines} lignes, {excel_valid} valides")
        
        # Étape 3: Préparation finale
        status_text.text("⚙️ Préparation des données pour rapprochement...")
        progress_bar.progress(85)
        time.sleep(0.5)
        
        # Étape 4: Finalisation
        status_text.text("✅ Traitement terminé!")
        progress_bar.progress(100)
        current_file.text("🎉 Tous les fichiers ont été traités avec succès!")
        
        # Animation de succès
        st.balloons()
        
        # Résumé final détaillé
        st.markdown("---")
        st.markdown("## 📊 Résumé du Traitement")
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.markdown("### 📄 PDFs")
            st.metric("Traités", len(pdf_results))
            st.metric("Réussis", pdf_success)
            st.metric("Erreurs", pdf_errors)
            
            if pdf_success > 0:
                # Montant total des PDFs
                total_pdf_amount = sum(r.get('total_net', 0) for r in pdf_results if r.get('success'))
                st.metric("Montant total", f"{total_pdf_amount:.2f} €")
        
        with col2:
            st.markdown("### 📊 Excel")  
            st.metric("Lignes traitées", excel_lines)
            st.metric("Lignes valides", excel_valid)
            st.metric("Taux validité", f"{(excel_valid/excel_lines)*100:.1f}%" if excel_lines > 0 else "0%")
            
            # Commandes uniques
            all_unique_orders = set()
            for r in excel_results:
                if r.get('order_number'):
                    all_unique_orders.add(str(r['order_number']))
            st.metric("Commandes uniques", len(all_unique_orders))
        
        with col3:
            st.markdown("### ⚖️ Prêt pour Rapprochement")
            st.metric("PDFs processables", pdf_success)
            st.metric("Commandes Excel", len(all_unique_orders))
            
            # Estimation du potentiel de matching
            pdf_orders = set(r.get('purchase_order') for r in pdf_results if r.get('success') and r.get('purchase_order'))
            common_orders = pdf_orders.intersection(all_unique_orders)
            st.metric("Correspondances potentielles", len(common_orders))
        
        # Étapes suivantes
        st.markdown("---")
        st.info("👉 **Étape suivante :** Utilisez la section **⚖️ Rapprochement** pour lancer l'analyse intelligente des correspondances.")
        
        # Bouton direct vers rapprochement
        if st.button("⚖️ Lancer le Rapprochement Maintenant", type="primary"):
            st.session_state.current_step = 3
            st.switch_page("Rapprochement") if hasattr(st, 'switch_page') else st.rerun()
        
    except Exception as e:
        st.error(f"❌ Erreur critique durant le traitement: {str(e)}")
        st.exception(e)
        
        # Diagnostic d'erreur
        st.markdown("### 🔧 Diagnostic")
        st.write("**Vérifiez :**")
        st.write("• Format des fichiers (PDF natifs, Excel lisibles)")
        st.write("• Taille des fichiers (< 50MB chacun)")
        st.write("• Encodage des CSV (UTF-8 recommandé)")
        st.write("• Permissions de lecture des fichiers")

def show_reconciliation_page(reconciliation_engine):
    """Page de rapprochement avec options avancées"""
    
    st.markdown("## ⚖️ Rapprochement Automatique")
    
    # Vérification des données
    if not st.session_state.pdf_data or not st.session_state.excel_data:
        st.warning("⚠️ **Données manquantes** - Veuillez d'abord traiter vos fichiers dans la section **📤 Upload Fichiers**")
        
        # Aide contextuelle
        st.markdown("""
        ### 🔄 Processus recommandé:
        1. **📤 Upload Fichiers** - Chargez et traitez vos PDFs et Excel
        2. **⚖️ Rapprochement** - Configurez et lancez l'analyse (vous êtes ici)
        3. **📊 Résultats** - Consultez les matches et écarts
        """)
        return
    
    # Statistiques des données chargées
    pdf_valid = len([p for p in st.session_state.pdf_data if p.get('success', False)])
    excel_valid = len([e for e in st.session_state.excel_data if e.get('is_valid', False)])
    
    st.markdown(f"""
    <div class="highlight-info">
    <strong>📊 Données chargées :</strong> {pdf_valid} PDFs valides • {excel_valid} lignes Excel valides<br>
    <strong>🎯 Stratégie :</strong> Rapprochement par N° commande + validation par références et centres de coût
    </div>
    """, unsafe_allow_html=True)
    
    # Configuration avancée du rapprochement
    st.markdown("### ⚙️ Configuration du Rapprochement")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown("#### 💰 Tolérance Montants")
        tolerance = st.slider(
            "Écart accepté (%)",
            min_value=0.0,
            max_value=10.0,
            value=1.0,
            step=0.1,
            help="Pourcentage d'écart accepté entre montant PDF et Excel pour considérer un rapprochement comme parfait"
        )
        
        st.markdown("#### 🎯 Seuil de Confiance")
        min_confidence = st.slider(
            "Confiance minimum (%)",
            min_value=50,
            max_value=95,
            value=60,
            step=5,
            help="Seuil minimum de confiance pour valider un rapprochement"
        ) / 100
    
    with col2:
        st.markdown("#### 🔍 Méthode de Rapprochement")
        matching_method = st.selectbox(
            "Stratégie",
            ["Intelligent", "Exact uniquement", "Partiel + Exact"],
            index=0,
            help="""
            • Intelligent: Multi-niveaux avec références
            • Exact uniquement: N° commande identique seulement  
            • Partiel + Exact: N° commande exact puis partiel
            """
        )
        
        st.markdown("#### 🔄 Options Avancées")
        enable_reference_matching = st.checkbox(
            "Matching par références",
            value=True,
            help="Active le rapprochement par références PDF (4949_65744) vs centres de coût Excel"
        )
        
        extended_tolerance = st.checkbox(
            "Tolérance élargie pour fuzzy",
            value=False,
            help="Utilise une tolérance 5x plus large pour le rapprochement par montant"
        )
    
    with col3:
        st.markdown("#### 📊 Analyse et Export")
        
        generate_detailed_report = st.checkbox(
            "Rapport détaillé",
            value=True,
            help="Génère un rapport Excel complet avec tous les onglets"
        )
        
        include_diagnostics = st.checkbox(
            "Diagnostics avancés", 
            value=True,
            help="Inclut l'analyse des échecs et recommandations"
        )
        
        st.markdown("#### 📧 Notification")
        st.info("💡 Notification email\n(Fonctionnalité désactivée)")
    
    # Aperçu de la configuration
    with st.expander("🔍 Aperçu de la Configuration"):
        config_preview = {
            "Tolérance montants": f"{tolerance}%",
            "Seuil confiance": f"{min_confidence*100}%", 
            "Méthode": matching_method,
            "Matching références": "✅" if enable_reference_matching else "❌",
            "Tolérance élargie": "✅" if extended_tolerance else "❌"
        }
        
        for key, value in config_preview.items():
            st.write(f"• **{key}:** {value}")
    
    # Lancement du rapprochement
    st.markdown("---")
    
    col_btn1, col_btn2, col_btn3 = st.columns([1, 2, 1])
    with col_btn2:
        if st.button("🚀 Lancer le Rapprochement Intelligent", type="primary", use_container_width=True):
            
            # Configuration des paramètres
            config = {
                'tolerance': tolerance / 100,
                'method': matching_method.lower().replace(" uniquement", "").replace(" + exact", "_partial"),
                'min_confidence': min_confidence,
                'enable_reference_matching': enable_reference_matching,
                'extended_tolerance_factor': 5 if extended_tolerance else 2,
                'generate_detailed_report': generate_detailed_report,
                'include_diagnostics': include_diagnostics
            }
            
            # Lancement avec feedback temps réel
            perform_reconciliation(reconciliation_engine, config)

def perform_reconciliation(reconciliation_engine, config):
    """Effectue le rapprochement avec suivi en temps réel"""
    
    st.markdown("### ⚖️ Rapprochement en Cours")
    
    # Container pour le suivi temps réel
    progress_container = st.container()
    with progress_container:
        progress_bar = st.progress(0)
        status_text = st.empty()
        phase_text = st.empty()
    
    results_container = st.container()
    
    try:
        with st.spinner("🧠 Initialisation du moteur de rapprochement..."):
            
            # Phase 1: Préparation
            status_text.text("📋 Préparation des données...")
            phase_text.text("Nettoyage et validation des données PDF et Excel")
            progress_bar.progress(10)
            time.sleep(0.5)
            
            # Phase 2: Lancement
            status_text.text("⚖️ Rapprochement intelligent en cours...")
            phase_text.text("Analyse multi-niveaux: N° commande → Références → Montants")
            progress_bar.progress(30)
            
            # Exécution du rapprochement
            results = reconciliation_engine.perform_reconciliation(
                st.session_state.pdf_data,
                st.session_state.excel_data,
                config
            )
            
            progress_bar.progress(80)
            
            # Phase 3: Post-traitement
            status_text.text("📊 Génération des rapports...")
            phase_text.text("Calcul des statistiques et préparation des exports")
            progress_bar.progress(90)
            time.sleep(0.3)
            
            # Phase 4: Finalisation
            status_text.text("✅ Rapprochement terminé!")
            phase_text.text("🎉 Analyse complète - Résultats disponibles")
            progress_bar.progress(100)
            
            # Sauvegarde des résultats
            st.session_state.reconciliation_results = results
            st.session_state.processing_complete = True
            
            # Animation de succès
            st.balloons()
            
            # Affichage immédiat des résultats clés
            with results_container:
                show_reconciliation_summary(results)
            
            # Navigation automatique
            st.info("👉 Consultez la section **📊 Résultats** pour l'analyse détaillée complète.")
            
    except Exception as e:
        st.error(f"❌ Erreur durant le rapprochement: {str(e)}")
        st.exception(e)
        
        # Diagnostic d'erreur
        st.markdown("### 🔧 Diagnostic d'Erreur")
        st.write("**Causes possibles :**")
        st.write("• Données PDF ou Excel corrompues")
        st.write("• Paramètres de configuration incompatibles")
        st.write("• Mémoire insuffisante pour le traitement")
        st.write("• Format de données non supporté")

def show_reconciliation_summary(results):
    """Affiche un résumé rapide des résultats de rapprochement"""
    
    st.markdown("### 📊 Résumé du Rapprochement")
    
    summary = results.get('summary', {})
    
    # Métriques principales en colonnes
    col1, col2, col3, col4, col5 = st.columns(5)
    
    with col1:
        st.metric(
            "✅ Matches",
            summary.get('perfect_matches', 0),
            help="Rapprochements parfaits sans écart"
        )
    
    with col2:
        st.metric(
            "⚠️ Écarts",
            summary.get('discrepancies', 0),
            help="Rapprochements avec différences de montant"
        )
    
    with col3:
        st.metric(
            "❌ PDFs seuls",
            summary.get('unmatched_pdf', 0),
            help="Factures PDF sans correspondance Excel"
        )
    
    with col4:
        st.metric(
            "📊 Excel seuls",
            summary.get('unmatched_excel', 0),
            help="Commandes Excel sans correspondance PDF"
        )
    
    with col5:
        matching_rate = summary.get('matching_rate', 0)
        st.metric(
            "🎯 Taux Réussite",
            f"{matching_rate:.1f}%",
            help="Pourcentage global de rapprochements réussis"
        )
    
    # Graphique de synthèse
    if summary.get('total_invoices', 0) > 0:
        
        # Données pour le graphique en secteurs
        labels = ['Matches Parfaits', 'Écarts Détectés', 'PDFs Non Matchés', 'Excel Non Matchés']
        values = [
            summary.get('perfect_matches', 0),
            summary.get('discrepancies', 0),
            summary.get('unmatched_pdf', 0),
            summary.get('unmatched_excel', 0)
        ]
        colors = ['#28a745', '#ffc107', '#dc3545', '#6c757d']
        
        # Graphique avec Plotly
        fig = go.Figure(data=[go.Pie(
            labels=labels,
            values=values,
            marker_colors=colors,
            textinfo='label+percent+value',
            hovertemplate='<b>%{label}</b><br>Nombre: %{value}<br>Pourcentage: %{percent}<extra></extra>',
            hole=.4
        )])
        
        fig.update_layout(
            title="Répartition des Résultats de Rapprochement",
            title_x=0.5,
            font=dict(size=11),
            height=350,
            annotations=[dict(text='Résultats', x=0.5, y=0.5, font_size=14, showarrow=False)]
        )
        
        st.plotly_chart(fig, use_container_width=True)
    
    # Message d'évaluation
    if matching_rate >= 90:
        st.success("🎉 **Excellent résultat !** La majorité des factures ont été rapprochées avec succès.")
    elif matching_rate >= 75:
        st.info("👍 **Bon résultat.** Quelques écarts à analyser mais globalement satisfaisant.")
    elif matching_rate >= 50:
        st.warning("⚠️ **Résultat moyen.** Vérifiez la qualité des données d'entrée.")
    else:
        st.error("⚠️ **Résultat insuffisant.** Contrôlez les paramètres et la cohérence des données.")
    
    # Performance du traitement
    processing_time = results.get('metadata', {}).get('processing_time', 0)
    if processing_time > 0:
        st.caption(f"⏱️ Traitement effectué en {processing_time:.2f} secondes")

def show_results_page():
    """Page de résultats détaillés avec tableaux de bord avancés"""
    
    st.markdown("## 📊 Résultats Détaillés")
    
    if not st.session_state.reconciliation_results:
        st.warning("⚠️ Aucun résultat de rapprochement disponible. Lancez d'abord un traitement complet.")
        
        # Guide de démarrage
        st.markdown("""
        ### 🚀 Pour obtenir des résultats :
        1. **📤 Upload Fichiers** - Chargez vos PDFs et Excel
        2. **⚖️ Rapprochement** - Configurez et lancez l'analyse
        3. **📊 Résultats** - Consultez cette page (vous y êtes !)
        """)
        return
    
    results = st.session_state.reconciliation_results
    
    # Tabs pour les différentes analyses
    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
        "📈 Dashboard", "✅ Matches", "⚠️ Écarts", "❌ Non Matchés", "📥 Téléchargements", "🔧 Diagnostics"
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
    
    with tab6:
        show_diagnostics_tab(results)

def show_dashboard_tab(results):
    """Tableau de bord interactif avec métriques et graphiques"""
    
    st.markdown("### 📈 Tableau de Bord Analytique")
    
    summary = results.get('summary', {})
    metadata = results.get('metadata', {})
    
    # Section des KPIs principaux
    st.markdown("#### 🎯 Indicateurs Clés de Performance")
    
    kpi_col1, kpi_col2, kpi_col3, kpi_col4, kpi_col5, kpi_col6 = st.columns(6)
    
    with kpi_col1:
        st.metric("📄 PDFs Traités", summary.get('total_invoices', 0))
    with kpi_col2:
        st.metric("📊 Lignes Excel", summary.get('total_excel_lines', 0))
    with kpi_col3:
        st.metric("🎯 Taux Réussite", f"{summary.get('matching_rate', 0):.1f}%")
    with kpi_col4:
        st.metric("💰 Montant Total", f"{summary.get('total_amount', 0):,.2f} €")
    with kpi_col5:
        st.metric("⚖️ Couverture", f"{summary.get('coverage_rate', 0):.1f}%")
    with kpi_col6:
        processing_time = metadata.get('processing_time', 0)
        st.metric("⏱️ Temps", f"{processing_time:.1f}s")
    
    # Section graphiques analytiques
    graph_col1, graph_col2 = st.columns(2)
    
    with graph_col1:
        # Graphique de répartition des résultats
        st.markdown("#### 📊 Répartition des Résultats")
        
        labels = ['Matches Parfaits', 'Écarts', 'Non Matchés']
        values = [
            summary.get('perfect_matches', 0),
            summary.get('discrepancies', 0),
            summary.get('unmatched_pdf', 0) + summary.get('unmatched_excel', 0)
        ]
        
        fig_pie = go.Figure(data=[go.Pie(
            labels=labels, 
            values=values, 
            hole=.3,
            marker_colors=['#28a745', '#ffc107', '#dc3545']
        )])
        fig_pie.update_layout(
            title="Distribution des Rapprochements",
            height=300,
            annotations=[dict(text='Total', x=0.5, y=0.5, font_size=14, showarrow=False)]
        )
        st.plotly_chart(fig_pie, use_container_width=True)
    
    with graph_col2:
        # Analyse des écarts de montants
        st.markdown("#### 💰 Analyse des Montants")
        
        if results.get('matches') or results.get('discrepancies'):
            all_items = results.get('matches', []) + results.get('discrepancies', [])
            amounts_pdf = [item.get('pdf_amount', 0) for item in all_items]
            amounts_excel = [item.get('excel_amount', 0) for item in all_items]
            
            fig_scatter = go.Figure()
            fig_scatter.add_trace(go.Scatter(
                x=amounts_pdf,
                y=amounts_excel,
                mode='markers',
                marker=dict(size=8, opacity=0.7),
                text=[f"N° {item.get('order_number', 'N/A')}" for item in all_items],
                hovertemplate='PDF: %{x:.2f}€<br>Excel: %{y:.2f}€<br>%{text}<extra></extra>'
            ))
            
            # Ligne de correspondance parfaite
            max_amount = max(max(amounts_pdf, default=0), max(amounts_excel, default=0))
            fig_scatter.add_trace(go.Scatter(
                x=[0, max_amount],
                y=[0, max_amount],
                mode='lines',
                line=dict(color='red', dash='dash'),
                name='Correspondance parfaite'
            ))
            
            fig_scatter.update_layout(
                title="Corrélation PDF vs Excel",
                xaxis_title="Montant PDF (€)",
                yaxis_title="Montant Excel (€)",
                height=300,
                showlegend=False
            )
            st.plotly_chart(fig_scatter, use_container_width=True)
    
    # Section performance par méthode
    if metadata.get('performance_stats', {}).get('method_performance'):
        st.markdown("#### 🔬 Performance par Méthode de Rapprochement")
        
        method_perf = metadata['performance_stats']['method_performance']
        
        method_names = list(method_perf.keys())
        method_counts = [method_perf[m]['count'] for m in method_names]
        method_confidence = [method_perf[m]['avg_confidence'] for m in method_names]
        
        perf_col1, perf_col2 = st.columns(2)
        
        with perf_col1:
            # Graphique en barres des comptages
            fig_bar = px.bar(
                x=method_names,
                y=method_counts,
                title="Utilisation des Méthodes",
                labels={'x': 'Méthode', 'y': 'Nombre d\'utilisations'}
            )
            fig_bar.update_layout(height=250)
            st.plotly_chart(fig_bar, use_container_width=True)
        
        with perf_col2:
            # Graphique de confiance moyenne
            fig_conf = px.bar(
                x=method_names,
                y=[c*100 for c in method_confidence],
                title="Confiance Moyenne par Méthode",
                labels={'x': 'Méthode', 'y': 'Confiance (%)'},
                color=method_confidence,
                color_continuous_scale='RdYlGn'
            )
            fig_conf.update_layout(height=250, showlegend=False)
            st.plotly_chart(fig_conf, use_container_width=True)
    
    # Section qualité globale
    if summary.get('quality_assessment'):
        st.markdown("#### 🏆 Évaluation de la Qualité")
        
        quality = summary['quality_assessment']
        
        qual_col1, qual_col2, qual_col3 = st.columns(3)
        
        with qual_col1:
            score = quality.get('score', 0)
            st.metric("📈 Score Qualité", f"{score:.1f}/100")
            
            # Barre de progression pour le score
            progress_color = "normal"
            if score >= 90:
                progress_color = "success"
            elif score >= 70:
                progress_color = "warning"
            else:
                progress_color = "error"
            
            st.progress(score/100)
        
        with qual_col2:
            grade = quality.get('grade', 'N/A')
            assessment = quality.get('assessment', 'Non évalué')
            st.metric("🎖️ Note", grade)
            st.caption(f"**{assessment}**")
        
        with qual_col3:
            recommendations = quality.get('recommendations', [])
            st.write("**💡 Recommandations:**")
            for i, rec in enumerate(recommendations[:3], 1):
                st.caption(f"{i}. {rec}")

def show_matches_tab(results):
    """Onglet des rapprochements parfaits avec détails enrichis"""
    
    st.markdown("### ✅ Rapprochements Parfaits")
    
    matches = results.get('matches', [])
    
    if not matches:
        st.info("ℹ️ Aucun rapprochement parfait identifié.")
        st.markdown("""
        **Causes possibles :**
        - Écarts de montants supérieurs à la tolérance
        - Numéros de commande non identiques
        - Données de mauvaise qualité
        """)
        return
    
    st.success(f"🎉 {len(matches)} rapprochement(s) parfait(s) identifié(s)")
    
    # Statistiques des matches
    total_amount = sum(m.get('pdf_amount', 0) for m in matches)
    avg_amount = total_amount / len(matches) if matches else 0
    
    match_col1, match_col2, match_col3 = st.columns(3)
    with match_col1:
        st.metric("💰 Montant Total", f"{total_amount:,.2f} €")
    with match_col2:
        st.metric("📊 Montant Moyen", f"{avg_amount:,.2f} €")
    with match_col3:
        perfect_rate = len(matches) / results.get('summary', {}).get('total_invoices', 1) * 100
        st.metric("🎯 Taux Parfait", f"{perfect_rate:.1f}%")
    
    # Tableau détaillé des matches
    matches_data = []
    for i, match in enumerate(matches, 1):
        matches_data.append({
            '#': i,
            'N° Commande': match.get('order_number', 'N/A'),
            'Fichier PDF': match.get('pdf_file', 'N/A'),
            'Montant PDF (€)': f"{match.get('pdf_amount', 0):.2f}",
            'Montant Excel (€)': f"{match.get('excel_amount', 0):.2f}",
            'Écart (€)': f"{match.get('difference', 0):.2f}",
            'Collaborateurs': match.get('collaborators', 'N/A')[:50] + ('...' if len(match.get('collaborators', '')) > 50 else ''),
            'Méthode': match.get('method', 'N/A'),
            'Confiance': f"{match.get('confidence', 0)*100:.1f}%"
        })
    
    matches_df = pd.DataFrame(matches_data)
    
    # Affichage avec options de tri
    sort_column = st.selectbox(
        "Trier par:",
        ["#", "N° Commande", "Montant PDF (€)", "Écart (€)", "Confiance"],
        index=0
    )
    
    if sort_column != "#":
        if sort_column in ["Montant PDF (€)", "Écart (€)"]:
            matches_df[sort_column] = pd.to_numeric(matches_df[sort_column], errors='coerce')
        elif sort_column == "Confiance":
            matches_df[sort_column] = pd.to_numeric(matches_df[sort_column].str.replace('%', ''), errors='coerce')
        
        matches_df = matches_df.sort_values(sort_column, ascending=False)
    
    st.dataframe(
        matches_df,
        use_container_width=True,
        hide_index=True
    )
    
    # Options d'export
    export_col1, export_col2 = st.columns(2)
    
    with export_col1:
        if st.button("📥 Exporter Matches (CSV)", use_container_width=True):
            csv = matches_df.to_csv(index=False)
            st.download_button(
                label="⬇️ Télécharger CSV",
                data=csv,
                file_name=f"matches_parfaits_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                mime="text/csv"
            )
    
    with export_col2:
        if st.button("📊 Rapport Excel Matches", use_container_width=True):
            # Fonction pour créer un rapport Excel des matches
            excel_data = create_matches_excel_report(matches_df)
            st.download_button(
                label="⬇️ Télécharger Excel",
                data=excel_data,
                file_name=f"rapport_matches_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

def show_discrepancies_tab(results):
    """Onglet des écarts avec analyse approfondie"""
    
    st.markdown("### ⚠️ Écarts Détectés")
    
    discrepancies = results.get('discrepancies', [])
    
    if not discrepancies:
        st.success("🎉 Félicitations ! Aucun écart détecté - Tous les rapprochements sont parfaits.")
        return
    
    st.warning(f"⚠️ {len(discrepancies)} écart(s) détecté(s) nécessitant une attention")
    
    # Analyse statistique des écarts
    discrepancy_analysis = results.get('discrepancy_analysis', {})
    
    if discrepancy_analysis:
        disc_col1, disc_col2, disc_col3, disc_col4 = st.columns(4)
        
        with disc_col1:
            st.metric("💰 Écart Total", f"{discrepancy_analysis.get('total_discrepancy', 0):.2f} €")
        with disc_col2:
            st.metric("📊 Écart Moyen", f"{discrepancy_analysis.get('average_discrepancy', 0):.2f} €")
        with disc_col3:
            st.metric("⚠️ Écart Maximum", f"{discrepancy_analysis.get('max_discrepancy', 0):.2f} €")
        with disc_col4:
            st.metric("📉 Écart Minimum", f"{discrepancy_analysis.get('min_discrepancy', 0):.2f} €")
    
    # Classification des écarts par priorité
    critical_discrepancies = []
    high_discrepancies = []
    medium_discrepancies = []
    low_discrepancies = []
    
    for disc in discrepancies:
        difference = disc.get('difference', 0)
        percentage = disc.get('difference_percent', 0)
        
        if difference > 1000 or percentage > 15:
            critical_discrepancies.append(disc)
        elif difference > 500 or percentage > 10:
            high_discrepancies.append(disc)
        elif difference > 100 or percentage > 5:
            medium_discrepancies.append(disc)
        else:
            low_discrepancies.append(disc)
    
    # Affichage par niveau de priorité
    priority_tab1, priority_tab2, priority_tab3, priority_tab4 = st.tabs([
        f"🔴 Critiques ({len(critical_discrepancies)})",
        f"🟡 Élevés ({len(high_discrepancies)})",
        f"🟠 Moyens ({len(medium_discrepancies)})",
        f"🟢 Faibles ({len(low_discrepancies)})"
    ])
    
    with priority_tab1:
        show_discrepancy_level(critical_discrepancies, "🔴 Critiques", "Ces écarts nécessitent une vérification immédiate")
    
    with priority_tab2:
        show_discrepancy_level(high_discrepancies, "🟡 Élevés", "Ces écarts doivent être analysés rapidement")
    
    with priority_tab3:
        show_discrepancy_level(medium_discrepancies, "🟠 Moyens", "Ces écarts peuvent être traités en différé")
    
    with priority_tab4:
        show_discrepancy_level(low_discrepancies, "🟢 Faibles", "Ces écarts sont probablement dus à des arrondis")
    
    # Visualisation des écarts
    if len(discrepancies) > 1:
        st.markdown("#### 📊 Visualisation des Écarts")
        
        viz_col1, viz_col2 = st.columns(2)
        
        with viz_col1:
            # Histogramme des écarts
            amounts = [d.get('difference', 0) for d in discrepancies]
            fig_hist = px.histogram(
                x=amounts,
                title="Distribution des Écarts",
                labels={'x': 'Écart (€)', 'y': 'Fréquence'},
                nbins=min(20, len(amounts))
            )
            st.plotly_chart(fig_hist, use_container_width=True)
        
        with viz_col2:
            # Graphique en barres par commande
            order_numbers = [d.get('order_number', f'Cmd {i+1}') for i, d in enumerate(discrepancies)]
            fig_bar = px.bar(
                x=order_numbers[:10],  # Top 10
                y=amounts[:10],
                title="Top 10 des Écarts",
                labels={'x': 'N° Commande', 'y': 'Écart (€)'},
                color=amounts[:10],
                color_continuous_scale="Reds"
            )
            fig_bar.update_layout(xaxis_tickangle=-45)
            st.plotly_chart(fig_bar, use_container_width=True)

def show_discrepancy_level(discrepancies, level_name, description):
    """Affiche les écarts d'un niveau de priorité donné"""
    
    if not discrepancies:
        st.info(f"✅ Aucun écart {level_name.lower()}")
        return
    
    st.write(f"**{description}**")
    
    # Tableau des écarts de ce niveau
    level_data = []
    for i, disc in enumerate(discrepancies, 1):
        level_data.append({
            '#': i,
            'N° Commande': disc.get('order_number', 'N/A'),
            'Fichier PDF': disc.get('pdf_file', 'N/A'),
            'PDF (€)': f"{disc.get('pdf_amount', 0):.2f}",
            'Excel (€)': f"{disc.get('excel_amount', 0):.2f}",
            'Écart (€)': f"{disc.get('difference', 0):.2f}",
            'Écart (%)': f"{disc.get('difference_percent', 0):.2f}%",
            'Collaborateurs': disc.get('collaborators', 'N/A')[:30] + ('...' if len(disc.get('collaborators', '')) > 30 else ''),
            'Action': "🔍 À vérifier"
        })
    
    if level_data:
        level_df = pd.DataFrame(level_data)
        st.dataframe(level_df, use_container_width=True, hide_index=True)

def show_unmatched_tab(results):
    """Onglet des éléments non rapprochés avec analyse détaillée"""
    
    st.markdown("### ❌ Éléments Non Rapprochés")
    
    unmatched_pdf = results.get('unmatched_pdf', [])
    unmatched_excel = results.get('unmatched_excel', [])
    
    # Vue d'ensemble
    overview_col1, overview_col2 = st.columns(2)
    
    with overview_col1:
        st.metric("📄 PDFs Non Matchés", len(unmatched_pdf))
    
    with overview_col2:
        st.metric("📊 Excel Non Matchés", len(unmatched_excel))
    
    # Sections détaillées
    unmatched_tab1, unmatched_tab2 = st.tabs([
        f"📄 PDFs Non Rapprochés ({len(unmatched_pdf)})",
        f"📊 Excel Non Rapprochés ({len(unmatched_excel)})"
    ])
    
    with unmatched_tab1:
        st.markdown("#### 📄 Factures PDF Sans Correspondance Excel")
        
        if not unmatched_pdf:
            st.success("✅ Toutes les factures PDF ont été rapprochées!")
        else:
            st.error(f"❌ {len(unmatched_pdf)} facture(s) PDF non rapprochée(s)")
            
            # Analyse des raisons
            reason_counts = {}
            for pdf in unmatched_pdf:
                reasons = pdf.get('reasons', ['Raison inconnue'])
                for reason in reasons:
                    reason_counts[reason] = reason_counts.get(reason, 0) + 1
            
            if reason_counts:
                st.markdown("**📊 Analyse des Causes:**")
                for reason, count in sorted(reason_counts.items(), key=lambda x: x[1], reverse=True):
                    st.write(f"• **{reason}:** {count} cas")
            
            # Tableau des PDFs non matchés
            pdf_data = []
            for i, pdf in enumerate(unmatched_pdf, 1):
                pdf_data.append({
                    '#': i,
                    'Fichier': pdf.get('filename', 'N/A'),
                    'N° Commande': pdf.get('order_number', 'Non trouvé'),
                    'Montant (€)': f"{pdf.get('amount', 0):.2f}",
                    'ID Facture': pdf.get('invoice_id', 'N/A'),
                    'Fournisseur': pdf.get('supplier', 'N/A'),
                    'Qualité': f"{pdf.get('data_quality_score', 0)*100:.0f}%",
                    'Raison Principale': pdf.get('reason', 'Non déterminée')
                })
            
            if pdf_data:
                pdf_df = pd.DataFrame(pdf_data)
                st.dataframe(pdf_df, use_container_width=True, hide_index=True)
                
                # Recommandations spécifiques
                st.markdown("**💡 Actions Recommandées:**")
                st.info("🔍 Vérifiez les numéros de commande dans les fichiers Excel")
                st.info("📅 Contrôlez les périodes de facturation")
                st.info("🔧 Analysez la qualité de l'extraction PDF")
    
    with unmatched_tab2:
        st.markdown("#### 📊 Commandes Excel Sans Correspondance PDF")
        
        if not unmatched_excel:
            st.success("✅ Toutes les commandes Excel ont été rapprochées!")
        else:
            st.error(f"❌ {len(unmatched_excel)} commande(s) Excel non rapprochée(s)")
            
            # Statistiques des commandes Excel
            total_amount = sum(e.get('total_amount', 0) for e in unmatched_excel)
            total_lines = sum(e.get('line_count', 0) for e in unmatched_excel)
            
            excel_stats_col1, excel_stats_col2 = st.columns(2)
            with excel_stats_col1:
                st.metric("💰 Montant Non Matché", f"{total_amount:,.2f} €")
            with excel_stats_col2:
                st.metric("📋 Lignes Concernées", total_lines)
            
            # Tableau des commandes Excel non matchées
            excel_data = []
            for i, excel in enumerate(unmatched_excel, 1):
                excel_data.append({
                    '#': i,
                    'N° Commande': excel.get('order_number', 'N/A'),
                    'Montant (€)': f"{excel.get('total_amount', 0):.2f}",
                    'Collaborateurs': excel.get('collaborators', 'N/A'),
                    'Centres Coût': excel.get('cost_centers', 'N/A'),
                    'Nb Lignes': excel.get('line_count', 0),
                    'Fichiers Sources': excel.get('source_files', 'N/A'),
                    'Raison': 'Aucun PDF correspondant trouvé'
                })
            
            if excel_data:
                excel_df = pd.DataFrame(excel_data)
                st.dataframe(excel_df, use_container_width=True, hide_index=True)
                
                # Recommandations spécifiques
                st.markdown("**💡 Actions Recommandées:**")
                st.info("📄 Recherchez les factures PDF manquantes")
                st.info("📋 Vérifiez les numéros de commande dans les PDFs")
                st.info("📞 Contactez les fournisseurs pour les factures manquantes")
    
    # Recommandations globales
    if unmatched_pdf or unmatched_excel:
        st.markdown("---")
        st.markdown("#### 🔧 Recommandations Globales d'Amélioration")
        
        recommendations = [
            "🔍 **Contrôle Qualité:** Vérifiez la cohérence des numéros de commande entre sources",
            "📅 **Synchronisation:** Alignez les périodes de facturation PDF et Excel", 
            "🔧 **Paramétrage:** Ajustez les seuils de tolérance pour augmenter les matches",
            "📊 **Données:** Améliorez la qualité des données sources (extraction PDF, saisie Excel)",
            "📞 **Communication:** Contactez les équipes métier pour clarifier les écarts",
            "🔄 **Processus:** Mettez en place un suivi régulier des rapprochements"
        ]
        
        for rec in recommendations:
            st.write(rec)

def show_downloads_tab(results):
    """Onglet des téléchargements avec options d'export avancées"""
    
    st.markdown("### 📥 Téléchargements et Exports")
    
    # Section principale des exports
    export_main_col1, export_main_col2 = st.columns(2)
    
    with export_main_col1:
        st.markdown("#### 📊 Rapports Complets")
        
        # Excel complet multi-onglets
        if st.button("📈 Rapport Excel Complet", use_container_width=True, type="primary"):
            with st.spinner("Génération du rapport Excel..."):
                excel_data = create_complete_excel_report(results)
                st.download_button(
                    label="⬇️ Télécharger Rapport Excel",
                    data=excel_data,
                    file_name=f"rapport_beeline_complet_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                st.success("✅ Rapport Excel généré!")
        
        # CSV des matches
        if results.get('matches') and st.button("✅ Matches Parfaits (CSV)", use_container_width=True):
            csv_data = create_csv_matches(results['matches'])
            st.download_button(
                label="⬇️ Télécharger CSV Matches",
                data=csv_data,
                file_name=f"matches_parfaits_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                mime="text/csv",
                use_container_width=True
            )
        
        # CSV des écarts
        if results.get('discrepancies') and st.button("⚠️ Écarts Détectés (CSV)", use_container_width=True):
            csv_data = create_csv_discrepancies(results['discrepancies'])
            st.download_button(
                label="⬇️ Télécharger CSV Écarts",
                data=csv_data,
                file_name=f"ecarts_detectes_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                mime="text/csv",
                use_container_width=True
            )
    
    with export_main_col2:
        st.markdown("#### 🔧 Exports Techniques")
        
        # JSON technique complet
        if st.button("🔗 Données JSON Complètes", use_container_width=True):
            json_data = json.dumps(results, indent=2, default=str, ensure_ascii=False)
            st.download_button(
                label="⬇️ Télécharger JSON",
                data=json_data.encode('utf-8'),
                file_name=f"resultats_technique_{datetime.now().strftime('%Y%m%d_%H%M')}.json",
                mime="application/json",
                use_container_width=True
            )
        
        # Résumé exécutif
        if st.button("📋 Résumé Exécutif (PDF)", use_container_width=True):
            st.info("🔧 Fonctionnalité en développement - Utilisez l'export Excel pour l'instant")
        
        # Archive complète
        if st.button("📦 Archive Complète (ZIP)", use_container_width=True):
            with st.spinner("Création de l'archive..."):
                zip_data = create_complete_archive(results)
                st.download_button(
                    label="⬇️ Télécharger Archive ZIP",
                    data=zip_data,
                    file_name=f"archive_beeline_{datetime.now().strftime('%Y%m%d_%H%M')}.zip",
                    mime="application/zip",
                    use_container_width=True
                )
    
    # Section partage et sauvegarde
    st.markdown("---")
    st.markdown("#### 🔗 Partage et Collaboration")
    
    share_col1, share_col2 = st.columns(2)
    
    with share_col1:
        st.markdown("**📧 Partage par Email**")
        st.info("📧 Fonctionnalité email temporairement désactivée pour la stabilité")
        
        # email_recipient = st.text_input("Adresse email:", placeholder="nom@entreprise.com")
        # if st.button("📤 Envoyer Rapport"):
        #     st.info("Envoi en cours...")
    
    with share_col2:
        st.markdown("**💾 Sauvegarde Projet**")
        
        if st.button("💾 Sauvegarder dans l'Historique", use_container_width=True):
            save_to_history(results)
            st.success("✅ Résultats sauvegardés dans l'historique!")
        
        if st.button("🔗 Générer Lien de Partage", use_container_width=True):
            # Simulation d'un lien de partage
            share_link = f"https://beeline-app.streamlit.app/results/{datetime.now().strftime('%Y%m%d%H%M')}"
            st.code(share_link, language=None)
            st.caption("🔗 Lien de partage généré (simulation - validité 24h)")
    
    # Informations sur les exports
    with st.expander("ℹ️ Informations sur les Exports"):
        st.markdown("""
        **📊 Rapport Excel Complet** - Contient :
        - Synthèse des résultats avec KPIs
        - Liste détaillée des matches parfaits
        - Analyse des écarts avec priorités
        - Éléments non rapprochés avec diagnostics
        - Graphiques et métriques de performance
        
        **📄 Exports CSV** - Formats :
        - Encodage UTF-8 pour la compatibilité internationale
        - Séparateur virgule standard
        - En-têtes explicites en français
        
        **🔧 Export JSON** - Données techniques :
        - Structure complète des résultats
        - Métadonnées de traitement
        - Configuration utilisée
        - Pour intégration avec d'autres systèmes
        """)

def show_diagnostics_tab(results):
    """Onglet de diagnostics avancés et recommandations"""
    
    st.markdown("### 🔧 Diagnostics Avancés")
    
    # Métriques de performance globale
    metadata = results.get('metadata', {})
    performance_stats = metadata.get('performance_stats', {})
    
    st.markdown("#### 📊 Métriques de Performance")
    
    perf_col1, perf_col2, perf_col3, perf_col4 = st.columns(4)
    
    with perf_col1:
        processing_time = metadata.get('processing_time', 0)
        st.metric("⏱️ Temps Total", f"{processing_time:.2f}s")
    
    with perf_col2:
        total_files = performance_stats.get('total_pdfs', 0) + len(st.session_state.uploaded_excels)
        throughput = total_files / processing_time if processing_time > 0 else 0
        st.metric("📈 Débit", f"{throughput:.1f} fichiers/s")
    
    with perf_col3:
        config_used = metadata.get('config_used', {})
        tolerance = config_used.get('tolerance', 0) * 100
        st.metric("🎯 Tolérance", f"{tolerance:.1f}%")
    
    with perf_col4:
        engine_version = metadata.get('engine_version', 'N/A')
        st.metric("🔧 Version Moteur", engine_version)
    
    # Analyse de la qualité des données
    st.markdown("#### 🔍 Analyse de la Qualité des Données")
    
    quality_tabs = st.tabs(["📄 Qualité PDFs", "📊 Qualité Excel", "⚖️ Qualité Rapprochement"])
    
    with quality_tabs[0]:
        analyze_pdf_quality(results)
    
    with quality_tabs[1]:
        analyze_excel_quality(results)
    
    with quality_tabs[2]:
        analyze_matching_quality(results)
    
    # Recommandations d'amélioration
    if results.get('summary', {}).get('quality_assessment', {}).get('recommendations'):
        st.markdown("#### 💡 Recommandations d'Amélioration")
        
        recommendations = results['summary']['quality_assessment']['recommendations']
        
        for i, rec in enumerate(recommendations, 1):
            with st.expander(f"💡 Recommandation {i}"):
                st.write(rec)
                
                # Actions suggérées basées sur la recommandation
                if "taux de rapprochement" in rec.lower():
                    st.markdown("**Actions possibles:**")
                    st.write("• Vérifiez la cohérence des numéros de commande")
                    st.write("• Ajustez les paramètres de tolérance")
                    st.write("• Contrôlez les périodes de facturation")
                
                elif "qualité" in rec.lower():
                    st.markdown("**Actions possibles:**")
                    st.write("• Améliorez l'extraction des données PDF")
                    st.write("• Validez les données Excel avant traitement")
                    st.write("• Nettoyez les données sources")
    
    # Configuration utilisée
    st.markdown("#### ⚙️ Configuration du Traitement")
    
    if config_used:
        config_df = pd.DataFrame([
            {"Paramètre": "Tolérance montants", "Valeur": f"{config_used.get('tolerance', 0)*100:.1f}%"},
            {"Paramètre": "Méthode rapprochement", "Valeur": config_used.get('method', 'N/A')},
            {"Paramètre": "Confiance minimum", "Valeur": f"{config_used.get('min_confidence', 0)*100:.1f}%"},
            {"Paramètre": "Matching références", "Valeur": "✅ Activé" if config_used.get('enable_reference_matching') else "❌ Désactivé"},
            {"Paramètre": "Rapport détaillé", "Valeur": "✅ Activé" if config_used.get('generate_detailed_report') else "❌ Désactivé"}
        ])
        
        st.dataframe(config_df, use_container_width=True, hide_index=True)

def analyze_pdf_quality(results):
    """Analyse la qualité des données PDF"""
    
    if not st.session_state.pdf_data:
        st.info("Aucune donnée PDF disponible pour l'analyse")
        return
    
    pdf_data = st.session_state.pdf_data
    
    # Statistiques globales
    total_pdfs = len(pdf_data)
    successful_extractions = len([p for p in pdf_data if p.get('success', False)])
    
    pdf_qual_col1, pdf_qual_col2 = st.columns(2)
    
    with pdf_qual_col1:
        st.metric("📄 PDFs Analysés", total_pdfs)
        st.metric("✅ Extractions Réussies", successful_extractions)
        success_rate = (successful_extractions / total_pdfs * 100) if total_pdfs > 0 else 0
        st.metric("📈 Taux de Réussite", f"{success_rate:.1f}%")
    
    with pdf_qual_col2:
        # Analyse des scores de qualité
        quality_scores = []
        for pdf in pdf_data:
            if pdf.get('success') and pdf.get('data_completeness'):
                quality_scores.append(pdf['data_completeness'].get('overall_score', 0))
        
        if quality_scores:
            avg_quality = sum(quality_scores) / len(quality_scores)
            st.metric("📊 Qualité Moyenne", f"{avg_quality:.1f}%")
            
            # Distribution des scores de qualité
            fig_quality = px.histogram(
                x=quality_scores,
                title="Distribution des Scores de Qualité PDF",
                labels={'x': 'Score de Qualité (%)', 'y': 'Nombre de PDFs'},
                nbins=10
            )
            st.plotly_chart(fig_quality, use_container_width=True)

def analyze_excel_quality(results):
    """Analyse la qualité des données Excel"""
    
    if not st.session_state.excel_data:
        st.info("Aucune donnée Excel disponible pour l'analyse")
        return
    
    excel_data = st.session_state.excel_data
    
    # Statistiques globales
    total_lines = len(excel_data)
    valid_lines = len([e for e in excel_data if e.get('is_valid', False)])
    
    excel_qual_col1, excel_qual_col2 = st.columns(2)
    
    with excel_qual_col1:
        st.metric("📊 Lignes Analysées", total_lines)
        st.metric("✅ Lignes Valides", valid_lines)
        validity_rate = (valid_lines / total_lines * 100) if total_lines > 0 else 0
        st.metric("📈 Taux de Validité", f"{validity_rate:.1f}%")
    
    with excel_qual_col2:
        # Analyse des commandes uniques
        unique_orders = set()
        for line in excel_data:
            if line.get('order_number'):
                unique_orders.add(str(line['order_number']))
        
        st.metric("🔢 Commandes Uniques", len(unique_orders))
        
        # Sources de fichiers
        source_files = set()
        for line in excel_data:
            if line.get('source_filename'):
                source_files.add(line['source_filename'])
        
        st.metric("📁 Fichiers Sources", len(source_files))

def analyze_matching_quality(results):
    """Analyse la qualité du processus de rapprochement"""
    
    summary = results.get('summary', {})
    
    # Métriques de rapprochement
    match_qual_col1, match_qual_col2, match_qual_col3 = st.columns(3)
    
    with match_qual_col1:
        matching_rate = summary.get('matching_rate', 0)
        st.metric("🎯 Taux Rapprochement", f"{matching_rate:.1f}%")
        
        coverage_rate = summary.get('coverage_rate', 0)
        st.metric("📊 Taux Couverture", f"{coverage_rate:.1f}%")
    
    with match_qual_col2:
        if results.get('discrepancy_analysis'):
            avg_discrepancy = results['discrepancy_analysis'].get('average_discrepancy', 0)
            st.metric("💰 Écart Moyen", f"{avg_discrepancy:.2f} €")
            
            max_discrepancy = results['discrepancy_analysis'].get('max_discrepancy', 0)
            st.metric("⚠️ Écart Maximum", f"{max_discrepancy:.2f} €")
    
    with match_qual_col3:
        # Performance par méthode
        method_performance = results.get('metadata', {}).get('performance_stats', {}).get('method_performance', {})
        
        if method_performance:
            best_method = max(method_performance.keys(), key=lambda k: method_performance[k]['avg_confidence'])
            best_confidence = method_performance[best_method]['avg_confidence']
            
            st.metric("🏆 Meilleure Méthode", best_method.title())
            st.metric("🎯 Confiance Max", f"{best_confidence*100:.1f}%")

def show_history_page():
    """Page d'historique des traitements avec analytics"""
    
    st.markdown("## 📈 Historique des Traitements")
    
    # Initialisation de l'historique si vide
    if 'processing_history' not in st.session_state:
        st.session_state.processing_history = []
    
    if not st.session_state.processing_history:
        st.info("📭 Aucun traitement dans l'historique pour le moment.")
        
        # Section d'aide
        st.markdown("""
        ### 🔍 Comment Alimenter l'Historique
        
        L'historique se remplit automatiquement lorsque vous :
        1. **Effectuez des rapprochements** via la section ⚖️ Rapprochement
        2. **Sauvegardez des résultats** via la section 📊 Résultats > 📥 Téléchargements
        3. **Finalisez des traitements** complets
        
        ### 📊 Informations Suivies
        - Date et heure du traitement
        - Nombre de fichiers traités (PDFs et Excel)
        - Résultats du rapprochement (matches, écarts)
        - Taux de réussite et métriques de performance
        - Configuration utilisée
        """)
        
        # Bouton pour ajouter un exemple de démonstration
        if st.button("🎲 Ajouter des Exemples de Démonstration"):
            add_demo_history_entries()
            st.success("✅ Exemples ajoutés à l'historique!")
            st.rerun()
        
        return
    
    # Affichage de l'historique existant
    st.success(f"📊 {len(st.session_state.processing_history)} traitement(s) dans l'historique")
    
    # Conversion en DataFrame pour l'analyse
    history_df = pd.DataFrame(st.session_state.processing_history)
    history_df['Date'] = pd.to_datetime(history_df['date']).dt.strftime('%d/%m/%Y %H:%M')
    
    # Vue tabulaire de l'historique
    st.markdown("### 📋 Vue Tabulaire")
    
    # Sélection des colonnes à afficher
    display_columns = st.multiselect(
        "Colonnes à afficher:",
        ['Date', 'pdf_count', 'excel_count', 'matches', 'discrepancies', 'matching_rate', 'total_amount'],
        default=['Date', 'pdf_count', 'excel_count', 'matches', 'matching_rate'],
        key="history_columns"
    )
    
    if display_columns:
        # Renommage des colonnes pour l'affichage
        column_names = {
            'pdf_count': 'PDFs',
            'excel_count': 'Excel',
            'matches': 'Matches',
            'discrepancies': 'Écarts',
            'matching_rate': 'Taux (%)',
            'total_amount': 'Montant (€)'
        }
        
        display_df = history_df[display_columns].copy()
        display_df = display_df.rename(columns=column_names)
        
        # Formatage des colonnes numériques
        if 'Taux (%)' in display_df.columns:
            display_df['Taux (%)'] = display_df['Taux (%)'].apply(lambda x: f"{x:.1f}%")
        if 'Montant (€)' in display_df.columns:
            display_df['Montant (€)'] = display_df['Montant (€)'].apply(lambda x: f"{x:,.2f} €")
        
        st.dataframe(display_df, use_container_width=True, hide_index=True)
    
    # Analyses graphiques
    if len(st.session_state.processing_history) > 1:
        st.markdown("### 📊 Analyses Graphiques")
        
        analysis_tab1, analysis_tab2, analysis_tab3 = st.tabs([
            "📈 Évolution Performance", "💰 Évolution Montants", "📊 Distribution Résultats"
        ])
        
        with analysis_tab1:
            # Graphique d'évolution du taux de rapprochement
            fig_evolution = px.line(
                history_df,
                x='date',
                y='matching_rate',
                title="Évolution du Taux de Rapprochement",
                labels={'matching_rate': 'Taux de Réussite (%)', 'date': 'Date'},
                markers=True
            )
            fig_evolution.update_layout(yaxis_range=[0, 100])
            st.plotly_chart(fig_evolution, use_container_width=True)
            
            # Corrélation nombre de fichiers vs performance
            fig_correlation = px.scatter(
                history_df,
                x=history_df['pdf_count'] + history_df['excel_count'],
                y='matching_rate',
                title="Relation Nombre de Fichiers vs Performance",
                labels={'x': 'Nombre Total de Fichiers', 'y': 'Taux de Réussite (%)'},
                hover_data=['Date']
            )
            st.plotly_chart(fig_correlation, use_container_width=True)
        
        with analysis_tab2:
            # Évolution des montants traités
            fig_amounts = px.bar(
                history_df,
                x='Date',
                y='total_amount',
                title="Évolution des Montants Traités",
                labels={'total_amount': 'Montant Total (€)', 'Date': 'Date'}
            )
            fig_amounts.update_layout(xaxis_tickangle=-45)
            st.plotly_chart(fig_amounts, use_container_width=True)
        
        with analysis_tab3:
            # Distribution des résultats
            avg_matches = history_df['matches'].mean()
            avg_discrepancies = history_df['discrepancies'].mean()
            
            distribution_data = pd.DataFrame({
                'Type': ['Matches Moyens', 'Écarts Moyens'],
                'Valeur': [avg_matches, avg_discrepancies]
            })
            
            fig_dist = px.pie(
                distribution_data,
                values='Valeur',
                names='Type',
                title="Distribution Moyenne des Résultats"
            )
            st.plotly_chart(fig_dist, use_container_width=True)
    
    # Actions sur l'historique
    st.markdown("---")
    st.markdown("### 🔧 Actions sur l'Historique")
    
    action_col1, action_col2, action_col3 = st.columns(3)
    
    with action_col1:
        if st.button("📥 Exporter Historique (CSV)", use_container_width=True):
            csv_data = history_df.to_csv(index=False)
            st.download_button(
                label="⬇️ Télécharger CSV",
                data=csv_data,
                file_name=f"historique_beeline_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                mime="text/csv"
            )
    
    with action_col2:
        if st.button("🗑️ Vider l'Historique", use_container_width=True):
            if st.session_state.get('confirm_clear_history', False):
                st.session_state.processing_history = []
                st.session_state.confirm_clear_history = False
                st.success("✅ Historique vidé!")
                st.rerun()
            else:
                st.session_state.confirm_clear_history = True
                st.warning("⚠️ Cliquez à nouveau pour confirmer la suppression")
    
    with action_col3:
        if st.button("🎲 Ajouter Exemples", use_container_width=True):
            add_demo_history_entries()
            st.success("✅ Exemples ajoutés!")
            st.rerun()

def add_demo_history_entries():
    """Ajoute des entrées de démonstration à l'historique"""
    
    from datetime import timedelta
    import random
    
    demo_entries = []
    base_date = datetime.now()
    
    for i in range(5):
        # Dates échelonnées sur les 30 derniers jours
        entry_date = base_date - timedelta(days=random.randint(1, 30))
        
        # Génération de données réalistes
        pdf_count = random.randint(3, 15)
        excel_count = random.randint(1, 5)
        matches = random.randint(int(pdf_count * 0.6), pdf_count)
        discrepancies = random.randint(0, pdf_count - matches)
        matching_rate = (matches / pdf_count) * 100 if pdf_count > 0 else 0
        total_amount = random.uniform(5000, 50000)
        
        demo_entries.append({
            'date': entry_date,
            'pdf_count': pdf_count,
            'excel_count': excel_count,
            'matches': matches,
            'discrepancies': discrepancies,
            'matching_rate': matching_rate,
            'total_amount': total_amount,
            'processing_time': random.uniform(5.0, 30.0),
            'demo': True  # Marquer comme données de démo
        })
    
    # Tri par date
    demo_entries.sort(key=lambda x: x['date'])
    
    # Ajout à l'historique existant
    st.session_state.processing_history.extend(demo_entries)

# Fonctions utilitaires pour les exports

def create_complete_excel_report(results):
    """Crée un rapport Excel complet multi-onglets"""
    
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        
        # Formats pour l'Excel
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'fg_color': '#D7E4BC',
            'border': 1
        })
        
        number_format = workbook.add_format({'num_format': '#,##0.00'})
        percent_format = workbook.add_format({'num_format': '0.0%'})
        
        # Onglet 1: Synthèse
        summary = results.get('summary', {})
        synthesis_data = {
            'Métrique': [
                'Total PDFs Traités',
                'Total Lignes Excel',
                'Matches Parfaits',
                'Écarts Détectés', 
                'PDFs Non Matchés',
                'Excel Non Matchés',
                'Taux de Réussite (%)',
                'Taux de Couverture (%)',
                'Montant Total (€)'
            ],
            'Valeur': [
                summary.get('total_invoices', 0),
                summary.get('total_excel_lines', 0),
                summary.get('perfect_matches', 0),
                summary.get('discrepancies', 0),
                summary.get('unmatched_pdf', 0),
                summary.get('unmatched_excel', 0),
                summary.get('matching_rate', 0),
                summary.get('coverage_rate', 0),
                summary.get('total_amount', 0)
            ]
        }
        
        synthesis_df = pd.DataFrame(synthesis_data)
        synthesis_df.to_excel(writer, sheet_name='Synthèse', index=False)
        
        # Onglet 2: Matches parfaits
        if results.get('matches'):
            matches_df = pd.DataFrame(results['matches'])
            matches_df.to_excel(writer, sheet_name='Matches Parfaits', index=False)
        
        # Onglet 3: Écarts
        if results.get('discrepancies'):
            discrepancies_df = pd.DataFrame(results['discrepancies'])
            discrepancies_df.to_excel(writer, sheet_name='Écarts', index=False)
        
        # Onglet 4: Non matchés PDFs
        if results.get('unmatched_pdf'):
            unmatched_pdf_df = pd.DataFrame(results['unmatched_pdf'])
            unmatched_pdf_df.to_excel(writer, sheet_name='PDFs Non Matchés', index=False)
        
        # Onglet 5: Non matchés Excel
        if results.get('unmatched_excel'):
            unmatched_excel_df = pd.DataFrame(results['unmatched_excel'])
            unmatched_excel_df.to_excel(writer, sheet_name='Excel Non Matchés', index=False)
        
        # Onglet 6: Configuration et métadonnées
        metadata = results.get('metadata', {})
        config_data = {
            'Paramètre': ['Version Moteur', 'Temps de Traitement (s)', 'Date de Traitement'],
            'Valeur': [
                metadata.get('engine_version', 'N/A'),
                metadata.get('processing_time', 0),
                metadata.get('reconciliation_timestamp', 'N/A')
            ]
        }
        
        config_df = pd.DataFrame(config_data)
        config_df.to_excel(writer, sheet_name='Configuration', index=False)
    
    return output.getvalue()

def create_matches_excel_report(matches_df):
    """Crée un rapport Excel spécifique aux matches"""
    
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        matches_df.to_excel(writer, sheet_name='Matches Parfaits', index=False)
        
        # Ajout de graphiques si possible
        workbook = writer.book
        worksheet = writer.sheets['Matches Parfaits']
        
        # Format des colonnes
        worksheet.set_column('A:I', 15)
    
    return output.getvalue()

def create_csv_matches(matches):
    """Crée un CSV des matches parfaits"""
    
    matches_data = []
    for match in matches:
        matches_data.append({
            'N° Commande': match.get('order_number', ''),
            'Fichier PDF': match.get('pdf_file', ''),
            'Montant PDF (€)': match.get('pdf_amount', 0),
            'Montant Excel (€)': match.get('excel_amount', 0),
            'Écart (€)': match.get('difference', 0),
            'Collaborateurs': match.get('collaborators', ''),
            'Supplier': match.get('supplier', ''),
            'Méthode': match.get('method', ''),
            'Confiance': match.get('confidence', 0)
        })
    
    matches_df = pd.DataFrame(matches_data)
    return matches_df.to_csv(index=False, encoding='utf-8')

def create_csv_discrepancies(discrepancies):
    """Crée un CSV des écarts"""
    
    discrepancies_data = []
    for disc in discrepancies:
        difference = disc.get('difference', 0)
        percentage = disc.get('difference_percent', 0)
        
        # Classification de priorité
        if difference > 1000 or percentage > 15:
            priority = "Critique"
        elif difference > 500 or percentage > 10:
            priority = "Élevée"
        elif difference > 100 or percentage > 5:
            priority = "Moyenne"
        else:
            priority = "Faible"
        
        discrepancies_data.append({
            'Priorité': priority,
            'N° Commande': disc.get('order_number', ''),
            'Fichier PDF': disc.get('pdf_file', ''),
            'Montant PDF (€)': disc.get('pdf_amount', 0),
            'Montant Excel (€)': disc.get('excel_amount', 0),
            'Écart (€)': difference,
            'Écart (%)': percentage,
            'Collaborateurs': disc.get('collaborators', ''),
            'Méthode': disc.get('method', ''),
            'Confiance': disc.get('confidence', 0)
        })
    
    discrepancies_df = pd.DataFrame(discrepancies_data)
    return discrepancies_df.to_csv(index=False, encoding='utf-8')

def create_complete_archive(results):
    """Crée une archive ZIP complète avec tous les exports"""
    
    zip_buffer = io.BytesIO()
    
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        
        # Rapport Excel complet
        excel_data = create_complete_excel_report(results)
        zip_file.writestr(
            f"rapport_complet_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            excel_data
        )
        
        # CSV des matches
        if results.get('matches'):
            csv_matches = create_csv_matches(results['matches'])
            zip_file.writestr(
                f"matches_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                csv_matches.encode('utf-8')
            )
        
        # CSV des écarts
        if results.get('discrepancies'):
            csv_discrepancies = create_csv_discrepancies(results['discrepancies'])
            zip_file.writestr(
                f"ecarts_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                csv_discrepancies.encode('utf-8')
            )
        
        # JSON technique complet
        json_data = json.dumps(results, indent=2, default=str, ensure_ascii=False)
        zip_file.writestr(
            f"donnees_techniques_{datetime.now().strftime('%Y%m%d_%H%M')}.json",
            json_data.encode('utf-8')
        )
        
        # Fichier README
        readme_content = f"""
# ARCHIVE BEELINE - Rapprochement PDF/Excel
Date de génération : {datetime.now().strftime('%d/%m/%Y %H:%M')}

## Contenu de l'archive

### Rapports Excel
- rapport_complet_*.xlsx : Rapport Excel multi-onglets avec toutes les analyses

### Exports CSV
- matches_*.csv : Liste des rapprochements parfaits
- ecarts_*.csv : Liste des écarts détectés avec priorités

### Données techniques
- donnees_techniques_*.json : Données brutes complètes pour intégration

## Statistiques du traitement
- Total PDFs traités : {results.get('summary', {}).get('total_invoices', 0)}
- Total lignes Excel : {results.get('summary', {}).get('total_excel_lines', 0)}
- Taux de réussite : {results.get('summary', {}).get('matching_rate', 0):.1f}%
- Temps de traitement : {results.get('metadata', {}).get('processing_time', 0):.2f}s

## Support
Pour toute question sur ces résultats, consultez la documentation de l'application Beeline.
        """
        
        zip_file.writestr("README.txt", readme_content.encode('utf-8'))
    
    return zip_buffer.getvalue()

def save_to_history(results):
    """Sauvegarde les résultats dans l'historique de session"""
    
    if 'processing_history' not in st.session_state:
        st.session_state.processing_history = []
    
    summary = results.get('summary', {})
    metadata = results.get('metadata', {})
    
    history_entry = {
        'date': datetime.now(),
        'pdf_count': summary.get('total_invoices', 0),
        'excel_count': summary.get('total_excel_lines', 0),
        'matches': summary.get('perfect_matches', 0),
        'discrepancies': summary.get('discrepancies', 0),
        'matching_rate': summary.get('matching_rate', 0),
        'total_amount': summary.get('total_amount', 0),
        'processing_time': metadata.get('processing_time', 0),
        'engine_version': metadata.get('engine_version', '2.1.0')
    }
    
    st.session_state.processing_history.append(history_entry)

# Point d'entrée principal
if __name__ == "__main__":
    main()

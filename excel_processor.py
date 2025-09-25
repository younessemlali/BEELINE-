"""
EXCEL_PROCESSOR.PY
Module de traitement des fichiers Excel pour l'application Beeline
Traitement natif des données Excel/CSV avec Pandas
"""

import pandas as pd
import numpy as np
import re
import logging
from typing import Dict, List, Optional, Any, Union
from datetime import datetime
import streamlit as st

class ExcelProcessor:
    """
    Processeur de fichiers Excel spécialisé pour les données Beeline
    """
    
    def __init__(self):
        self.setup_logging()
        
        # Configuration des colonnes attendues (flexible)
        self.expected_columns = {
            'order_number': [
                'N° commande', 'Numero commande', 'Order Number', 'Purchase Order',
                'Commande', 'PO', 'Bon de commande'
            ],
            'net_amount': [
                'Montant net à payer au fournisseur', 'Net Amount', 'Montant Net',
                'Amount', 'Total', 'Montant'
            ],
            'collaborator': [
                'Collaborateur', 'Employee', 'Worker', 'Nom', 'Name',
                'Consultant', 'Contractor'
            ],
            'statement_date': [
                'Statement Date', 'Date', 'Invoice Date', 'Billing Date',
                'Date facture', 'Période'
            ],
            'invoice_number': [
                'Invoice Number', 'N° facture', 'Facture', 'Invoice ID'
            ],
            'supplier': [
                'Supplier', 'Fournisseur', 'Vendor', 'Company'
            ],
            'billing_period': [
                'Billing Period', 'Période facturation', 'Period', 'Semaine'
            ],
            'units': [
                'Unités', 'Units', 'Hours', 'Heures', 'Quantity', 'Quantité'
            ],
            'rate': [
                'Taux de facturation', 'Rate', 'Unit Price', 'Prix unitaire', 'Hourly Rate'
            ],
            'gross_amount': [
                'Montant brut', 'Gross Amount', 'Total Gross', 'Brut'
            ]
        }
        
        # Configuration de validation
        self.validation_rules = {
            'order_number': r'^[0-9]{8,12}$',  # 8 à 12 chiffres
            'amount': r'^-?[0-9]+\.?[0-9]*$',   # Nombre décimal (peut être négatif)
            'date': [
                '%d/%m/%Y', '%Y-%m-%d', '%d-%m-%Y', '%Y/%m/%d',
                '%d/%m/%Y %H:%M:%S', '%Y-%m-%d %H:%M:%S'
            ]
        }
        
        # Seuils de validation
        self.thresholds = {
            'min_amount': -10000,     # Montant minimum (peut être négatif)
            'max_amount': 50000,      # Montant maximum
            'min_units': 0,           # Unités minimum
            'max_units': 200,         # Unités maximum (heures/semaine)
            'min_rate': 0,            # Taux minimum
            'max_rate': 200           # Taux maximum par heure
        }
    
    def setup_logging(self):
        """Configure le logging pour le module"""
        logging.basicConfig(level=logging.INFO)
        self.logger = logging.getLogger(__name__)
    
    def process_excel_file(self, excel_file) -> List[Dict[str, Any]]:
        """
        Traite un fichier Excel/CSV complet
        
        Args:
            excel_file: Fichier Excel uploadé (Streamlit UploadedFile)
            
        Returns:
            Liste des lignes traitées
        """
        try:
            self.logger.info(f"Début traitement Excel: {excel_file.name}")
            
            # Lecture du fichier selon son type
            df = self.read_excel_file(excel_file)
            
            if df is None or df.empty:
                raise ValueError("Fichier vide ou illisible")
            
            # Nettoyage et normalisation
            df = self.clean_dataframe(df)
            
            # Mapping des colonnes
            column_mapping = self.map_columns(df.columns.tolist())
            df = self.rename_columns(df, column_mapping)
            
            # Validation des données
            validation_results = self.validate_dataframe(df)
            
            # Conversion en liste de dictionnaires
            processed_data = self.dataframe_to_records(df, excel_file.name, validation_results)
            
            # Métadonnées de traitement
            metadata = {
                'filename': excel_file.name,
                'file_size': excel_file.size,
                'total_rows': len(df),
                'valid_rows': len([r for r in processed_data if r.get('is_valid', False)]),
                'columns_found': list(column_mapping.values()),
                'processing_timestamp': datetime.now().isoformat(),
                'processor_version': '2.0.0'
            }
            
            # Ajout des métadonnées à chaque enregistrement
            for record in processed_data:
                record['processing_metadata'] = metadata
            
            self.logger.info(f"Traitement réussi: {len(processed_data)} lignes traitées")
            return processed_data
            
        except Exception as e:
            self.logger.error(f"Erreur traitement Excel {excel_file.name}: {str(e)}")
            return [{
                'success': False,
                'error': str(e),
                'filename': excel_file.name,
                'processing_timestamp': datetime.now().isoformat()
            }]
    
    def read_excel_file(self, excel_file) -> Optional[pd.DataFrame]:
        """
        Lit un fichier Excel/CSV selon son extension
        
        Args:
            excel_file: Fichier à lire
            
        Returns:
            DataFrame pandas ou None
        """
        filename = excel_file.name.lower()
        
        try:
            if filename.endswith('.csv'):
                # Lecture CSV avec détection automatique du délimiteur
                return self.read_csv_with_detection(excel_file)
            
            elif filename.endswith(('.xlsx', '.xls')):
                # Lecture Excel
                return pd.read_excel(excel_file, engine='openpyxl' if filename.endswith('.xlsx') else None)
            
            else:
                raise ValueError(f"Format de fichier non supporté: {filename}")
                
        except Exception as e:
            self.logger.error(f"Erreur lecture fichier {filename}: {str(e)}")
            return None
    
    def read_csv_with_detection(self, csv_file) -> pd.DataFrame:
        """
        Lit un CSV avec détection automatique des paramètres
        
        Args:
            csv_file: Fichier CSV
            
        Returns:
            DataFrame pandas
        """
        # Tentatives avec différents délimiteurs
        delimiters = [',', ';', '\t', '|']
        encodings = ['utf-8', 'latin-1', 'cp1252', 'iso-8859-1']
        
        for encoding in encodings:
            for delimiter in delimiters:
                try:
                    # Reset du pointeur de fichier
                    csv_file.seek(0)
                    
                    df = pd.read_csv(
                        csv_file,
                        delimiter=delimiter,
                        encoding=encoding,
                        skipinitialspace=True,
                        na_values=['', 'NULL', 'null', 'N/A', 'n/a', '-']
                    )
                    
                    # Vérification de la qualité de la lecture
                    if len(df.columns) > 1 and len(df) > 0:
                        self.logger.info(f"CSV lu avec succès: délimiteur='{delimiter}', encoding='{encoding}'")
                        return df
                        
                except Exception as e:
                    continue
        
        # Fallback: lecture basique
        csv_file.seek(0)
        return pd.read_csv(csv_file)
    
    def clean_dataframe(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Nettoie et normalise un DataFrame
        
        Args:
            df: DataFrame à nettoyer
            
        Returns:
            DataFrame nettoyé
        """
        df_cleaned = df.copy()
        
        # Suppression des lignes entièrement vides
        df_cleaned = df_cleaned.dropna(how='all')
        
        # Suppression des colonnes entièrement vides
        df_cleaned = df_cleaned.dropna(axis=1, how='all')
        
        # Nettoyage des noms de colonnes
        df_cleaned.columns = [
            str(col).strip().replace('\n', ' ').replace('\r', ' ')
            for col in df_cleaned.columns
        ]
        
        # Suppression des espaces multiples dans les noms de colonnes
        df_cleaned.columns = [re.sub(r'\s+', ' ', col) for col in df_cleaned.columns]
        
        # Nettoyage des données texte
        for col in df_cleaned.select_dtypes(include=['object']).columns:
            df_cleaned[col] = df_cleaned[col].astype(str).str.strip()
            df_cleaned[col] = df_cleaned[col].replace(['nan', 'NaN', 'None'], '')
        
        # Conversion des montants (détection automatique)
        for col in df_cleaned.columns:
            if self.is_amount_column(col, df_cleaned[col]):
                df_cleaned[col] = self.clean_amount_column(df_cleaned[col])
        
        return df_cleaned
    
    def is_amount_column(self, col_name: str, series: pd.Series) -> bool:
        """
        Détermine si une colonne contient des montants
        
        Args:
            col_name: Nom de la colonne
            series: Série pandas
            
        Returns:
            True si c'est une colonne de montants
        """
        # Détection par nom de colonne
        amount_keywords = ['montant', 'amount', 'total', 'price', 'cost', 'taux', 'rate']
        col_lower = col_name.lower()
        
        if any(keyword in col_lower for keyword in amount_keywords):
            return True
        
        # Détection par contenu (échantillon)
        sample = series.dropna().head(100).astype(str)
        if len(sample) == 0:
            return False
        
        # Compter les valeurs qui ressemblent à des montants
        amount_pattern = r'^-?[0-9]{1,}[,.]?[0-9]*$'
        amount_like = sum(1 for val in sample if re.match(amount_pattern, val.replace(' ', '')))
        
        # Si plus de 70% des valeurs ressemblent à des montants
        return (amount_like / len(sample)) > 0.7
    
    def clean_amount_column(self, series: pd.Series) -> pd.Series:
        """
        Nettoie une colonne de montants
        
        Args:
            series: Série à nettoyer
            
        Returns:
            Série nettoyée et convertie en float
        """
        def parse_amount(val):
            if pd.isna(val):
                return 0.0
            
            val_str = str(val).strip()
            
            # Suppression des caractères non numériques sauf , . -
            cleaned = re.sub(r'[^\d,.-]', '', val_str)
            
            if not cleaned or cleaned in ['-', '.', ',']:
                return 0.0
            
            # Gestion des formats européens vs anglo-saxons
            if ',' in cleaned and '.' in cleaned:
                # Format européen: 1.234,56 -> 1234.56
                if cleaned.rfind(',') > cleaned.rfind('.'):
                    cleaned = cleaned.replace('.', '').replace(',', '.')
                # Format US: 1,234.56 -> 1234.56
                else:
                    cleaned = cleaned.replace(',', '')
            elif ',' in cleaned:
                # Détection du contexte
                comma_pos = cleaned.rfind(',')
                after_comma = cleaned[comma_pos + 1:]
                
                if len(after_comma) <= 2:  # Probablement des décimales
                    cleaned = cleaned.replace(',', '.')
                else:  # Séparateur de milliers
                    cleaned = cleaned.replace(',', '')
            
            try:
                return float(cleaned)
            except (ValueError, TypeError):
                return 0.0
        
        return series.apply(parse_amount)
    
    def map_columns(self, column_names: List[str]) -> Dict[str, str]:
        """
        Mappe les noms de colonnes trouvés aux noms standardisés
        
        Args:
            column_names: Liste des noms de colonnes du fichier
            
        Returns:
            Dictionnaire de mapping {colonne_fichier: colonne_standard}
        """
        mapping = {}
        
        for standard_name, possible_names in self.expected_columns.items():
            best_match = None
            best_score = 0
            
            for col_name in column_names:
                if col_name in mapping.values():  # Déjà mappée
                    continue
                
                # Score de correspondance
                score = self.calculate_column_match_score(col_name, possible_names)
                
                if score > best_score and score > 0.7:  # Seuil de confiance
                    best_match = col_name
                    best_score = score
            
            if best_match:
                mapping[best_match] = standard_name
        
        return mapping
    
    def calculate_column_match_score(self, col_name: str, expected_names: List[str]) -> float:
        """
        Calcule un score de correspondance entre un nom de colonne et les noms attendus
        
        Args:
            col_name: Nom de colonne à évaluer
            expected_names: Liste des noms attendus
            
        Returns:
            Score entre 0 et 1
        """
        col_lower = col_name.lower().strip()
        max_score = 0
        
        for expected in expected_names:
            expected_lower = expected.lower()
            
            # Correspondance exacte
            if col_lower == expected_lower:

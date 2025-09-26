"""
EXCEL_PROCESSOR.PY
Module de traitement des fichiers Excel optimisé pour les données Beeline
Traitement natif des données Excel/CSV avec Pandas - Version adaptée aux colonnes réelles
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
    Processeur de fichiers Excel spécialisé pour les données Beeline réelles
    """
    
    def __init__(self):
        self.setup_logging()
        
        # Configuration des colonnes attendues - adaptée à vos fichiers réels Beeline
        self.expected_columns = {
            'order_number': [
                'N° commande', 'Numero commande', 'Order Number', 'Purchase Order',
                'Commande', 'PO', 'Bon de commande'
            ],
            'cost_center': [
                'Centre de coût', 'Centre de cout', 'Cost Center', 'Code coût',
                'Cost Centre', 'Centre', 'CC'
            ],
            'net_amount': [
                'Montant net à payer au fournisseur', 'Net Amount', 'Montant Net',
                'Amount', 'Total', 'Montant', 'Unités'  # Basé sur votre colonne M
            ],
            'collaborator': [
                'Collaborateur', 'Employee', 'Worker', 'Nom', 'Name',
                'Consultant', 'Contractor'
            ],
            'supplier': [
                'Supplier', 'Fournisseur', 'Vendor', 'Company'
            ],
            'billing_period': [
                'Billing Period', 'Période facturation', 'Period', 'Semaine'
            ],
            'statement_date': [
                'Statement Date', 'Date', 'Invoice Date', 'Billing Date',
                'Date facture', 'Période'
            ],
            'remit_to': [
                'Remit To', 'Remit', 'Adresse', 'Address'
            ],
            'project': [
                'Project', 'Projet', 'Code projet'
            ],
            'code_rubrique': [
                'Code rubrique', 'Rubrique', 'Category Code'
            ],
            'units': [
                'Unités', 'Units', 'Hours', 'Heures', 'Quantity', 'Quantité'
            ],
            'rate': [
                'Taux de facturation', 'Rate', 'Unit Price', 'Prix unitaire', 'Hourly Rate'
            ]
        }
        
        # Configuration de validation
        self.validation_rules = {
            'order_number': r'^[0-9]{10}$',  # Exactement 10 chiffres
            'cost_center': r'^[0-9A-Za-z_]{1,15}$',  # Alphanumériques et underscore
            'amount': r'^-?[0-9]+\.?[0-9]*$',   # Nombre décimal (peut être négatif)
            'date': [
                '%d/%m/%Y', '%Y-%m-%d', '%d-%m-%Y', '%Y/%m/%d',
                '%d/%m/%Y %H:%M:%S', '%Y-%m-%d %H:%M:%S'
            ]
        }
        
        # Seuils de validation adaptés à vos données
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
                'processor_version': '2.1.0'
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
        """
        filename = excel_file.name.lower()
        
        try:
            if filename.endswith('.csv'):
                # Lecture CSV avec détection automatique du délimiteur
                return self.read_csv_with_detection(excel_file)
            
            elif filename.endswith(('.xlsx', '.xls')):
                # Lecture Excel avec gestion d'erreurs
                try:
                    return pd.read_excel(excel_file, engine='openpyxl')
                except:
                    # Fallback pour les anciens formats
                    return pd.read_excel(excel_file, engine=None)
            
            else:
                raise ValueError(f"Format de fichier non supporté: {filename}")
                
        except Exception as e:
            self.logger.error(f"Erreur lecture fichier {filename}: {str(e)}")
            return None
    
    def read_csv_with_detection(self, csv_file) -> pd.DataFrame:
        """
        Lit un CSV avec détection automatique des paramètres
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
                        
                except Exception:
                    continue
        
        # Fallback: lecture basique
        csv_file.seek(0)
        return pd.read_csv(csv_file)
    
    def clean_dataframe(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Nettoie et normalise un DataFrame
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
        """
        # Détection par nom de colonne
        amount_keywords = ['montant', 'amount', 'total', 'price', 'cost', 'taux', 'rate', 'unités']
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
        """
        mapping = {}
        
        for standard_name, possible_names in self.expected_columns.items():
            best_match = None
            best_score = 0
            
            for col_name in column_names:
                if col_name in [v for v in mapping.keys()]:  # Déjà mappée
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
        """
        col_lower = col_name.lower().strip()
        max_score = 0
        
        for expected in expected_names:
            expected_lower = expected.lower()
            
            # Correspondance exacte
            if col_lower == expected_lower:
                return 1.0
            
            # Correspondance partielle
            if expected_lower in col_lower or col_lower in expected_lower:
                partial_score = min(len(expected_lower), len(col_lower)) / max(len(expected_lower), len(col_lower))
                max_score = max(max_score, partial_score * 0.9)
            
            # Correspondance par mots-clés
            expected_words = expected_lower.split()
            col_words = col_lower.split()
            
            matching_words = sum(1 for word in expected_words if word in col_words)
            if matching_words > 0:
                word_score = matching_words / max(len(expected_words), len(col_words))
                max_score = max(max_score, word_score * 0.8)
        
        return max_score
    
    def rename_columns(self, df: pd.DataFrame, mapping: Dict[str, str]) -> pd.DataFrame:
        """
        Renomme les colonnes du DataFrame selon le mapping
        """
        df_renamed = df.copy()
        df_renamed = df_renamed.rename(columns=mapping)
        
        # Log des mappings effectués
        for original, standard in mapping.items():
            self.logger.info(f"Colonne mappée: '{original}' -> '{standard}'")
        
        return df_renamed
    
    def validate_dataframe(self, df: pd.DataFrame) -> Dict[str, Any]:
        """
        Valide un DataFrame complet
        """
        validation_results = {
            'total_rows': len(df),
            'valid_rows': 0,
            'invalid_rows': 0,
            'validation_errors': [],
            'field_validations': {},
            'data_quality_score': 0
        }
        
        # Validation ligne par ligne
        valid_count = 0
        for index, row in df.iterrows():
            row_validation = self.validate_row(row, index)
            if row_validation['is_valid']:
                valid_count += 1
            else:
                validation_results['validation_errors'].extend(row_validation['errors'])
        
        validation_results['valid_rows'] = valid_count
        validation_results['invalid_rows'] = len(df) - valid_count
        
        # Score de qualité global
        if len(df) > 0:
            validation_results['data_quality_score'] = (valid_count / len(df)) * 100
        
        # Validation des colonnes critiques
        critical_columns = ['order_number', 'net_amount']
        for col in critical_columns:
            if col in df.columns:
                col_validation = self.validate_column(df[col], col)
                validation_results['field_validations'][col] = col_validation
        
        return validation_results
    
    def validate_row(self, row: pd.Series, row_index: int) -> Dict[str, Any]:
        """
        Valide une ligne individuelle
        """
        validation = {
            'is_valid': True,
            'errors': [],
            'warnings': []
        }
        
        # Validation du numéro de commande (critique)
        if 'order_number' in row:
            order_validation = self.validate_order_number(row['order_number'])
            if not order_validation['valid']:
                validation['is_valid'] = False
                validation['errors'].extend([f"Ligne {row_index}: {err}" for err in order_validation['errors']])
        
        # Validation du montant (important mais pas critique)
        if 'net_amount' in row:
            amount_validation = self.validate_amount(row['net_amount'])
            if not amount_validation['valid']:
                validation['warnings'].extend([f"Ligne {row_index}: {err}" for err in amount_validation['errors']])
        
        # Validation de la date
        if 'statement_date' in row and pd.notna(row['statement_date']):
            date_validation = self.validate_date_field(row['statement_date'])
            if not date_validation['valid']:
                validation['warnings'].extend([f"Ligne {row_index}: {err}" for err in date_validation['errors']])
        
        return validation
    
    def validate_column(self, series: pd.Series, column_name: str) -> Dict[str, Any]:
        """
        Valide une colonne complète
        """
        validation = {
            'column_name': column_name,
            'total_values': len(series),
            'valid_values': 0,
            'invalid_values': 0,
            'null_values': series.isna().sum(),
            'unique_values': series.nunique(),
            'validation_rate': 0
        }
        
        valid_count = 0
        for value in series.dropna():
            if column_name == 'order_number':
                is_valid = self.validate_order_number(value)['valid']
            elif column_name in ['net_amount', 'units']:
                is_valid = self.validate_amount(value)['valid']
            elif column_name == 'statement_date':
                is_valid = self.validate_date_field(value)['valid']
            else:
                is_valid = True  # Pas de validation spécifique
            
            if is_valid:
                valid_count += 1
        
        validation['valid_values'] = valid_count
        validation['invalid_values'] = len(series.dropna()) - valid_count
        
        if len(series.dropna()) > 0:
            validation['validation_rate'] = (valid_count / len(series.dropna())) * 100
        
        return validation
    
    def validate_order_number(self, order_number: Any) -> Dict[str, Any]:
        """Valide un numéro de commande"""
        validation = {'valid': True, 'errors': []}
        
        if pd.isna(order_number):
            validation['valid'] = False
            validation['errors'].append("Numéro de commande manquant")
            return validation
        
        order_str = str(order_number).strip()
        
        # Suppression des points décimaux pour les nombres Excel (5600002101.0 -> 5600002101)
        if '.' in order_str and order_str.replace('.', '').replace('0', '').isdigit():
            order_str = order_str.split('.')[0]
        
        # Doit être numérique et exactement 10 chiffres
        if not re.match(self.validation_rules['order_number'], order_str):
            validation['valid'] = False
            validation['errors'].append(f"Format numéro de commande invalide: {order_str} (attendu: 10 chiffres)")
        
        return validation
    
    def validate_amount(self, amount: Any) -> Dict[str, Any]:
        """Valide un montant"""
        validation = {'valid': True, 'errors': []}
        
        if pd.isna(amount):
            validation['valid'] = False
            validation['errors'].append("Montant manquant")
            return validation
        
        try:
            amount_float = float(amount)
            
            # Vérification des seuils - tolérance élargie pour vos données
            if amount_float < self.thresholds['min_amount']:
                validation['errors'].append(f"Montant trop faible: {amount_float}")
            elif amount_float > self.thresholds['max_amount']:
                validation['errors'].append(f"Montant trop élevé: {amount_float}")
            # Accepter les montants à 0 (contrairement à la version précédente)
                
        except (ValueError, TypeError):
            validation['valid'] = False
            validation['errors'].append(f"Montant non numérique: {amount}")
        
        return validation
    
    def validate_date_field(self, date_value: Any) -> Dict[str, Any]:
        """Valide une date"""
        validation = {'valid': True, 'errors': []}
        
        if pd.isna(date_value):
            return validation  # Date optionnelle
        
        date_str = str(date_value).strip()
        
        # Tentative de parsing avec différents formats
        for date_format in self.validation_rules['date']:
            try:
                parsed_date = datetime.strptime(date_str, date_format)
                
                # Vérification de plausibilité
                now = datetime.now()
                if parsed_date > now:
                    validation['errors'].append(f"Date dans le futur: {date_str}")
                elif parsed_date.year < 2020:
                    validation['errors'].append(f"Date trop ancienne: {date_str}")
                
                return validation  # Date valide trouvée
                
            except ValueError:
                continue
        
        # Aucun format reconnu
        validation['valid'] = False
        validation['errors'].append(f"Format de date non reconnu: {date_str}")
        
        return validation
    
    def dataframe_to_records(self, df: pd.DataFrame, filename: str, 
                           validation_results: Dict[str, Any]) -> List[Dict[str, Any]]:
        """
        Convertit un DataFrame en liste de dictionnaires
        """
        records = []
        
        for index, row in df.iterrows():
            record = {}
            
            # Données de base
            for col in df.columns:
                value = row[col]
                if pd.isna(value):
                    record[col] = None
                else:
                    # Nettoyage spécial pour les numéros de commande
                    if col == 'order_number' and isinstance(value, (int, float)):
                        record[col] = str(int(value))  # Conversion 5600002101.0 -> "5600002101"
                    else:
                        record[col] = value
            
            # Métadonnées de la ligne
            record['source_filename'] = filename
            record['row_index'] = index
            record['processing_timestamp'] = datetime.now().isoformat()
            
            # Validation de la ligne
            row_validation = self.validate_row(row, index)
            record['is_valid'] = row_validation['is_valid']
            record['validation_errors'] = row_validation['errors']
            record['validation_warnings'] = row_validation['warnings']
            
            records.append(record)
        
        return records
    
    def aggregate_by_order_number(self, excel_data: List[Dict[str, Any]]) -> Dict[str, Dict[str, Any]]:
        """
        Agrège les données Excel par numéro de commande
        """
        aggregated = {}
        
        for record in excel_data:
            if not record.get('is_valid', False):
                continue
            
            order_number = record.get('order_number')
            if not order_number:
                continue
            
            order_key = str(order_number).strip()
            
            if order_key not in aggregated:
                aggregated[order_key] = {
                    'order_number': order_key,
                    'total_amount': 0,
                    'line_count': 0,
                    'collaborators': [],
                    'cost_centers': [],  # Nouveau pour vos données
                    'statement_dates': [],
                    'suppliers': [],
                    'source_files': [],
                    'raw_lines': [],
                    'projects': [],  # Nouveau
                    'validation_summary': {
                        'total_lines': 0,
                        'valid_lines': 0,
                        'invalid_lines': 0,
                        'validity_rate': 0
                    }
                }
            
            order_data = aggregated[order_key]
            
            # Agrégation des montants
            net_amount = record.get('net_amount', 0)
            if isinstance(net_amount, (int, float)) and not pd.isna(net_amount):
                order_data['total_amount'] += net_amount
            
            # Comptage des lignes
            order_data['line_count'] += 1
            
            # Collecte des collaborateurs
            collaborator = record.get('collaborator')
            if collaborator and collaborator not in order_data['collaborators']:
                order_data['collaborators'].append(str(collaborator).strip())
            
            # Collecte des centres de coût (nouveau)
            cost_center = record.get('cost_center')
            if cost_center and cost_center not in order_data['cost_centers']:
                order_data['cost_centers'].append(str(cost_center).strip())
            
            # Collecte des dates
            statement_date = record.get('statement_date')
            if statement_date and statement_date not in order_data['statement_dates']:
                order_data['statement_dates'].append(statement_date)
            
            # Collecte des fournisseurs
            supplier = record.get('supplier')
            if supplier and supplier not in order_data['suppliers']:
                order_data['suppliers'].append(str(supplier).strip())
            
            # Collecte des projets (nouveau)
            project = record.get('project')
            if project and project not in order_data['projects']:
                order_data['projects'].append(str(project).strip())
            
            # Fichiers sources
            source_file = record.get('source_filename')
            if source_file and source_file not in order_data['source_files']:
                order_data['source_files'].append(source_file)
            
            # Lignes brutes pour debug
            order_data['raw_lines'].append(record)
            
            # Statistiques de validation
            order_data['validation_summary']['total_lines'] += 1
            if record.get('is_valid', False):
                order_data['validation_summary']['valid_lines'] += 1
            else:
                order_data['validation_summary']['invalid_lines'] += 1
        
        # Calcul des taux de validité
        for order_data in aggregated.values():
            total = order_data['validation_summary']['total_lines']
            valid = order_data['validation_summary']['valid_lines']
            
            if total > 0:
                order_data['validation_summary']['validity_rate'] = (valid / total) * 100
        
        self.logger.info(f"Agrégation terminée: {len(aggregated)} commandes uniques")
        return aggregated
    
    def get_processing_summary(self, processed_data: List[Dict[str, Any]]) -> Dict[str, Any]:
        """
        Génère un résumé du traitement Excel
        """
        total_records = len(processed_data)
        valid_records = sum(1 for r in processed_data if r.get('is_valid', False))
        
        # Analyse des erreurs communes
        error_types = {}
        for record in processed_data:
            for error in record.get('validation_errors', []):
                error_type = error.split(':')[1].strip() if ':' in error else error
                error_types[error_type] = error_types.get(error_type, 0) + 1
        
        # Analyse des colonnes trouvées
        columns_found = set()
        for record in processed_data:
            columns_found.update([k for k in record.keys() if not k.startswith(('source_', 'row_', 'processing_', 'is_', 'validation_'))])
        
        return {
            'total_records': total_records,
            'valid_records': valid_records,
            'invalid_records': total_records - valid_records,
            'validity_rate': (valid_records / total_records) * 100 if total_records > 0 else 0,
            'columns_found': list(columns_found),
            'common_errors': error_types,
            'processing_timestamp': datetime.now().isoformat()
        }

# Fonctions utilitaires pour les tests
def test_excel_processor():
    """Fonction de test pour le module Excel"""
    processor = ExcelProcessor()
    
    # Test de mapping de colonnes avec vos données réelles
    test_columns = ['N° commande', 'Centre de coût', 'Collaborateur', 'Supplier', 'Unités']
    mapping = processor.map_columns(test_columns)
    
    print("Test mapping colonnes Beeline:")
    for original, mapped in mapping.items():
        print(f"  '{original}' -> '{mapped}'")
    
    # Test de validation d'un numéro de commande
    test_orders = ["5600002101", "5600002101.0", "invalid"]
    for test_order in test_orders:
        validation = processor.validate_order_number(test_order)
        print(f"Test validation commande '{test_order}': {validation['valid']}")
    
    # Test de nettoyage d'un montant
    test_amounts = ["1.234,56", "0,75", "0", "-10.5"]
    for test_amount in test_amounts:
        cleaned_series = processor.clean_amount_column(pd.Series([test_amount]))
        print(f"Test montant '{test_amount}' -> {cleaned_series.iloc[0]}")
    
    return mapping

if __name__ == "__main__":
    # Test du module si exécuté directement
    test_excel_processor()

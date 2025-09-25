"""
PDF_EXTRACTOR.PY
Module d'extraction des données PDF pour l'application Beeline
Extraction native des factures PDF Randstad
"""

import pdfplumber
import re
import logging
from typing import Dict, List, Optional, Any
from datetime import datetime
import streamlit as st

class PDFExtractor:
    """
    Extracteur de données PDF spécialisé pour les factures Randstad/Beeline
    """
    
    def __init__(self):
        self.setup_logging()
        
        # Patterns d'extraction spécifiques aux factures Randstad
        self.patterns = {
            'invoice_id': [
                r'Invoice\s+ID[\/\s]*Number[:\s]*([A-Z0-9]+)',
                r'Invoice\s+ID[:\s]*([A-Z0-9]+)',
                r'Facture\s*n°[:\s]*([A-Z0-9]+)',
                r'N°\s*facture[:\s]*([A-Z0-9]+)'
            ],
            'purchase_order': [
                r'Purchase\s+Order[:\s]*([0-9]{10})',
                r'Bon\s+de\s+commande[:\s]*([0-9]{10})',
                r'N°\s*commande[:\s]*([0-9]{10})',
                r'Order[:\s]*([0-9]{10})'
            ],
            'invoice_date': [
                r'Invoice\s+Date[:\s]*([0-9]{4}[\/\-][0-9]{2}[\/\-][0-9]{2})',
                r'Date\s+facture[:\s]*([0-9]{4}[\/\-][0-9]{2}[\/\-][0-9]{2})',
                r'Date[:\s]*([0-9]{4}[\/\-][0-9]{2}[\/\-][0-9]{2})'
            ],
            'supplier': [
                r'Facture\s+émise\s+par\s*[:]\s*([^\n\r]+)',
                r'Supplier[:\s]*([^\n\r]+)',
                r'Fournisseur[:\s]*([^\n\r]+)'
            ],
            'client': [
                r'Mars\s+Wrigley[^\n\r]*',
                r'Au\s+nom\s+et\s+pour\s+le\s+compte\s+de[:\s]*([^\n\r]+)'
            ],
            'total_net': [
                r'Invoice\s+Total.*?EUR[:\s]*([0-9,]+\.?[0-9]*)',
                r'Montant\s*Net[:\s]*([0-9,]+\.?[0-9]*)',
                r'Total\s*HT[:\s]*([0-9,]+\.?[0-9]*)',
                r'Net\s*Amount[:\s]*([0-9,]+\.?[0-9]*)'
            ],
            'total_vat': [
                r'TVA[:\s]*([0-9,]+\.?[0-9]*)',
                r'VAT[:\s]*([0-9,]+\.?[0-9]*)',
                r'Tax[:\s]*([0-9,]+\.?[0-9]*)'
            ],
            'total_gross': [
                r'Total\s*TTC[:\s]*([0-9,]+\.?[0-9]*)',
                r'Gross\s*Total[:\s]*([0-9,]+\.?[0-9]*)',
                r'Invoice\s+Total.*?EUR.*?[0-9,]+\.?[0-9]*\s+[0-9,]+\.?[0-9]*\s+([0-9,]+\.?[0-9]*)'
            ]
        }
        
        # Validation des formats
        self.validation_rules = {
            'purchase_order': r'^[0-9]{10}$',
            'invoice_id': r'^[A-Z0-9]{6,}$',
            'amount': r'^[0-9]+(\.[0-9]{2})?$'
        }
    
    def setup_logging(self):
        """Configure le logging pour le module"""
        logging.basicConfig(level=logging.INFO)
        self.logger = logging.getLogger(__name__)
    
    def extract_single_pdf(self, pdf_file) -> Dict[str, Any]:
        """
        Extrait les données d'un seul fichier PDF
        
        Args:
            pdf_file: Fichier PDF uploadé (Streamlit UploadedFile)
            
        Returns:
            Dict contenant les données extraites
        """
        try:
            self.logger.info(f"Début extraction PDF: {pdf_file.name}")
            
            # Lecture du PDF avec pdfplumber
            with pdfplumber.open(pdf_file) as pdf:
                
                # Extraction du texte de toutes les pages
                full_text = ""
                tables_data = []
                
                for page_num, page in enumerate(pdf.pages):
                    # Texte de la page
                    page_text = page.extract_text()
                    if page_text:
                        full_text += f"\n--- Page {page_num + 1} ---\n" + page_text
                    
                    # Tableaux de la page
                    tables = page.extract_tables()
                    if tables:
                        tables_data.extend(tables)
                
                # Parsing des données
                extracted_data = self.parse_pdf_content(full_text, pdf_file.name, tables_data)
                
                # Validation des données extraites
                validation_results = self.validate_extracted_data(extracted_data)
                extracted_data['validation'] = validation_results
                
                # Métadonnées
                extracted_data['extraction_metadata'] = {
                    'filename': pdf_file.name,
                    'file_size': pdf_file.size,
                    'pages_count': len(pdf.pages),
                    'extraction_timestamp': datetime.now().isoformat(),
                    'extractor_version': '2.0.0',
                    'text_length': len(full_text),
                    'tables_count': len(tables_data)
                }
                
                self.logger.info(f"Extraction réussie: {pdf_file.name}")
                return extracted_data
                
        except Exception as e:
            self.logger.error(f"Erreur extraction PDF {pdf_file.name}: {str(e)}")
            return {
                'success': False,
                'error': str(e),
                'filename': pdf_file.name,
                'extraction_timestamp': datetime.now().isoformat()
            }
    
    def parse_pdf_content(self, text: str, filename: str, tables: List = None) -> Dict[str, Any]:
        """
        Parse le contenu texte du PDF pour extraire les données structurées
        
        Args:
            text: Texte complet du PDF
            filename: Nom du fichier
            tables: Tableaux extraits (optionnel)
            
        Returns:
            Dict avec les données parsées
        """
        extracted_data = {
            'success': True,
            'filename': filename,
            'raw_text_sample': text[:500] if text else "",  # Échantillon pour debug
        }
        
        # Extraction des champs principaux
        for field_name, patterns in self.patterns.items():
            extracted_value = self.extract_field_with_patterns(text, patterns, field_name)
            extracted_data[field_name] = extracted_value
        
        # Post-traitement des montants
        extracted_data = self.post_process_amounts(extracted_data)
        
        # Extraction des détails de facturation (depuis les tableaux)
        if tables:
            billing_details = self.extract_billing_details(tables)
            extracted_data['billing_details'] = billing_details
        
        # Calcul de la complétude des données
        extracted_data['data_completeness'] = self.calculate_data_completeness(extracted_data)
        
        return extracted_data
    
    def extract_field_with_patterns(self, text: str, patterns: List[str], field_name: str) -> Optional[str]:
        """
        Extrait un champ en utilisant plusieurs patterns de regex
        
        Args:
            text: Texte à analyser
            patterns: Liste des patterns regex à tester
            field_name: Nom du champ pour les logs
            
        Returns:
            Valeur extraite ou None
        """
        if not text:
            return None
        
        for pattern in patterns:
            try:
                match = re.search(pattern, text, re.IGNORECASE | re.MULTILINE)
                if match:
                    value = match.group(1).strip()
                    
                    # Nettoyage de la valeur
                    value = self.clean_extracted_value(value, field_name)
                    
                    if value:  # Si la valeur n'est pas vide après nettoyage
                        self.logger.debug(f"Champ {field_name} extrait: {value}")
                        return value
                        
            except re.error as e:
                self.logger.warning(f"Erreur regex pour {field_name}: {e}")
                continue
        
        self.logger.warning(f"Champ {field_name} non trouvé dans {text[:100]}...")
        return None
    
    def clean_extracted_value(self, value: str, field_type: str) -> str:
        """
        Nettoie une valeur extraite selon son type
        
        Args:
            value: Valeur à nettoyer
            field_type: Type du champ
            
        Returns:
            Valeur nettoyée
        """
        if not value:
            return ""
        
        # Nettoyage général
        value = value.strip().replace('\n', ' ').replace('\r', ' ')
        value = re.sub(r'\s+', ' ', value)  # Multiples espaces -> un seul
        
        # Nettoyage spécifique selon le type
        if 'amount' in field_type or 'total' in field_type:
            # Nettoyage des montants: garder seulement chiffres, virgules, points
            value = re.sub(r'[^\d,.]', '', value)
            # Normalisation: virgule -> point pour les décimales
            if ',' in value and '.' not in value:
                value = value.replace(',', '.')
            elif ',' in value and '.' in value:
                # Format européen: 1.234,56 -> 1234.56
                if value.rfind(',') > value.rfind('.'):
                    value = value.replace('.', '').replace(',', '.')
        
        elif field_type == 'purchase_order':
            # Numéros de commande: que des chiffres
            value = re.sub(r'[^\d]', '', value)
        
        elif field_type == 'invoice_date':
            # Dates: normalisation du format
            value = re.sub(r'[^\d\/\-]', '', value)
        
        elif field_type in ['supplier', 'client']:
            # Texte: nettoyage des caractères spéciaux en trop
            value = re.sub(r'[^\w\s\-\.,&]', '', value)
        
        return value.strip()
    
    def post_process_amounts(self, data: Dict[str, Any]) -> Dict[str, Any]:
        """
        Post-traite les montants extraits pour les convertir en float
        
        Args:
            data: Données extraites
            
        Returns:
            Données avec montants convertis
        """
        amount_fields = ['total_net', 'total_vat', 'total_gross']
        
        for field in amount_fields:
            if field in data and data[field]:
                try:
                    # Conversion en float
                    amount_str = str(data[field])
                    amount_float = self.parse_amount_string(amount_str)
                    
                    # Stockage des deux formats
                    data[f"{field}_original"] = data[field]  # Valeur originale
                    data[field] = amount_float  # Valeur numérique
                    
                except (ValueError, TypeError) as e:
                    self.logger.warning(f"Erreur conversion montant {field}: {e}")
                    data[f"{field}_error"] = str(e)
        
        return data
    
    def parse_amount_string(self, amount_str: str) -> float:
        """
        Convertit une chaîne de montant en float
        
        Args:
            amount_str: Chaîne représentant un montant
            
        Returns:
            Montant en float
        """
        if not amount_str:
            return 0.0
        
        # Nettoyage et normalisation
        clean_amount = str(amount_str).strip()
        clean_amount = re.sub(r'[^\d,.]', '', clean_amount)
        
        if not clean_amount:
            return 0.0
        
        # Gestion des formats européens vs anglo-saxons
        if ',' in clean_amount and '.' in clean_amount:
            # Format européen: 1.234,56
            if clean_amount.rfind(',') > clean_amount.rfind('.'):
                clean_amount = clean_amount.replace('.', '').replace(',', '.')
            # Format US: 1,234.56 (déjà correct)
        elif ',' in clean_amount and '.' not in clean_amount:
            # Peut être 1234,56 (européen) ou 1,234 (US sans décimales)
            comma_pos = clean_amount.rfind(',')
            after_comma = clean_amount[comma_pos + 1:]
            
            if len(after_comma) <= 2:  # Probablement des décimales
                clean_amount = clean_amount.replace(',', '.')
            else:  # Probablement un séparateur de milliers
                clean_amount = clean_amount.replace(',', '')
        
        try:
            return float(clean_amount)
        except ValueError:
            raise ValueError(f"Impossible de convertir '{amount_str}' en montant")
    
    def extract_billing_details(self, tables: List) -> List[Dict[str, Any]]:
        """
        Extrait les détails de facturation depuis les tableaux
        
        Args:
            tables: Liste des tableaux extraits du PDF
            
        Returns:
            Liste des lignes de facturation
        """
        billing_details = []
        
        for table in tables:
            if not table or len(table) < 2:  # Ignorer les tableaux vides ou sans données
                continue
            
            # Détection de l'en-tête du tableau de facturation
            header_row = table[0] if table else []
            header_text = ' '.join([str(cell) if cell else '' for cell in header_row]).lower()
            
            # Vérifier si c'est un tableau de facturation
            if any(keyword in header_text for keyword in ['description', 'montant', 'amount', 'quantity', 'hours']):
                
                # Parsing des lignes de données
                for row in table[1:]:  # Ignorer l'en-tête
                    if not row or all(cell is None or str(cell).strip() == '' for cell in row):
                        continue
                    
                    # Extraction des données de la ligne
                    line_data = {}
                    
                    # Mapping basique des colonnes (peut être amélioré)
                    for i, cell in enumerate(row):
                        if cell is not None and str(cell).strip():
                            if i == 0:  # Première colonne généralement description
                                line_data['description'] = str(cell).strip()
                            elif self.is_amount_cell(cell):  # Détection automatique des montants
                                amount_key = f'amount_{len([k for k in line_data.keys() if k.startswith("amount")])}'
                                line_data[amount_key] = self.parse_amount_string(str(cell))
                            elif self.is_quantity_cell(cell):  # Détection des quantités
                                line_data['quantity'] = self.parse_amount_string(str(cell))
                            else:
                                line_data[f'column_{i}'] = str(cell).strip()
                    
                    if line_data:  # Si on a extrait des données
                        billing_details.append(line_data)
        
        return billing_details
    
    def is_amount_cell(self, cell) -> bool:
        """Détermine si une cellule contient un montant"""
        if not cell:
            return False
        
        cell_str = str(cell).strip()
        # Pattern pour détecter les montants
        amount_pattern = r'^[0-9,.\s]+
        return bool(re.match(amount_pattern, cell_str)) and (',' in cell_str or '.' in cell_str)
    
    def is_quantity_cell(self, cell) -> bool:
        """Détermine si une cellule contient une quantité"""
        if not cell:
            return False
        
        cell_str = str(cell).strip()
        # Pattern pour quantités simples
        quantity_pattern = r'^[0-9]+(\.[0-9]{1,2})?
        return bool(re.match(quantity_pattern, cell_str))
    
    def calculate_data_completeness(self, data: Dict[str, Any]) -> Dict[str, Any]:
        """
        Calcule le pourcentage de complétude des données extraites
        
        Args:
            data: Données extraites
            
        Returns:
            Statistiques de complétude
        """
        # Champs critiques pour le rapprochement
        critical_fields = ['purchase_order', 'total_net']
        important_fields = ['invoice_id', 'invoice_date', 'supplier']
        optional_fields = ['total_vat', 'total_gross', 'client']
        
        completeness = {
            'critical_score': 0,
            'important_score': 0,
            'optional_score': 0,
            'overall_score': 0,
            'missing_critical': [],
            'missing_important': [],
            'processable': False
        }
        
        # Évaluation des champs critiques
        critical_found = 0
        for field in critical_fields:
            if data.get(field):
                critical_found += 1
            else:
                completeness['missing_critical'].append(field)
        
        completeness['critical_score'] = (critical_found / len(critical_fields)) * 100
        
        # Évaluation des champs importants
        important_found = 0
        for field in important_fields:
            if data.get(field):
                important_found += 1
            else:
                completeness['missing_important'].append(field)
        
        completeness['important_score'] = (important_found / len(important_fields)) * 100
        
        # Évaluation des champs optionnels
        optional_found = sum(1 for field in optional_fields if data.get(field))
        completeness['optional_score'] = (optional_found / len(optional_fields)) * 100
        
        # Score global pondéré
        completeness['overall_score'] = (
            completeness['critical_score'] * 0.6 +  # 60% pour les critiques
            completeness['important_score'] * 0.3 +  # 30% pour les importants
            completeness['optional_score'] * 0.1     # 10% pour les optionnels
        )
        
        # Détermination si les données sont processables
        completeness['processable'] = (
            completeness['critical_score'] >= 100 and  # Tous les champs critiques
            completeness['important_score'] >= 50      # Au moins 50% des importants
        )
        
        return completeness
    
    def validate_extracted_data(self, data: Dict[str, Any]) -> Dict[str, Any]:
        """
        Valide les données extraites selon les règles métier
        
        Args:
            data: Données extraites à valider
            
        Returns:
            Résultats de validation
        """
        validation_results = {
            'is_valid': True,
            'errors': [],
            'warnings': [],
            'field_validations': {}
        }
        
        # Validation du numéro de commande
        if data.get('purchase_order'):
            po_validation = self.validate_purchase_order(data['purchase_order'])
            validation_results['field_validations']['purchase_order'] = po_validation
            if not po_validation['valid']:
                validation_results['errors'].extend(po_validation['errors'])
                validation_results['is_valid'] = False
        
        # Validation de l'ID facture
        if data.get('invoice_id'):
            id_validation = self.validate_invoice_id(data['invoice_id'])
            validation_results['field_validations']['invoice_id'] = id_validation
            if not id_validation['valid']:
                validation_results['warnings'].extend(id_validation['errors'])
        
        # Validation des montants
        amount_validation = self.validate_amounts(data)
        validation_results['field_validations']['amounts'] = amount_validation
        if not amount_validation['valid']:
            validation_results['errors'].extend(amount_validation['errors'])
            validation_results['is_valid'] = False
        
        # Validation de la date
        if data.get('invoice_date'):
            date_validation = self.validate_date(data['invoice_date'])
            validation_results['field_validations']['invoice_date'] = date_validation
            if not date_validation['valid']:
                validation_results['warnings'].extend(date_validation['errors'])
        
        # Validation du fournisseur
        if data.get('supplier'):
            supplier_validation = self.validate_supplier(data['supplier'])
            validation_results['field_validations']['supplier'] = supplier_validation
            if not supplier_validation['valid']:
                validation_results['warnings'].extend(supplier_validation['errors'])
        
        return validation_results
    
    def validate_purchase_order(self, po: str) -> Dict[str, Any]:
        """Valide un numéro de commande"""
        validation = {'valid': True, 'errors': []}
        
        if not po:
            validation['valid'] = False
            validation['errors'].append("Numéro de commande vide")
            return validation
        
        po_str = str(po).strip()
        
        # Doit contenir exactement 10 chiffres
        if not re.match(r'^[0-9]{10}, po_str):
            validation['valid'] = False
            validation['errors'].append(f"Format numéro de commande incorrect: {po_str} (attendu: 10 chiffres)")
        
        return validation
    
    def validate_invoice_id(self, invoice_id: str) -> Dict[str, Any]:
        """Valide un ID de facture"""
        validation = {'valid': True, 'errors': []}
        
        if not invoice_id:
            validation['valid'] = False
            validation['errors'].append("ID facture vide")
            return validation
        
        id_str = str(invoice_id).strip()
        
        # Doit contenir au moins 4 caractères alphanumériques
        if len(id_str) < 4 or not re.match(r'^[A-Z0-9]+, id_str):
            validation['valid'] = False
            validation['errors'].append(f"Format ID facture suspect: {id_str}")
        
        return validation
    
    def validate_amounts(self, data: Dict[str, Any]) -> Dict[str, Any]:
        """Valide les montants extraits"""
        validation = {'valid': True, 'errors': []}
        
        # Vérification du montant net (critique)
        total_net = data.get('total_net')
        if not total_net or total_net <= 0:
            validation['valid'] = False
            validation['errors'].append("Montant net manquant ou invalide")
        elif total_net > 100000:  # Plus de 100k euros, suspect
            validation['errors'].append(f"Montant net très élevé: {total_net}€")
        
        # Cohérence entre HT, TVA et TTC
        total_vat = data.get('total_vat', 0)
        total_gross = data.get('total_gross', 0)
        
        if total_net and total_vat and total_gross:
            expected_gross = total_net + total_vat
            difference = abs(total_gross - expected_gross)
            
            if difference > 0.02:  # Tolérance de 2 centimes
                validation['errors'].append(
                    f"Incohérence montants: HT({total_net}) + TVA({total_vat}) ≠ TTC({total_gross})"
                )
        
        return validation
    
    def validate_date(self, date_str: str) -> Dict[str, Any]:
        """Valide une date de facture"""
        validation = {'valid': True, 'errors': []}
        
        if not date_str:
            validation['valid'] = False
            validation['errors'].append("Date de facture manquante")
            return validation
        
        # Tentative de parsing de la date
        date_formats = ['%Y/%m/%d', '%d/%m/%Y', '%Y-%m-%d', '%d-%m-%Y']
        parsed_date = None
        
        for fmt in date_formats:
            try:
                parsed_date = datetime.strptime(str(date_str).strip(), fmt)
                break
            except ValueError:
                continue
        
        if not parsed_date:
            validation['valid'] = False
            validation['errors'].append(f"Format de date non reconnu: {date_str}")
        else:
            # Vérification de la plausibilité de la date
            now = datetime.now()
            
            if parsed_date > now:
                validation['errors'].append(f"Date de facture dans le futur: {date_str}")
            elif parsed_date.year < 2020:
                validation['errors'].append(f"Date de facture très ancienne: {date_str}")
        
        return validation
    
    def validate_supplier(self, supplier: str) -> Dict[str, Any]:
        """Valide le nom du fournisseur"""
        validation = {'valid': True, 'errors': []}
        
        if not supplier:
            validation['valid'] = False
            validation['errors'].append("Nom fournisseur manquant")
            return validation
        
        supplier_str = str(supplier).strip().lower()
        
        # Pour Beeline, on s'attend principalement à Randstad
        expected_suppliers = ['randstad', 'select t.t.']
        
        if not any(expected in supplier_str for expected in expected_suppliers):
            validation['errors'].append(f"Fournisseur non reconnu: {supplier}")
        
        return validation
    
    def extract_multiple_pdfs(self, pdf_files: List) -> List[Dict[str, Any]]:
        """
        Extrait les données de plusieurs fichiers PDF
        
        Args:
            pdf_files: Liste des fichiers PDF
            
        Returns:
            Liste des résultats d'extraction
        """
        results = []
        
        for i, pdf_file in enumerate(pdf_files):
            try:
                self.logger.info(f"Traitement PDF {i+1}/{len(pdf_files)}: {pdf_file.name}")
                
                # Mise à jour de la progress bar si dans Streamlit
                if hasattr(st, 'progress'):
                    progress = (i + 1) / len(pdf_files)
                    # Note: La progress bar sera gérée depuis l'app principale
                
                result = self.extract_single_pdf(pdf_file)
                result['batch_index'] = i
                results.append(result)
                
            except Exception as e:
                self.logger.error(f"Erreur traitement PDF {pdf_file.name}: {str(e)}")
                results.append({
                    'success': False,
                    'error': str(e),
                    'filename': pdf_file.name,
                    'batch_index': i
                })
        
        return results
    
    def get_extraction_summary(self, extraction_results: List[Dict[str, Any]]) -> Dict[str, Any]:
        """
        Génère un résumé des extractions
        
        Args:
            extraction_results: Résultats des extractions
            
        Returns:
            Résumé statistique
        """
        total_files = len(extraction_results)
        successful_extractions = sum(1 for r in extraction_results if r.get('success', False))
        processable_files = sum(1 for r in extraction_results 
                               if r.get('data_completeness', {}).get('processable', False))
        
        # Analyse des erreurs communes
        error_types = {}
        for result in extraction_results:
            if not result.get('success', True):
                error = result.get('error', 'Erreur inconnue')
                error_types[error] = error_types.get(error, 0) + 1
        
        # Analyse de la complétude moyenne
        completeness_scores = [
            r.get('data_completeness', {}).get('overall_score', 0)
            for r in extraction_results
            if r.get('success', False)
        ]
        
        avg_completeness = sum(completeness_scores) / len(completeness_scores) if completeness_scores else 0
        
        return {
            'total_files': total_files,
            'successful_extractions': successful_extractions,
            'failed_extractions': total_files - successful_extractions,
            'processable_files': processable_files,
            'success_rate': (successful_extractions / total_files) * 100 if total_files > 0 else 0,
            'processable_rate': (processable_files / total_files) * 100 if total_files > 0 else 0,
            'average_completeness': avg_completeness,
            'common_errors': error_types,
            'extraction_timestamp': datetime.now().isoformat()
        }

# Fonctions utilitaires pour les tests
def test_pdf_extractor():
    """Fonction de test pour le module PDF"""
    extractor = PDFExtractor()
    
    # Test des patterns
    test_text = """
    Invoice ID: 5118S0004
    Purchase Order: 5600013960
    Invoice Date: 2025/09/22
    Total Net Amount: 1,014.20 EUR
    Supplier: Randstad France
    """
    
    result = extractor.parse_pdf_content(test_text, "test.pdf")
    print("Test extraction:")
    for key, value in result.items():
        if not key.startswith('raw_text'):
            print(f"  {key}: {value}")
    
    return result

if __name__ == "__main__":
    # Test du module si exécuté directement
    test_pdf_extractor()

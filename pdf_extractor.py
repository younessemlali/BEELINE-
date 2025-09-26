"""
PDF_EXTRACTOR.PY - VERSION CORRIGÉE
Corrections pour la gestion des formats de montants européens (1,234.56 vs 1.234,56)
"""

import pdfplumber
import re
import logging
from typing import Dict, List, Optional, Any
from datetime import datetime
import streamlit as st

class PDFExtractor:
    """
    Extracteur de données PDF corrigé pour les formats européens
    """
    
    def __init__(self):
        self.setup_logging()
        
        # Patterns d'extraction optimisés
        self.patterns = {
            'invoice_id': [
                r'Invoice\s+ID/Number[^A-Z0-9]*([A-Z0-9]+)',
                r'Numéro[^A-Z0-9]*([A-Z0-9]+)',
                r'(\d{4}S\d{4})',  # Pattern spécifique 4949S0001, 4950S0001
                r'Invoice\s+ID[^A-Z0-9]*([A-Z0-9]+)'
            ],
            'purchase_order': [
                r'Purchase\s+Order[^0-9]*([0-9]{10})',
                r'Bon\s+de\s+commande[^0-9]*([0-9]{10})',
                r'commande[^0-9]*([0-9]{10})',
                r'Purchase\s+Order\s*/\s*Bon\s+de[^0-9]*([0-9]{10})',
                r'([0-9]{10})\s+[\d/]+'  # Numéro suivi de date
            ],
            'invoice_date': [
                r'Invoice\s+Date[^0-9]*([0-9]{4}/[0-9]{2}/[0-9]{2})',
                r'Date[^0-9]*([0-9]{4}/[0-9]{2}/[0-9]{2})',
                r'([0-9]{4}/[0-9]{2}/[0-9]{2})',
                r'(\d{4}/\d{2}/\d{2})'
            ],
            'supplier': [
                r'Facture\s+émise\s+par\s*[:]\s*([^\n\r]+)',
                r'^([^0-9\n\r]{2,})',  # Première ligne non numérique
                r'Select\s+T\.T',  # Pattern spécifique
                r'Randstad'
            ],
            'total_net': [
                # Pattern pour les PDFs avec "Invoice Total (EUR)"
                r'Invoice\s+Total\s*\(EUR\)[^0-9,]*([0-9]{1,3}(?:[,.]?\d{3})*[,.]?\d{2})',
                r'Net\s+Amount[^0-9,]*([0-9]{1,3}(?:[,.]?\d{3})*[,.]?\d{2})',
                r'Montant\s*Net[^0-9,]*([0-9]{1,3}(?:[,.]?\d{3})*[,.]?\d{2})',
                # Pattern plus flexible pour capturer différents formats
                r'(\d{1,3}(?:[,.]?\d{3})*[,.]?\d{2})\s*(?:EUR|€|$)',
                # Pattern de fin de ligne pour tableaux
                r'(\d+[,.]\d{2})\s+\d+[,.]\d{2}\s+\d+[,.]\d{2}\s*$'
            ],
            'total_vat': [
                r'VAT\s+Amount[^0-9,]*([0-9]{1,3}(?:[,.]?\d{3})*[,.]?\d{2})',
                r'Montant\s+TVA[^0-9,]*([0-9]{1,3}(?:[,.]?\d{3})*[,.]?\d{2})',
                r'TVA[^0-9,]*([0-9]{1,3}(?:[,.]?\d{3})*[,.]?\d{2})'
            ],
            'total_gross': [
                r'Gross\s+Amount[^0-9,]*([0-9]{1,3}(?:[,.]?\d{3})*[,.]?\d{2})',
                r'Montant\s+brut[^0-9,]*([0-9]{1,3}(?:[,.]?\d{3})*[,.]?\d{2})',
                r'Total.*?([0-9,]+\.[0-9]{2})$'
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
        Extrait les données d'un seul fichier PDF avec gestion d'erreurs améliorée
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
                    try:
                        page_text = page.extract_text()
                        if page_text:
                            full_text += f"\n--- Page {page_num + 1} ---\n" + page_text
                    except Exception as e:
                        self.logger.warning(f"Erreur extraction texte page {page_num + 1}: {e}")
                    
                    # Tableaux de la page
                    try:
                        tables = page.extract_tables()
                        if tables:
                            tables_data.extend(tables)
                    except Exception as e:
                        self.logger.warning(f"Erreur extraction tables page {page_num + 1}: {e}")
                
                # Parsing des données
                extracted_data = self.parse_pdf_content(full_text, pdf_file.name, tables_data)
                
                # Validation des données extraites
                validation_results = self.validate_extracted_data(extracted_data)
                extracted_data['validation'] = validation_results
                
                # Métadonnées
                extracted_data['extraction_metadata'] = {
                    'filename': pdf_file.name,
                    'file_size': getattr(pdf_file, 'size', 0),
                    'pages_count': len(pdf.pages),
                    'extraction_timestamp': datetime.now().isoformat(),
                    'extractor_version': '2.1.1',  # Version corrigée
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
                'filename': getattr(pdf_file, 'name', 'Unknown'),
                'extraction_timestamp': datetime.now().isoformat()
            }
    
    def parse_pdf_content(self, text: str, filename: str, tables: List = None) -> Dict[str, Any]:
        """
        Parse le contenu texte du PDF pour extraire les données structurées
        """
        extracted_data = {
            'success': True,
            'filename': filename,
            'raw_text_sample': text[:500] if text else "",
        }
        
        # Extraction des champs principaux
        for field_name, patterns in self.patterns.items():
            extracted_value = self.extract_field_with_patterns(text, patterns, field_name)
            extracted_data[field_name] = extracted_value
        
        # Extraction des références détaillées pour Select T.T
        invoice_references = self.extract_invoice_references(text)
        extracted_data['invoice_references'] = invoice_references
        
        # Compilation des références pour le rapprochement
        if invoice_references:
            main_ref = invoice_references[0]
            extracted_data['main_reference'] = main_ref['reference_key']
            extracted_data['batch_id'] = main_ref['batch_id']
            extracted_data['assignment_id'] = main_ref['assignment_id']
            extracted_data['all_references'] = [ref['reference_key'] for ref in invoice_references]
        
        # Post-traitement des montants CORRIGÉ
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
        
        self.logger.warning(f"Champ {field_name} non trouvé")
        return None
    
    def clean_extracted_value(self, value: str, field_type: str) -> str:
        """
        Nettoie une valeur extraite selon son type
        """
        if not value:
            return ""
        
        # Nettoyage général
        value = value.strip().replace('\n', ' ').replace('\r', ' ')
        value = re.sub(r'\s+', ' ', value)  # Multiples espaces -> un seul
        
        # Nettoyage spécifique selon le type
        if 'amount' in field_type or 'total' in field_type:
            # Pour les montants, on garde le format original pour parsing ultérieur
            value = re.sub(r'[^\d,.\s]', '', value).strip()
        
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
        Post-traite les montants extraits avec gestion des formats européens CORRIGÉ
        """
        amount_fields = ['total_net', 'total_vat', 'total_gross']
        
        for field in amount_fields:
            if field in data and data[field]:
                try:
                    # Conversion en float avec gestion des formats européens
                    amount_str = str(data[field])
                    amount_float = self.parse_amount_string_corrected(amount_str)
                    
                    # Stockage des deux formats
                    data[f"{field}_original"] = data[field]  # Valeur originale
                    data[field] = amount_float  # Valeur numérique
                    
                    self.logger.debug(f"Montant {field} parsé: '{amount_str}' -> {amount_float}")
                    
                except (ValueError, TypeError) as e:
                    self.logger.warning(f"Erreur conversion montant {field}: {e}")
                    data[f"{field}_error"] = str(e)
                    # Garder une valeur par défaut au lieu de laisser vide
                    data[field] = 0.0
        
        return data
    
    def parse_amount_string_corrected(self, amount_str: str) -> float:
        """
        Convertit une chaîne de montant en float - VERSION CORRIGÉE pour formats européens
        """
        if not amount_str:
            return 0.0
        
        # Nettoyage et normalisation
        clean_amount = str(amount_str).strip()
        
        # Supprimer les espaces internes
        clean_amount = clean_amount.replace(' ', '')
        
        # Debug log pour voir ce qu'on traite
        self.logger.debug(f"Parsing montant: '{clean_amount}'")
        
        if not clean_amount:
            return 0.0
        
        # Gestion des différents formats européens et anglo-saxons
        try:
            # Format européen avec point comme séparateur de milliers et virgule décimale
            # Ex: "1.234,56" -> 1234.56
            if ',' in clean_amount and '.' in clean_amount:
                if clean_amount.rfind(',') > clean_amount.rfind('.'):
                    # Format européen: 1.234,56
                    clean_amount = clean_amount.replace('.', '').replace(',', '.')
                    self.logger.debug(f"Format européen détecté: {clean_amount}")
                else:
                    # Format US: 1,234.56 - enlever la virgule des milliers
                    clean_amount = clean_amount.replace(',', '')
                    self.logger.debug(f"Format US détecté: {clean_amount}")
            
            elif ',' in clean_amount and '.' not in clean_amount:
                # Peut être 1234,56 (européen) ou 1,234 (US sans décimales)
                comma_pos = clean_amount.rfind(',')
                after_comma = clean_amount[comma_pos + 1:]
                
                if len(after_comma) <= 2:  # Probablement des décimales
                    clean_amount = clean_amount.replace(',', '.')
                    self.logger.debug(f"Format européen simple détecté: {clean_amount}")
                else:  # Probablement un séparateur de milliers
                    clean_amount = clean_amount.replace(',', '')
                    self.logger.debug(f"Séparateur milliers détecté: {clean_amount}")
            
            # Nettoyage final - ne garder que chiffres, point et signe moins
            clean_amount = re.sub(r'[^\d.-]', '', clean_amount)
            
            if not clean_amount or clean_amount in ['.', '-', '.-']:
                return 0.0
            
            result = float(clean_amount)
            self.logger.debug(f"Montant final parsé: {result}")
            return result
            
        except ValueError as e:
            # En cas d'erreur, essayer d'extraire au moins les chiffres
            digits_only = re.sub(r'[^\d]', '', clean_amount)
            if digits_only:
                try:
                    # Traiter comme des centimes si plus de 2 chiffres
                    if len(digits_only) > 2:
                        result = float(digits_only[:-2] + '.' + digits_only[-2:])
                    else:
                        result = float(digits_only) / 100  # Traiter comme centimes
                    
                    self.logger.warning(f"Parsing de secours réussi pour '{amount_str}': {result}")
                    return result
                except:
                    pass
            
            self.logger.error(f"Impossible de convertir '{amount_str}' en montant: {e}")
            raise ValueError(f"Impossible de convertir '{amount_str}' en montant")
    
    def extract_invoice_references(self, text: str) -> List[Dict[str, Any]]:
        """
        Extrait les références ligne par ligne des factures Select T.T
        """
        references = []
        
        # Pattern pour capturer les lignes de service
        service_patterns = [
            r'(\d{4}_\d{5}_[^0-9\n\r]+?)[\s]+([\d/]+)[\s]+(Each|Hours)[\s]+([\d,.]+)',
            r'(\d{4}_\d{5}_[^\n\r]+)',  # Pattern plus simple
            r'(\d{4})_(\d{5})_([^0-9\n\r]+)',  # Pattern décomposé
        ]
        
        for pattern in service_patterns:
            try:
                matches = re.findall(pattern, text, re.MULTILINE | re.IGNORECASE)
                
                for match in matches:
                    if isinstance(match, tuple):
                        if len(match) >= 3 and '_' not in match[1]:  # Pattern décomposé
                            batch_id, assignment_id, service_desc = match[0], match[1], match[2]
                            full_ref = f"{batch_id}_{assignment_id}_{service_desc.strip()}"
                        elif len(match) >= 1:  # Pattern complet
                            full_ref = match[0].strip()
                            # Décomposition de la référence
                            ref_parts = full_ref.split('_', 2)
                            if len(ref_parts) >= 3:
                                batch_id = ref_parts[0]
                                assignment_id = ref_parts[1]
                                service_desc = ref_parts[2].strip()
                            else:
                                continue
                        else:
                            continue
                        
                        references.append({
                            'full_reference': full_ref,
                            'batch_id': batch_id,
                            'assignment_id': assignment_id,
                            'service_description': service_desc,
                            'reference_key': f"{batch_id}_{assignment_id}"
                        })
                        
            except Exception as e:
                self.logger.warning(f"Erreur extraction références avec pattern {pattern}: {e}")
                continue
        
        # Supprimer les doublons
        seen = set()
        unique_references = []
        for ref in references:
            ref_key = ref.get('reference_key', '')
            if ref_key and ref_key not in seen:
                seen.add(ref_key)
                unique_references.append(ref)
        
        return unique_references
    
    def extract_billing_details(self, tables: List) -> List[Dict[str, Any]]:
        """
        Extrait les détails de facturation depuis les tableaux avec gestion d'erreurs
        """
        billing_details = []
        
        for table in tables:
            try:
                if not table or len(table) < 2:
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
                        
                        # Mapping basique des colonnes
                        for i, cell in enumerate(row):
                            if cell is not None and str(cell).strip():
                                if i == 0:  # Première colonne généralement description
                                    line_data['description'] = str(cell).strip()
                                elif self.is_amount_cell(cell):
                                    amount_key = f'amount_{len([k for k in line_data.keys() if k.startswith("amount")])}'
                                    try:
                                        line_data[amount_key] = self.parse_amount_string_corrected(str(cell))
                                    except:
                                        line_data[amount_key] = 0.0
                                elif self.is_quantity_cell(cell):
                                    try:
                                        line_data['quantity'] = self.parse_amount_string_corrected(str(cell))
                                    except:
                                        line_data['quantity'] = 0.0
                                else:
                                    line_data[f'column_{i}'] = str(cell).strip()
                        
                        if line_data:
                            billing_details.append(line_data)
                            
            except Exception as e:
                self.logger.warning(f"Erreur extraction billing details pour table: {e}")
                continue
        
        return billing_details
    
    def is_amount_cell(self, cell) -> bool:
        """Détermine si une cellule contient un montant"""
        if not cell:
            return False
        
        cell_str = str(cell).strip()
        # Pattern pour détecter les montants avec formats européens
        amount_pattern = r'^[0-9]{1,3}(?:[,.]?\d{3})*[,.]\d{2}$|^[0-9]+[,.]?\d{0,2}
        return bool(re.match(amount_pattern, cell_str))
    
    def is_quantity_cell(self, cell) -> bool:
        """Détermine si une cellule contient une quantité"""
        if not cell:
            return False
        
        cell_str = str(cell).strip()
        # Pattern pour quantités simples
        quantity_pattern = r'^[0-9]+([,.][0-9]{1,2})?
        return bool(re.match(quantity_pattern, cell_str))
    
    def calculate_data_completeness(self, data: Dict[str, Any]) -> Dict[str, Any]:
        """
        Calcule le pourcentage de complétude des données extraites
        """
        # Champs critiques pour le rapprochement
        critical_fields = ['purchase_order', 'total_net']
        important_fields = ['invoice_id', 'invoice_date', 'supplier']
        optional_fields = ['total_vat', 'total_gross', 'main_reference']
        
        completeness = {
            'critical_score': 0,
            'important_score': 0,
            'optional_score': 0,
            'overall_score': 0,
            'missing_critical': [],
            'missing_important': [],
            'processable': False
        }
        
        try:
            # Évaluation des champs critiques
            critical_found = 0
            for field in critical_fields:
                field_value = data.get(field)
                if field_value and str(field_value).strip() and field_value != 0:
                    critical_found += 1
                else:
                    completeness['missing_critical'].append(field)
            
            completeness['critical_score'] = (critical_found / len(critical_fields)) * 100
            
            # Évaluation des champs importants
            important_found = 0
            for field in important_fields:
                field_value = data.get(field)
                if field_value and str(field_value).strip():
                    important_found += 1
                else:
                    completeness['missing_important'].append(field)
            
            completeness['important_score'] = (important_found / len(important_fields)) * 100
            
            # Évaluation des champs optionnels
            optional_found = 0
            for field in optional_fields:
                field_value = data.get(field)
                if field_value and str(field_value).strip():
                    optional_found += 1
            
            completeness['optional_score'] = (optional_found / len(optional_fields)) * 100
            
            # Score global pondéré
            completeness['overall_score'] = (
                completeness['critical_score'] * 0.6 +
                completeness['important_score'] * 0.3 +
                completeness['optional_score'] * 0.1
            )
            
            # Détermination si les données sont processables
            completeness['processable'] = (
                completeness['critical_score'] >= 100 and
                completeness['important_score'] >= 50
            )
            
        except Exception as e:
            self.logger.error(f"Erreur calcul completeness: {e}")
            completeness['processable'] = False
        
        return completeness
    
    def validate_extracted_data(self, data: Dict[str, Any]) -> Dict[str, Any]:
        """
        Valide les données extraites selon les règles métier
        """
        validation_results = {
            'is_valid': True,
            'errors': [],
            'warnings': [],
            'field_validations': {}
        }
        
        try:
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
            
            # Validation des montants CORRIGÉE
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
                    
        except Exception as e:
            self.logger.error(f"Erreur validation: {e}")
            validation_results['is_valid'] = False
            validation_results['errors'].append(f"Erreur validation: {str(e)}")
        
        return validation_results
    
    def validate_purchase_order(self, po: str) -> Dict[str, Any]:
        """Valide un numéro de commande"""
        validation = {'valid': True, 'errors': []}
        
        try:
            if not po:
                validation['valid'] = False
                validation['errors'].append("Numéro de commande vide")
                return validation
            
            po_str = str(po).strip()
            
            # Doit contenir exactement 10 chiffres
            if not re.match(r'^[0-9]{10}, po_str):
                validation['valid'] = False
                validation['errors'].append(f"Format numéro de commande incorrect: {po_str} (attendu: 10 chiffres)")
                
        except Exception as e:
            validation['valid'] = False
            validation['errors'].append(f"Erreur validation PO: {e}")
        
        return validation
    
    def validate_invoice_id(self, invoice_id: str) -> Dict[str, Any]:
        """Valide un ID de facture"""
        validation = {'valid': True, 'errors': []}
        
        try:
            if not invoice_id:
                validation['valid'] = False
                validation['errors'].append("ID facture vide")
                return validation
            
            id_str = str(invoice_id).strip()
            
            # Doit contenir au moins 4 caractères alphanumériques
            if len(id_str) < 4 or not re.match(r'^[A-Z0-9]+, id_str):
                validation['valid'] = False
                validation['errors'].append(f"Format ID facture suspect: {id_str}")
                
        except Exception as e:
            validation['valid'] = False
            validation['errors'].append(f"Erreur validation invoice_id: {e}")
        
        return validation
    
    def validate_amounts(self, data: Dict[str, Any]) -> Dict[str, Any]:
        """Valide les montants extraits - VERSION CORRIGÉE"""
        validation = {'valid': True, 'errors': []}
        
        try:
            # Vérification du montant net (critique)
            total_net = data.get('total_net')
            
            # Le montant peut maintenant être 0.0 (au lieu de None ou string)
            if total_net is None:
                validation['valid'] = False
                validation['errors'].append("Montant net manquant")
            elif not isinstance(total_net, (int, float)):
                validation['valid'] = False
                validation['errors'].append(f"Montant net doit être numérique, reçu: {type(total_net)}")
            elif total_net < 0:
                validation['valid'] = False
                validation['errors'].append(f"Montant net négatif: {total_net}€")
            elif total_net > 100000:  # Plus de 100k euros, suspect
                validation['errors'].append(f"Montant net très élevé: {total_net}€")
            
            # Cohérence entre HT, TVA et TTC
            total_vat = data.get('total_vat')
            total_gross = data.get('total_gross')
            
            if (isinstance(total_net, (int, float)) and isinstance(total_vat, (int, float)) and 
                isinstance(total_gross, (int, float)) and all([total_net > 0, total_vat >= 0, total_gross > 0])):
                
                expected_gross = total_net + total_vat
                difference = abs(total_gross - expected_gross)
                
                if difference > 0.02:  # Tolérance de 2 centimes
                    validation['errors'].append(
                        f"Incohérence montants: HT({total_net}) + TVA({total_vat}) ≠ TTC({total_gross})"
                    )
                    
        except Exception as e:
            self.logger.error(f"Erreur validation montants: {e}")
            validation['valid'] = False
            validation['errors'].append(f"Erreur validation montants: {e}")
        
        return validation
    
    def validate_date(self, date_str: str) -> Dict[str, Any]:
        """Valide une date de facture"""
        validation = {'valid': True, 'errors': []}
        
        try:
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
                    
        except Exception as e:
            validation['valid'] = False
            validation['errors'].append(f"Erreur validation date: {e}")
        
        return validation
    
    def validate_supplier(self, supplier: str) -> Dict[str, Any]:
        """Valide le nom du fournisseur"""
        validation = {'valid': True, 'errors': []}
        
        try:
            if not supplier:
                validation['valid'] = False
                validation['errors'].append("Nom fournisseur manquant")
                return validation
            
            supplier_str = str(supplier).strip().lower()
            
            # Fournisseurs acceptés basés sur les données réelles
            expected_suppliers = [
                'randstad', 'select t.t.', 'select t.t', 'select tt',
                'randstad france', 'select t.t. (randstad france)'
            ]
            
            if not any(expected in supplier_str for expected in expected_suppliers):
                validation['errors'].append(f"Fournisseur non reconnu: {supplier}")
                
        except Exception as e:
            validation['valid'] = False
            validation['errors'].append(f"Erreur validation supplier: {e}")
        
        return validation
    
    def extract_multiple_pdfs(self, pdf_files: List) -> List[Dict[str, Any]]:
        """
        Extrait les données de plusieurs fichiers PDF
        """
        results = []
        
        for i, pdf_file in enumerate(pdf_files):
            try:
                self.logger.info(f"Traitement PDF {i+1}/{len(pdf_files)}: {pdf_file.name}")
                
                result = self.extract_single_pdf(pdf_file)
                result['batch_index'] = i
                results.append(result)
                
            except Exception as e:
                self.logger.error(f"Erreur traitement PDF {pdf_file.name}: {str(e)}")
                results.append({
                    'success': False,
                    'error': str(e),
                    'filename': getattr(pdf_file, 'name', 'Unknown'),
                    'batch_index': i
                })
        
        return results
    
    def get_extraction_summary(self, extraction_results: List[Dict[str, Any]]) -> Dict[str, Any]:
        """
        Génère un résumé des extractions
        """
        try:
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
            
        except Exception as e:
            self.logger.error(f"Erreur génération summary: {e}")
            return {
                'total_files': len(extraction_results) if extraction_results else 0,
                'successful_extractions': 0,
                'failed_extractions': len(extraction_results) if extraction_results else 0,
                'success_rate': 0,
                'error': str(e)
            }


# Fonctions utilitaires pour les tests
def test_pdf_extractor():
    """Fonction de test pour le module PDF corrigé"""
    extractor = PDFExtractor()
    
    # Test des montants européens
    test_amounts = [
        "1,234.56",  # Format US
        "1.234,56",  # Format européen
        "4,047.24",  # Problème du log
        "28,308.73", # Gros montant du log
        "9.84",      # Simple
        "1,014.54"   # Moyen
    ]
    
    print("Test parsing montants corrigé:")
    for amount in test_amounts:
        try:
            result = extractor.parse_amount_string_corrected(amount)
            print(f"  '{amount}' -> {result}")
        except Exception as e:
            print(f"  '{amount}' -> ERREUR: {e}")
    
    # Test avec texte de facture simulé
    test_text = """
    Invoice ID/Number: 4949S0001
    Purchase Order / Bon de commande: 5600025054
    Invoice Date: 2025/03/10
    4949_65744_Temporary employees - Expense
    4950_65744_Temporary employees - Timesheet
    Invoice Total (EUR): 4,047.24
    Supplier: Select T.T
    """
    
    result = extractor.parse_pdf_content(test_text, "test.pdf")
    print("\nTest extraction corrigée:")
    for key, value in result.items():
        if not key.startswith('raw_text'):
            print(f"  {key}: {value}")
    
    return result

if __name__ == "__main__":
    # Test du module si exécuté directement
    test_pdf_extractor()

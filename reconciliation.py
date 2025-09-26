"""
RECONCILIATION.PY - VERSION CORRIGÉE
Corrections des erreurs critiques détectées dans les logs
"""

import pandas as pd
import numpy as np
import re
import logging
from typing import Dict, List, Optional, Any, Tuple
from datetime import datetime, timedelta
from dataclasses import dataclass
from enum import Enum
import difflib
import streamlit as st

class MatchType(Enum):
    """Types de rapprochement possibles"""
    PERFECT_MATCH = "perfect_match"
    DISCREPANCY = "discrepancy"
    UNMATCHED = "unmatched"

class MatchMethod(Enum):
    """Méthodes de rapprochement"""
    EXACT_ORDER = "exact_order"
    PARTIAL_ORDER = "partial_order"
    AMOUNT_FUZZY = "amount_fuzzy"
    REFERENCE_CROSS = "reference_cross"
    COST_CENTER = "cost_center"
    INTELLIGENT = "intelligent"

@dataclass
class MatchResult:
    """Résultat d'un rapprochement"""
    match_type: MatchType
    method: MatchMethod
    confidence: float
    pdf_data: Dict[str, Any]
    excel_data: Dict[str, Any]
    differences: Dict[str, Any]
    metadata: Dict[str, Any]

class ReconciliationEngine:
    """
    Moteur de rapprochement intelligent pour Beeline - Version corrigée
    """
    
    def __init__(self, config: Optional[Dict[str, Any]] = None):
        self.setup_logging()
        
        # Configuration par défaut optimisée
        self.config = {
            'tolerance': 0.01,
            'method': 'intelligent',
            'fuzzy_threshold': 0.8,
            'date_tolerance_days': 30,
            'amount_weights': {
                'exact': 1.0,
                'close': 0.8,
                'distant': 0.3
            },
            'enable_reference_matching': True,
            'min_confidence': 0.6
        }
        
        if config:
            self.config.update(config)
        
        # Statistiques de performance
        self.stats = {
            'total_pdfs': 0,
            'total_excel_orders': 0,
            'perfect_matches': 0,
            'discrepancies': 0,
            'unmatched_pdf': 0,
            'unmatched_excel': 0,
            'processing_time': 0,
            'method_performance': {}
        }
    
    def setup_logging(self):
        """Configure le logging pour le moteur"""
        logging.basicConfig(level=logging.INFO)
        self.logger = logging.getLogger(__name__)
    
    def perform_reconciliation(self, pdf_data: List[Dict[str, Any]], 
                             excel_data: List[Dict[str, Any]], 
                             config: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
        """
        Effectue le rapprochement principal entre PDFs et Excel
        """
        start_time = datetime.now()
        self.logger.info("Démarrage du rapprochement intelligent optimisé")
        
        if config:
            self.config.update(config)
        
        try:
            # 1. Validation et nettoyage des données d'entrée
            if not pdf_data or not excel_data:
                raise ValueError("Données PDF ou Excel manquantes")
            
            # S'assurer que pdf_data est une liste de dictionnaires
            if not isinstance(pdf_data, list):
                raise ValueError(f"pdf_data doit être une liste, reçu: {type(pdf_data)}")
            
            if not isinstance(excel_data, list):
                raise ValueError(f"excel_data doit être une liste, reçu: {type(excel_data)}")
            
            # Filtrer les données invalides
            valid_pdf_data = []
            for item in pdf_data:
                if isinstance(item, dict) and item.get('success', False):
                    valid_pdf_data.append(item)
                elif isinstance(item, dict):
                    self.logger.warning(f"PDF invalide ignoré: {item.get('filename', 'Unknown')}")
            
            valid_excel_data = []
            for item in excel_data:
                if isinstance(item, dict) and item.get('is_valid', False):
                    valid_excel_data.append(item)
                elif isinstance(item, dict):
                    self.logger.warning(f"Ligne Excel invalide ignorée: {item.get('order_number', 'Unknown')}")
            
            self.logger.info(f"Données valides: {len(valid_pdf_data)} PDFs, {len(valid_excel_data)} lignes Excel")
            
            # 2. Préparation des données
            prepared_pdfs = self.prepare_pdf_data(valid_pdf_data)
            aggregated_excel = self.prepare_excel_data(valid_excel_data)
            
            # 3. Statistiques initiales
            self.stats['total_pdfs'] = len(prepared_pdfs)
            self.stats['total_excel_orders'] = len(aggregated_excel)
            
            self.logger.info(f"PDFs préparés: {len(prepared_pdfs)}")
            self.logger.info(f"Excel agrégés: {len(aggregated_excel)}")
            
            # 4. Rapprochement intelligent multi-niveaux
            results = self.intelligent_reconciliation(prepared_pdfs, aggregated_excel)
            
            # 5. Post-traitement et analyse
            results = self.post_process_results(results, prepared_pdfs, aggregated_excel)
            
            # 6. Génération du résumé
            results['summary'] = self.generate_summary(results)
            
            # 7. Métadonnées de traitement
            processing_time = (datetime.now() - start_time).total_seconds()
            self.stats['processing_time'] = processing_time
            
            results['metadata'] = {
                'processing_time': processing_time,
                'reconciliation_timestamp': datetime.now().isoformat(),
                'engine_version': '2.1.0',
                'config_used': self.config.copy(),
                'performance_stats': self.stats.copy()
            }
            
            self.logger.info(f"Rapprochement terminé en {processing_time:.2f}s")
            return results
            
        except Exception as e:
            self.logger.error(f"Erreur durant le rapprochement: {str(e)}")
            # Retourner un résultat d'erreur structuré au lieu de lever l'exception
            return {
                'success': False,
                'error': str(e),
                'matches': [],
                'discrepancies': [],
                'unmatched_pdf': [],
                'unmatched_excel': [],
                'summary': {
                    'total_invoices': 0,
                    'perfect_matches': 0,
                    'discrepancies': 0,
                    'unmatched_pdf': 0,
                    'unmatched_excel': 0,
                    'matching_rate': 0,
                    'coverage_rate': 0,
                    'total_amount': 0
                },
                'metadata': {
                    'processing_time': (datetime.now() - start_time).total_seconds(),
                    'reconciliation_timestamp': datetime.now().isoformat(),
                    'engine_version': '2.1.0',
                    'error_occurred': True
                }
            }
    
    def prepare_pdf_data(self, pdf_data: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """
        Prépare et nettoie les données PDF pour le rapprochement
        """
        prepared = []
        
        for i, pdf in enumerate(pdf_data):
            if not isinstance(pdf, dict):
                self.logger.warning(f"PDF {i} n'est pas un dictionnaire, ignoré")
                continue
                
            if not pdf.get('success', True):
                continue
            
            # Extraction des champs essentiels avec validation
            prepared_pdf = {
                'index': i,
                'filename': pdf.get('filename', f'pdf_{i}'),
                'purchase_order': self.clean_order_number(pdf.get('purchase_order')),
                'invoice_id': pdf.get('invoice_id'),
                'total_net': self.safe_parse_amount(pdf.get('total_net', 0)),
                'invoice_date': self.parse_date(pdf.get('invoice_date')),
                'supplier': self.clean_supplier_name(pdf.get('supplier', '')),
                'client': pdf.get('client', ''),
                'validation': pdf.get('validation', {}),
                'data_completeness': pdf.get('data_completeness', {}),
                'raw_data': pdf
            }
            
            # Ajout des références pour matching avancé
            prepared_pdf['main_reference'] = pdf.get('main_reference', '')
            prepared_pdf['batch_id'] = pdf.get('batch_id', '')
            prepared_pdf['assignment_id'] = pdf.get('assignment_id', '')
            prepared_pdf['invoice_references'] = pdf.get('invoice_references', [])
            
            # Calcul de la qualité des données
            prepared_pdf['data_quality_score'] = self.calculate_pdf_quality_score(prepared_pdf)
            
            # Inclure seulement les PDFs avec données minimales
            if (prepared_pdf['purchase_order'] or prepared_pdf['main_reference']) and prepared_pdf['total_net'] > 0:
                prepared.append(prepared_pdf)
            else:
                self.logger.warning(f"PDF {prepared_pdf['filename']} exclu: données insuffisantes")
        
        self.logger.info(f"PDFs préparés: {len(prepared)} sur {len(pdf_data)} fournis")
        return prepared
    
    def prepare_excel_data(self, excel_data: List[Dict[str, Any]]) -> Dict[str, Dict[str, Any]]:
        """
        Prépare et agrège les données Excel par numéro de commande
        """
        aggregated = {}
        
        for record in excel_data:
            if not isinstance(record, dict):
                continue
                
            if not record.get('is_valid', False):
                continue
            
            order_number = record.get('order_number')
            if not order_number:
                continue
            
            # Nettoyage du numéro de commande
            clean_order = self.clean_order_number(order_number)
            if not clean_order:
                continue
            
            if clean_order not in aggregated:
                aggregated[clean_order] = {
                    'order_number': clean_order,
                    'total_amount': 0,
                    'line_count': 0,
                    'collaborators': [],
                    'cost_centers': [],
                    'statement_dates': [],
                    'suppliers': [],
                    'source_files': [],
                    'raw_lines': [],
                    'projects': [],
                    'validation_summary': {
                        'total_lines': 0,
                        'valid_lines': 0,
                        'invalid_lines': 0,
                        'validity_rate': 0
                    }
                }
            
            order_data = aggregated[clean_order]
            
            # Agrégation des montants avec sécurité
            net_amount = self.safe_parse_amount(record.get('net_amount', 0))
            order_data['total_amount'] += net_amount
            
            # Comptage des lignes
            order_data['line_count'] += 1
            
            # Collecte des collaborateurs
            collaborator = record.get('collaborator')
            if collaborator and str(collaborator).strip() not in order_data['collaborators']:
                order_data['collaborators'].append(str(collaborator).strip())
            
            # Collecte des centres de coût
            cost_center = record.get('cost_center')
            if cost_center and str(cost_center).strip() not in order_data['cost_centers']:
                order_data['cost_centers'].append(str(cost_center).strip())
            
            # Autres collectes...
            statement_date = record.get('statement_date')
            if statement_date and statement_date not in order_data['statement_dates']:
                order_data['statement_dates'].append(statement_date)
            
            supplier = record.get('supplier')
            if supplier and str(supplier).strip() not in order_data['suppliers']:
                order_data['suppliers'].append(str(supplier).strip())
            
            project = record.get('project')
            if project and str(project).strip() not in order_data['projects']:
                order_data['projects'].append(str(project).strip())
            
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
    
    def safe_parse_amount(self, amount: Any) -> float:
        """Parse un montant en float avec gestion d'erreurs robuste"""
        if not amount:
            return 0.0
        
        try:
            if isinstance(amount, (int, float)):
                return float(amount)
            
            # Nettoyage du string
            amount_str = str(amount).replace(',', '.').replace(' ', '')
            amount_clean = re.sub(r'[^\d.-]', '', amount_str)
            
            return float(amount_clean) if amount_clean else 0.0
            
        except (ValueError, TypeError) as e:
            self.logger.warning(f"Erreur parsing montant {amount}: {e}")
            return 0.0
    
    def clean_order_number(self, order_number: Any) -> Optional[str]:
        """Nettoie un numéro de commande avec validation robuste"""
        if not order_number:
            return None
        
        try:
            # Conversion en string et nettoyage
            order_str = str(order_number).strip()
            
            # Gestion des nombres Excel (5600002101.0 -> 5600002101)
            if '.' in order_str:
                order_str = order_str.split('.')[0]
            
            # Suppression des caractères non numériques
            clean_order = re.sub(r'[^\d]', '', order_str)
            
            # Validation de la longueur (accepter 8 à 12 chiffres)
            if 8 <= len(clean_order) <= 12:
                return clean_order
            
            return None
            
        except Exception as e:
            self.logger.warning(f"Erreur nettoyage order_number {order_number}: {e}")
            return None
    
    def intelligent_reconciliation(self, pdf_data: List[Dict[str, Any]], 
                                 excel_data: Dict[str, Dict[str, Any]]) -> Dict[str, Any]:
        """
        Rapprochement intelligent multi-niveaux optimisé avec gestion d'erreurs
        """
        self.logger.info("Lancement du rapprochement intelligent multi-niveaux")
        
        results = {
            'matches': [],
            'discrepancies': [],
            'unmatched_pdf': [],
            'unmatched_excel': list(excel_data.keys()),
            'match_details': []
        }
        
        if not pdf_data or not excel_data:
            self.logger.warning("Données vides pour le rapprochement")
            return results
        
        # Phase 1: Rapprochement EXACT par numéro de commande
        self.logger.info("   Phase 1: Rapprochement exact par N° commande")
        for pdf in pdf_data:
            try:
                match_result = self.try_exact_order_match(pdf, excel_data)
                
                if match_result:
                    self.process_match_result(match_result, results)
                else:
                    pdf['match_attempts'] = []
                    
            except Exception as e:
                self.logger.error(f"Erreur phase 1 pour PDF {pdf.get('filename', 'Unknown')}: {e}")
                pdf['match_attempts'] = []
        
        # Phase 2: Rapprochement par références croisées
        self.logger.info("   Phase 2: Rapprochement par références croisées")
        unmatched_pdfs = [pdf for pdf in pdf_data if 'match_attempts' in pdf]
        
        for pdf in unmatched_pdfs:
            try:
                match_result = self.try_reference_cross_match(pdf, excel_data, results['unmatched_excel'])
                
                if match_result and match_result.confidence >= self.config['min_confidence']:
                    self.process_match_result(match_result, results)
                else:
                    if match_result:
                        pdf['match_attempts'].append(('reference', match_result.confidence))
                        
            except Exception as e:
                self.logger.error(f"Erreur phase 2 pour PDF {pdf.get('filename', 'Unknown')}: {e}")
        
        # Phase 3: Finalisation des non-matchés
        final_unmatched = [pdf for pdf in pdf_data if 'match_attempts' in pdf]
        for pdf in final_unmatched:
            try:
                results['unmatched_pdf'].append(self.format_unmatched_pdf(pdf))
            except Exception as e:
                self.logger.error(f"Erreur formatage PDF non matché {pdf.get('filename', 'Unknown')}: {e}")
        
        self.logger.info(f"   Résultats: {len(results['matches'])} matches, {len(results['discrepancies'])} écarts")
        return results
    
    # ... [Autres méthodes avec gestion d'erreurs similaire]
    
    def try_exact_order_match(self, pdf: Dict[str, Any], 
                             excel_data: Dict[str, Dict[str, Any]]) -> Optional[MatchResult]:
        """Rapprochement exact par numéro de commande avec validation"""
        try:
            pdf_order = pdf.get('purchase_order')
            if not pdf_order:
                return None
            
            # Recherche exacte dans Excel
            if pdf_order in excel_data:
                excel_order = excel_data[pdf_order]
                
                # Calcul des différences de montants
                pdf_amount = pdf.get('total_net', 0)
                excel_amount = excel_order.get('total_amount', 0)
                amount_diff = abs(pdf_amount - excel_amount)
                amount_tolerance = pdf_amount * self.config['tolerance'] if pdf_amount > 0 else 0.01
                
                # Détermination du type de match
                if amount_diff <= amount_tolerance:
                    match_type = MatchType.PERFECT_MATCH
                    confidence = 1.0
                else:
                    match_type = MatchType.DISCREPANCY
                    confidence = max(0.3, 1.0 - (amount_diff / pdf_amount)) if pdf_amount > 0 else 0.3
                
                # Création du résultat
                match_result = MatchResult(
                    match_type=match_type,
                    method=MatchMethod.EXACT_ORDER,
                    confidence=confidence,
                    pdf_data=pdf,
                    excel_data=excel_order,
                    differences={
                        'amount_difference': amount_diff,
                        'amount_percentage': (amount_diff / pdf_amount) * 100 if pdf_amount > 0 else 0,
                        'order_match': 'exact'
                    },
                    metadata={
                        'match_timestamp': datetime.now().isoformat(),
                        'order_number': pdf_order,
                        'tolerance_used': self.config['tolerance']
                    }
                )
                
                return match_result
            
        except Exception as e:
            self.logger.error(f"Erreur exact match pour PDF {pdf.get('filename', 'Unknown')}: {e}")
        
        return None
    
    def process_match_result(self, match_result: MatchResult, results: Dict[str, Any]):
        """Traite et classe un résultat de match avec gestion d'erreurs"""
        try:
            # Retirer de la liste des Excel non matchés
            excel_order_number = match_result.excel_data.get('order_number')
            if excel_order_number and excel_order_number in results['unmatched_excel']:
                results['unmatched_excel'].remove(excel_order_number)
            
            # Marquer le PDF comme matché
            if 'match_attempts' in match_result.pdf_data:
                del match_result.pdf_data['match_attempts']
            
            # Formater le résultat pour l'output
            formatted_result = self.format_match_result(match_result)
            
            # Classifier le résultat
            if match_result.match_type == MatchType.PERFECT_MATCH:
                results['matches'].append(formatted_result)
                self.stats['perfect_matches'] += 1
            else:
                results['discrepancies'].append(formatted_result)
                self.stats['discrepancies'] += 1
            
            # Enregistrer les détails du match
            results['match_details'].append({
                'method': match_result.method.value,
                'confidence': match_result.confidence,
                'type': match_result.match_type.value,
                'pdf_file': match_result.pdf_data.get('filename', 'Unknown'),
                'excel_order': excel_order_number or 'Unknown'
            })
            
        except Exception as e:
            self.logger.error(f"Erreur processing match result: {e}")
    
    # ... [Reste des méthodes avec gestion d'erreurs similaire]
    
    def format_match_result(self, match_result: MatchResult) -> Dict[str, Any]:
        """Formatage enrichi des résultats avec gestion d'erreurs"""
        try:
            base_result = {
                'type': match_result.match_type.value,
                'method': match_result.method.value,
                'confidence': round(match_result.confidence, 3),
                'order_number': match_result.excel_data.get('order_number', 'Unknown'),
                'pdf_file': match_result.pdf_data.get('filename', 'Unknown'),
                'pdf_amount': match_result.pdf_data.get('total_net', 0),
                'excel_amount': match_result.excel_data.get('total_amount', 0),
                'difference': match_result.differences.get('amount_difference', 0),
                'difference_percent': round(match_result.differences.get('amount_percentage', 0), 2),
                'collaborators': ', '.join(match_result.excel_data.get('collaborators', [])),
                'supplier': match_result.pdf_data.get('supplier', ''),
                'invoice_id': match_result.pdf_data.get('invoice_id', ''),
                'invoice_date': match_result.pdf_data.get('invoice_date', ''),
                'excel_line_count': match_result.excel_data.get('line_count', 0),
            }
            
            return base_result
            
        except Exception as e:
            self.logger.error(f"Erreur formatage match result: {e}")
            return {
                'type': 'error',
                'method': 'unknown',
                'confidence': 0,
                'error': str(e)
            }
    
    # ... [Méthodes utilitaires avec gestion d'erreurs]
    
    def post_process_results(self, results: Dict[str, Any], 
                           pdf_data: List[Dict[str, Any]], 
                           excel_data: Dict[str, Dict[str, Any]]) -> Dict[str, Any]:
        """Post-traite les résultats avec gestion d'erreurs"""
        try:
            # Ajout des Excel non matchés formatés
            unmatched_excel_formatted = []
            for excel_order_key in results['unmatched_excel']:
                if excel_order_key in excel_data:
                    excel_order = excel_data[excel_order_key]
                    unmatched_excel_formatted.append({
                        'order_number': excel_order_key,
                        'total_amount': excel_order.get('total_amount', 0),
                        'collaborators': ', '.join(excel_order.get('collaborators', [])),
                        'cost_centers': ', '.join(excel_order.get('cost_centers', [])),
                        'line_count': excel_order.get('line_count', 0),
                        'source_files': ', '.join(excel_order.get('source_files', [])),
                        'reason': 'Aucun PDF correspondant trouvé'
                    })
            
            results['unmatched_excel'] = unmatched_excel_formatted
            
            return results
            
        except Exception as e:
            self.logger.error(f"Erreur post-processing: {e}")
            return results
    
    def generate_summary(self, results: Dict[str, Any]) -> Dict[str, Any]:
        """Génère un résumé des résultats avec gestion d'erreurs"""
        try:
            total_pdfs = self.stats['total_pdfs']
            total_excel = self.stats['total_excel_orders']
            matches = len(results.get('matches', []))
            discrepancies = len(results.get('discrepancies', []))
            unmatched_pdf = len(results.get('unmatched_pdf', []))
            unmatched_excel = len(results.get('unmatched_excel', []))
            
            # Calcul des taux
            matching_rate = (matches / total_pdfs * 100) if total_pdfs > 0 else 0
            discrepancy_rate = (discrepancies / total_pdfs * 100) if total_pdfs > 0 else 0
            coverage_rate = ((matches + discrepancies) / total_pdfs * 100) if total_pdfs > 0 else 0
            
            summary = {
                'total_invoices': total_pdfs,
                'total_excel_orders': total_excel,
                'perfect_matches': matches,
                'discrepancies': discrepancies,
                'unmatched_pdf': unmatched_pdf,
                'unmatched_excel': unmatched_excel,
                'matching_rate': round(matching_rate, 2),
                'discrepancy_rate': round(discrepancy_rate, 2),
                'coverage_rate': round(coverage_rate, 2),
                'total_amount': sum(m.get('pdf_amount', 0) for m in results.get('matches', [])),
                'method_performance': self.stats['method_performance'],
            }
            
            return summary
            
        except Exception as e:
            self.logger.error(f"Erreur génération summary: {e}")
            return {
                'total_invoices': 0,
                'perfect_matches': 0,
                'discrepancies': 0,
                'matching_rate': 0
            }
    
    # ... [Autres méthodes utilitaires]
    
    def calculate_pdf_quality_score(self, pdf: Dict[str, Any]) -> float:
        """Calcule un score de qualité pour un PDF avec gestion d'erreurs"""
        try:
            score = 0.0
            weights = {
                'purchase_order': 0.3,
                'total_net': 0.3,
                'invoice_id': 0.1,
                'supplier': 0.1,
                'invoice_date': 0.1,
                'main_reference': 0.1
            }
            
            for field, weight in weights.items():
                if pdf.get(field):
                    score += weight
            
            return min(1.0, score)
            
        except Exception as e:
            self.logger.warning(f"Erreur calcul quality score: {e}")
            return 0.0
    
    def parse_date(self, date_value: Any) -> Optional[datetime]:
        """Parse une date avec gestion d'erreurs robuste"""
        if not date_value:
            return None
        
        if isinstance(date_value, datetime):
            return date_value
        
        try:
            date_str = str(date_value).strip()
            date_formats = ['%Y/%m/%d', '%d/%m/%Y', '%Y-%m-%d', '%d-%m-%Y']
            
            for fmt in date_formats:
                try:
                    return datetime.strptime(date_str, fmt)
                except ValueError:
                    continue
                    
        except Exception as e:
            self.logger.warning(f"Erreur parsing date {date_value}: {e}")
        
        return None
    
    def clean_supplier_name(self, supplier: str) -> str:
        """Nettoie un nom de fournisseur avec gestion d'erreurs"""
        try:
            if not supplier:
                return ""
            
            cleaned = str(supplier).strip().lower()
            
            if 'randstad' in cleaned:
                return 'Randstad'
            elif 'select' in cleaned and 't.t' in cleaned:
                return 'Select T.T.'
            
            return str(supplier).strip()
            
        except Exception as e:
            self.logger.warning(f"Erreur cleaning supplier {supplier}: {e}")
            return ""
    
    def format_unmatched_pdf(self, pdf: Dict[str, Any]) -> Dict[str, Any]:
        """Formate un PDF non matché avec diagnostic détaillé"""
        try:
            reasons = []
            
            if not pdf.get('purchase_order'):
                reasons.append("Numéro de commande manquant ou invalide")
            
            if pdf.get('total_net', 0) <= 0:
                reasons.append("Montant invalide ou manquant")
            
            if pdf.get('data_quality_score', 0) < 0.5:
                reasons.append("Qualité des données insuffisante")
            
            # Analyser les tentatives de match
            match_attempts = pdf.get('match_attempts', [])
            if match_attempts:
                best_attempt = max(match_attempts, key=lambda x: x[1])
                reasons.append(f"Meilleure tentative: {best_attempt[0]} (confiance: {best_attempt[1]:.2f})")
            else:
                reasons.append("Aucune correspondance trouvée")
            
            return {
                'filename': pdf.get('filename', 'Unknown'),
                'order_number': pdf.get('purchase_order', 'Non trouvé'),
                'amount': pdf.get('total_net', 0),
                'invoice_id': pdf.get('invoice_id', 'Non trouvé'),
                'invoice_date': pdf.get('invoice_date', 'Non trouvé'),
                'supplier': pdf.get('supplier', 'Non trouvé'),
                'main_reference': pdf.get('main_reference', 'Non trouvé'),
                'data_quality_score': pdf.get('data_quality_score', 0),
                'reasons': reasons,
                'reason': ' | '.join(reasons)
            }
            
        except Exception as e:
            self.logger.error(f"Erreur formatage PDF non matché: {e}")
            return {
                'filename': pdf.get('filename', 'Unknown'),
                'order_number': 'Erreur',
                'amount': 0,
                'reason': f'Erreur formatage: {str(e)}'
            }
    
    def try_reference_cross_match(self, pdf: Dict[str, Any], 
                                excel_data: Dict[str, Dict[str, Any]], 
                                available_orders: List[str]) -> Optional[MatchResult]:
        """
        Rapprochement croisé par références avec gestion d'erreurs
        """
        try:
            pdf_references = pdf.get('invoice_references', [])
            if not pdf_references:
                return None
            
            best_match = None
            best_confidence = 0
            
            for excel_order_key in available_orders:
                if excel_order_key not in excel_data:
                    continue
                    
                excel_order = excel_data[excel_order_key]
                excel_lines = excel_order.get('raw_lines', [])
                
                # Score de correspondance par références
                reference_score = 0
                matching_details = []
                
                for pdf_ref in pdf_references:
                    if not isinstance(pdf_ref, dict):
                        continue
                        
                    batch_id = pdf_ref.get('batch_id', '')
                    assignment_id = pdf_ref.get('assignment_id', '')
                    ref_key = pdf_ref.get('reference_key', '')
                    
                    for line in excel_lines:
                        if not isinstance(line, dict):
                            continue
                            
                        cost_center = str(line.get('cost_center', ''))
                        
                        # Différents types de correspondance
                        if ref_key == cost_center:
                            reference_score += 1.0  # Match exact
                            matching_details.append(f"Exact: {ref_key}")
                        elif batch_id in cost_center and assignment_id in cost_center:
                            reference_score += 0.8  # Match partiel fort
                            matching_details.append(f"Partiel: {batch_id}_{assignment_id}")
                        elif batch_id in cost_center or assignment_id in cost_center:
                            reference_score += 0.5  # Match partiel faible
                            matching_details.append(f"ID: {batch_id if batch_id in cost_center else assignment_id}")
                
                # Normalisation du score
                max_possible_score = len(pdf_references)
                normalized_score = reference_score / max_possible_score if max_possible_score > 0 else 0
                
                if normalized_score > best_confidence and normalized_score >= 0.4:
                    best_confidence = normalized_score
                    
                    # Calcul des montants
                    pdf_amount = pdf.get('total_net', 0)
                    excel_amount = excel_order.get('total_amount', 0)
                    amount_diff = abs(pdf_amount - excel_amount)
                    
                    # Validation par montant
                    amount_factor = self.calculate_amount_confidence(pdf_amount, excel_amount)
                    combined_confidence = (normalized_score * 0.6) + (amount_factor * 0.4)
                    
                    match_type = (MatchType.PERFECT_MATCH 
                                if amount_diff <= pdf_amount * (self.config['tolerance'] * 2)
                                else MatchType.DISCREPANCY)
                    
                    best_match = MatchResult(
                        match_type=match_type,
                        method=MatchMethod.REFERENCE_CROSS,
                        confidence=combined_confidence,
                        pdf_data=pdf,
                        excel_data=excel_order,
                        differences={
                            'amount_difference': amount_diff,
                            'amount_percentage': (amount_diff / pdf_amount) * 100 if pdf_amount > 0 else 0,
                            'reference_score': reference_score,
                            'matching_details': matching_details,
                            'match_method': 'reference_cross'
                        },
                        metadata={
                            'match_timestamp': datetime.now().isoformat(),
                            'cross_reference_matching': True,
                            'reference_confidence': normalized_score
                        }
                    )
            
            return best_match
            
        except Exception as e:
            self.logger.error(f"Erreur reference cross match pour PDF {pdf.get('filename', 'Unknown')}: {e}")
            return None
    
    def calculate_amount_confidence(self, amount1: float, amount2: float) -> float:
        """Calcule la confiance basée sur la proximité des montants"""
        try:
            if amount1 == 0 and amount2 == 0:
                return 1.0
            
            if amount1 == 0 or amount2 == 0:
                return 0.0
            
            # Différence relative
            max_amount = max(amount1, amount2)
            difference = abs(amount1 - amount2)
            
            if difference == 0:
                return 1.0
            
            # Confiance inversement proportionnelle à l'écart relatif
            relative_diff = difference / max_amount
            return max(0.0, 1.0 - relative_diff * 2)
            
        except Exception as e:
            self.logger.warning(f"Erreur calcul amount confidence: {e}")
            return 0.0


# Fonctions utilitaires pour les tests
def test_reconciliation_engine():
    """Test du moteur de rapprochement corrigé"""
    
    # Données de test basées sur les exemples réels
    test_pdf_data = [
        {
            'success': True,
            'filename': 'test1.pdf',
            'purchase_order': '5600025054',
            'total_net': 9.84,
            'invoice_id': '4949S0001',
            'supplier': 'Select T.T',
            'invoice_date': '2025/03/10',
            'main_reference': '4949_65744',
            'invoice_references': [{
                'batch_id': '4949',
                'assignment_id': '65744',
                'reference_key': '4949_65744'
            }]
        }
    ]
    
    test_excel_data = [
        {
            'is_valid': True,
            'order_number': '5600025054',
            'net_amount': 9.84,
            'cost_center': '0025033402',
            'collaborator': 'Test User',
            'source_filename': 'test.xlsx'
        }
    ]
    
    # Test du moteur corrigé
    engine = ReconciliationEngine()
    results = engine.perform_reconciliation(test_pdf_data, test_excel_data)
    
    print("Test rapprochement corrigé:")
    print(f"  Succès: {results.get('success', True)}")
    print(f"  Matches: {len(results.get('matches', []))}")
    print(f"  Écarts: {len(results.get('discrepancies', []))}")
    print(f"  Taux réussite: {results.get('summary', {}).get('matching_rate', 0):.1f}%")
    
    return results

if __name__ == "__main__":
    # Test du module si exécuté directement
    test_reconciliation_engine()

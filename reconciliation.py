"""
RECONCILIATION.PY
Moteur de rapprochement intelligent pour l'application Beeline
Rapprochement avancé entre données PDF et Excel avec algorithmes multiples
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
    SUPPLIER_DATE = "supplier_date"
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
    Moteur de rapprochement intelligent pour Beeline
    """
    
    def __init__(self, config: Optional[Dict[str, Any]] = None):
        self.setup_logging()
        
        # Configuration par défaut
        self.config = {
            'tolerance': 0.01,  # 1% de tolérance sur les montants
            'method': 'intelligent',  # Méthode de rapprochement
            'fuzzy_threshold': 0.8,  # Seuil pour correspondance floue
            'date_tolerance_days': 30,  # Tolérance en jours pour les dates
            'amount_weights': {
                'exact': 1.0,
                'close': 0.8,
                'distant': 0.3
            },
            'enable_machine_learning': False,  # ML pour améliorer les matches
            'min_confidence': 0.7  # Confiance minimum pour valider un match
        }
        
        # Mise à jour avec la config fournie
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
        
        Args:
            pdf_data: Données extraites des PDFs
            excel_data: Données traitées des Excel
            config: Configuration optionnelle
            
        Returns:
            Résultats complets du rapprochement
        """
        start_time = datetime.now()
        self.logger.info("🚀 Démarrage du rapprochement intelligent")
        
        # Mise à jour de la configuration si fournie
        if config:
            self.config.update(config)
        
        try:
            # 1. Préparation des données
            prepared_pdfs = self.prepare_pdf_data(pdf_data)
            aggregated_excel = self.prepare_excel_data(excel_data)
            
            # 2. Statistiques initiales
            self.stats['total_pdfs'] = len(prepared_pdfs)
            self.stats['total_excel_orders'] = len(aggregated_excel)
            
            # 3. Rapprochement selon la méthode configurée
            if self.config['method'] == 'intelligent':
                results = self.intelligent_reconciliation(prepared_pdfs, aggregated_excel)
            elif self.config['method'] == 'exact':
                results = self.exact_reconciliation(prepared_pdfs, aggregated_excel)
            elif self.config['method'] == 'partiel':
                results = self.partial_reconciliation(prepared_pdfs, aggregated_excel)
            else:
                results = self.intelligent_reconciliation(prepared_pdfs, aggregated_excel)
            
            # 4. Post-traitement et analyse
            results = self.post_process_results(results, prepared_pdfs, aggregated_excel)
            
            # 5. Génération du résumé
            results['summary'] = self.generate_summary(results)
            
            # 6. Métadonnées de traitement
            processing_time = (datetime.now() - start_time).total_seconds()
            self.stats['processing_time'] = processing_time
            
            results['metadata'] = {
                'processing_time': processing_time,
                'reconciliation_timestamp': datetime.now().isoformat(),
                'engine_version': '2.0.0',
                'config_used': self.config.copy(),
                'performance_stats': self.stats.copy()
            }
            
            self.logger.info(f"✅ Rapprochement terminé en {processing_time:.2f}s")
            return results
            
        except Exception as e:
            self.logger.error(f"❌ Erreur durant le rapprochement: {str(e)}")
            raise Exception(f"Erreur rapprochement: {str(e)}")
    
    def prepare_pdf_data(self, pdf_data: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """
        Prépare et nettoie les données PDF pour le rapprochement
        
        Args:
            pdf_data: Données brutes des PDFs
            
        Returns:
            Données PDF préparées
        """
        prepared = []
        
        for i, pdf in enumerate(pdf_data):
            if not pdf.get('success', True):
                continue
            
            # Extraction des champs essentiels
            prepared_pdf = {
                'index': i,
                'filename': pdf.get('filename', f'pdf_{i}'),
                'order_number': self.clean_order_number(pdf.get('purchase_order')),
                'invoice_id': pdf.get('invoice_id'),
                'amount': self.parse_amount(pdf.get('total_net', 0)),
                'invoice_date': self.parse_date(pdf.get('invoice_date')),
                'supplier': self.clean_supplier_name(pdf.get('supplier', '')),
                'client': pdf.get('client', ''),
                'validation': pdf.get('validation', {}),
                'data_completeness': pdf.get('data_completeness', {}),
                'raw_data': pdf  # Garder les données originales
            }
            
            # Calcul de la qualité des données
            prepared_pdf['data_quality_score'] = self.calculate_pdf_quality_score(prepared_pdf)
            
            # Seulement inclure les PDFs avec données suffisantes
            if prepared_pdf['order_number'] and prepared_pdf['amount'] > 0:
                prepared.append(prepared_pdf)
            else:
                self.logger.warning(f"PDF {prepared_pdf['filename']} exclu: données insuffisantes")
        
        self.logger.info(f"📄 {len(prepared)} PDFs préparés sur {len(pdf_data)} fournis")
        return prepared
    
    def prepare_excel_data(self, excel_data: List[Dict[str, Any]]) -> Dict[str, Dict[str, Any]]:
        """
        Prépare et agrège les données Excel par numéro de commande
        
        Args:
            excel_data: Données brutes des Excel
            
        Returns:
            Données Excel agrégées par commande
        """
        # Import du processeur Excel pour l'agrégation
        from excel_processor import ExcelProcessor
        processor = ExcelProcessor()
        
        # Agrégation par numéro de commande
        aggregated = processor.aggregate_by_order_number(excel_data)
        
        # Enrichissement des données agrégées
        for order_number, order_data in aggregated.items():
            # Nettoyage du numéro de commande
            clean_order = self.clean_order_number(order_number)
            if clean_order != order_number:
                aggregated[clean_order] = aggregated.pop(order_number)
                order_data['order_number'] = clean_order
            
            # Calcul de métriques supplémentaires
            order_data['data_quality_score'] = self.calculate_excel_quality_score(order_data)
            order_data['amount_per_collaborator'] = (
                order_data['total_amount'] / len(order_data['collaborators'])
                if order_data['collaborators'] else 0
            )
            
            # Détection de la période de facturation
            if order_data['statement_dates']:
                dates = [self.parse_date(d) for d in order_data['statement_dates']]
                valid_dates = [d for d in dates if d]
                if valid_dates:
                    order_data['billing_period_start'] = min(valid_dates)
                    order_data['billing_period_end'] = max(valid_dates)
        
        self.logger.info(f"📊 {len(aggregated)} commandes Excel préparées")
        return aggregated
    
    def intelligent_reconciliation(self, pdf_data: List[Dict[str, Any]], 
                                 excel_data: Dict[str, Dict[str, Any]]) -> Dict[str, Any]:
        """
        Rapprochement intelligent multi-niveaux
        
        Args:
            pdf_data: PDFs préparés
            excel_data: Excel agrégés
            
        Returns:
            Résultats du rapprochement
        """
        self.logger.info("🧠 Lancement du rapprochement intelligent")
        
        results = {
            'matches': [],
            'discrepancies': [],
            'unmatched_pdf': [],
            'unmatched_excel': list(excel_data.keys()),
            'match_details': []
        }
        
        # Phase 1: Rapprochement exact par numéro de commande
        self.logger.info("   Phase 1: Rapprochement exact")
        for pdf in pdf_data:
            match_result = self.try_exact_match(pdf, excel_data)
            
            if match_result:
                self.process_match_result(match_result, results)
            else:
                # Stocker temporairement pour les phases suivantes
                pdf['match_attempts'] = []
        
        # Phase 2: Rapprochement partiel pour les non-matchés
        self.logger.info("   Phase 2: Rapprochement partiel")
        unmatched_pdfs = [pdf for pdf in pdf_data if 'match_attempts' in pdf]
        
        for pdf in unmatched_pdfs:
            match_result = self.try_partial_match(pdf, excel_data, results['unmatched_excel'])
            
            if match_result and match_result.confidence >= self.config['min_confidence']:
                self.process_match_result(match_result, results)
            else:
                if match_result:
                    pdf['match_attempts'].append(('partial', match_result.confidence))
        
        # Phase 3: Rapprochement par montant (fuzzy)
        self.logger.info("   Phase 3: Rapprochement par montant")
        still_unmatched = [pdf for pdf in pdf_data if 'match_attempts' in pdf]
        
        for pdf in still_unmatched:
            match_result = self.try_amount_fuzzy_match(pdf, excel_data, results['unmatched_excel'])
            
            if match_result and match_result.confidence >= self.config['min_confidence']:
                self.process_match_result(match_result, results)
            else:
                if match_result:
                    pdf['match_attempts'].append(('amount_fuzzy', match_result.confidence))
        
        # Phase 4: Rapprochement contextuel (fournisseur + date)
        self.logger.info("   Phase 4: Rapprochement contextuel")
        final_unmatched = [pdf for pdf in pdf_data if 'match_attempts' in pdf]
        
        for pdf in final_unmatched:
            match_result = self.try_contextual_match(pdf, excel_data, results['unmatched_excel'])
            
            if match_result and match_result.confidence >= self.config['min_confidence']:
                self.process_match_result(match_result, results)
            else:
                # Définitivement non matché
                results['unmatched_pdf'].append(self.format_unmatched_pdf(pdf))
        
        self.logger.info(f"   🎯 Résultats: {len(results['matches'])} matches, {len(results['discrepancies'])} écarts")
        return results
    
    def try_exact_match(self, pdf: Dict[str, Any], 
                       excel_data: Dict[str, Dict[str, Any]]) -> Optional[MatchResult]:
        """
        Tente un rapprochement exact par numéro de commande
        
        Args:
            pdf: Données PDF
            excel_data: Données Excel agrégées
            
        Returns:
            Résultat du match ou None
        """
        pdf_order = pdf.get('order_number')
        if not pdf_order:
            return None
        
        # Recherche exacte
        if pdf_order in excel_data:
            excel_order = excel_data[pdf_order]
            
            # Calcul des différences
            amount_diff = abs(pdf['amount'] - excel_order['total_amount'])
            amount_tolerance = pdf['amount'] * self.config['tolerance']
            
            # Détermination du type de match
            if amount_diff <= amount_tolerance:
                match_type = MatchType.PERFECT_MATCH
                confidence = 1.0
            else:
                match_type = MatchType.DISCREPANCY
                # Confiance inversement proportionnelle à l'écart
                confidence = max(0.1, 1.0 - (amount_diff / pdf['amount']))
            
            return MatchResult(
                match_type=match_type,
                method=MatchMethod.EXACT_ORDER,
                confidence=confidence,
                pdf_data=pdf,
                excel_data=excel_order,
                differences={
                    'amount_difference': amount_diff,
                    'amount_percentage': (amount_diff / pdf['amount']) * 100 if pdf['amount'] > 0 else 0
                },
                metadata={
                    'match_timestamp': datetime.now().isoformat(),
                    'tolerance_used': self.config['tolerance']
                }
            )
        
        return None
    
    def try_partial_match(self, pdf: Dict[str, Any], 
                         excel_data: Dict[str, Dict[str, Any]], 
                         available_orders: List[str]) -> Optional[MatchResult]:
        """
        Tente un rapprochement partiel (début de numéro de commande)
        
        Args:
            pdf: Données PDF
            excel_data: Données Excel agrégées
            available_orders: Commandes Excel encore disponibles
            
        Returns:
            Résultat du match ou None
        """
        pdf_order = str(pdf.get('order_number', ''))
        if len(pdf_order) < 4:  # Trop court pour match partiel
            return None
        
        best_match = None
        best_confidence = 0
        
        for excel_order_key in available_orders:
            excel_order_str = str(excel_order_key)
            
            # Correspondance par début de chaîne
            similarity = self.calculate_string_similarity(pdf_order, excel_order_str)
            
            if similarity >= self.config['fuzzy_threshold']:
                excel_order = excel_data[excel_order_key]
                
                # Facteur de confiance basé sur la similarité et la différence de montant
                amount_factor = self.calculate_amount_confidence(pdf['amount'], excel_order['total_amount'])
                combined_confidence = (similarity * 0.6) + (amount_factor * 0.4)
                
                if combined_confidence > best_confidence:
                    best_confidence = combined_confidence
                    
                    amount_diff = abs(pdf['amount'] - excel_order['total_amount'])
                    match_type = (MatchType.PERFECT_MATCH 
                                if amount_diff <= pdf['amount'] * self.config['tolerance']
                                else MatchType.DISCREPANCY)
                    
                    best_match = MatchResult(
                        match_type=match_type,
                        method=MatchMethod.PARTIAL_ORDER,
                        confidence=combined_confidence,
                        pdf_data=pdf,
                        excel_data=excel_order,
                        differences={
                            'amount_difference': amount_diff,
                            'amount_percentage': (amount_diff / pdf['amount']) * 100 if pdf['amount'] > 0 else 0,
                            'order_similarity': similarity
                        },
                        metadata={
                            'match_timestamp': datetime.now().isoformat(),
                            'similarity_threshold': self.config['fuzzy_threshold']
                        }
                    )
        
        return best_match
    
    def try_amount_fuzzy_match(self, pdf: Dict[str, Any], 
                              excel_data: Dict[str, Dict[str, Any]], 
                              available_orders: List[str]) -> Optional[MatchResult]:
        """
        Tente un rapprochement par montant approximatif
        
        Args:
            pdf: Données PDF
            excel_data: Données Excel agrégées  
            available_orders: Commandes Excel encore disponibles
            
        Returns:
            Résultat du match ou None
        """
        pdf_amount = pdf['amount']
        if pdf_amount <= 0:
            return None
        
        best_match = None
        best_confidence = 0
        
        # Tolérance élargie pour le fuzzy matching
        extended_tolerance = pdf_amount * (self.config['tolerance'] * 5)  # 5x la tolérance normale
        
        for excel_order_key in available_orders:
            excel_order = excel_data[excel_order_key]
            excel_amount = excel_order['total_amount']
            
            amount_diff = abs(pdf_amount - excel_amount)
            
            if amount_diff <= extended_tolerance:
                # Confiance basée sur la proximité des montants
                confidence = 1.0 - (amount_diff / extended_tolerance)
                
                # Bonus si dates cohérentes
                if self.dates_are_coherent(pdf, excel_order):
                    confidence *= 1.2  # Boost de 20%
                
                # Bonus si fournisseur cohérent
                if self.suppliers_are_coherent(pdf, excel_order):
                    confidence *= 1.1  # Boost de 10%
                
                confidence = min(1.0, confidence)  # Cap à 1.0
                
                if confidence > best_confidence:
                    best_confidence = confidence
                    
                    match_type = (MatchType.PERFECT_MATCH 
                                if amount_diff <= pdf_amount * self.config['tolerance']
                                else MatchType.DISCREPANCY)
                    
                    best_match = MatchResult(
                        match_type=match_type,
                        method=MatchMethod.AMOUNT_FUZZY,
                        confidence=confidence,
                        pdf_data=pdf,
                        excel_data=excel_order,
                        differences={
                            'amount_difference': amount_diff,
                            'amount_percentage': (amount_diff / pdf_amount) * 100,
                            'extended_tolerance_used': extended_tolerance
                        },
                        metadata={
                            'match_timestamp': datetime.now().isoformat(),
                            'fuzzy_matching': True
                        }
                    )
        
        return best_match
    
    def try_contextual_match(self, pdf: Dict[str, Any], 
                           excel_data: Dict[str, Dict[str, Any]], 
                           available_orders: List[str]) -> Optional[MatchResult]:
        """
        Tente un rapprochement contextuel (fournisseur + période)
        
        Args:
            pdf: Données PDF
            excel_data: Données Excel agrégées
            available_orders: Commandes Excel encore disponibles
            
        Returns:
            Résultat du match ou None
        """
        pdf_supplier = pdf.get('supplier', '')
        pdf_date = pdf.get('invoice_date')
        
        if not pdf_supplier and not pdf_date:
            return None
        
        best_match = None
        best_confidence = 0
        
        for excel_order_key in available_orders:
            excel_order = excel_data[excel_order_key]
            confidence_factors = []
            
            # Facteur fournisseur
            if pdf_supplier and excel_order.get('suppliers'):
                supplier_similarity = max([
                    self.calculate_string_similarity(pdf_supplier, excel_supplier)
                    for excel_supplier in excel_order['suppliers']
                ])
                confidence_factors.append(('supplier', supplier_similarity, 0.4))
            
            # Facteur date
            if pdf_date:
                date_similarity = self.calculate_date_similarity(pdf_date, excel_order)
                confidence_factors.append(('date', date_similarity, 0.3))
            
            # Facteur montant (moins important dans cette méthode)
            amount_similarity = self.calculate_amount_confidence(pdf['amount'], excel_order['total_amount'])
            confidence_factors.append(('amount', amount_similarity, 0.3))
            
            # Calcul de la confiance pondérée
            if confidence_factors:
                total_weight = sum(weight for _, _, weight in confidence_factors)
                combined_confidence = sum(
                    score * (weight / total_weight) 
                    for _, score, weight in confidence_factors
                )
                
                if combined_confidence > best_confidence and combined_confidence >= 0.6:
                    best_confidence = combined_confidence
                    
                    amount_diff = abs(pdf['amount'] - excel_order['total_amount'])
                    match_type = (MatchType.PERFECT_MATCH 
                                if amount_diff <= pdf['amount'] * self.config['tolerance']
                                else MatchType.DISCREPANCY)
                    
                    best_match = MatchResult(
                        match_type=match_type,
                        method=MatchMethod.SUPPLIER_DATE,
                        confidence=combined_confidence,
                        pdf_data=pdf,
                        excel_data=excel_order,
                        differences={
                            'amount_difference': amount_diff,
                            'amount_percentage': (amount_diff / pdf['amount']) * 100 if pdf['amount'] > 0 else 0,
                            'confidence_factors': dict((name, score) for name, score, _ in confidence_factors)
                        },
                        metadata={
                            'match_timestamp': datetime.now().isoformat(),
                            'contextual_matching': True
                        }
                    )
        
        return best_match
    
    def process_match_result(self, match_result: MatchResult, results: Dict[str, Any]):
        """
        Traite et classe un résultat de match
        
        Args:
            match_result: Résultat du match
            results: Structure des résultats à mettre à jour
        """
        # Retirer de la liste des Excel non matchés
        excel_order_number = match_result.excel_data['order_number']
        if excel_order_number in results['unmatched_excel']:
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
            'pdf_file': match_result.pdf_data['filename'],
            'excel_order': excel_order_number
        })
        
        # Mise à jour des stats par méthode
        method_name = match_result.method.value
        if method_name not in self.stats['method_performance']:
            self.stats['method_performance'][method_name] = {'count': 0, 'avg_confidence': 0}
        
        method_stats = self.stats['method_performance'][method_name]
        method_stats['count'] += 1
        method_stats['avg_confidence'] = (
            (method_stats['avg_confidence'] * (method_stats['count'] - 1) + match_result.confidence)
            / method_stats['count']
        )
    
    def format_match_result(self, match_result: MatchResult) -> Dict[str, Any]:
        """
        Formate un résultat de match pour l'affichage
        
        Args:
            match_result: Résultat à formater
            
        Returns:
            Résultat formaté
        """
        return {
            'type': match_result.match_type.value,
            'method': match_result.method.value,
            'confidence': round(match_result.confidence, 3),
            'order_number': match_result.excel_data['order_number'],
            'pdf_file': match_result.pdf_data['filename'],
            'pdf_amount': match_result.pdf_data['amount'],
            'excel_amount': match_result.excel_data['total_amount'],
            'difference': match_result.differences.get('amount_difference', 0),
            'difference_percent': round(match_result.differences.get('amount_percentage', 0), 2),
            'collaborators': ', '.join(match_result.excel_data.get('collaborators', [])),
            'supplier': match_result.pdf_data.get('supplier', ''),
            'invoice_id': match_result.pdf_data.get('invoice_id', ''),
            'invoice_date': match_result.pdf_data.get('invoice_date', ''),
            'excel_line_count': match_result.excel_data.get('line_count', 0),
            'match_metadata': match_result.metadata
        }
    
    def format_unmatched_pdf(self, pdf: Dict[str, Any]) -> Dict[str, Any]:
        """
        Formate un PDF non matché
        
        Args:
            pdf: Données PDF
            
        Returns:
            PDF formaté avec raisons du non-match
        """
        # Analyse des raisons du non-match
        reasons = []
        
        if not pdf.get('order_number'):
            reasons.append("Numéro de commande manquant ou invalide")
        
        if pdf.get('amount', 0) <= 0:
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
            'filename': pdf['filename'],
            'order_number': pdf.get('order_number', 'Non trouvé'),
            'amount': pdf.get('amount', 0),
            'invoice_id': pdf.get('invoice_id', 'Non trouvé'),
            'invoice_date': pdf.get('invoice_date', 'Non trouvé'),
            'supplier': pdf.get('supplier', 'Non trouvé'),
            'data_quality_score': pdf.get('data_quality_score', 0),
            'reasons': reasons,
            'reason': ' | '.join(reasons)
        }
    
    # Méthodes utilitaires de calcul
    
    def clean_order_number(self, order_number: Any) -> Optional[str]:
        """Nettoie un numéro de commande"""
        if not order_number:
            return None
        
        # Conversion en string et nettoyage
        order_str = str(order_number).strip()
        
        # Suppression des caractères non numériques
        clean_order = re.sub(r'[^\d]', '', order_str)
        
        # Validation de la longueur
        if 8 <= len(clean_order) <= 12:
            return clean_order
        
        return None
    
    def clean_supplier_name(self, supplier: str) -> str:
        """Nettoie un nom de fournisseur"""
        if not supplier:
            return ""
        
        # Nettoyage basique
        cleaned = str(supplier).strip().lower()
        
        # Normalisation des noms connus
        if 'randstad' in cleaned:
            return 'Randstad'
        elif 'select' in cleaned and 't.t' in cleaned:
            return 'Select T.T.'
        
        return supplier.strip()
    
    def parse_amount(self, amount: Any) -> float:
        """Parse un montant en float"""
        if not amount:
            return 0.0
        
        try:
            if isinstance(amount, (int, float)):
                return float(amount)
            
            # Nettoyage du string
            amount_str = str(amount).replace(',', '.').replace(' ', '')
            amount_clean = re.sub(r'[^\d.-]', '', amount_str)
            
            return float(amount_clean) if amount_clean else 0.0
            
        except (ValueError, TypeError):
            return 0.0
    
    def parse_date(self, date_value: Any) -> Optional[datetime]:
        """Parse une date en objet datetime"""
        if not date_value:
            return None
        
        if isinstance(date_value, datetime):
            return date_value
        
        date_str = str(date_value).strip()
        date_formats = ['%Y/%m/%d', '%d/%m/%Y', '%Y-%m-%d', '%d-%m-%Y']
        
        for fmt in date_formats:
            try:
                return datetime.strptime(date_str, fmt)
            except ValueError:
                continue
        
        return None
    
    def calculate_string_similarity(self, str1: str, str2: str) -> float:
        """Calcule la similarité entre deux chaînes"""
        if not str1 or not str2:
            return 0.0
        
        return difflib.SequenceMatcher(None, str1.lower(), str2.lower()).ratio()
    
    def calculate_amount_confidence(self, amount1: float, amount2: float) -> float:
        """Calcule la confiance basée sur la proximité des montants"""
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
        return max(0.0, 1.0 - relative_diff * 2)  # *2 pour être plus strict
    
    def dates_are_coherent(self, pdf: Dict[str, Any], excel_order: Dict[str, Any]) -> bool:
        """Vérifie si les dates PDF et Excel sont cohérentes"""
        pdf_date = pdf.get('invoice_date')
        if not pdf_date:
            return False
        
        excel_dates = []
        if excel_order.get('billing_period_start'):
            excel_dates.append(excel_order['billing_period_start'])
        if excel_order.get('billing_period_end'):
            excel_dates.append(excel_order['billing_period_end'])
        
        if not excel_dates:
            return False
        
        # Vérifier si la date PDF est dans la tolérance des dates Excel
        tolerance_days = self.config['date_tolerance_days']
        
        for excel_date in excel_dates:
            if isinstance(excel_date, datetime):
                diff_days = abs((pdf_date - excel_date).days)
                if diff_days <= tolerance_days:
                    return True
        
        return False
    
    def suppliers_are_coherent(self, pdf: Dict[str, Any], excel_order: Dict[str, Any]) -> bool:
        """Vérifie si les fournisseurs PDF et Excel sont cohérents"""
        pdf_supplier = pdf.get('supplier', '').lower()
        excel_suppliers = [s.lower() for s in excel_order.get('suppliers', [])]
        
        if not pdf_supplier or not excel_suppliers:
            return False
        
        # Recherche de correspondance
        for excel_supplier in excel_suppliers:
            if self.calculate_string_similarity(pdf_supplier, excel_supplier) > 0.8:
                return True
        
        return False
    
    def calculate_date_similarity(self, pdf_date: datetime, excel_order: Dict[str, Any]) -> float:
        """Calcule la similarité temporelle"""
        if not pdf_date:
            return 0.0
        
        excel_dates = []
        if excel_order.get('billing_period_start'):
            excel_dates.append(excel_order['billing_period_start'])
        if excel_order.get('billing_period_end'):
            excel_dates.append(excel_order['billing_period_end'])
        
        if not excel_dates:
            return 0.0
        
        # Trouver la date Excel la plus proche
        min_diff_days = float('inf')
        for excel_date in excel_dates:
            if isinstance(excel_date, datetime):
                diff_days = abs((pdf_date - excel_date).days)
                min_diff_days = min(min_diff_days, diff_days)
        
        if min_diff_days == float('inf'):
            return 0.0
        
        # Similarité inversement proportionnelle à la différence
        max_tolerance = self.config['date_tolerance_days']
        if min_diff_days <= max_tolerance:
            return 1.0 - (min_diff_days / max_tolerance)
        else:
            return 0.0
    
    def calculate_pdf_quality_score(self, pdf: Dict[str, Any]) -> float:
        """Calcule un score de qualité pour un PDF"""
        score = 0.0
        weights = {
            'order_number': 0.4,
            'amount': 0.3,
            'invoice_id': 0.1,
            'supplier': 0.1,
            'invoice_date': 0.1
        }
        
        for field, weight in weights.items():
            if pdf.get(field):
                score += weight
        
        # Bonus pour validation réussie
        if pdf.get('validation', {}).get('is_valid', False):
            score *= 1.2
        
        # Bonus pour complétude élevée
        completeness = pdf.get('data_completeness', {}).get('overall_score', 0)
        score *= (0.8 + (completeness / 100) * 0.2)  # Entre 0.8 et 1.0
        
        return min(1.0, score)
    
    def calculate_excel_quality_score(self, excel_order: Dict[str, Any]) -> float:
        """Calcule un score de qualité pour une commande Excel"""
        score = 0.0
        
        # Critères de base
        if excel_order.get('order_number'):
            score += 0.4
        if excel_order.get('total_amount', 0) > 0:
            score += 0.3
        if excel_order.get('collaborators'):
            score += 0.2
        if excel_order.get('line_count', 0) > 0:
            score += 0.1
        
        # Bonus pour cohérence des données
        validity_rate = excel_order.get('validation_summary', {}).get('validity_rate', 100)
        score *= (validity_rate / 100)
        
        return min(1.0, score)
    
    # Méthodes de rapprochement alternatives
    
    def exact_reconciliation(self, pdf_data: List[Dict[str, Any]], 
                           excel_data: Dict[str, Dict[str, Any]]) -> Dict[str, Any]:
        """Rapprochement exact seulement"""
        self.logger.info("🎯 Rapprochement exact uniquement")
        
        results = {
            'matches': [],
            'discrepancies': [],
            'unmatched_pdf': [],
            'unmatched_excel': list(excel_data.keys()),
            'match_details': []
        }
        
        for pdf in pdf_data:
            match_result = self.try_exact_match(pdf, excel_data)
            
            if match_result:
                self.process_match_result(match_result, results)
            else:
                results['unmatched_pdf'].append(self.format_unmatched_pdf(pdf))
        
        return results
    
    def partial_reconciliation(self, pdf_data: List[Dict[str, Any]], 
                             excel_data: Dict[str, Dict[str, Any]]) -> Dict[str, Any]:
        """Rapprochement exact + partiel"""
        self.logger.info("🔍 Rapprochement exact + partiel")
        
        results = {
            'matches': [],
            'discrepancies': [],
            'unmatched_pdf': [],
            'unmatched_excel': list(excel_data.keys()),
            'match_details': []
        }
        
        # Phase 1: Exact
        for pdf in pdf_data:
            match_result = self.try_exact_match(pdf, excel_data)
            
            if match_result:
                self.process_match_result(match_result, results)
            else:
                pdf['needs_partial_match'] = True
        
        # Phase 2: Partiel pour les non-matchés
        unmatched_pdfs = [pdf for pdf in pdf_data if pdf.get('needs_partial_match')]
        
        for pdf in unmatched_pdfs:
            match_result = self.try_partial_match(pdf, excel_data, results['unmatched_excel'])
            
            if match_result and match_result.confidence >= self.config['min_confidence']:
                self.process_match_result(match_result, results)
            else:
                results['unmatched_pdf'].append(self.format_unmatched_pdf(pdf))
        
        return results
    
    # Post-traitement et analyse
    
    def post_process_results(self, results: Dict[str, Any], 
                           pdf_data: List[Dict[str, Any]], 
                           excel_data: Dict[str, Dict[str, Any]]) -> Dict[str, Any]:
        """Post-traite les résultats pour les enrichir"""
        
        # Ajout des Excel non matchés formatés
        unmatched_excel_formatted = []
        for excel_order_key in results['unmatched_excel']:
            if excel_order_key in excel_data:
                excel_order = excel_data[excel_order_key]
                unmatched_excel_formatted.append({
                    'order_number': excel_order_key,
                    'total_amount': excel_order['total_amount'],
                    'collaborators': ', '.join(excel_order.get('collaborators', [])),
                    'line_count': excel_order.get('line_count', 0),
                    'source_files': ', '.join(excel_order.get('source_files', [])),
                    'reason': 'Aucun PDF correspondant trouvé'
                })
        
        results['unmatched_excel'] = unmatched_excel_formatted
        
        # Calcul des totaux
        total_pdf_amount = sum(pdf['amount'] for pdf in pdf_data if pdf['amount'] > 0)
        total_excel_amount = sum(excel['total_amount'] for excel in excel_data.values())
        
        results['totals'] = {
            'total_pdf_amount': total_pdf_amount,
            'total_excel_amount': total_excel_amount,
            'difference': abs(total_pdf_amount - total_excel_amount)
        }
        
        # Analyse de la distribution des écarts
        if results['discrepancies']:
            discrepancy_amounts = [d['difference'] for d in results['discrepancies']]
            results['discrepancy_analysis'] = {
                'total_discrepancy': sum(discrepancy_amounts),
                'average_discrepancy': sum(discrepancy_amounts) / len(discrepancy_amounts),
                'max_discrepancy': max(discrepancy_amounts),
                'min_discrepancy': min(discrepancy_amounts)
            }
        
        return results
    
    def generate_summary(self, results: Dict[str, Any]) -> Dict[str, Any]:
        """Génère un résumé des résultats"""
        
        total_pdfs = self.stats['total_pdfs']
        total_excel = self.stats['total_excel_orders']
        matches = len(results['matches'])
        discrepancies = len(results['discrepancies'])
        unmatched_pdf = len(results['unmatched_pdf'])
        unmatched_excel = len(results['unmatched_excel'])
        
        # Calcul des taux
        matching_rate = (matches / total_pdfs * 100) if total_pdfs > 0 else 0
        discrepancy_rate = (discrepancies / total_pdfs * 100) if total_pdfs > 0 else 0
        coverage_rate = ((matches + discrepancies) / total_pdfs * 100) if total_pdfs > 0 else 0
        
        summary = {
            # Statistiques de base
            'total_invoices': total_pdfs,
            'total_excel_lines': sum(
                excel.get('line_count', 0) 
                for excel in results.get('matches', []) + results.get('discrepancies', [])
                if excel.get('excel_line_count')
            ),
            'total_excel_orders': total_excel,
            
            # Résultats du rapprochement
            'perfect_matches': matches,
            'discrepancies': discrepancies,
            'unmatched_pdf': unmatched_pdf,
            'unmatched_excel': unmatched_excel,
            
            # Taux de performance
            'matching_rate': round(matching_rate, 2),
            'discrepancy_rate': round(discrepancy_rate, 2),
            'coverage_rate': round(coverage_rate, 2),
            
            # Montants
            'total_amount': results.get('totals', {}).get('total_pdf_amount', 0),
            'total_discrepancy_amount': results.get('discrepancy_analysis', {}).get('total_discrepancy', 0),
            
            # Performance des méthodes
            'method_performance': self.stats['method_performance'],
            
            # Qualité globale
            'quality_assessment': self.assess_overall_quality(matching_rate, coverage_rate, results)
        }
        
        # Mise à jour des statistiques finales
        self.stats.update({
            'unmatched_pdf': unmatched_pdf,
            'unmatched_excel': unmatched_excel
        })
        
        return summary
    
    def assess_overall_quality(self, matching_rate: float, coverage_rate: float, 
                             results: Dict[str, Any]) -> Dict[str, Any]:
        """Évalue la qualité globale du rapprochement"""
        
        # Calcul du score de qualité (0-100)
        quality_score = 0
        
        # 40 points pour le taux de matching
        quality_score += min(40, matching_rate * 0.4)
        
        # 30 points pour la couverture
        quality_score += min(30, coverage_rate * 0.3)
        
        # 20 points pour la précision (peu d'écarts importants)
        if results.get('discrepancy_analysis'):
            avg_discrepancy = results['discrepancy_analysis'].get('average_discrepancy', 0)
            total_amount = results.get('totals', {}).get('total_pdf_amount', 1)
            discrepancy_rate = (avg_discrepancy / total_amount) * 100 if total_amount > 0 else 0
            
            if discrepancy_rate <= 1:  # Moins de 1% d'écart moyen
                quality_score += 20
            elif discrepancy_rate <= 5:  # Moins de 5%
                quality_score += 15
            elif discrepancy_rate <= 10:  # Moins de 10%
                quality_score += 10
        else:
            quality_score += 20  # Pas d'écarts = parfait
        
        # 10 points pour la confiance moyenne
        if results.get('match_details'):
            avg_confidence = sum(
                detail['confidence'] for detail in results['match_details']
            ) / len(results['match_details'])
            quality_score += avg_confidence * 10
        
        # Grade basé sur le score
        if quality_score >= 90:
            grade = 'A'
            assessment = 'Excellent'
        elif quality_score >= 80:
            grade = 'B'
            assessment = 'Très bon'
        elif quality_score >= 70:
            grade = 'C'
            assessment = 'Correct'
        elif quality_score >= 60:
            grade = 'D'
            assessment = 'Passable'
        else:
            grade = 'F'
            assessment = 'Insuffisant'
        
        return {
            'score': round(quality_score, 1),
            'grade': grade,
            'assessment': assessment,
            'recommendations': self.generate_recommendations(matching_rate, coverage_rate, results)
        }
    
    def generate_recommendations(self, matching_rate: float, coverage_rate: float, 
                               results: Dict[str, Any]) -> List[str]:
        """Génère des recommandations d'amélioration"""
        recommendations = []
        
        if matching_rate < 80:
            recommendations.append(
                "Taux de rapprochement faible. Vérifiez la cohérence des numéros de commande entre PDF et Excel."
            )
        
        if coverage_rate < 90:
            recommendations.append(
                "Couverture insuffisante. Contrôlez la complétude des données sources."
            )
        
        if len(results.get('unmatched_pdf', [])) > 0:
            recommendations.append(
                f"{len(results['unmatched_pdf'])} PDF(s) non rapproché(s). Vérifiez l'extraction des données PDF."
            )
        
        if len(results.get('unmatched_excel', [])) > 0:
            recommendations.append(
                f"{len(results['unmatched_excel'])} commande(s) Excel non rapprochée(s). Vérifiez la synchronisation des données."
            )
        
        if results.get('discrepancy_analysis'):
            avg_discrepancy = results['discrepancy_analysis'].get('average_discrepancy', 0)
            if avg_discrepancy > 50:
                recommendations.append(
                    f"Écarts importants détectés (moyenne: {avg_discrepancy:.2f}€). Contrôlez les calculs et taux appliqués."
                )
        
        # Recommandations positives
        if matching_rate >= 95 and coverage_rate >= 95:
            recommendations.append(
                "Excellent taux de rapprochement ! Le processus fonctionne de manière optimale."
            )
        
        if not recommendations:
            recommendations.append("Processus de rapprochement fonctionnel. Aucune action prioritaire requise.")
        
        return recommendations

# Fonctions utilitaires pour les tests
def test_reconciliation_engine():
    """Test du moteur de rapprochement"""
    
    # Données de test
    test_pdf_data = [
        {
            'success': True,
            'filename': 'test1.pdf',
            'purchase_order': '5600013960',
            'total_net': 1000.50,
            'invoice_id': '5118S0001',
            'supplier': 'Randstad',
            'invoice_date': '2025/09/22'
        }
    ]
    
    test_excel_data = [
        {
            'is_valid': True,
            'order_number': '5600013960',
            'net_amount': 1000.50,
            'collaborator': 'Test User',
            'source_filename': 'test.xlsx'
        }
    ]
    
    # Test du moteur
    engine = ReconciliationEngine()
    results = engine.perform_reconciliation(test_pdf_data, test_excel_data)
    
    print("Test rapprochement:")
    print(f"  Matches: {len(results['matches'])}")
    print(f"  Écarts: {len(results['discrepancies'])}")
    print(f"  Taux réussite: {results['summary']['matching_rate']:.1f}%")
    
    return results

if __name__ == "__main__":
    # Test du module si exécuté directement
    test_reconciliation_engine()

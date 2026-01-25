"""
text_corrector.py - Module de correction orthographique et grammaticale
"""

from typing import List, Optional
import logging

logger = logging.getLogger(__name__)

class TextCorrector:
    """Classe pour la correction orthographique et grammaticale."""
    
    def __init__(self, config):
        self.config = config
        self.cohere_client = None
        self._init_cohere()
    
    def _init_cohere(self):
        """Initialise le client Cohere."""
        try:
            self.cohere_client = self.config.get_cohere_client()
        except Exception as e:
            logger.warning(f"Impossible d'initialiser Cohere: {e}")
            self.cohere_client = None
    
    def correct_paragraphs_batch(
        self, 
        paragraphs: List, 
        batch_size: Optional[int] = None,
        whitelist: Optional[List[str]] = None
    ) -> List:
        """
        Corrige les paragraphes par lots pour optimiser les appels API.
        
        Args:
            paragraphs: Liste des objets paragraphes à corriger
            batch_size: Taille des lots (utilise la config par défaut si None)
            whitelist: Liste de mots à ne pas corriger
        
        Returns:
            Liste des paragraphes corrigés
        """
        if not self.cohere_client:
            logger.warning("Client Cohere non disponible, retour des paragraphes non corrigés")
            return paragraphs
        
        batch_size = batch_size or self.config.BATCH_SIZE
        whitelist = whitelist or self.config.CORRECTION_WHITELIST
        
        # Filtrer les paragraphes non vides
        valid_paragraphs = [(i, p) for i, p in enumerate(paragraphs) if p.text.strip()]
        
        if not valid_paragraphs:
            return paragraphs
        
        corrected_count = 0
        total_batches = (len(valid_paragraphs) + batch_size - 1) // batch_size
        
        for batch_idx in range(0, len(valid_paragraphs), batch_size):
            batch = valid_paragraphs[batch_idx:batch_idx + batch_size]
            current_batch_num = (batch_idx // batch_size) + 1
            
            logger.info(f"Traitement du lot {current_batch_num}/{total_batches}")
            
            # Corriger le lot
            corrected_batch = self._correct_batch(batch, whitelist)
            corrected_count += len(corrected_batch)
        
        logger.info(f"Correction terminée : {corrected_count} paragraphes traités")
        return paragraphs
    
    def _correct_batch(self, batch: List[tuple], whitelist: List[str]) -> List:
        """
        Corrige un lot de paragraphes.
        
        Args:
            batch: Liste de tuples (index, paragraphe)
            whitelist: Liste de mots à ne pas corriger
        
        Returns:
            Liste des paragraphes corrigés
        """
        if not batch:
            return []
        
        # Extraire les textes
        texts = [para.text.strip() for _, para in batch]
        
        # Joindre les textes avec le séparateur
        joined_text = self.config.BATCH_SEPARATOR.join(texts)
        
        # Corriger via l'API
        corrected_text = self._call_correction_api(joined_text, whitelist)
        
        if not corrected_text:
            return []
        
        # Séparer les textes corrigés
        corrected_texts = corrected_text.split(self.config.BATCH_SEPARATOR)
        
        # Appliquer les corrections aux paragraphes originaux
        corrected_paragraphs = []
        for i, corrected in enumerate(corrected_texts):
            if i < len(batch):
                _, para = batch[i]
                self._update_paragraph_text(para, corrected.strip())
                corrected_paragraphs.append(para)
        
        return corrected_paragraphs
    
    def _call_correction_api(self, text: str, whitelist: List[str]) -> Optional[str]:
        """
        Appelle l'API de correction.
        
        Args:
            text: Texte à corriger
            whitelist: Liste de mots à ne pas corriger
        
        Returns:
            Texte corrigé ou None en cas d'erreur
        """
        if not self.cohere_client:
            return None
        
        # Mettre à jour la whitelist dans la config temporairement
        original_whitelist = self.config.CORRECTION_WHITELIST
        self.config.CORRECTION_WHITELIST = whitelist
        
        try:
            prompt = self.config.get_correction_prompt(text)
            
            response = self.cohere_client.chat(
                model=self.config.COHERE_MODEL,
                message=prompt,
                temperature=0.0
            )
            
            return response.text.strip()
            
        except Exception as e:
            logger.error(f"Erreur lors de la correction: {e}")
            logger.debug(f"Texte non corrigé: {text[:200]}...")
            return text
        finally:
            # Restaurer la whitelist originale
            self.config.CORRECTION_WHITELIST = original_whitelist
    
    def _update_paragraph_text(self, paragraph, new_text: str):
        """
        Met à jour le texte d'un paragraphe en conservant le formatage.
        
        Args:
            paragraph: Objet paragraphe à mettre à jour
            new_text: Nouveau texte corrigé
        """
        runs = paragraph.runs
        
        if len(runs) == 1:
            # Un seul run : remplacement direct
            runs[0].text = new_text
        elif len(runs) == 0:
            # Aucun run : créer un nouveau run
            paragraph.add_run(new_text)
        else:
            # Plusieurs runs : répartir le texte en conservant le formatage
            self._distribute_text_across_runs(runs, new_text)
    
    def _distribute_text_across_runs(self, runs, new_text: str):
        """
        Distribue le texte corrigé à travers les runs existants.
        
        Args:
            runs: Liste des runs du paragraphe
            new_text: Texte à distribuer
        """
        # Calculer la longueur proportionnelle pour chaque run
        total_length = sum(len(run.text) for run in runs)
        if total_length == 0:
            runs[0].text = new_text
            for run in runs[1:]:
                run.text = ""
            return
        
        remaining_text = new_text
        for i, run in enumerate(runs):
            if not remaining_text:
                run.text = ""
                continue
            
            # Calculer la proportion de texte pour ce run
            if i == len(runs) - 1:
                # Dernier run : prendre tout le reste
                run.text = remaining_text
                remaining_text = ""
            else:
                proportion = len(run.text) / total_length
                chars_to_take = int(len(new_text) * proportion)
                run.text = remaining_text[:chars_to_take]
                remaining_text = remaining_text[chars_to_take:]
    
    def correct_single_text(self, text: str, whitelist: Optional[List[str]] = None) -> str:
        """
        Corrige un texte unique.
        
        Args:
            text: Texte à corriger
            whitelist: Liste de mots à ne pas corriger
        
        Returns:
            Texte corrigé
        """
        if not text.strip():
            return text
        
        whitelist = whitelist or self.config.CORRECTION_WHITELIST
        corrected = self._call_correction_api(text, whitelist)
        
        return corrected if corrected else text
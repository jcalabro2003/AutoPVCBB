"""
utils.py - Fonctions utilitaires pour le traitement du texte et LaTeX
"""

import re
from typing import Dict, List, Optional

class TextProcessor:
    """Classe pour le traitement et la transformation du texte."""
    
    def __init__(self, config):
        self.config = config
    
    def escape_latex(self, text: str) -> str:
        """Échappe les caractères spéciaux LaTeX pour éviter les erreurs."""
        if not text:
            return text
            
        for char, replacement in self.config.LATEX_SPECIAL_CHARS.items():
            text = text.replace(char, replacement)
        return text
    
    def replace_abbreviations(self, text: str, type=["begin", "end"]) -> str:
        """Remplace les abréviations par leur forme complète."""
        if not text:
            return text
        
        # Ajoute une ponctuation finale si nécessaire
        if "end" in type and not text.rstrip().endswith((".", "?", "!")):
            text = text.rstrip() + "."
        
        # Met en majuscule la première lettre
        if text and "begin" in type:
            text.strip()
            text = text[0].upper() + text[1:]
        
        # Remplace les abréviations
        for pattern, replacement in self.config.ABBREVIATIONS.items():
            def replace_match(match):
                word = match.group(0)
                return replacement.capitalize() if word[0].isupper() else replacement
            
            text = re.sub(pattern, replace_match, text, flags=re.IGNORECASE)
        
        return text
    
    def capitalize_first_letter(self, text: str) -> str:
        """Met en majuscule la première lettre d'un texte."""
        if text:
            return text[0].upper() + text[1:]
        return text
    
    def ensure_punctuation(self, text: str) -> str:
        """S'assure que le texte se termine par une ponctuation."""
        if text and not text.rstrip().endswith((".", "?", "!")):
            return text.rstrip() + "."
        return text


class DocumentParser:
    """Classe pour parser les éléments du document."""
    
    @staticmethod
    def parse_title(title: str) -> Dict[str, str]:
        """
        Parse le titre du document et extrait les composants.
        
        Args:
            title: Titre du document, ex: "PV RC 7 - Anno LIX - 2025-01-27"
        
        Returns:
            Dict contenant les composants extraits
        
        Raises:
            ValueError: Si le format du titre n'est pas reconnu
        """
        pattern = r"PV RC (\d+) - Anno (LIX|[IVXLCDM]+) - (\d{4})-(\d{2})-(\d{2})"
        match = re.match(pattern, title)
        
        if not match:
            raise ValueError(f"Le format du titre '{title}' n'est pas reconnu")
        
        numero, anno, annee, mois, jour = match.groups()
        
        return {
            'numero_reunion': numero,
            'anno': anno,
            'annee': annee,
            'mois': mois,
            'jour': jour,
            'date_formatee': f"{jour}/{mois}/{annee}"
        }
    
    @staticmethod
    def calculate_academic_year(year: str, month: str) -> str:
        """Calcule l'année académique basée sur l'année et le mois."""
        year_int = int(year)
        month_int = int(month)
        
        if month_int > 6:
            return f"{year} - {year_int + 1}"
        else:
            return f"{year_int - 1} - {year}"
    
    @staticmethod
    def is_section_header(text: str) -> bool:
        """Détermine si un texte est un en-tête de section."""
        return bool(re.match(r"^\d+\)", text))
    
    @staticmethod
    def is_subsection_header(text: str) -> bool:
        """Détermine si un texte est un en-tête de sous-section."""
        return bool(re.match(r"^[a-zA-Z]\)", text))
    
    @staticmethod
    def extract_section_title(text: str) -> str:
        """Extrait le titre d'une section à partir du texte."""
        if DocumentParser.is_section_header(text) or DocumentParser.is_subsection_header(text):
            return text.split(")", 1)[1].strip()
        return text
    
    @staticmethod
    def extract_sections_list(paragraphs) -> List[str]:
        """Extrait la liste de toutes les sections du document."""
        sections = []
        subsections = []
        for para in paragraphs:
            text = para.text.strip()
            if DocumentParser.is_section_header(text):
                section_title = DocumentParser.extract_section_title(text)
                sections.append(section_title)
            elif DocumentParser.is_subsection_header(text):
                subsection_title = DocumentParser.extract_section_title(text)
                subsections.append(section_title)
                subsections.append(subsection_title)
        return sections, subsections
    



class TableProcessor:
    """Classe pour le traitement des tableaux."""
    
    @staticmethod
    def extract_table_data(table, text_processor: TextProcessor) -> List[List[str]]:
        """
        Extrait les données d'un tableau avec formatage.
        
        Args:
            table: Objet table de python-docx
            text_processor: Instance de TextProcessor pour l'échappement LaTeX
        
        Returns:
            Liste de listes représentant les données du tableau
        """
        table_data = []
        
        for row in table.rows:
            row_data = []
            for cell in row.cells:
                if cell.paragraphs:
                    formatted_texts = []
                    for para in cell.paragraphs:
                        formatted_text = TableProcessor._format_paragraph(para, text_processor)
                        if formatted_text:
                            formatted_texts.append(formatted_text)
                    cell_text = " ".join(formatted_texts)
                else:
                    cell_text = ""
                row_data.append(cell_text)
            table_data.append(row_data)
        
        return table_data
    
    @staticmethod
    def _format_paragraph(para, text_processor: TextProcessor) -> str:
        """Formate un paragraphe en conservant le style (gras, italique)."""
        latex_text = ""
        for run in para.runs:
            run_text = text_processor.escape_latex(run.text.strip())
            if run.bold:
                run_text = f"\\textbf{{{run_text}}}"
            if run.italic:
                run_text = f"\\textit{{{run_text}}}"
            if run_text:
                latex_text += run_text + " "
        return latex_text.strip()
    
    @staticmethod
    def remove_duplicate_columns(table_data: List[List[str]]) -> List[List[str]]:
        """
        Supprime les colonnes en double dans un tableau.
        
        Args:
            table_data: Données du tableau
        
        Returns:
            Données du tableau sans colonnes dupliquées
        """
        if not table_data or len(table_data[0]) <= 1:
            return table_data
        
        num_cols = len(table_data[0])
        columns_to_remove = set()
        
        # Identifier les colonnes dupliquées
        for col1 in range(num_cols - 1):
            for col2 in range(col1 + 1, num_cols):
                if all(row[col1] == row[col2] for row in table_data):
                    columns_to_remove.add(col2)
        
        # Filtrer les colonnes
        return [
            [cell for i, cell in enumerate(row) if i not in columns_to_remove]
            for row in table_data
        ]
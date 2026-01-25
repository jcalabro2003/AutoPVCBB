"""
latex_generator.py - Générateur de code LaTeX
"""

from typing import List, Dict, Optional
from utils import TextProcessor, DocumentParser
import logging


logger = logging.getLogger(__name__)

class LaTeXGenerator:
    """Classe pour générer le code LaTeX."""
    
    def __init__(self, config, text_processor: TextProcessor):
        self.config = config
        self.text_processor = text_processor
    
    def generate_document_header(self) -> str:
        """Génère l'en-tête du document LaTeX."""
        lines = [
            "\\documentclass{article}",
            *self.config.LATEX_PACKAGES,
            *self.config.LATEX_SETTINGS,
            "\\begin{document}\n"
        ]
        return "\n".join(lines)
    
    def generate_document_footer(self) -> str:
        """Génère le pied de page du document LaTeX."""
        return "\\end{document}\n"
    
    def generate_title_header(self, title: str) -> str:
        """
        Génère l'en-tête LaTeX avec les informations du titre.
        
        Args:
            title: Titre du document
        
        Returns:
            Code LaTeX pour l'en-tête
        """
        try:
            components = DocumentParser.parse_title(title)
            annee_univ = DocumentParser.calculate_academic_year(
                components['annee'], 
                components['mois']
            )
            
            return (
                f"\\fancyhead[L]{{CBB - Anno {components['anno']} \\hfill {annee_univ} \n\n}} "
                f"\\fancyhead[R]{{Réunion Comité n° {components['numero_reunion']} "
                f"\\hfill {components['date_formatee']}}}"
            )
        except ValueError as e:
            # En cas d'erreur de parsing, retourner un en-tête simple
            logger.error(e)
            return f"\\fancyhead[C]{{{self.text_processor.escape_latex(title)}}}"
    
    def generate_title_section(self, title: str) -> str:
        """Génère le titre principal centré."""
        escaped_title = self.text_processor.escape_latex(title)
        return (
            "\\begin{center}\n"
            f"\\LARGE \\textbf{{{escaped_title}}}\\\\ \n"
            "\\end{center}\n\n"
        )
    
    def generate_toc(self, sections: List[str], subsections: List[str]) -> str:
        """
        Génère la table des matières (ordre du jour).
        
        Args:
            sections: Liste des titres de sections

        
        Returns:
            Code LaTeX pour la table des matières
        """
        if not sections:
            return ""
        
        lines = [
            "\\begin{center}",
            "\\section*{\\hspace{-1.5cm}Ordre du jour}",
            "\\hspace*{-0.5cm}\\begin{varwidth}{\\textwidth}",
        ]
        
        for section in sections:
            escaped_section = self.text_processor.escape_latex(section)
            escaped_section = self.text_processor.capitalize_first_letter(escaped_section)
            lines.append(f"\\textbf{{- {escaped_section}}}\\\\ ")
            # Ajouter les sous-sections associées
            while subsections and subsections[0].startswith(section):
                subsections.pop(0)
                sub = subsections.pop(0)     
                escaped_sub = self.text_processor.escape_latex(sub)
                lines.append(f"\\hspace*{{0.8cm}} - \\textbf{{{escaped_sub}}}\\\\ ")
        
        lines.append("\\end{varwidth}")

        lines.append("\\end{center}\n\n")
        
        return "\n".join(lines)
    
    def generate_present_section(self, names: List[str]) -> str:
        """
        Génère la section des personnes présentes.
        
        Args:
            names: Liste des noms des personnes présentes
        
        Returns:
            Code LaTeX pour la section des présents
        """
        if not names:
            return ""
        
        lines = [
            "\\section*{Camarades présents :}",
            "",
            "\\begin{multicols}{2}"
        ]
        for name in names:
            escaped_name = self.text_processor.escape_latex(name.strip())
            if escaped_name:
                if names.index(name) < len(names) - 1:
                    lines.append(f" {escaped_name}\\\\ ")
                else:
                    lines.append(f" {escaped_name} ")

        
        lines.append("\\end{multicols}\n")
        
        return "\n".join(lines)
    
    def generate_section(self, title: str) -> str:
        """Génère une section."""
        escaped_title = self.text_processor.escape_latex(title)
        escaped_title = self.text_processor.capitalize_first_letter(escaped_title)
        return f"\\section{{{escaped_title}}}\n\n"
    
    def generate_subsection(self, title: str) -> str:
        """Génère une sous-section."""
        escaped_title = self.text_processor.escape_latex(title)
        escaped_title = self.text_processor.capitalize_first_letter(escaped_title)
        return f"\\subsection*{{{escaped_title}}}\n\n"
    
    def generate_paragraph(self, text: str, runs=None) -> str:
        """
        Génère un paragraphe avec formatage optionnel.
        
        Args:
            text: Texte du paragraphe
            runs: Runs du paragraphe pour conserver le formatage
        
        Returns:
            Code LaTeX pour le paragraphe
        """
        if not text:
            return ""
        
        # Vérifier si c'est un paragraphe avec définition (contient ":")
        if ":" in text and runs is None:
            parts = text.split(":", 1)
            bold_part = f"\\textbf{{{self.text_processor.escape_latex(parts[0].strip())}}}"
            remaining = self.text_processor.escape_latex(parts[1].strip())
            remaining = self.text_processor.replace_abbreviations(remaining)
            return f"{bold_part} : {remaining}\n\n"
        
        # Paragraphe normal avec formatage des runs si disponible
        if runs:
            return self._format_runs(runs) + "\n\n"
        
        # Paragraphe simple
        processed = self.text_processor.replace_abbreviations(text)
        return self.text_processor.escape_latex(processed) + "\n\n"
    
    def _format_runs(self, runs) -> str:
        """Formate les runs d'un paragraphe en conservant le style."""
        latex_text = ""
        for run in runs:
            run_text = self.text_processor.escape_latex(run.text.strip())
            run_text = self.text_processor.replace_abbreviations(run_text)
            
            if run.bold:
                run_text = f"\\textbf{{{run_text}}}"
            if run.italic:
                run_text = f"\\textit{{{run_text}}}"
            
            if run_text:
                latex_text += run_text + " "
        
        return latex_text.strip()
    
    def generate_table(self, table_data: List[List[str]]) -> str:
        """
        Génère le code LaTeX pour un tableau.
        
        Args:
            table_data: Données du tableau
        
        Returns:
            Code LaTeX pour le tableau
        """
        if not table_data or not table_data[0]:
            return ""
        
        num_cols = len(table_data[0])
        
        lines = [
            "\\begin{table}[h]",
            "\\centering",
            "\\begin{tabular}{|" + " | ".join(["c"] * num_cols) + " |}",
            "\\hline"
        ]
        
        for row in table_data:
            lines.append(" & ".join(row) + " \\\\")
            lines.append("\\hline")
        
        lines.extend([
            "\\end{tabular}",
            "\\end{table}\n"
        ])
        
        return "\n".join(lines)
    
    def generate_logo_section(self) -> str:
        """Génère la section avec le logo en bas de page."""
        return (
            "\\vspace{\\fill}\n"
            "\\begin{center}\n"
            f"\\includegraphics[height=\\dimexpr\\textheight-\\pagetotal\\relax]{{{self.config.DEFAULT_LOGO_PATH}}}\n"
            "\\end{center}\n"
            "\\newpage\n"
        )
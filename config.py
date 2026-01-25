"""
config.py - Configuration centralisée pour le convertisseur
"""

import os
from typing import Dict, List, Optional

class Config:
    """Configuration centralisée pour le convertisseur DocX vers LaTeX."""
    
    def __init__(self):
        # API Configuration
        self.COHERE_API_KEY = os.getenv('COHERE_API_KEY', 'nlyRQZmMg67jyS0RmN3wofNlBG74V12gIjP0EV8L')
        self.COHERE_MODEL = 'command-a-03-2025'
        
        # Batch processing
        self.BATCH_SIZE = 100
        self.BATCH_SEPARATOR = "#SEP#"
        
        # LaTeX settings
        self.LATEX_PACKAGES = [
            "\\usepackage[T1]{fontenc}",
            "\\usepackage[utf8]{inputenc}",
            "\\usepackage[margin=1.2in]{geometry}",
            "\\geometry{a4paper}",
            "\\usepackage{fancyhdr}",
            "\\usepackage{multicol}",
            "\\usepackage{graphicx}",
            "\\usepackage{float}",
            "\\usepackage{varwidth}",
            "\\usepackage{textcomp}",
            "\\usepackage{csquotes}",
            "\\usepackage[gen]{eurosym}"
        ]
        
        self.LATEX_SETTINGS = [
            "\\pagestyle{fancy}",
            "\\setlength{\\headheight}{22.5pt}",
            "\\setlength{\\parindent}{0pt}",
            "\\setlength{\\parskip}{1em}"
        ]
        
        # Text replacements (loaded from file `abbreviations.txt`)
        self.ABBREVIATIONS: Dict[str, str] = self._read_kv_file('abbreviations.txt', default={
            r'\bitw\b': 'interview',
            r'\bdeleg\b': 'délégation',
            r'\bdéleg\b': 'délégation',
            r'\bdélég\b': 'délégation',
            r'\bqqch\b': 'quelque chose',
            r'\bqqun\b': "quelqu'un",
            r'\bpcq\b': "parce que",
            r'\bprez\b': "président",
            r'\btrez\b': "trésorier",
            r'\bvp\b': "vice-président"
        })

        # Special LaTeX characters (loaded from file `special_chars.txt`)
        self.LATEX_SPECIAL_CHARS: Dict[str, str] = self._read_kv_file('special_chars.txt', default={
            '&': r'\&',
            '%': r'\%',
            '$': r'\$',
            '#': r'\#',
            '_': r'\_',
            '{': r'\{',
            '}': r'\}',
            '~': r'\textasciitilde{}',
            '^': r'\textasciicircum{}',
            '€': r'\euro{}',
            '<': r'\textless{}',
            '>': r'\textgreater{}',
            '°': r'\textdegree{}'
        })
        # Document paths
        self.DEFAULT_LOGO_PATH = "../logo.png"
        
        # Whitelist words for correction (loaded from file `whitelist.txt`)
        self.CORRECTION_WHITELIST: List[str] = self._read_list_file('whitelist.txt', default=[
            "Cm !", "Cs !", "CM !", "CS !", "F.", "le X", "CBB", "io vivat",
            "PGCA", "CBBQ", "CMF", "XX", "XXX", "XXXX", "FM", "tapette",
            "réunion ex", "band", "peye", "Sam’saoule"
        ])

        # Prompt template (loaded from file `prompt.txt`)
        self.CORRECTION_PROMPT_TEMPLATE: Optional[str] = self._read_prompt_file('prompt.txt', default=None)
    
    def get_cohere_client(self):
        """Retourne une instance du client Cohere."""
        try:
            import cohere
            return cohere.Client(self.COHERE_API_KEY)
        except ImportError:
            raise ImportError("Le module 'cohere' n'est pas installé. Installez-le avec: pip install cohere")
    
    def get_correction_prompt(self, text: str) -> str:
        """Génère le prompt pour la correction orthographique."""
        whitelist_str = ", ".join(self.CORRECTION_WHITELIST) if self.CORRECTION_WHITELIST else "aucun"

        template = self.CORRECTION_PROMPT_TEMPLATE or (
            "Corrige le texte suivant en français :\n"
            "- Corrige les fautes d'orthographe, de grammaire, de ponctuation, de conjugaison et la concordance des temps.\n"
            "- Améliore la syntaxe et la clarté.\n"
            "- Ne modifie pas les noms propres, les anglicismes ni le latin (ex: io vivat).\n"
            "- Ne modifie pas les mots suivants : {whitelist}.\n\n"
            "Ta réponse doit UNIQUEMENT contenir le texte corrigé, sans explications ni commentaires.\n\n"
            "Texte à corriger :\n\n{text}\n"
        )

        return template.format(text=text, whitelist=whitelist_str)

    # --- Helpers pour charger les fichiers éditables ---
    def _resource_path(self, filename: str) -> str:
        return os.path.join(os.path.dirname(__file__), filename)

    def _read_kv_file(self, filename: str, default: Dict[str, str]) -> Dict[str, str]:
        path = self._resource_path(filename)
        if not os.path.exists(path):
            return default.copy()
        result: Dict[str, str] = {}
        try:
            with open(path, 'r', encoding='utf-8') as f:
                for line in f:
                    line = line.strip()
                    if not line or line.startswith('#'):
                        continue
                    sep = '=>'
                    if sep in line:
                        left, right = line.split(sep, 1)
                    elif ':' in line:
                        left, right = line.split(':', 1)
                    elif '=' in line:
                        left, right = line.split('=', 1)
                    else:
                        continue
                    result[left.strip()] = right.strip()
        except Exception:
            return default.copy()
        return result if result else default.copy()

    def _read_list_file(self, filename: str, default: List[str]) -> List[str]:
        path = self._resource_path(filename)
        if not os.path.exists(path):
            return list(default)
        result: List[str] = []
        try:
            with open(path, 'r', encoding='utf-8') as f:
                for line in f:
                    line = line.strip()
                    if not line or line.startswith('#'):
                        continue
                    result.append(line)
        except Exception:
            return list(default)
        return result if result else list(default)

    def _read_prompt_file(self, filename: str, default: Optional[str]) -> Optional[str]:
        path = self._resource_path(filename)
        if not os.path.exists(path):
            return default
        try:
            with open(path, 'r', encoding='utf-8') as f:
                return f.read()
        except Exception:
            return default

    def reload_files(self) -> None:
        """Recharge les fichiers éditables (abbreviations, special chars, whitelist, prompt)."""
        self.ABBREVIATIONS = self._read_kv_file('abbreviations.txt', default=self.ABBREVIATIONS)
        self.LATEX_SPECIAL_CHARS = self._read_kv_file('special_chars.txt', default=self.LATEX_SPECIAL_CHARS)
        self.CORRECTION_WHITELIST = self._read_list_file('whitelist.txt', default=self.CORRECTION_WHITELIST)
        self.CORRECTION_PROMPT_TEMPLATE = self._read_prompt_file('prompt.txt', default=self.CORRECTION_PROMPT_TEMPLATE)
#!/usr/bin/env python3
"""
app.py - Point d'entrée principal de l'application avec interface graphique
"""

import sys
import os
from pathlib import Path

# Ajouter le répertoire courant au path pour les imports
sys.path.insert(0, str(Path(__file__).parent))

def check_dependencies():
    """Vérifie que toutes les dépendances sont installées."""
    missing_deps = []
    
    # Vérifier les dépendances
    try:
        import docx
    except ImportError:
        missing_deps.append('python-docx')
    
    try:
        import cohere
    except ImportError:
        missing_deps.append('cohere')
    
    try:
        import tkinterdnd2
    except ImportError:
        missing_deps.append('tkinterdnd2')
    
    if missing_deps:
        print("Dépendances manquantes détectées!")
        print(f"   Veuillez installer: {', '.join(missing_deps)}")
        print("\n Pour installer toutes les dépendances:")
        print("   pip install -r requirements.txt")
        print("\n   ou individuellement:")
        for dep in missing_deps:
            print(f"   pip install {dep}")
        return False
    
    return True

def main():
    """Lance l'application avec interface graphique."""
    
    # Vérifier les dépendances
    if not check_dependencies():
        print("\n L'application ne peut pas démarrer sans les dépendances requises.")
        input("\nAppuyez sur Entrée pour quitter...")
        sys.exit(1)
    
    try:
        # Importer et lancer l'interface graphique
        from gui import ConverterGUI
        
        print(" Lancement de l'interface graphique...")
        app = ConverterGUI()
        app.run()
        
    except Exception as e:
        print(f"\n Erreur lors du lancement de l'application: {e}")
        import traceback
        traceback.print_exc()
        input("\nAppuyez sur Entrée pour quitter...")
        sys.exit(1)

if __name__ == "__main__":
    main()
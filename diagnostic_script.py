#!/usr/bin/env python3
"""
test_write_access.py - Script de diagnostic pour tester les permissions d'écriture
"""

import sys
import os
from pathlib import Path
import tempfile

def test_write_access(directory):
    """Teste si un dossier est accessible en écriture."""
    try:
        directory = Path(directory)
        directory.mkdir(parents=True, exist_ok=True)
        
        # Tenter d'écrire un fichier test
        test_file = directory / ".write_test"
        test_file.write_text("test")
        content = test_file.read_text()
        test_file.unlink()
        
        print(f" {directory}")
        return True
    except Exception as e:
        print(f" {directory}")
        print(f"   Erreur: {e}")
        return False

def main():
    print("=" * 60)
    print("DIAGNOSTIC DES PERMISSIONS D'ÉCRITURE")
    print("=" * 60)
    print()
    
    print(f"Système: {sys.platform}")
    print(f"Python: {sys.version}")
    print(f"Exécutable: {sys.executable}")
    print(f"Frozen: {getattr(sys, 'frozen', False)}")
    print()
    
    print("=" * 60)
    print("TEST DES EMPLACEMENTS POSSIBLES")
    print("=" * 60)
    print()
    
    home = Path.home()
    
    test_locations = []
    
    if sys.platform == 'darwin':  # macOS
        test_locations = [
            ("Documents", home / "Documents" / "ConvertisseurDocxLatex"),
            ("Desktop", home / "Desktop" / "ConvertisseurDocxLatex"),
            ("Downloads", home / "Downloads" / "ConvertisseurDocxLatex"),
            ("Temp système", Path(tempfile.gettempdir()) / "ConvertisseurDocxLatex"),
        ]
        
        # Si on est dans un bundle
        if getattr(sys, 'frozen', False):
            bundle_dir = Path(sys.executable).parent
            test_locations.insert(0, ("Bundle (NE DEVRAIT PAS MARCHER)", bundle_dir))
    
    elif sys.platform == 'win32':  # Windows
        if getattr(sys, 'frozen', False):
            exe_dir = Path(sys.executable).parent
            test_locations = [
                ("À côté de l'exe", exe_dir),
            ]
        else:
            script_dir = Path(__file__).parent
            test_locations = [
                ("À côté du script", script_dir),
            ]
        
        test_locations.append(("Documents", home / "Documents" / "ConvertisseurDocxLatex"))
        test_locations.append(("Temp système", Path(tempfile.gettempdir()) / "ConvertisseurDocxLatex"))
    
    results = []
    for name, location in test_locations:
        print(f"Test: {name}")
        success = test_write_access(location)
        results.append((name, location, success))
        print()
    
    print("=" * 60)
    print("RÉSUMÉ")
    print("=" * 60)
    print()
    
    working_locations = [r for r in results if r[2]]
    
    if working_locations:
        print(" Emplacements accessibles en écriture:")
        for name, location, _ in working_locations:
            print(f"   • {name}: {location}")
        
        print()
        print(f" RECOMMANDATION: Utiliser {working_locations[0][1]}")
    else:
        print(" AUCUN emplacement accessible!")
        print("   C'est très inhabituel. Vérifiez les permissions système.")
    
    print()
    print("=" * 60)
    
    # Sur macOS, donner des instructions supplémentaires
    if sys.platform == 'darwin' and getattr(sys, 'frozen', False):
        print()
        print("CONSEILS POUR macOS:")
        print()
        print("1. Si l'app est dans un dossier de translocation:")
        print("   • Déplacez l'app dans /Applications")
        print("   • OU déplacez-la sur le Bureau")
        print("   • Puis relancez-la")
        print()
        print("2. Pour supprimer les attributs de quarantaine:")
        print("   • Ouvrez le Terminal")
        print("   • Tapez: xattr -cr ")
        print("   • Glissez l'app dans le Terminal")
        print("   • Appuyez sur Entrée")
        print()
        print("3. Donnez les permissions complètes:")
        print("   • Préférences Système > Sécurité et confidentialité")
        print("   • Onglet Confidentialité")
        print("   • Accès complet au disque")
        print("   • Ajoutez l'application")
        print()

if __name__ == "__main__":
    main()
    
    if sys.platform == 'darwin':
        input("\nAppuyez sur Entrée pour quitter...")

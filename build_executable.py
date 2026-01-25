#!/usr/bin/env python3
"""
build_executable.py - Script pour créer des exécutables pour Windows et macOS
"""

import os
import sys
import subprocess
import platform
import shutil
from pathlib import Path

class ExecutableBuilder:
    """Classe pour construire les exécutables multi-plateformes."""
    
    def __init__(self):
        self.root_dir = Path(__file__).parent
        self.dist_dir = self.root_dir / "dist"
        self.build_dir = self.root_dir / "build"
        self.system = platform.system()
        
    def clean_previous_builds(self):
        """Nettoie les builds précédents."""
        print(" Nettoyage des builds précédents...")
        
        for directory in [self.dist_dir, self.build_dir]:
            if directory.exists():
                shutil.rmtree(directory)
                print(f"    Supprimé: {directory}")
        
        # Supprimer les fichiers spec
        for spec_file in self.root_dir.glob("*.spec"):
            spec_file.unlink()
            print(f"    Supprimé: {spec_file}")
    
    def check_pyinstaller(self):
        """Vérifie que PyInstaller est installé."""
        print("\n Vérification de PyInstaller...")
        
        try:
            import PyInstaller
            print(f"    PyInstaller {PyInstaller.__version__} détecté")
            return True
        except ImportError:
            print("    PyInstaller non détecté")
            print("\n   Installation de PyInstaller...")
            subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
            print("    PyInstaller installé")
            return True
    
    def create_launcher_script(self):
        """Crée un script de lancement qui gère LaTeX."""
        print("\n Création du script de lancement...")
        
        launcher_content = '''#!/usr/bin/env python3
"""
Launcher - Point d'entrée avec vérification LaTeX
"""

import sys
import os
import subprocess
import platform
from pathlib import Path

def check_latex():
    """Vérifie si LaTeX est installé."""
    try:
        subprocess.run(['pdflatex', '--version'], 
                      capture_output=True, 
                      check=True)
        return True
    except (subprocess.CalledProcessError, FileNotFoundError):
        return False

def show_latex_warning():
    """Affiche un avertissement si LaTeX n'est pas installé."""
    import tkinter as tk
    from tkinter import messagebox
    
    root = tk.Tk()
    root.withdraw()
    
    system = platform.system()
    
    if system == "Windows":
        message = (
            "LaTeX (MiKTeX) n'est pas installé sur votre système.\\n\\n"
            "L'application fonctionnera mais ne pourra pas générer de PDF.\\n\\n"
            "Pour installer MiKTeX :\\n"
            "1. Téléchargez depuis : https://miktex.org/download\\n"
            "2. Installez avec les options par défaut\\n"
            "3. Redémarrez l'application\\n\\n"
            "Voulez-vous continuer sans génération de PDF ?"
        )
    else:  # macOS
        message = (
            "LaTeX (BasicTeX) n'est pas installé sur votre système.\\n\\n"
            "L'application fonctionnera mais ne pourra pas générer de PDF.\\n\\n"
            "Pour installer BasicTeX :\\n"
            "1. Installez Homebrew : https://brew.sh\\n"
            "2. Exécutez : brew install --cask basictex\\n"
            "3. Redémarrez l'application\\n\\n"
            "Voulez-vous continuer sans génération de PDF ?"
        )
    
    result = messagebox.askyesno(
        "LaTeX non détecté",
        message,
        icon='warning'
    )
    
    root.destroy()
    return result

def main():
    """Point d'entrée principal."""
    
    # Vérifier LaTeX
    if not check_latex():
        if not show_latex_warning():
            sys.exit(0)
    
    # Importer et lancer l'application
    from gui import ConverterGUI
    
    app = ConverterGUI()
    app.run()

if __name__ == "__main__":
    main()
'''
        
        launcher_path = self.root_dir / "launcher.py"
        with open(launcher_path, 'w', encoding='utf-8') as f:
            f.write(launcher_content)
        
        print(f"    Script créé: {launcher_path}")
        return launcher_path
    
    def build_windows(self):
        """Construit l'exécutable Windows."""
        print("\n Construction de l'exécutable Windows...")
        
        cmd = [
            sys.executable,
            '-m', 'PyInstaller',
            '--name=ConvertisseurDocxLatex',
            '--onefile',
            '--windowed',
            '--add-data=config.py;.',
            '--add-data=converter.py;.',
            '--add-data=gui.py;.',
            '--add-data=latex_generator.py;.',
            '--add-data=text_corrector.py;.',
            '--add-data=utils.py;.',
            '--hidden-import=tkinter',
            '--hidden-import=tkinterdnd2',
            '--hidden-import=docx',
            '--hidden-import=cohere',
            '--collect-all=tkinterdnd2',
            'launcher.py'
        ]
        
        subprocess.check_call(cmd)
        print("    Exécutable Windows créé avec succès!")
    
    def build_macos(self):
        """Construit l'exécutable macOS."""
        print("\n Construction de l'application macOS...")
        
        cmd = [
            sys.executable,
            '-m', 'PyInstaller',
            '--name=ConvertisseurDocxLatex',
            '--windowed',
            '--add-data=config.py:.',
            '--add-data=converter.py:.',
            '--add-data=gui.py:.',
            '--add-data=latex_generator.py:.',
            '--add-data=text_corrector.py:.',
            '--add-data=utils.py:.',
            '--hidden-import=tkinter',
            '--hidden-import=tkinterdnd2',
            '--hidden-import=docx',
            '--hidden-import=cohere',
            '--collect-all=tkinterdnd2',
            '--osx-bundle-identifier=com.cbb.convertisseur',
            'launcher.py'
        ]
        
        subprocess.check_call(cmd)
        print("    Application macOS créée avec succès!")
    
    def create_readme(self):
        """Crée un fichier README pour les utilisateurs."""
        print("\n Création du README...")
        
        readme_content = """# Convertisseur DocX vers LaTeX/PDF

## Installation de LaTeX (REQUIS pour la génération de PDF)

### Windows
1. Téléchargez MiKTeX : https://miktex.org/download
2. Installez avec les options par défaut
3. Lors de la première utilisation, MiKTeX installera automatiquement les packages manquants

### macOS
1. Installez Homebrew si ce n'est pas déjà fait : https://brew.sh
2. Exécutez dans le Terminal :
   ```
   brew install --cask basictex
   ```
3. Ajoutez au PATH (ajoutez à ~/.zshrc ou ~/.bash_profile) :
   ```
   export PATH="/Library/TeX/texbin:$PATH"
   ```
4. Installez les packages essentiels :
   ```
   sudo tlmgr update --self
   sudo tlmgr install geometry fancyhdr multicol graphics float varwidth eurosym
   ```

## Utilisation

### Windows
- Double-cliquez sur `ConvertisseurDocxLatex.exe`

### macOS
- Double-cliquez sur `ConvertisseurDocxLatex.app`
- Si macOS bloque l'application :
  1. Ouvrez Préférences Système > Sécurité et confidentialité
  2. Cliquez sur "Ouvrir quand même"
  
  OU
  
  1. Faites clic droit sur l'application > Ouvrir
  2. Confirmez l'ouverture

## Fonctionnalités

- Glissez-déposez vos fichiers .docx dans l'interface
- Conversion automatique en LaTeX
- Compilation automatique en PDF (si LaTeX est installé)
- Correction orthographique et grammaticale automatique
- Traitement par lots

## Support

En cas de problème, vérifiez que :
1. LaTeX est correctement installé
2. Vous avez les droits d'accès aux fichiers
3. Les fichiers .docx ne sont pas corrompus

## Note importante

L'application fonctionne sans LaTeX mais ne générera que des fichiers .tex (sans PDF).
Pour une expérience complète, l'installation de LaTeX est fortement recommandée.
"""
        
        readme_path = self.dist_dir / "README.txt"
        readme_path.parent.mkdir(parents=True, exist_ok=True)
        
        with open(readme_path, 'w', encoding='utf-8') as f:
            f.write(readme_content)
        
        print(f"    README créé: {readme_path}")
    
    def build(self):
        """Lance le processus de build complet."""
        print("=" * 60)
        print(" Construction de l'exécutable")
        print("=" * 60)
        
        # Nettoyage
        self.clean_previous_builds()
        
        # Vérification PyInstaller
        self.check_pyinstaller()
        
        # Création du launcher
        launcher = self.create_launcher_script()
        
        try:
            # Build selon la plateforme
            if self.system == "Windows":
                self.build_windows()
            elif self.system == "Darwin":  # macOS
                self.build_macos()
            else:
                print(f"\n Système non supporté: {self.system}")
                print("   Ce script supporte uniquement Windows et macOS")
                return False
            
            # Création du README
            self.create_readme()
            
            print("\n" + "=" * 60)
            print(" BUILD RÉUSSI!")
            print("=" * 60)
            print(f"\n Exécutable disponible dans: {self.dist_dir}")
            
            if self.system == "Windows":
                print("   → ConvertisseurDocxLatex.exe")
            else:
                print("   → ConvertisseurDocxLatex.app")
            
            print("\n  IMPORTANT: LaTeX doit être installé séparément!")
            print("   Consultez le fichier README.txt pour les instructions.")
            
            return True
            
        except subprocess.CalledProcessError as e:
            print(f"\n Erreur lors du build: {e}")
            return False
        finally:
            # Nettoyage du launcher temporaire
            if launcher.exists():
                launcher.unlink()

def main():
    """Point d'entrée principal."""
    builder = ExecutableBuilder()
    
    print("\n  Système détecté:", platform.system())
    print(" Python version:", sys.version.split()[0])
    print()
    
    success = builder.build()
    
    if success:
        print("Success!\n")
        sys.exit(0)
    else:
        print("Failed!\n")
        sys.exit(1)

if __name__ == "__main__":
    main()
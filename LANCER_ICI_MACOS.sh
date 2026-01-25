#!/bin/bash

echo "============================================"
echo "  Convertisseur DocX vers LaTeX/PDF"
echo "============================================"
echo "Répertoire de travail: $(pwd)"
echo ""

# Se placer dans le répertoire du script
cd "$(dirname "$0")"

# Couleurs pour les messages
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

# Fonction: Installer Homebrew
install_homebrew() {
    echo -e "${YELLOW}Installation de Homebrew...${NC}"
    echo "Homebrew est nécessaire pour installer Python et LaTeX"
    echo ""
    
    /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
    
    if [ $? -ne 0 ]; then
        echo -e "${RED}[ERREUR] L'installation de Homebrew a échoué${NC}"
        return 1
    fi
    
    # Ajouter Homebrew au PATH (pour Apple Silicon)
    if [[ $(uname -m) == 'arm64' ]]; then
        echo 'eval "$(/opt/homebrew/bin/brew shellenv)"' >> ~/.zprofile
        eval "$(/opt/homebrew/bin/brew shellenv)"
    else
        echo 'eval "$(/usr/local/bin/brew shellenv)"' >> ~/.zprofile
        eval "$(/usr/local/bin/brew shellenv)"
    fi
    
    echo -e "${GREEN}Homebrew installé avec succès!${NC}"
    return 0
}

# Fonction: Installer Python
install_python() {
    echo -e "${YELLOW}Installation de Python 3...${NC}"
    echo "Cette opération peut prendre quelques minutes..."
    echo ""
    
    brew install python@3.12
    
    if [ $? -ne 0 ]; then
        echo -e "${RED}[ERREUR] L'installation de Python a échoué${NC}"
        return 1
    fi
    
    # Créer un lien symbolique si nécessaire
    brew link python@3.12
    
    echo -e "${GREEN}Python installé avec succès!${NC}"
    return 0
}

# Fonction: Installer MacTeX (BasicTeX pour version légère)
install_latex() {
    echo -e "${YELLOW}Installation de BasicTeX (distribution LaTeX légère)...${NC}"
    echo "Cette opération peut prendre 10-15 minutes..."
    echo "Taille du téléchargement: ~100 MB"
    echo ""
    
    # Installer BasicTeX via Homebrew Cask
    brew install --cask basictex
    
    if [ $? -ne 0 ]; then
        echo -e "${RED}[ERREUR] L'installation de BasicTeX a échoué${NC}"
        return 1
    fi
    
    # Ajouter BasicTeX au PATH
    export PATH="/Library/TeX/texbin:$PATH"
    echo 'export PATH="/Library/TeX/texbin:$PATH"' >> ~/.zprofile
    echo 'export PATH="/Library/TeX/texbin:$PATH"' >> ~/.bash_profile
    
    echo ""
    echo "Configuration de LaTeX..."
    
    # Mettre à jour tlmgr (TeX Live Manager)
    sudo tlmgr update --self
    
    # Installer les packages essentiels
    echo "Installation des packages LaTeX essentiels..."
    sudo tlmgr install geometry fancyhdr multicol graphics float varwidth eurosym
    
    if [ $? -eq 0 ]; then
        echo -e "${GREEN}BasicTeX et packages installés avec succès!${NC}"
    else
        echo -e "${YELLOW}[ATTENTION] Certains packages LaTeX n'ont pas pu être installés${NC}"
        echo "L'application fonctionnera mais certaines fonctionnalités PDF peuvent manquer"
    fi
    
    return 0
}

# Vérifier si Homebrew est installé
if ! command -v brew &> /dev/null; then
    echo -e "${YELLOW}[ATTENTION] Homebrew n'est pas installé${NC}"
    echo ""
    read -p "Voulez-vous installer Homebrew maintenant? (o/n) " -n 1 -r
    echo ""
    if [[ $REPLY =~ ^[Oo]$ ]]; then
        install_homebrew
        if [ $? -ne 0 ]; then
            echo -e "${RED}Impossible de continuer sans Homebrew${NC}"
            exit 1
        fi
    else
        echo -e "${RED}Homebrew est requis pour l'installation automatique${NC}"
        echo "Veuillez installer Homebrew manuellement: https://brew.sh"
        exit 1
    fi
fi

# Vérifier si Python est installé
if ! command -v python3 &> /dev/null; then
    echo -e "${YELLOW}[ATTENTION] Python n'est pas installé${NC}"
    echo ""
    
    read -p "Voulez-vous installer Python maintenant? (o/n) " -n 1 -r
    echo ""
    if [[ $REPLY =~ ^[Oo]$ ]]; then
        install_python
        if [ $? -ne 0 ]; then
            echo -e "${RED}Installation de Python échouée${NC}"
            exit 1
        fi
    else
        echo -e "${RED}Python est requis pour cette application${NC}"
        exit 1
    fi
else
    echo "Python détecté:"
    python3 --version
    echo ""
fi

# Vérifier si LaTeX est installé
if ! command -v pdflatex &> /dev/null; then
    echo -e "${YELLOW}[ATTENTION] LaTeX n'est pas installé${NC}"
    echo ""
    
    read -p "Voulez-vous installer BasicTeX maintenant? (o/n) " -n 1 -r
    echo ""
    if [[ $REPLY =~ ^[Oo]$ ]]; then
        install_latex
        if [ $? -ne 0 ]; then
            echo -e "${YELLOW}[ATTENTION] LaTeX non installé${NC}"
            echo "L'application fonctionnera mais sans génération de PDF"
            sleep 3
        fi
    else
        echo -e "${YELLOW}L'application fonctionnera sans génération de PDF${NC}"
        echo "Pour installer LaTeX plus tard: brew install --cask basictex"
        sleep 3
    fi
else
    echo "LaTeX détecté:"
    pdflatex --version | head -n 1
    echo ""
fi

# Créer l'environnement virtuel si nécessaire
if [ ! -d "venv" ]; then
    echo "Création de l'environnement virtuel..."
    python3 -m venv venv
    if [ $? -ne 0 ]; then
        echo -e "${RED}[ERREUR] Impossible de créer l'environnement virtuel${NC}"
        exit 1
    fi
    echo ""
fi

# Activer l'environnement virtuel
echo "Activation de l'environnement virtuel..."
source venv/bin/activate
if [ $? -ne 0 ]; then
    echo -e "${RED}[ERREUR] Impossible d'activer l'environnement virtuel${NC}"
    exit 1
fi

# Installer/Mettre à jour les dépendances
echo ""
echo "Vérification des dépendances..."
pip install --quiet --upgrade pip

if [ ! -f "requirements.txt" ]; then
    echo -e "${YELLOW}[ATTENTION] Fichier requirements.txt introuvable${NC}"
    echo "Création d'un requirements.txt minimal..."
    cat > requirements.txt << EOF
python-docx>=0.8.11
cohere>=5.0
tkinterdnd2>=0.3.0
EOF
fi

echo "Installation des dépendances depuis requirements.txt..."
pip install -r requirements.txt
if [ $? -ne 0 ]; then
    echo -e "${RED}[ERREUR] Échec de l'installation des dépendances${NC}"
    echo "Vérifiez le contenu de requirements.txt"
    echo ""
    echo "Contenu actuel de requirements.txt:"
    cat requirements.txt
    echo ""
    read -p "Appuyez sur Entrée pour quitter..."
    exit 1
fi

# Lancer l'application
echo ""
echo "============================================"
echo "  Lancement de l'application..."
echo "============================================"
echo ""
python3 app.py

# Pause si erreur
if [ $? -ne 0 ]; then
    echo ""
    echo -e "${RED}[ERREUR] L'application s'est terminée avec une erreur.${NC}"
    read -p "Appuyez sur Entrée pour quitter..."
fi

# Désactiver l'environnement virtuel
deactivate

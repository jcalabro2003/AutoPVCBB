# Convertisseur PV DocX vers LaTeX

## Description
Convertisseur automatique de procès-verbaux de réunions du format DocX vers LaTeX avec correction orthographique intégrée.

## Structure du projet

```
pv_converter/
├── app.py              # Point d'entrée principal
├── config.py            # Configuration centralisée
├── converter.py         # Logique principale de conversion
├── utils.py            # Fonctions utilitaires
├── latex_generator.py   # Génération du code LaTeX
├── text_corrector.py    # Correction orthographique
├── requirements.txt     # Dépendances Python
└── README.md           # Documentation
```

## Installation

1. **Cloner le projet**
```bash
git clone <votre-repo>
cd pv_converter
```

2. **Créer un environnement virtuel** (recommandé)
```bash
python -m venv venv
source venv/bin/activate  # Linux/Mac
# ou
venv\Scripts\activate  # Windows
```

3. **Installer les dépendances**
```bash
pip install -r requirements.txt
```

4. **Configuration de l'API Cohere**
   - Obtenir une clé API sur [Cohere](https://cohere.ai)
   - Définir la variable d'environnement :
   ```bash
   export COHERE_API_KEY="votre_cle_api"
   ```
   - Ou modifier directement dans `config.py`

## Utilisation

### Utilisation basique
```python
python main.py
```

### Utilisation programmatique
```python
from converter import DocxToLatexConverter
from config import Config

config = Config()
converter = DocxToLatexConverter(config)
converter.convert("mon_pv.docx", "mon_pv.tex")
```

### Configuration personnalisée
```python
from config import Config

config = Config()
config.BATCH_SIZE = 20  # Augmenter la taille des lots
config.ABBREVIATIONS['nouveau'] = 'remplacement'
config.CORRECTION_WHITELIST = ['mot1', 'mot2']
```

## Architecture modulaire

### `config.py`
- Configuration centralisée
- Paramètres API
- Dictionnaires de remplacement
- Paramètres LaTeX

### `converter.py`
- Orchestrateur principal
- Gestion du flux de conversion
- Coordination des modules

### `utils.py`
- **TextProcessor** : Traitement du texte (échappement LaTeX, abréviations)
- **DocumentParser** : Parsing des éléments du document
- **TableProcessor** : Traitement des tableaux

### `latex_generator.py`
- Génération du code LaTeX
- Templates pour chaque type d'élément
- Formatage et mise en page

### `text_corrector.py`
- Correction orthographique via Cohere
- Traitement par lots optimisé
- Conservation du formatage

## Format attendu du document

Le document DocX doit suivre cette structure :
1. **Titre** : Format `PV RC X - Anno Y - AAAA-MM-JJ`
2. **Liste des présents** : Section commençant par "Présents"
3. **Sections** : Numérotées avec format `X) Titre`
4. **Contenu** : Paragraphes avec formatage (gras, italique)

## Fonctionnalités

- ✅ Conversion complète DocX → LaTeX
- ✅ Correction orthographique automatique
- ✅ Conservation du formatage (gras, italique)
- ✅ Génération automatique de la table des matières
- ✅ Traitement des tableaux
- ✅ Remplacement des abréviations
- ✅ Mise en page professionnelle

## Optimisations implémentées

1. **Traitement par lots** : Correction de plusieurs paragraphes en un seul appel API
2. **Cache de configuration** : Évite les recalculs répétitifs
3. **Lazy loading** : Chargement des modules à la demande
4. **Structure modulaire** : Code réutilisable et maintenable

## Logging

Pour activer les logs détaillés :
```python
import logging
logging.basicConfig(level=logging.INFO)
```

## Personnalisation

### Ajouter une nouvelle abréviation
```python
config.ABBREVIATIONS[r'\bnouveau\b'] = 'remplacement'
```

### Modifier le template LaTeX
Éditer les méthodes dans `latex_generator.py`

### Changer le service de correction
Implémenter une nouvelle classe dans `text_corrector.py`

## Dépannage

**Problème** : Erreur API Cohere
- Vérifier la clé API
- Vérifier la connexion internet
- Les paragraphes seront retournés non corrigés

**Problème** : Format de titre non reconnu
- Vérifier le format : `PV RC X - Anno Y - AAAA-MM-JJ`
- Un en-tête simple sera généré en cas d'erreur

**Problème** : Caractères spéciaux dans LaTeX
- Vérifier `LATEX_SPECIAL_CHARS` dans `config.py`
- Ajouter les caractères manquants si nécessaire

## Contribution

Pour contribuer :
1. Fork le projet
2. Créer une branche feature
3. Commiter les changements
4. Pusher vers la branche
5. Ouvrir une Pull Request

## Licence

[Votre licence]

## Contact

[Vos informations de contact]
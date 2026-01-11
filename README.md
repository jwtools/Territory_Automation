# Territory Automation

🚀 **Automatisation complète** de la saisie et l'import de territoires dans **New World Scheduler 7.9**.

## 🎯 Fonctionnalités

### Import et saisie automatisée
- ✅ Import automatique de territoires depuis un fichier Excel/CSV
- ✅ Remplissage automatique des formulaires (numéro, suffixe, type, notes, etc.)
- ✅ **Catégories configurables** via `data/options.json`
- ✅ Support de multiples types de territoires (En présentiel, Courrier, Téléphone, Entreprise)
- ✅ **Villes configurables** via `data/options.json`
- ✅ Import automatique des fichiers PDF associés

### Gestion intelligente
- 💾 Sauvegarde automatique de la progression (reprise après interruption)
- 📊 Vérification des données et PDFs avant exécution
- 📝 Logging détaillé des actions et erreurs
- 🎭 Mode "dry-run" pour simulation (sans exécuter les actions)
- 🔍 Mode "no-save" pour validation visuelle (remplit sans sauvegarder)
- 🛑 Arrêt d'urgence (fail-safe)

### Outils de calibration
- 🎯 Assistant de calibration guidé (recommandé)
- 🖱️ Outil de capture manuelle de coordonnées
- ✅ Test de calibration avec modes survol et clic
- 🔄 Recalibration facile si besoin

## 📋 Prérequis

- **Python 3.10+** (testé avec Python 3.10, 3.11, 3.12)
- **Windows 10/11** (obligatoire pour pywinauto)
- **New World Scheduler 7.9** installé et configuré
- **[uv](https://docs.astral.sh/uv/)** (gestionnaire de paquets Python recommandé)
- **Résolution d'écran stable** (pour la calibration des coordonnées)

## 🔧 Installation

### Installation avec uv

```bash
# Créer l'environnement virtuel et installer les dépendances
uv sync

# Ou installation classique avec pip
pip install -r requirements.txt
```

## Configuration

1. **Modifier `config.py`** avec le chemin de votre installation NWS :
   ```python
   NWS_EXE_PATH = r"C:\Program Files\New World Scheduler\NWScheduler.exe"
   ```

2. **Calibrer les coordonnées** (voir section Calibration)

3. **Préparer vos données** :
   - Créez votre fichier Excel à partir du template
   - Placez vos PDFs dans `data/pdfs/`

## Utilisation

### Créer le template Excel

```bash
uv run python tools/create_template.py
```

### Vérifier les prérequis

```bash
uv run python tools/test_connection.py
```

### Lancer l'automatisation

```bash
# Mode normal (exécution complète)
uv run python main.py

# Mode simulation (affiche les actions sans les exécuter)
uv run python main.py --dry-run

# Mode validation (remplit les champs sans sauvegarder)
uv run python main.py --no-save

# Vérifier les données et PDFs avant exécution
uv run python main.py --verify

# Réinitialiser la progression (recommencer depuis le début)
uv run python main.py --reset

# Commencer à partir d'un index spécifique
uv run python main.py --start-from 10

# Utiliser un fichier de données personnalisé
uv run python main.py --data-file data/custom.xlsx
```

### Mode validation (--no-save)

Ce mode permet de vérifier visuellement ce que l'automatisation va faire :
- Les champs sont remplis dans NWS mais **non sauvegardés**
- Le script attend que vous appuyiez sur Entrée avant de passer au territoire suivant
- Vous pouvez vérifier les données dans l'interface NWS
- Aucune modification n'est enregistrée dans la base de données

## Calibration des coordonnées

Les coordonnées des boutons dépendent de votre résolution d'écran. Un assistant de calibration guidé est disponible.

### Calibration guidée (recommandé)

```bash
uv run python tools/calibration.py
```

L'assistant vous guide étape par étape :
1. Affiche le nom et la description de chaque élément
2. Vous demande de positionner la souris dessus
3. Appuyez sur `C` pour capturer, `S` pour passer
4. Les coordonnées sont sauvegardées automatiquement dans `data/calibration.json`

### Test des coordonnées

Après calibration, vérifiez que les coordonnées sont correctes :

```bash
# Mode survol : la souris se déplace sur chaque élément
uv run python tools/test_calibration.py

# Mode clic : clique réellement sur chaque élément
uv run python tools/test_calibration.py --click

# Avec pause entre chaque élément
uv run python tools/test_calibration.py --pause

# Tester un seul élément
uv run python tools/test_calibration.py --element btn_new_territory
```

### Outil de capture libre

Pour capturer des coordonnées manuellement :
```bash
uv run python tools/coordinate_finder.py
```

## 📁 Structure du projet

```
Territoy_Automation/
├── main.py                     # 🚀 Script principal d'automatisation
├── config.py                   # ⚙️ Configuration (chemins, délais, coordonnées)
├── pyproject.toml              # 📦 Configuration uv et dépendances
├── requirements.txt            # 📦 Dépendances Python (fallback pip)
│
├── territory_automation/       # 🔧 Modules Python core
│   ├── __init__.py
│   ├── automation.py           # Logique d'automatisation NWS (pywinauto + pyautogui)
│   ├── data_loader.py          # Chargement Excel/CSV + gestion progression
│   └── logger_setup.py         # Configuration des logs rotatifs
│
├── tools/                      # 🛠️ Outils de calibration et tests
│   ├── calibration.py          # Assistant de calibration guidé (recommandé)
│   ├── coordinate_finder.py    # Capture manuelle de coordonnées
│   ├── test_calibration.py     # Test des coordonnées calibrées
│   ├── test_connection.py      # Test de connexion à NWS
│   └── create_template.py      # Génération du template Excel
│
├── data/                       # 📊 Données d'automatisation
│   ├── territories.xlsx        # Fichier de données (à créer)
│   ├── options.json            # ⚙️ Configuration catégories/villes
│   ├── progress.json           # Suivi de progression (auto-généré)
│   ├── calibration.json        # Coordonnées calibrées (auto-généré)
│   └── pdfs/                   # 📄 Fichiers PDF des territoires
│
├── logs/                       # 📝 Journaux d'exécution
│   └── automation_*.log        # Logs horodatés de chaque exécution
│
└── docs/                       # 📚 Documentation
    └── GUIDE.md                # Guide détaillé d'installation et utilisation
```

## 📊 Format des données Excel

### Colonnes du fichier Excel

| Colonne | Type | Description | Obligatoire | Exemple |
|---------|------|-------------|-------------|----------|
| **Numero** | Texte | Numéro unique du territoire | ✅ Oui | `SAR-1-01` |
| **Suffixe** | Texte | Suffixe du territoire | ❌ Non | `A`, `B` |
| **Categorie** | Liste | Catégorie du territoire | ❌ Non | `SAR` |
| **Type** | Liste | Type de territoire | ❌ Non | `En présentiel` |
| **Ville** | Liste | Ville du territoire | ❌ Non | `SARTROUVILLE` |
| **Lien_GPS** | URL | Lien Google Maps | ❌ Non | `https://maps.google.com/...` |
| **Notes** | Texte | Notes générales | ❌ Non | `Zone résidentielle` |
| **Ne_Pas_Visiter** | Texte | Adresses à éviter | ❌ Non | `Apt 3B, 15 rue...` |
| **Notes_Proclamateur** | Texte | Notes pour proclamateurs | ❌ Non | `Prévoir 2h` |
| **PDF_Filename** | Texte | Nom du fichier PDF | ❌ Non | `custom.pdf` |

### Valeurs acceptées

**Type** (menu déroulant dans NWS) :
- `En présentiel` (défaut)
- `Courrier`
- `Téléphone`
- `Entreprise`
- `Aucun` (ou laisser vide)

**Catégorie et Ville** : Configurables via `data/options.json`

### Règles de nommage des PDFs

1. **Par défaut** : Le fichier PDF doit avoir le même nom que le numéro du territoire
   - Exemple : `SAR-1-01.pdf` pour le territoire `SAR-1-01`

2. **Personnalisé** : Si vous utilisez un nom différent, remplissez la colonne `PDF_Filename`
   - Exemple : `custom_map_01.pdf`

3. **Placement** : Tous les PDFs doivent être dans le dossier `data/pdfs/`

## ⚙️ Configuration des catégories et villes

Les catégories et villes sont configurables via le fichier `data/options.json` :

```json
{
  "categories": {
    "SAR": "dropdown_option_sar",
    "AUTRE": "dropdown_option_autre"
  },
  "villes": {
    "SARTROUVILLE": "dropdown_ville_sartrouville",
    "MAISONS-LAFFITTE": "dropdown_ville_maisons"
  }
}
```

### Ajouter une nouvelle catégorie ou ville

1. Ajouter l'entrée dans `data/options.json`
2. Calibrer la coordonnée correspondante : `uv run python tools/calibration.py`

> 💡 **Note** : Si la colonne `Categorie` est vide dans Excel, la première catégorie du fichier options.json est utilisée par défaut.

## Arrêt d'urgence

- **Déplacez la souris dans le coin supérieur gauche** de l'écran pour arrêter immédiatement l'automatisation (fail-safe pyautogui)
- Ou appuyez sur **Ctrl+C** dans le terminal

## Documentation détaillée

Voir [docs/GUIDE.md](docs/GUIDE.md) pour le guide complet d'installation et d'utilisation.

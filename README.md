# Territory Automation

Automatisation de la saisie et l'import de territoires dans **New World Scheduler 7.9**.

## Fonctionnalités

- Import automatique de territoires depuis un fichier Excel/CSV
- Remplissage automatique des formulaires (numéro, suffixe, type, notes, etc.)
- Import des fichiers PDF associés à chaque territoire
- Gestion de la progression (reprise après interruption)
- Logging détaillé des actions et erreurs
- Mode "dry-run" pour simulation
- Mode "no-save" pour valider les saisies avant enregistrement

## Installation

### Prérequis

- Python 3.10+
- Windows 10/11
- New World Scheduler 7.9 installé
- [uv](https://docs.astral.sh/uv/) (gestionnaire de paquets Python)

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
# Mode normal
uv run python main.py

# Mode simulation (sans exécuter)
uv run python main.py --dry-run

# Mode validation (remplit les champs sans sauvegarder)
uv run python main.py --no-save

# Réinitialiser la progression
uv run python main.py --reset

# Vérifier les données uniquement
uv run python main.py --verify
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

## Structure des fichiers

```
Territoy_Automation/
├── main.py                 # Script principal
├── config.py               # Configuration (chemins, coordonnées)
├── pyproject.toml          # Configuration projet et dépendances (uv)
├── requirements.txt        # Dépendances Python (pip fallback)
├── territory_automation/   # Modules Python
│   ├── automation.py       # Logique d'automatisation
│   ├── data_loader.py      # Chargement Excel/CSV
│   └── logger_setup.py     # Configuration logs
├── tools/                  # Outils utilitaires
│   ├── coordinate_finder.py
│   ├── create_template.py
│   └── test_connection.py
├── data/                   # Données
│   ├── territories.xlsx    # Votre fichier de données
│   └── pdfs/               # Fichiers PDF des territoires
└── logs/                   # Fichiers de log
```

## Format des données Excel

| Colonne | Description | Obligatoire |
|---------|-------------|-------------|
| Numero | Numéro du territoire (ex: SAR-1-01) | Oui |
| Suffixe | Suffixe (ex: A, B) | Non |
| Type | "En présentiel", "Courrier", "Téléphone" ou "Entreprise" | Non |
| Ville | SARTROUVILLE, MAISONS-LAFFITTE, MONTESSON, MESNIL LE ROI, CARRIERE S/ BOIS | Non |
| Lien_GPS | URL Google Maps | Non |
| Notes | Notes générales | Non |
| Ne_Pas_Visiter | Adresses à éviter | Non |
| Notes_Proclamateur | Notes pour les proclamateurs | Non |
| PDF_Filename | Nom du fichier PDF (sinon: Numero.pdf) | Non |

> **Note** : La catégorie "SAR" est sélectionnée automatiquement pour tous les territoires.

## Arrêt d'urgence

- **Déplacez la souris dans le coin supérieur gauche** de l'écran pour arrêter immédiatement l'automatisation (fail-safe pyautogui)
- Ou appuyez sur **Ctrl+C** dans le terminal

## Documentation détaillée

Voir [docs/GUIDE.md](docs/GUIDE.md) pour le guide complet d'installation et d'utilisation.

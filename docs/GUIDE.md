# Guide d'Installation et d'Utilisation

Guide complet pour l'automatisation de New World Scheduler 7.9.

## Table des matières

1. [Installation](#installation)
2. [Configuration](#configuration)
3. [Calibration des coordonnées](#calibration-des-coordonnées)
4. [Préparation des données](#préparation-des-données)
5. [Utilisation](#utilisation)
6. [Dépannage](#dépannage)

---

## Installation

### 1. Installer Python

Téléchargez Python 3.10+ depuis [python.org](https://www.python.org/downloads/)

Lors de l'installation, cochez **"Add Python to PATH"**.

### 2. Installer uv

uv est un gestionnaire de paquets Python ultra-rapide. Installez-le :

```bash
# Windows (PowerShell)
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"

# Ou avec pip
pip install uv
```

### 3. Installer les dépendances

Ouvrez un terminal (PowerShell ou CMD) dans le dossier du projet :

```bash
cd C:\chemin\vers\Territoy_Automation
uv sync
```

> **Alternative avec pip** : `pip install -r requirements.txt`

### 4. Vérifier l'installation

```bash
uv run python tools/test_connection.py
```

---

## Configuration

### Fichier `config.py`

Modifiez les paramètres selon votre environnement :

```python
# Chemin vers NWS (modifiez selon votre installation)
NWS_EXE_PATH = r"C:\Program Files\New World Scheduler\NWScheduler.exe"

# Délais (augmentez si l'application est lente)
DELAY_AFTER_CLICK = 0.3
DELAY_APP_LAUNCH = 5.0
```

---

## Calibration des coordonnées

### Pourquoi calibrer ?

Les coordonnées des boutons dépendent de :
- Votre résolution d'écran
- La position de la fenêtre NWS
- L'échelle d'affichage Windows (100%, 125%, 150%)

### Procédure de calibration

1. **Lancez New World Scheduler** et positionnez la fenêtre

2. **Lancez l'outil de capture** :
   ```bash
   uv run python tools/coordinate_finder.py
   ```

3. **Capturez chaque élément** :
   - Positionnez la souris sur l'élément
   - Appuyez sur `C` pour capturer
   - Notez le numéro et les coordonnées

4. **Éléments à capturer** (dans l'ordre) :

   | # | Élément | Description |
   |---|---------|-------------|
   | 1 | btn_new_territory | Bouton "+" (nouveau territoire) |
   | 2 | field_numero | Champ "Numéro de territoire" |
   | 3 | field_suffixe | Champ "Suffixe" |
   | 4 | dropdown_type | Dropdown "Type" |
   | 5 | dropdown_option_presentiel | Option "En présentiel" (après ouverture dropdown) |
   | 6 | dropdown_option_aucun | Option "Aucun" (après ouverture dropdown) |
   | 7 | field_lien_gps | Champ "Lien GPS" |
   | 8 | field_notes | Champ "Notes" |
   | 9 | field_ne_pas_visiter | Champ "Ne pas visiter" |
   | 10 | field_notes_proclamateur | Champ "Notes du proclamateur" |
   | 11 | btn_import_pdf | Bouton d'import PDF |
   | 12 | btn_save | Bouton Sauvegarder |

5. **Mettez à jour `config.py`** :
   ```python
   COORDINATES = {
       "btn_new_territory": (50, 150),      # Vos coordonnées
       "field_numero": (800, 200),
       # ... etc
   }
   ```

### Conseils pour la calibration

- **Maximisez la fenêtre NWS** pour des coordonnées cohérentes
- **Cliquez au centre** des champs/boutons
- Pour les dropdowns : capturez d'abord le dropdown fermé, puis ouvrez-le et capturez les options
- **Testez avec `--dry-run`** avant l'exécution réelle

---

## Préparation des données

### Créer le fichier Excel

1. Générez le template :
   ```bash
   uv run python tools/create_template.py
   ```

2. Ouvrez `data/territories_template.xlsx`

3. Remplissez avec vos données

4. Renommez en `territories.xlsx` (ou modifiez `DATA_FILE_PATH` dans config.py)

### Format des colonnes

| Colonne | Type | Exemple | Notes |
|---------|------|---------|-------|
| Numero | Texte | SAR-1-01 | **Obligatoire**. Identifiant unique |
| Suffixe | Texte | A | Optionnel |
| Type | Texte | En présentiel | "En présentiel" ou "Aucun" |
| Lien_GPS | URL | https://maps.google.com/... | Optionnel |
| Notes | Texte | Zone résidentielle | Optionnel |
| Ne_Pas_Visiter | Texte | Apt 3B | Optionnel |
| Notes_Proclamateur | Texte | Prévoir 2h | Optionnel |
| PDF_Filename | Texte | custom.pdf | Optionnel (sinon: Numero.pdf) |

### Préparer les fichiers PDF

1. Placez vos PDFs dans `data/pdfs/`

2. Nommage :
   - Par défaut : `{Numero}.pdf` (ex: `SAR-1-01.pdf`)
   - Personnalisé : remplissez la colonne `PDF_Filename`

3. Vérifiez avec :
   ```bash
   uv run python main.py --verify
   ```

---

## Utilisation

### Commandes disponibles

```bash
# Lancer l'automatisation
uv run python main.py

# Mode simulation (affiche sans exécuter)
uv run python main.py --dry-run

# Vérifier les données et PDFs
uv run python main.py --verify

# Réinitialiser la progression
uv run python main.py --reset

# Commencer à partir d'un index spécifique
uv run python main.py --start-from 10

# Utiliser un fichier de données différent
uv run python main.py --data-file chemin/vers/fichier.xlsx
```

### Workflow recommandé

1. **Préparer les données** : Excel + PDFs
2. **Vérifier** : `uv run python main.py --verify`
3. **Simuler** : `uv run python main.py --dry-run`
4. **Exécuter** : `uv run python main.py`
5. **Consulter les logs** : dossier `logs/`

### Reprise après interruption

Le script sauvegarde automatiquement la progression dans `data/progress.json`.

- Pour **continuer** : relancez simplement `uv run python main.py`
- Pour **recommencer** : `uv run python main.py --reset`

### Arrêt d'urgence

Deux méthodes :
1. **Fail-safe** : déplacez la souris dans le coin supérieur gauche
2. **Ctrl+C** dans le terminal

---

## Dépannage

### "Exécutable non trouvé"

Vérifiez le chemin dans `config.py` :
```python
NWS_EXE_PATH = r"C:\Votre\Chemin\Vers\NWScheduler.exe"
```

### "Fenêtre non trouvée"

- Lancez NWS manuellement avant l'automatisation
- Vérifiez que le titre contient "New World Scheduler"

### Clics au mauvais endroit

- Recalibrez les coordonnées
- Vérifiez que la fenêtre NWS est à la même position/taille

### L'automatisation est trop rapide

Augmentez les délais dans `config.py` :
```python
DELAY_AFTER_CLICK = 0.5  # Augmenter
DELAY_AFTER_SAVE = 2.0   # Augmenter
```

### Caractères spéciaux mal saisis

Le script utilise le presse-papiers (Ctrl+V) pour gérer les caractères spéciaux. Si problème :
- Vérifiez que votre clavier est en layout français
- Testez manuellement le copier-coller dans NWS

### Erreur "pyautogui.FailSafeException"

Vous avez déclenché l'arrêt d'urgence (souris en coin supérieur gauche).
C'est normal et prévu pour la sécurité.

### Les logs sont où ?

Dans le dossier `logs/`. Chaque exécution crée un fichier avec timestamp :
```
logs/automation_20240115_143022.log
```

---

## Support

En cas de problème :
1. Consultez les logs dans `logs/`
2. Essayez en mode `--dry-run` pour diagnostiquer
3. Vérifiez la calibration des coordonnées

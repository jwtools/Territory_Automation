# Guide d'Installation et d'Utilisation

📖 Guide complet pour l'automatisation de New World Scheduler 7.9.

## Table des matières

1. [Installation](#installation)
2. [Configuration](#configuration)
3. [Calibration des coordonnées](#calibration-des-coordonnées)
4. [Préparation des données](#préparation-des-données)
5. [Utilisation](#utilisation)
6. [Architecture technique](#architecture-technique)
7. [Dépannage](#dépannage)
8. [FAQ](#faq)

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

Les coordonnées des boutons et champs dépendent de :
- 🖥️ **Résolution d'écran** (1920x1080, 2560x1440, etc.)
- 📐 **Position et taille de la fenêtre NWS** (maximisée recommandée)
- 🔍 **Échelle d'affichage Windows** (100%, 125%, 150%, 175%)
- 🎨 **Thème et DPI** de l'interface

> ⚠️ **Important** : Une calibration précise est essentielle pour le bon fonctionnement de l'automatisation.

### Méthode 1 : Assistant guidé (recommandé)

L'assistant de calibration vous guide pas à pas pour capturer tous les éléments nécessaires.

#### Procédure complète

1. **Préparez l'environnement** :
   ```bash
   # Lancez New World Scheduler
   # Maximisez la fenêtre (recommandé)
   # Naviguez vers l'écran des territoires
   ```

2. **Lancez l'assistant de calibration** :
   ```bash
   uv run python tools/calibration.py
   ```

3. **Suivez les instructions à l'écran** :
   - L'assistant affiche le nom et la description de chaque élément
   - Positionnez votre souris sur l'élément indiqué
   - Appuyez sur `C` pour capturer les coordonnées
   - Appuyez sur `S` pour passer (si élément non disponible)
   - Appuyez sur `Q` pour quitter

4. **Éléments calibrés** (23 au total) :

   **Navigation (2 éléments)** :
   - Menu Territoires
   - Liste des territoires

   **Création (1 élément)** :
   - Bouton "+" (nouveau territoire)

   **Formulaire - Menus déroulants (8 éléments)** :
   - Catégorie (menu + option SAR)
   - Type (menu + 4 options : En présentiel, Courrier, Téléphone, Entreprise)
   - Ville (menu + 5 options)

   **Formulaire - Champs de saisie (6 éléments)** :
   - Numéro
   - Suffixe
   - Lien GPS
   - Notes
   - Ne pas visiter
   - Notes proclamateur

   **Actions (2 éléments)** :
   - Bouton Import PDF
   - Bouton Sauvegarder
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

## Architecture technique

### Technologies utilisées

**Automatisation Windows** :
- `pywinauto` : Contrôle de l'application Windows (gestion des fenêtres, focus)
- `pyautogui` : Contrôle de la souris et du clavier (clics, saisie)
- `keyboard` : Détection des touches clavier pour la calibration

**Traitement de données** :
- `pandas` : Lecture et manipulation des fichiers Excel/CSV
- `openpyxl` : Création de fichiers Excel (templates)
- `pyperclip` : Gestion du presse-papiers (copier-coller)

**Gestion de projet** :
- `uv` : Gestionnaire de paquets et environnements virtuels rapide
- `pyproject.toml` : Configuration moderne du projet Python

### Flux d'exécution

```
┌─────────────────────────────────────────────────────────────┐
│ 1. Chargement des données                                  │
│    - Lecture du fichier Excel                               │
│    - Validation des colonnes obligatoires                   │
│    - Vérification des fichiers PDF                          │
└────────────────────┬────────────────────────────────────────┘
                     │
                     ▼
┌─────────────────────────────────────────────────────────────┐
│ 2. Connexion à NWS                                          │
│    - Lancement de l'application (si nécessaire)             │
│    - Recherche et focus de la fenêtre                       │
│    - Fermeture des dialogues de démarrage                   │
└────────────────────┬────────────────────────────────────────┘
                     │
                     ▼
┌─────────────────────────────────────────────────────────────┐
│ 3. Navigation vers l'écran des territoires                  │
│    - Clic sur menu Territoires                              │
│    - Clic sur Liste des territoires                         │
└────────────────────┬────────────────────────────────────────┘
                     │
                     ▼
┌─────────────────────────────────────────────────────────────┐
│ 4. Pour chaque territoire (boucle)                          │
│    ┌───────────────────────────────────────────────────┐   │
│    │ a. Création nouveau territoire (clic sur +)       │   │
│    └───────────────┬───────────────────────────────────┘   │
│                    │                                         │
│    ┌───────────────▼───────────────────────────────────┐   │
│    │ b. Remplissage du formulaire                      │   │
│    │    - Catégorie (SAR)                              │   │
│    │    - Numéro                                       │   │
│    │    - Suffixe (optionnel)                          │   │
│    │    - Type (optionnel)                             │   │
│    │    - Ville (optionnel)                            │   │
│    │    - Lien GPS (optionnel)                         │   │
│    │    - Notes (optionnel)                            │   │
│    │    - Ne pas visiter (optionnel)                   │   │
│    │    - Notes proclamateur (optionnel)               │   │
│    └───────────────┬───────────────────────────────────┘   │
│                    │                                         │
│    ┌───────────────▼───────────────────────────────────┐   │
│    │ c. Import du PDF                                  │   │
│    │    - Clic sur bouton import                       │   │
│    │    - Sélection du fichier                         │   │
│    └───────────────┬───────────────────────────────────┘   │
│                    │                                         │
│    ┌───────────────▼───────────────────────────────────┐   │
│    │ d. Sauvegarde                                     │   │
│    │    - Clic sur bouton Sauvegarder                  │   │
│    │    - Attente de la sauvegarde                     │   │
│    └───────────────┬───────────────────────────────────┘   │
│                    │                                         │
│    ┌───────────────▼───────────────────────────────────┐   │
│    │ e. Mise à jour de la progression                  │   │
│    │    - Sauvegarde dans progress.json                │   │
│    └───────────────────────────────────────────────────┘   │
└─────────────────────────────────────────────────────────────┘
```

### Gestion des erreurs

**Stratégies de résilience** :
- ✅ **Fail-safe** : Arrêt immédiat si souris en coin supérieur gauche
- ✅ **Retry logic** : Tentatives multiples pour les actions critiques
- ✅ **Progress tracking** : Sauvegarde après chaque territoire
- ✅ **Logging détaillé** : Enregistrement de toutes les actions et erreurs
- ✅ **Validation** : Vérification des données avant exécution

**Types d'erreurs gérées** :
- Fenêtre NWS introuvable
- Fichier PDF manquant
- Coordonnées incorrectes
- Délai d'attente dépassé
- Données Excel invalides

### Modes d'exécution

1. **Mode normal** : Exécution complète avec sauvegarde
2. **Mode dry-run** : Affichage des actions sans exécution
3. **Mode no-save** : Remplissage sans sauvegarde (validation visuelle)
4. **Mode verify** : Vérification des données et PDFs uniquement

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

### PDF non importé

- Vérifiez que le fichier existe dans `data/pdfs/`
- Vérifiez le nom du fichier (sensible à la casse)
- Vérifiez que c'est bien un PDF valide
- Utilisez `--verify` pour lister les PDFs manquants

---

## FAQ

### Puis-je utiliser un autre format que Excel ?

Oui, les fichiers CSV sont également supportés. Assurez-vous que les colonnes ont les bons noms.

### Combien de temps prend l'automatisation ?

Environ 20-30 secondes par territoire en moyenne, selon :
- La complexité des données
- La taille du PDF
- La vitesse de votre ordinateur
- Les délais configurés

### L'automatisation fonctionne-t-elle en arrière-plan ?

Non, la fenêtre NWS doit rester visible et au premier plan. Ne minimisez pas la fenêtre pendant l'exécution.

### Puis-je modifier les coordonnées manuellement ?

Oui, éditez le fichier `data/calibration.json` :
```json
{
  "btn_new_territory": [100, 200],
  "field_numero": [300, 250]
}
```

### Comment changer la catégorie par défaut (SAR) ?

Modifiez le fichier `territory_automation/automation.py`, fonction `fill_territory_form()`, section catégorie.

### L'automatisation supporte-t-elle plusieurs catégories ?

Actuellement non, mais vous pouvez facilement modifier le code pour ajouter une colonne "Catégorie" dans Excel et adapter la logique.

### Puis-je exécuter plusieurs instances en parallèle ?

Non recommandé. Cela causerait des conflits avec le contrôle de la souris et du clavier.

### Comment sauvegarder ma configuration ?

Sauvegardez ces fichiers :
- `data/calibration.json` (coordonnées)
- `config.py` (configuration personnalisée)
- `data/territories.xlsx` (vos données)

### Le projet est-il open source ?

Oui, vous pouvez modifier et adapter le code selon vos besoins. Le code est documenté et modulaire.

---

## Support

En cas de problème :

1. 📝 **Consultez les logs** dans `logs/` pour identifier l'erreur exacte
2. 🎭 **Testez en mode dry-run** : `uv run python main.py --dry-run`
3. 🔍 **Vérifiez les données** : `uv run python main.py --verify`
4. 🎯 **Recalibrez** si nécessaire : `uv run python tools/calibration.py`
5. 🔄 **Testez la calibration** : `uv run python tools/test_calibration.py`

### Checklist de dépannage

- [ ] NWS est-il lancé et visible ?
- [ ] La fenêtre est-elle maximisée ?
- [ ] Les coordonnées sont-elles calibrées ?
- [ ] Le fichier Excel est-il valide ?
- [ ] Les PDFs sont-ils dans `data/pdfs/` ?
- [ ] Les logs montrent-ils une erreur spécifique ?
- [ ] Avez-vous testé avec `--dry-run` ?

**Pour aller plus loin** : Consultez le code source dans `territory_automation/` pour comprendre le fonctionnement interne.

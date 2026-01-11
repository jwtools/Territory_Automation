"""
Configuration pour l'automatisation de New World Scheduler 7.9
Modifiez ces valeurs selon votre environnement.
"""

import json
from pathlib import Path


def _load_calibration() -> dict:
    """
    Charge les coordonnées calibrées depuis le fichier JSON.
    Retourne un dict vide si le fichier n'existe pas.
    """
    calibration_file = Path(__file__).parent / "data" / "calibration.json"
    if calibration_file.exists():
        try:
            with open(calibration_file, "r", encoding="utf-8") as f:
                data = json.load(f)
                # Convertir les listes en tuples
                return {k: tuple(v) if isinstance(v, list) else v for k, v in data.items()}
        except (json.JSONDecodeError, IOError):
            pass
    return {}


def _load_options() -> dict:
    """
    Charge les options de catégories et villes depuis le fichier JSON.
    Retourne un dict avec les valeurs par défaut si le fichier n'existe pas.
    """
    options_file = Path(__file__).parent / "data" / "options.json"
    default_options = {
        "categories": {"SAR": "dropdown_option_sar"},
        "villes": {
            "AUCUN": "dropdown_ville_aucun",
            "": "dropdown_ville_aucun",
            "CARRIERE S/ BOIS": "dropdown_ville_carrieres",
            "CARRIERES": "dropdown_ville_carrieres",
            "MAISONS-LAFFITTE": "dropdown_ville_maisons",
            "MAISONS LAFFITTE": "dropdown_ville_maisons",
            "MESNIL LE ROI": "dropdown_ville_mesnil",
            "MONTESSON": "dropdown_ville_montesson",
            "SARTROUVILLE": "dropdown_ville_sartrouville",
        }
    }
    if options_file.exists():
        try:
            with open(options_file, "r", encoding="utf-8") as f:
                data = json.load(f)
                # Fusionner avec les valeurs par défaut
                return {
                    "categories": data.get("categories", default_options["categories"]),
                    "villes": data.get("villes", default_options["villes"]),
                }
        except (json.JSONDecodeError, IOError):
            pass
    return default_options

# =============================================================================
# CHEMINS DES FICHIERS
# =============================================================================

# Chemin vers l'exécutable de New World Scheduler
NWS_EXE_PATH = r"C:\Program Files\New World Scheduler\NW Scheduler.exe"

# Chemin vers le fichier Excel/CSV contenant les données des territoires
DATA_FILE_PATH = Path(__file__).parent / "data" / "territories.xlsx"

# Dossier contenant les fichiers PDF des territoires
PDF_FOLDER_PATH = Path(__file__).parent / "data" / "pdfs"

# Dossier pour les fichiers de log
LOG_FOLDER_PATH = Path(__file__).parent / "logs"

# Fichier de progression (pour reprendre après interruption)
PROGRESS_FILE_PATH = Path(__file__).parent / "data" / "progress.json"

# =============================================================================
# PARAMÈTRES D'AUTOMATISATION
# =============================================================================

# Délais en secondes
DELAY_AFTER_CLICK = 0.3          # Délai après chaque clic
DELAY_AFTER_TYPE = 0.1           # Délai après chaque saisie
DELAY_APP_LAUNCH = 10.0          # Délai pour le lancement de l'application
DELAY_AFTER_SAVE = 1.0           # Délai après sauvegarde
DELAY_BETWEEN_TERRITORIES = 0.5  # Délai entre chaque territoire

# Timeout maximum pour attendre un élément (en secondes)
ELEMENT_TIMEOUT = 10

# Nombre de tentatives en cas d'échec
MAX_RETRIES = 3

# =============================================================================
# COORDONNÉES DE L'INTERFACE
# =============================================================================
# Ces coordonnées sont chargées depuis data/calibration.json si disponible.
# Utilisez l'outil de calibration pour les définir:
#     uv run python tools/calibration.py
#
# Les valeurs ci-dessous sont des valeurs par défaut (à remplacer).

# Valeurs par défaut (seront écrasées par la calibration si elle existe)
_DEFAULT_COORDINATES = {
    # Navigation
    "btn_menu_territoires": (100, 80),
    "btn_liste_territoires": (150, 120),

    # Création
    "btn_new_territory": (50, 150),

    # Formulaire (ordre de saisie)
    "dropdown_categorie": (800, 175),      # 1. Catégorie
    "dropdown_option_sar": (800, 195),     # Option SAR dans Catégorie
    "field_numero": (800, 200),            # 2. Numéro
    "field_suffixe": (800, 225),           # 3. Suffixe
    "dropdown_type": (800, 250),           # 4. Type
    "dropdown_option_presentiel": (800, 270),
    "dropdown_option_courrier": (800, 290),
    "dropdown_option_telephone": (800, 310),
    "dropdown_option_entreprise": (800, 330),
    "btn_confirm_type": (500, 400),        # Bouton "Oui" modal confirmation type
    "dropdown_ville": (800, 375),          # 5. Ville
    "dropdown_ville_aucun": (800, 395),
    "dropdown_ville_carrieres": (800, 415),
    "dropdown_ville_maisons": (800, 435),
    "dropdown_ville_mesnil": (800, 455),
    "dropdown_ville_montesson": (800, 475),
    "dropdown_ville_sartrouville": (800, 495),
    "field_lien_gps": (800, 400),          # 6. Lien GPS
    "field_notes": (800, 400),
    "field_ne_pas_visiter": (800, 450),
    "field_notes_proclamateur": (800, 500),
    "btn_carte": (900, 50),                # Bouton/onglet Carte

    # Actions
    "btn_import_pdf": (800, 550),
}

# Charger la calibration et fusionner avec les valeurs par défaut
_calibrated = _load_calibration()
COORDINATES = {**_DEFAULT_COORDINATES, **_calibrated}

# Charger les options de catégories et villes
_options = _load_options()
CATEGORIES = _options["categories"]
VILLES = _options["villes"]

# Afficher un avertissement si pas de calibration
if not _calibrated:
    import sys
    print("[ATTENTION] Aucune calibration trouvee. Executez:", file=sys.stderr)
    print("    uv run python tools/calibration.py", file=sys.stderr)

# =============================================================================
# TITRE DE LA FENÊTRE
# =============================================================================

# Titre de la fenêtre de l'application (pour la détecter)
NWS_WINDOW_TITLE = "NW Scheduler"

# =============================================================================
# GESTION DES DIALOGUES DE DÉMARRAGE (Astuces, Tips, etc.)
# =============================================================================

# Titres possibles des fenêtres de dialogue à fermer au démarrage
STARTUP_DIALOG_TITLES = [
    "Astuce",
    "Tip",
    "Conseil",
    "Bienvenue",
    "Welcome",
    "Did you know",
    "Le saviez-vous",
]

# Méthode pour fermer les dialogues: "escape", "enter", "click_close", "click_ok"
STARTUP_DIALOG_CLOSE_METHOD = "escape"

# Coordonnées du bouton "Fermer" ou "OK" dans la fenêtre d'astuce (si click_close/click_ok)
# À calibrer avec coordinate_finder.py si nécessaire
STARTUP_DIALOG_CLOSE_BUTTON = (960, 600)  # Centre de l'écran approximatif

# Délai d'attente pour les dialogues de démarrage (secondes)
STARTUP_DIALOG_WAIT = 2.0

# =============================================================================
# OPTIONS DE MAPPING DES COLONNES EXCEL
# =============================================================================

# Mapping entre les colonnes Excel et les champs du formulaire
EXCEL_COLUMNS = {
    "numero": "Numero",
    "suffixe": "Suffixe",
    "categorie": "Categorie",  # Configurable (voir data/options.json)
    "type": "Type",
    "ville": "Ville",  # Configurable (voir data/options.json)
    "lien_gps": "Lien_GPS",
    "notes": "Notes",
    "ne_pas_visiter": "Ne_Pas_Visiter",
    "notes_proclamateur": "Notes_Proclamateur",
    "pdf_filename": "PDF_Filename",  # Optionnel: nom du fichier PDF si différent
}

# Valeurs acceptées pour le champ "Type"
TYPE_VALUES = {
    "presentiel": "En présentiel",
    "en présentiel": "En présentiel",
    "courrier": "Courrier",
    "telephone": "Téléphone",
    "téléphone": "Téléphone",
    "entreprise": "Entreprise",
}

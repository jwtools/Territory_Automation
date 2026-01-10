#!/usr/bin/env python3
"""
Outil de calibration guidé pour New World Scheduler.

Guide l'utilisateur étape par étape pour capturer les coordonnées
de chaque élément de l'interface.

Usage:
    uv run python tools/calibration.py
"""

import json
import sys
import time
from pathlib import Path

try:
    import pyautogui
except ImportError:
    print("ERREUR: pyautogui n'est pas installe.")
    print("Installez-le avec: uv sync")
    sys.exit(1)

try:
    import keyboard
except ImportError:
    print("ERREUR: keyboard n'est pas installe.")
    print("Installez-le avec: uv sync")
    sys.exit(1)


# Chemin du fichier de calibration
CALIBRATION_FILE = Path(__file__).parent.parent / "data" / "calibration.json"

# Éléments à calibrer avec leurs descriptions
ELEMENTS_TO_CALIBRATE = [
    # Navigation
    {
        "id": "btn_menu_territoires",
        "name": "Menu Territoires",
        "description": "Le bouton ou menu 'Territoires' dans la barre de navigation",
        "group": "Navigation"
    },
    {
        "id": "btn_liste_territoires",
        "name": "Liste des territoires",
        "description": "Le bouton 'Liste des territoires' (apres avoir clique sur Territoires)",
        "group": "Navigation"
    },
    # Création territoire
    {
        "id": "btn_new_territory",
        "name": "Bouton Nouveau (+)",
        "description": "Le bouton '+' pour creer un nouveau territoire",
        "group": "Creation"
    },
    # Champs du formulaire (ordre de saisie)
    # 1. Catégorie
    {
        "id": "dropdown_categorie",
        "name": "Menu deroulant Categorie",
        "description": "Le menu deroulant pour selectionner la categorie (1er champ)",
        "group": "Formulaire"
    },
    {
        "id": "dropdown_option_sar",
        "name": "Option 'SAR'",
        "description": "L'option 'SAR' dans le menu Categorie (ouvrez-le d'abord)",
        "group": "Formulaire"
    },
    # 2. Numéro
    {
        "id": "field_numero",
        "name": "Champ Numero",
        "description": "Le champ de saisie du numero de territoire",
        "group": "Formulaire"
    },
    # 3. Suffixe
    {
        "id": "field_suffixe",
        "name": "Champ Suffixe",
        "description": "Le champ de saisie du suffixe (A, B, etc.)",
        "group": "Formulaire"
    },
    # 4. Type
    {
        "id": "dropdown_type",
        "name": "Menu deroulant Type",
        "description": "Le menu deroulant pour selectionner le type",
        "group": "Formulaire"
    },
    {
        "id": "dropdown_option_presentiel",
        "name": "Option 'En presentiel'",
        "description": "L'option 'En presentiel' dans le menu deroulant Type (ouvrez-le d'abord)",
        "group": "Formulaire"
    },
    {
        "id": "dropdown_option_courrier",
        "name": "Option 'Courrier'",
        "description": "L'option 'Courrier' dans le menu deroulant Type",
        "group": "Formulaire"
    },
    {
        "id": "dropdown_option_telephone",
        "name": "Option 'Telephone'",
        "description": "L'option 'Telephone' dans le menu deroulant Type",
        "group": "Formulaire"
    },
    {
        "id": "dropdown_option_entreprise",
        "name": "Option 'Entreprise'",
        "description": "L'option 'Entreprise' dans le menu deroulant Type",
        "group": "Formulaire"
    },
    {
        "id": "btn_confirm_type",
        "name": "Bouton 'Oui' confirmation Type",
        "description": "Le bouton 'Oui' du modal qui apparait quand on change de type (pas En presentiel)",
        "group": "Formulaire"
    },
    # 5. Ville
    {
        "id": "dropdown_ville",
        "name": "Menu deroulant Ville",
        "description": "Le menu deroulant pour selectionner la ville (apres Type)",
        "group": "Formulaire"
    },
    {
        "id": "dropdown_ville_aucun",
        "name": "Option 'Aucun'",
        "description": "L'option 'Aucun' dans le menu Ville (ouvrez-le d'abord)",
        "group": "Formulaire"
    },
    {
        "id": "dropdown_ville_carrieres",
        "name": "Option 'CARRIERE S/ BOIS'",
        "description": "L'option 'CARRIERE S/ BOIS' dans le menu Ville",
        "group": "Formulaire"
    },
    {
        "id": "dropdown_ville_maisons",
        "name": "Option 'MAISONS-LAFFITTE'",
        "description": "L'option 'MAISONS-LAFFITTE' dans le menu Ville",
        "group": "Formulaire"
    },
    {
        "id": "dropdown_ville_mesnil",
        "name": "Option 'MESNIL LE ROI'",
        "description": "L'option 'MESNIL LE ROI' dans le menu Ville",
        "group": "Formulaire"
    },
    {
        "id": "dropdown_ville_montesson",
        "name": "Option 'MONTESSON'",
        "description": "L'option 'MONTESSON' dans le menu Ville",
        "group": "Formulaire"
    },
    {
        "id": "dropdown_ville_sartrouville",
        "name": "Option 'SARTROUVILLE'",
        "description": "L'option 'SARTROUVILLE' dans le menu Ville",
        "group": "Formulaire"
    },
    # 6. Lien GPS
    {
        "id": "field_lien_gps",
        "name": "Champ Lien GPS",
        "description": "Le champ de saisie du lien GPS",
        "group": "Formulaire"
    },
    {
        "id": "field_notes",
        "name": "Champ Notes",
        "description": "Le champ de saisie des notes generales",
        "group": "Formulaire"
    },
    {
        "id": "field_ne_pas_visiter",
        "name": "Champ Ne pas visiter",
        "description": "Le champ pour les adresses a ne pas visiter",
        "group": "Formulaire"
    },
    {
        "id": "field_notes_proclamateur",
        "name": "Champ Notes proclamateur",
        "description": "Le champ pour les notes du proclamateur",
        "group": "Formulaire"
    },
    # Navigation formulaire
    {
        "id": "btn_carte",
        "name": "Bouton Carte",
        "description": "L'onglet ou bouton 'Carte' pour passer a l'ecran de la carte",
        "group": "Formulaire"
    },
    # Actions
    {
        "id": "btn_import_pdf",
        "name": "Bouton Ajouter fichier",
        "description": "Le bouton 'Ajouter fichier' pour importer un PDF",
        "group": "Actions"
    },
]


def clear_screen():
    """Efface l'écran du terminal."""
    print("\033[2J\033[H", end="")


def print_header():
    """Affiche l'en-tête."""
    print("=" * 60)
    print("   CALIBRATION DE NEW WORLD SCHEDULER")
    print("=" * 60)
    print()


def print_instructions():
    """Affiche les instructions générales."""
    print("INSTRUCTIONS:")
    print("-" * 60)
    print("  1. Ouvrez New World Scheduler")
    print("  2. Naviguez vers l'ecran approprie pour chaque element")
    print("  3. Placez votre souris sur l'element indique")
    print("  4. Appuyez sur [C] pour capturer les coordonnees")
    print("  5. Appuyez sur [S] pour passer un element")
    print("  6. Appuyez sur [Q] pour quitter et sauvegarder")
    print("-" * 60)
    print()


def load_existing_calibration() -> dict:
    """Charge la calibration existante si elle existe."""
    if CALIBRATION_FILE.exists():
        try:
            with open(CALIBRATION_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except (json.JSONDecodeError, IOError):
            pass
    return {}


def save_calibration(coordinates: dict):
    """Sauvegarde la calibration dans le fichier JSON."""
    CALIBRATION_FILE.parent.mkdir(parents=True, exist_ok=True)

    with open(CALIBRATION_FILE, "w", encoding="utf-8") as f:
        json.dump(coordinates, f, indent=2, ensure_ascii=False)

    print(f"\nCalibration sauvegardee dans: {CALIBRATION_FILE}")


def capture_element(element: dict, existing_coords: dict) -> tuple:
    """
    Capture les coordonnées d'un élément.

    Returns:
        Tuple (x, y) ou None si passé
    """
    element_id = element["id"]
    existing = existing_coords.get(element_id)

    print()
    print(f"[{element['group']}] {element['name']}")
    print(f"   {element['description']}")

    if existing:
        print(f"   Valeur actuelle: ({existing[0]}, {existing[1]})")

    print()
    print("   Placez votre souris sur cet element...")
    print("   [C] Capturer  |  [S] Passer  |  [R] Reprendre valeur actuelle")
    print()

    captured = None
    waiting = True

    def on_capture():
        nonlocal captured, waiting
        x, y = pyautogui.position()
        captured = (x, y)
        waiting = False

    def on_skip():
        nonlocal waiting
        waiting = False

    def on_reuse():
        nonlocal captured, waiting
        if existing:
            captured = tuple(existing)
        waiting = False

    keyboard.on_press_key('c', lambda _: on_capture())
    keyboard.on_press_key('s', lambda _: on_skip())
    keyboard.on_press_key('r', lambda _: on_reuse())

    try:
        while waiting:
            x, y = pyautogui.position()
            sys.stdout.write(f"\r   Position actuelle: ({x:4d}, {y:4d})    ")
            sys.stdout.flush()
            time.sleep(0.1)
    finally:
        keyboard.unhook_all()

    if captured:
        print(f"\n   [OK] Capture: ({captured[0]}, {captured[1]})")
    else:
        print(f"\n   [--] Passe")

    return captured


def run_calibration():
    """Exécute la calibration guidée."""
    clear_screen()
    print_header()
    print_instructions()

    # Charger la calibration existante
    coordinates = load_existing_calibration()

    if coordinates:
        print(f"Calibration existante trouvee ({len(coordinates)} elements)")
        print("Les valeurs existantes seront proposees par defaut.")
        print()

    input("Appuyez sur Entree pour commencer...")

    # Grouper les éléments
    current_group = None
    total = len(ELEMENTS_TO_CALIBRATE)

    for i, element in enumerate(ELEMENTS_TO_CALIBRATE):
        # Afficher le groupe si changement
        if element["group"] != current_group:
            current_group = element["group"]
            print()
            print("=" * 60)
            print(f"  SECTION: {current_group.upper()}")
            print("=" * 60)

        print(f"\n[{i+1}/{total}]", end="")

        result = capture_element(element, coordinates)

        if result:
            coordinates[element["id"]] = list(result)

    # Sauvegarder
    print()
    print("=" * 60)
    print("  CALIBRATION TERMINEE")
    print("=" * 60)

    save_calibration(coordinates)

    # Afficher le résumé
    print()
    print("Resume des coordonnees capturees:")
    print("-" * 60)
    for element in ELEMENTS_TO_CALIBRATE:
        eid = element["id"]
        if eid in coordinates:
            coords = coordinates[eid]
            print(f"  {eid}: ({coords[0]}, {coords[1]})")
        else:
            print(f"  {eid}: (non defini)")

    print()
    print("Ces coordonnees seront utilisees automatiquement par le script.")
    print()


def main():
    """Point d'entrée principal."""
    try:
        run_calibration()
    except KeyboardInterrupt:
        print("\n\nCalibration interrompue.")
        # Sauvegarder quand même ce qui a été capturé
        coordinates = load_existing_calibration()
        if coordinates:
            save_calibration(coordinates)
    except Exception as e:
        print(f"\nErreur: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()

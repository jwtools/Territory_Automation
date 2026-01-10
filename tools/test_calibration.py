#!/usr/bin/env python3
"""
Outil de test des coordonnées calibrées.

Permet de vérifier visuellement que les coordonnées capturées
sont correctes en déplaçant la souris ou en cliquant sur chaque élément.

Usage:
    uv run python tools/test_calibration.py           # Mode survol (déplace la souris)
    uv run python tools/test_calibration.py --click   # Mode clic (clique sur chaque élément)
"""

import argparse
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

# Désactiver le failsafe pendant le test (optionnel)
# pyautogui.FAILSAFE = False

# Chemin du fichier de calibration
CALIBRATION_FILE = Path(__file__).parent.parent / "data" / "calibration.json"

# Descriptions des éléments pour l'affichage
ELEMENT_DESCRIPTIONS = {
    "btn_menu_territoires": "Menu Territoires",
    "btn_liste_territoires": "Liste des territoires",
    "btn_new_territory": "Bouton Nouveau (+)",
    "dropdown_categorie": "Menu deroulant Categorie",
    "dropdown_option_sar": "Option 'SAR'",
    "field_numero": "Champ Numero",
    "field_suffixe": "Champ Suffixe",
    "dropdown_type": "Menu deroulant Type",
    "dropdown_option_presentiel": "Option 'En presentiel'",
    "dropdown_option_courrier": "Option 'Courrier'",
    "dropdown_option_telephone": "Option 'Telephone'",
    "dropdown_option_entreprise": "Option 'Entreprise'",
    "btn_confirm_type": "Bouton 'Oui' confirmation Type",
    "dropdown_ville": "Menu deroulant Ville",
    "dropdown_ville_aucun": "Option 'Aucun'",
    "dropdown_ville_carrieres": "Option 'CARRIERE S/ BOIS'",
    "dropdown_ville_maisons": "Option 'MAISONS-LAFFITTE'",
    "dropdown_ville_mesnil": "Option 'MESNIL LE ROI'",
    "dropdown_ville_montesson": "Option 'MONTESSON'",
    "dropdown_ville_sartrouville": "Option 'SARTROUVILLE'",
    "field_lien_gps": "Champ Lien GPS",
    "field_notes": "Champ Notes",
    "field_ne_pas_visiter": "Champ Ne pas visiter",
    "field_notes_proclamateur": "Champ Notes proclamateur",
    "btn_carte": "Bouton Carte",
    "btn_import_pdf": "Bouton Ajouter fichier",
}

# Ordre de test (pour suivre le workflow logique)
TEST_ORDER = [
    "btn_menu_territoires",
    "btn_liste_territoires",
    "btn_new_territory",
    "dropdown_categorie",
    "dropdown_option_sar",
    "field_numero",
    "field_suffixe",
    "dropdown_type",
    "dropdown_option_presentiel",
    "dropdown_option_courrier",
    "dropdown_option_telephone",
    "dropdown_option_entreprise",
    "btn_confirm_type",
    "dropdown_ville",
    "dropdown_ville_aucun",
    "dropdown_ville_carrieres",
    "dropdown_ville_maisons",
    "dropdown_ville_mesnil",
    "dropdown_ville_montesson",
    "dropdown_ville_sartrouville",
    "field_lien_gps",
    "field_notes",
    "field_ne_pas_visiter",
    "field_notes_proclamateur",
    "btn_carte",
    "btn_import_pdf",
]


def load_calibration() -> dict:
    """Charge les coordonnées calibrées."""
    if not CALIBRATION_FILE.exists():
        print(f"ERREUR: Fichier de calibration non trouve: {CALIBRATION_FILE}")
        print("Executez d'abord: uv run python tools/calibration.py")
        sys.exit(1)

    with open(CALIBRATION_FILE, "r", encoding="utf-8") as f:
        return json.load(f)


def parse_args():
    """Parse les arguments de la ligne de commande."""
    parser = argparse.ArgumentParser(
        description="Test des coordonnees calibrees pour New World Scheduler"
    )
    parser.add_argument(
        "--click",
        action="store_true",
        help="Mode clic: clique sur chaque element (au lieu de juste survoler)"
    )
    parser.add_argument(
        "--delay",
        type=float,
        default=1.5,
        help="Delai entre chaque element en secondes (defaut: 1.5)"
    )
    parser.add_argument(
        "--pause",
        action="store_true",
        help="Attendre une touche entre chaque element"
    )
    parser.add_argument(
        "--element",
        type=str,
        help="Tester uniquement cet element (ex: btn_new_territory)"
    )
    parser.add_argument(
        "--list",
        action="store_true",
        help="Afficher la liste des elements calibres et quitter"
    )
    return parser.parse_args()


def highlight_position(x: int, y: int, duration: float = 0.3):
    """
    Fait un petit mouvement circulaire autour de la position
    pour la mettre en évidence visuellement.
    """
    import math
    radius = 20
    steps = 16

    for i in range(steps + 1):
        angle = (2 * math.pi * i) / steps
        new_x = x + int(radius * math.cos(angle))
        new_y = y + int(radius * math.sin(angle))
        pyautogui.moveTo(new_x, new_y, duration=duration/steps)

    # Revenir au centre
    pyautogui.moveTo(x, y, duration=0.1)


def test_element(element_id: str, coords: list, click_mode: bool, pause_mode: bool):
    """Teste un élément en déplaçant la souris ou en cliquant."""
    x, y = coords[0], coords[1]
    description = ELEMENT_DESCRIPTIONS.get(element_id, element_id)

    print(f"\n  [{element_id}]")
    print(f"  {description}")
    print(f"  Coordonnees: ({x}, {y})")

    # Déplacer la souris vers la position
    print(f"  -> Deplacement vers ({x}, {y})...")
    pyautogui.moveTo(x, y, duration=0.5)

    # Mettre en évidence avec un cercle
    highlight_position(x, y, duration=0.4)

    if click_mode:
        print(f"  -> Clic!")
        pyautogui.click(x, y)
        time.sleep(0.3)

    if pause_mode:
        input("  Appuyez sur Entree pour continuer...")

    return True


def main():
    """Point d'entrée principal."""
    args = parse_args()

    # Charger la calibration
    calibration = load_calibration()

    print("=" * 60)
    print("  TEST DES COORDONNEES CALIBREES")
    print("=" * 60)
    print()

    # Mode liste uniquement
    if args.list:
        print("Elements calibres:")
        print("-" * 60)
        for element_id in TEST_ORDER:
            if element_id in calibration:
                coords = calibration[element_id]
                desc = ELEMENT_DESCRIPTIONS.get(element_id, "")
                print(f"  {element_id}: ({coords[0]}, {coords[1]}) - {desc}")
            else:
                print(f"  {element_id}: (non calibre)")
        print()
        return

    # Informations sur le mode
    mode = "CLIC" if args.click else "SURVOL"
    print(f"Mode: {mode}")
    print(f"Delai entre elements: {args.delay}s")
    if args.pause:
        print("Pause entre chaque element: OUI")
    print()

    if args.click:
        print("[ATTENTION] Mode clic actif!")
        print("Assurez-vous que NWS est ouvert et au premier plan.")
        print("Deplacez la souris en haut a gauche pour arreter (failsafe).")
        print()

    print("-" * 60)

    # Tester un seul élément si spécifié
    if args.element:
        if args.element not in calibration:
            print(f"ERREUR: Element '{args.element}' non trouve dans la calibration.")
            print(f"Elements disponibles: {list(calibration.keys())}")
            sys.exit(1)

        input(f"Appuyez sur Entree pour tester '{args.element}'...")
        test_element(args.element, calibration[args.element], args.click, args.pause)
        print("\nTest termine!")
        return

    # Tester tous les éléments
    input("Appuyez sur Entree pour commencer le test...")

    elements_to_test = [e for e in TEST_ORDER if e in calibration]
    total = len(elements_to_test)

    print(f"\nTest de {total} elements...")

    for i, element_id in enumerate(elements_to_test):
        print(f"\n[{i+1}/{total}]", end="")

        try:
            test_element(
                element_id,
                calibration[element_id],
                args.click,
                args.pause
            )
        except pyautogui.FailSafeException:
            print("\n\n[!] Arret par failsafe (souris en coin superieur gauche)")
            break
        except KeyboardInterrupt:
            print("\n\n[!] Arret par l'utilisateur (Ctrl+C)")
            break

        if not args.pause:
            time.sleep(args.delay)

    print("\n" + "=" * 60)
    print("  TEST TERMINE")
    print("=" * 60)
    print()
    print("Si certaines coordonnees sont incorrectes, relancez la calibration:")
    print("  uv run python tools/calibration.py")
    print()


if __name__ == "__main__":
    main()

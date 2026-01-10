#!/usr/bin/env python3
"""
Outil pour trouver les coordonnées des éléments de l'interface.

Utilisez cet outil pour déterminer les coordonnées des boutons et champs
de New World Scheduler sur votre écran.

Usage:
    python tools/coordinate_finder.py

Instructions:
    1. Lancez ce script
    2. Positionnez votre souris sur l'élément à capturer
    3. Appuyez sur 'c' pour capturer les coordonnées
    4. Appuyez sur 'q' pour quitter
    5. Copiez les coordonnées dans config.py
"""

import sys
import time

try:
    import pyautogui
except ImportError:
    print("ERREUR: pyautogui n'est pas installé.")
    print("Installez-le avec: pip install pyautogui")
    sys.exit(1)

try:
    import keyboard
except ImportError:
    print("ERREUR: keyboard n'est pas installé.")
    print("Installez-le avec: pip install keyboard")
    sys.exit(1)


def main():
    """Point d'entrée principal."""
    print("=" * 60)
    print("       OUTIL DE CAPTURE DE COORDONNÉES")
    print("=" * 60)
    print()
    print("Instructions:")
    print("  1. Ouvrez New World Scheduler")
    print("  2. Positionnez votre souris sur un élément")
    print("  3. Appuyez sur 'C' pour capturer les coordonnées")
    print("  4. Appuyez sur 'Q' pour quitter")
    print()
    print("Les coordonnées seront affichées et prêtes à copier.")
    print("-" * 60)
    print()

    captured = []
    running = True

    def on_capture():
        x, y = pyautogui.position()
        captured.append((x, y))
        print(f"  [{len(captured)}] Coordonnées capturées: ({x}, {y})")

    def on_quit():
        nonlocal running
        running = False

    # Enregistrer les raccourcis
    keyboard.on_press_key('c', lambda _: on_capture())
    keyboard.on_press_key('q', lambda _: on_quit())

    print("Position actuelle de la souris:")
    print("(Appuyez sur C pour capturer, Q pour quitter)")
    print()

    try:
        while running:
            x, y = pyautogui.position()
            # Affichage en temps réel
            sys.stdout.write(f"\r  Souris: ({x:4d}, {y:4d})   ")
            sys.stdout.flush()
            time.sleep(0.1)
    except KeyboardInterrupt:
        pass
    finally:
        keyboard.unhook_all()

    print("\n")
    print("-" * 60)
    print("COORDONNÉES CAPTURÉES:")
    print("-" * 60)

    if captured:
        print()
        print("# Copiez ces lignes dans config.py COORDINATES = {")
        for i, (x, y) in enumerate(captured, 1):
            print(f'    "element_{i}": ({x}, {y}),')
        print("# }")
        print()

        # Générer un format plus utile
        print("Format Python:")
        for i, (x, y) in enumerate(captured, 1):
            print(f"    \"element_{i}\": ({x}, {y}),")
    else:
        print("Aucune coordonnée capturée.")

    print()
    print("Terminé!")


if __name__ == "__main__":
    main()

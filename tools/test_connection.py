#!/usr/bin/env python3
"""
Test de connexion à New World Scheduler.

Ce script vérifie si NWS est accessible et permet de tester
les interactions de base.

Usage:
    python tools/test_connection.py
"""

import sys
import time
from pathlib import Path

# Ajouter le dossier parent au path
sys.path.insert(0, str(Path(__file__).parent.parent))

try:
    import pyautogui
except ImportError:
    print("ERREUR: pyautogui n'est pas installé.")
    print("Installez-le avec: pip install pyautogui")
    sys.exit(1)

try:
    from pywinauto import Application, Desktop
    from pywinauto.findwindows import ElementNotFoundError
    PYWINAUTO_AVAILABLE = True
except ImportError:
    PYWINAUTO_AVAILABLE = False
    print("AVERTISSEMENT: pywinauto n'est pas installé.")
    print("Certaines fonctionnalités seront limitées.")

from config import NWS_EXE_PATH, NWS_WINDOW_TITLE


def check_pyautogui():
    """Vérifie le fonctionnement de pyautogui."""
    print("\n[1/4] Test de pyautogui...")
    try:
        x, y = pyautogui.position()
        screen_size = pyautogui.size()
        print(f"  [OK] pyautogui fonctionne")
        print(f"    Position souris: ({x}, {y})")
        print(f"    Taille écran: {screen_size.width}x{screen_size.height}")
        return True
    except Exception as e:
        print(f"  [X] Erreur: {e}")
        return False


def check_pywinauto():
    """Vérifie le fonctionnement de pywinauto."""
    print("\n[2/4] Test de pywinauto...")
    if not PYWINAUTO_AVAILABLE:
        print("  [!] pywinauto non installé (optionnel)")
        return True

    try:
        desktop = Desktop(backend="uia")
        windows = desktop.windows()
        print(f"  [OK] pywinauto fonctionne")
        print(f"    Fenêtres ouvertes: {len(windows)}")
        return True
    except Exception as e:
        print(f"  [X] Erreur: {e}")
        return False


def check_nws_executable():
    """Vérifie l'existence de l'exécutable NWS."""
    print("\n[3/4] Vérification de l'exécutable NWS...")
    exe_path = Path(NWS_EXE_PATH)

    if exe_path.exists():
        print(f"  [OK] Exécutable trouvé: {exe_path}")
        return True
    else:
        print(f"  [X] Exécutable non trouvé: {exe_path}")
        print("    Modifiez NWS_EXE_PATH dans config.py")
        return False


def check_nws_window():
    """Vérifie si une fenêtre NWS est ouverte."""
    print("\n[4/4] Recherche de la fenêtre NWS...")

    # Recherche avec pyautogui
    windows = pyautogui.getWindowsWithTitle(NWS_WINDOW_TITLE)

    if windows:
        print(f"  [OK] Fenêtre trouvée: {windows[0].title}")
        return True

    # Recherche avec pywinauto
    if PYWINAUTO_AVAILABLE:
        try:
            app = Application(backend="uia").connect(
                title_re=f".*{NWS_WINDOW_TITLE}.*",
                timeout=2
            )
            print(f"  [OK] Fenêtre trouvée via pywinauto")
            return True
        except Exception:
            pass

    print(f"  [!] Fenêtre non trouvée (NWS n'est peut-être pas lancé)")
    print(f"    Recherche: '{NWS_WINDOW_TITLE}'")
    return False


def list_all_windows():
    """Liste toutes les fenêtres ouvertes."""
    print("\n" + "=" * 50)
    print("FENÊTRES OUVERTES:")
    print("=" * 50)

    windows = pyautogui.getAllWindows()
    for w in windows:
        if w.title:  # Ignorer les fenêtres sans titre
            print(f"  - {w.title}")


def main():
    """Point d'entrée principal."""
    print("=" * 50)
    print("  TEST DE CONNEXION - New World Scheduler")
    print("=" * 50)

    results = {
        "pyautogui": check_pyautogui(),
        "pywinauto": check_pywinauto(),
        "executable": check_nws_executable(),
        "window": check_nws_window(),
    }

    # Afficher le résumé
    print("\n" + "=" * 50)
    print("RÉSUMÉ:")
    print("=" * 50)

    all_ok = True
    for name, ok in results.items():
        status = "[OK] OK" if ok else "[X] ERREUR"
        print(f"  {name}: {status}")
        if not ok and name not in ["window"]:  # window n'est pas critique
            all_ok = False

    if all_ok:
        print("\n[OK] Tous les prérequis sont satisfaits!")
        if not results["window"]:
            print("  (Lancez New World Scheduler avant l'automatisation)")
    else:
        print("\n[X] Certains prérequis ne sont pas satisfaits.")
        print("  Corrigez les erreurs avant de lancer l'automatisation.")

    # Optionnel: lister les fenêtres
    print("\nVoulez-vous voir la liste des fenêtres ouvertes? (o/n): ", end="")
    try:
        response = input().strip().lower()
        if response in ["o", "oui", "y", "yes"]:
            list_all_windows()
    except (EOFError, KeyboardInterrupt):
        pass

    print()


if __name__ == "__main__":
    main()

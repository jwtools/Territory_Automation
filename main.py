#!/usr/bin/env python3
"""
Script principal pour l'automatisation de New World Scheduler 7.9.

Usage:
    python main.py                  # Lancer l'automatisation
    python main.py --dry-run        # Mode simulation (sans exécuter)
    python main.py --no-save        # Remplir les champs sans sauvegarder (validation)
    python main.py --reset          # Réinitialiser la progression
    python main.py --verify         # Vérifier les fichiers sans exécuter
"""

import argparse
import sys
from pathlib import Path

# Ajouter le dossier racine au path
sys.path.insert(0, str(Path(__file__).parent))

from config import (
    NWS_EXE_PATH,
    DATA_FILE_PATH,
    PDF_FOLDER_PATH,
    LOG_FOLDER_PATH,
    PROGRESS_FILE_PATH,
    NWS_WINDOW_TITLE,
    COORDINATES,
    EXCEL_COLUMNS,
    DELAY_AFTER_CLICK,
    DELAY_AFTER_TYPE,
    DELAY_APP_LAUNCH,
    DELAY_AFTER_SAVE,
    DELAY_BETWEEN_TERRITORIES,
    MAX_RETRIES,
    STARTUP_DIALOG_TITLES,
    STARTUP_DIALOG_CLOSE_METHOD,
    STARTUP_DIALOG_CLOSE_BUTTON,
    STARTUP_DIALOG_WAIT,
)

from territory_automation.logger_setup import setup_logger
from territory_automation.data_loader import DataLoader, ProgressTracker
from territory_automation.automation import NWSAutomator, AutomationError


def parse_args():
    """Parse les arguments de la ligne de commande."""
    parser = argparse.ArgumentParser(
        description="Automatisation de l'import de territoires dans New World Scheduler"
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Mode simulation - affiche les actions sans les exécuter"
    )
    parser.add_argument(
        "--no-save",
        action="store_true",
        help="Mode validation - remplit les champs mais ne sauvegarde pas (attend confirmation)"
    )
    parser.add_argument(
        "--reset",
        action="store_true",
        help="Réinitialise la progression (recommence depuis le début)"
    )
    parser.add_argument(
        "--verify",
        action="store_true",
        help="Vérifie les données et les PDFs sans exécuter l'automatisation"
    )
    parser.add_argument(
        "--data-file",
        type=Path,
        default=DATA_FILE_PATH,
        help=f"Chemin vers le fichier Excel/CSV (défaut: {DATA_FILE_PATH})"
    )
    parser.add_argument(
        "--pdf-folder",
        type=Path,
        default=PDF_FOLDER_PATH,
        help=f"Dossier contenant les PDFs (défaut: {PDF_FOLDER_PATH})"
    )
    parser.add_argument(
        "--start-from",
        type=int,
        default=0,
        help="Index du territoire à partir duquel commencer (0 = premier)"
    )
    return parser.parse_args()


def verify_prerequisites(logger, data_file: Path, pdf_folder: Path) -> bool:
    """
    Vérifie que tous les prérequis sont en place.

    Returns:
        True si tout est OK, False sinon
    """
    errors = []

    # Vérifier le fichier de données
    if not data_file.exists():
        errors.append(f"Fichier de données non trouvé: {data_file}")

    # Vérifier le dossier des PDFs
    if not pdf_folder.exists():
        errors.append(f"Dossier des PDFs non trouvé: {pdf_folder}")

    # Vérifier l'exécutable NWS
    if not Path(NWS_EXE_PATH).exists():
        logger.warning(f"Exécutable NWS non trouvé: {NWS_EXE_PATH}")
        logger.warning("Assurez-vous que le chemin dans config.py est correct")

    if errors:
        for error in errors:
            logger.error(error)
        return False

    return True


def verify_data(logger, loader: DataLoader, pdf_folder: Path) -> dict:
    """
    Vérifie les données et les fichiers PDF associés.

    Returns:
        Dictionnaire avec les statistiques de vérification
    """
    territories = loader.get_all_territories()
    stats = {
        "total": len(territories),
        "with_pdf": 0,
        "without_pdf": 0,
        "missing_pdfs": []
    }

    for territory in territories:
        territory_id = territory.get("numero", "INCONNU")
        pdf_filename = territory.get("pdf_filename", "") or f"{territory_id}.pdf"
        pdf_path = pdf_folder / pdf_filename

        if pdf_path.exists():
            stats["with_pdf"] += 1
        else:
            stats["without_pdf"] += 1
            stats["missing_pdfs"].append(pdf_filename)

    logger.info(f"=== Vérification des données ===")
    logger.info(f"Total territoires: {stats['total']}")
    logger.info(f"Avec PDF: {stats['with_pdf']}")
    logger.info(f"Sans PDF: {stats['without_pdf']}")

    if stats["missing_pdfs"]:
        logger.warning(f"PDFs manquants ({len(stats['missing_pdfs'])}):")
        for pdf in stats["missing_pdfs"][:10]:  # Afficher les 10 premiers
            logger.warning(f"  - {pdf}")
        if len(stats["missing_pdfs"]) > 10:
            logger.warning(f"  ... et {len(stats['missing_pdfs']) - 10} autres")

    return stats


def run_automation(
    logger,
    loader: DataLoader,
    tracker: ProgressTracker,
    automator: NWSAutomator,
    dry_run: bool = False,
    no_save: bool = False,
    start_from: int = 0
):
    """
    Exécute l'automatisation pour tous les territoires.

    Args:
        logger: Logger
        loader: Chargeur de données
        tracker: Tracker de progression
        automator: Automatiseur NWS
        dry_run: Mode simulation
        no_save: Mode validation (remplit sans sauvegarder)
        start_from: Index de départ
    """
    territories = loader.get_all_territories()
    total = len(territories)

    logger.info(f"=== Démarrage de l'automatisation ===")
    logger.info(f"Territoires à traiter: {total}")
    logger.info(f"Mode dry-run: {dry_run}")
    if no_save:
        logger.info("MODE VALIDATION: Les champs seront remplis mais NON sauvegardés")
        logger.info("Appuyez sur Entrée après chaque territoire pour continuer, ou Ctrl+C pour arrêter")

    if not dry_run:
        # Lancer l'application
        if not automator.launch_application():
            logger.error("Impossible de lancer New World Scheduler")
            return

    processed = 0
    failed = 0

    for i, territory in enumerate(territories):
        if i < start_from:
            continue

        territory_id = territory.get("numero", f"INDEX_{i}")

        # Vérifier si déjà traité
        if tracker.is_processed(territory_id):
            logger.info(f"[{i+1}/{total}] {territory_id} - Déjà traité, ignoré")
            continue

        logger.info(f"[{i+1}/{total}] Traitement de: {territory_id}")

        if dry_run:
            # Mode simulation
            pdf_exists, pdf_path = automator.verify_pdf_exists(territory)
            logger.info(f"  -> Numéro: {territory.get('numero', '')}")
            logger.info(f"  -> Suffixe: {territory.get('suffixe', '')}")
            logger.info(f"  -> Type: {territory.get('type', '')}")
            logger.info(f"  -> PDF: {pdf_path.name} ({'OK' if pdf_exists else 'MANQUANT'})")
            processed += 1
            continue

        # Exécution réelle
        success = False
        for attempt in range(MAX_RETRIES):
            try:
                if automator.process_territory(territory, no_save=no_save):
                    if no_save:
                        # Mode validation: attendre confirmation utilisateur
                        logger.info(f"  -> Territoire {territory_id} rempli (NON sauvegardé)")
                        logger.info("  -> Vérifiez les champs dans NWS, puis appuyez sur Entrée...")
                        try:
                            input()
                        except EOFError:
                            pass
                    else:
                        tracker.mark_processed(territory_id)
                    processed += 1
                    success = True
                    break
            except AutomationError as e:
                logger.warning(f"Tentative {attempt+1}/{MAX_RETRIES} échouée: {e}")
                if attempt < MAX_RETRIES - 1:
                    logger.info("Nouvelle tentative...")
                    continue

        if not success:
            tracker.mark_failed(territory_id, "Échec après plusieurs tentatives")
            failed += 1

    # Résumé final
    logger.info(f"=== Automatisation terminée ===")
    logger.info(f"Traités avec succès: {processed}")
    logger.info(f"Échecs: {failed}")

    summary = tracker.get_summary()
    if summary["failed_territories"]:
        logger.warning("Territoires en échec:")
        for fail in summary["failed_territories"]:
            logger.warning(f"  - {fail['id']}: {fail['error']}")


def main():
    """Point d'entrée principal."""
    args = parse_args()

    # Initialiser le logging
    logger = setup_logger(LOG_FOLDER_PATH)
    logger.info("=== Territory Automation pour New World Scheduler ===")

    # Vérifier les prérequis
    if not verify_prerequisites(logger, args.data_file, args.pdf_folder):
        logger.error("Prérequis non satisfaits. Arrêt.")
        sys.exit(1)

    # Charger les données
    try:
        loader = DataLoader(args.data_file, EXCEL_COLUMNS)
        loader.load()
    except Exception as e:
        logger.error(f"Erreur lors du chargement des données: {e}")
        sys.exit(1)

    # Initialiser le tracker de progression
    tracker = ProgressTracker(PROGRESS_FILE_PATH)

    if args.reset:
        logger.info("Réinitialisation de la progression...")
        tracker.reset()

    # Mode vérification uniquement
    if args.verify:
        verify_data(logger, loader, args.pdf_folder)
        sys.exit(0)

    # Préparer les délais
    delays = {
        "after_click": DELAY_AFTER_CLICK,
        "after_type": DELAY_AFTER_TYPE,
        "app_launch": DELAY_APP_LAUNCH,
        "after_save": DELAY_AFTER_SAVE,
        "between_territories": DELAY_BETWEEN_TERRITORIES,
    }

    # Configuration des dialogues de démarrage
    startup_dialog_config = {
        "titles": STARTUP_DIALOG_TITLES,
        "close_method": STARTUP_DIALOG_CLOSE_METHOD,
        "close_button": STARTUP_DIALOG_CLOSE_BUTTON,
        "wait_time": STARTUP_DIALOG_WAIT,
    }

    # Créer l'automatiseur
    automator = NWSAutomator(
        exe_path=NWS_EXE_PATH,
        window_title=NWS_WINDOW_TITLE,
        coordinates=COORDINATES,
        delays=delays,
        pdf_folder=args.pdf_folder,
        startup_dialog_config=startup_dialog_config
    )

    # Lancer l'automatisation
    try:
        run_automation(
            logger=logger,
            loader=loader,
            tracker=tracker,
            automator=automator,
            dry_run=args.dry_run,
            no_save=args.no_save,
            start_from=args.start_from
        )
    except KeyboardInterrupt:
        logger.warning("Interruption par l'utilisateur (Ctrl+C)")
        logger.info("La progression a été sauvegardée. Relancez pour continuer.")
    except Exception as e:
        logger.exception(f"Erreur inattendue: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()

"""
Configuration du système de logging.
"""

import logging
from datetime import datetime
from pathlib import Path
import sys


def setup_logger(log_folder: Path, name: str = "territory_automation") -> logging.Logger:
    """
    Configure et retourne un logger avec sortie fichier et console.

    Args:
        log_folder: Dossier où stocker les fichiers de log
        name: Nom du logger

    Returns:
        Logger configuré
    """
    # Créer le dossier de logs s'il n'existe pas
    log_folder.mkdir(parents=True, exist_ok=True)

    # Nom du fichier avec timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = log_folder / f"automation_{timestamp}.log"

    # Créer le logger
    logger = logging.getLogger(name)
    logger.setLevel(logging.DEBUG)

    # Éviter les handlers en double
    if logger.handlers:
        logger.handlers.clear()

    # Format des messages
    formatter = logging.Formatter(
        "%(asctime)s | %(levelname)-8s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S"
    )

    # Handler fichier (DEBUG et plus)
    file_handler = logging.FileHandler(log_file, encoding="utf-8")
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    # Handler console (INFO et plus)
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

    logger.info(f"Logging initialisé - Fichier: {log_file}")

    return logger


def get_logger(name: str = "territory_automation") -> logging.Logger:
    """Récupère un logger existant."""
    return logging.getLogger(name)

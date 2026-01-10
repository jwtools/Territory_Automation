"""
Module de chargement des données depuis Excel/CSV.
"""

import pandas as pd
from pathlib import Path
from typing import Optional
import json

from .logger_setup import get_logger


class DataLoader:
    """Charge et valide les données des territoires depuis Excel ou CSV."""

    def __init__(self, file_path: Path, column_mapping: dict):
        """
        Initialise le loader.

        Args:
            file_path: Chemin vers le fichier Excel ou CSV
            column_mapping: Mapping des colonnes (clé interne -> nom colonne Excel)
        """
        self.file_path = Path(file_path)
        self.column_mapping = column_mapping
        self.logger = get_logger()
        self.data: Optional[pd.DataFrame] = None

    def load(self) -> pd.DataFrame:
        """
        Charge les données depuis le fichier.

        Returns:
            DataFrame avec les données des territoires

        Raises:
            FileNotFoundError: Si le fichier n'existe pas
            ValueError: Si le format n'est pas supporté
        """
        if not self.file_path.exists():
            raise FileNotFoundError(f"Fichier non trouvé: {self.file_path}")

        suffix = self.file_path.suffix.lower()

        if suffix in [".xlsx", ".xls"]:
            self.logger.info(f"Chargement du fichier Excel: {self.file_path}")
            self.data = pd.read_excel(self.file_path)
        elif suffix == ".csv":
            self.logger.info(f"Chargement du fichier CSV: {self.file_path}")
            self.data = pd.read_csv(self.file_path, encoding="utf-8-sig")
        else:
            raise ValueError(f"Format non supporté: {suffix}. Utilisez .xlsx, .xls ou .csv")

        self._validate_columns()
        self._clean_data()

        self.logger.info(f"Données chargées: {len(self.data)} territoires")
        return self.data

    def _validate_columns(self):
        """Vérifie que les colonnes requises sont présentes."""
        required = ["numero"]  # Seul le numéro est obligatoire

        for key in required:
            col_name = self.column_mapping.get(key)
            if col_name and col_name not in self.data.columns:
                raise ValueError(
                    f"Colonne requise manquante: '{col_name}'. "
                    f"Colonnes disponibles: {list(self.data.columns)}"
                )

    def _clean_data(self):
        """Nettoie les données (remplace NaN par chaînes vides)."""
        self.data = self.data.fillna("")

        # Convertir les colonnes numériques en string si nécessaire
        for col in self.data.columns:
            if self.data[col].dtype in ["int64", "float64"]:
                # Garder les entiers sans décimales
                self.data[col] = self.data[col].apply(
                    lambda x: str(int(x)) if x != "" and pd.notna(x) and x == int(x) else str(x) if x != "" else ""
                )

    def get_territory(self, index: int) -> dict:
        """
        Récupère les données d'un territoire par son index.

        Args:
            index: Index de la ligne dans le DataFrame

        Returns:
            Dictionnaire avec les données du territoire
        """
        if self.data is None:
            raise RuntimeError("Données non chargées. Appelez load() d'abord.")

        row = self.data.iloc[index]

        territory = {}
        for key, col_name in self.column_mapping.items():
            if col_name in self.data.columns:
                value = row[col_name]
                territory[key] = str(value) if value != "" else ""
            else:
                territory[key] = ""

        return territory

    def get_all_territories(self) -> list[dict]:
        """
        Récupère tous les territoires.

        Returns:
            Liste de dictionnaires avec les données
        """
        if self.data is None:
            raise RuntimeError("Données non chargées. Appelez load() d'abord.")

        return [self.get_territory(i) for i in range(len(self.data))]

    def __len__(self) -> int:
        """Retourne le nombre de territoires."""
        return len(self.data) if self.data is not None else 0


class ProgressTracker:
    """Suit la progression et permet de reprendre après interruption."""

    def __init__(self, progress_file: Path):
        """
        Initialise le tracker.

        Args:
            progress_file: Chemin vers le fichier JSON de progression
        """
        self.progress_file = Path(progress_file)
        self.logger = get_logger()
        self.processed: list[str] = []
        self.failed: list[dict] = []
        self._load()

    def _load(self):
        """Charge la progression depuis le fichier."""
        if self.progress_file.exists():
            try:
                with open(self.progress_file, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    self.processed = data.get("processed", [])
                    self.failed = data.get("failed", [])
                    self.logger.info(
                        f"Progression chargée: {len(self.processed)} traités, "
                        f"{len(self.failed)} en erreur"
                    )
            except (json.JSONDecodeError, IOError) as e:
                self.logger.warning(f"Impossible de charger la progression: {e}")

    def save(self):
        """Sauvegarde la progression dans le fichier."""
        self.progress_file.parent.mkdir(parents=True, exist_ok=True)

        with open(self.progress_file, "w", encoding="utf-8") as f:
            json.dump({
                "processed": self.processed,
                "failed": self.failed
            }, f, indent=2, ensure_ascii=False)

    def mark_processed(self, territory_id: str):
        """Marque un territoire comme traité."""
        if territory_id not in self.processed:
            self.processed.append(territory_id)
            self.save()

    def mark_failed(self, territory_id: str, error: str):
        """Marque un territoire comme en erreur."""
        self.failed.append({
            "id": territory_id,
            "error": error
        })
        self.save()

    def is_processed(self, territory_id: str) -> bool:
        """Vérifie si un territoire a déjà été traité."""
        return territory_id in self.processed

    def reset(self):
        """Réinitialise la progression."""
        self.processed = []
        self.failed = []
        if self.progress_file.exists():
            self.progress_file.unlink()
        self.logger.info("Progression réinitialisée")

    def get_summary(self) -> dict:
        """Retourne un résumé de la progression."""
        return {
            "processed_count": len(self.processed),
            "failed_count": len(self.failed),
            "failed_territories": self.failed
        }

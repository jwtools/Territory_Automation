"""
Module principal d'automatisation de New World Scheduler.
Utilise pywinauto pour le contrôle de l'application Windows.
"""

import time
import subprocess
from pathlib import Path
from typing import Optional

import pyautogui
import pyperclip

try:
    from pywinauto import Application, Desktop
    from pywinauto.findwindows import ElementNotFoundError
    from pywinauto.timings import TimeoutError as PywinautoTimeout
    PYWINAUTO_AVAILABLE = True
except ImportError:
    PYWINAUTO_AVAILABLE = False

from .logger_setup import get_logger


# Configuration de pyautogui
pyautogui.FAILSAFE = True  # Coin supérieur gauche = arrêt d'urgence
pyautogui.PAUSE = 0.1  # Pause entre chaque action


class AutomationError(Exception):
    """Exception pour les erreurs d'automatisation."""
    pass


class NWSAutomator:
    """
    Automatise les interactions avec New World Scheduler 7.9.

    Utilise une combinaison de pywinauto (pour la gestion des fenêtres)
    et pyautogui (pour les clics et la saisie).
    """

    def __init__(
        self,
        exe_path: str,
        window_title: str,
        coordinates: dict,
        delays: dict,
        pdf_folder: Path,
        startup_dialog_config: Optional[dict] = None
    ):
        """
        Initialise l'automatiseur.

        Args:
            exe_path: Chemin vers l'exécutable NWS
            window_title: Titre de la fenêtre à rechercher
            coordinates: Dictionnaire des coordonnées des éléments
            delays: Dictionnaire des délais
            pdf_folder: Dossier contenant les PDFs
            startup_dialog_config: Configuration pour gérer les dialogues de démarrage
        """
        self.exe_path = exe_path
        self.window_title = window_title
        self.coords = coordinates
        self.delays = delays
        self.pdf_folder = Path(pdf_folder)
        self.logger = get_logger()
        self.app: Optional[Application] = None
        self.main_window = None

        # Configuration des dialogues de démarrage
        self.startup_config = startup_dialog_config or {
            "titles": [],
            "close_method": "escape",
            "close_button": (960, 600),
            "wait_time": 2.0
        }

    def launch_application(self) -> bool:
        """
        Lance New World Scheduler s'il n'est pas déjà ouvert.

        Returns:
            True si l'application est prête, False sinon
        """
        try:
            # Essayer de se connecter à une instance existante (plusieurs tentatives)
            for attempt in range(3):
                if self._connect_to_existing():
                    self.logger.info("Connecté à une instance existante de NWS")
                    self._dismiss_startup_dialogs()
                    self._navigate_to_territory_screen()
                    return True
                if attempt < 2:
                    time.sleep(1)

            # Lancer une nouvelle instance
            self.logger.info(f"Lancement de NWS: {self.exe_path}")

            if not Path(self.exe_path).exists():
                raise AutomationError(f"Exécutable non trouvé: {self.exe_path}")

            subprocess.Popen(self.exe_path)

            # Attendre que l'application démarre (avec plusieurs tentatives)
            launch_delay = self.delays.get("app_launch", 10.0)
            max_attempts = 5
            delay_per_attempt = launch_delay / max_attempts

            for attempt in range(max_attempts):
                self.logger.debug(f"Tentative de connexion {attempt + 1}/{max_attempts}...")
                time.sleep(delay_per_attempt)

                if self._connect_to_existing():
                    self.logger.info("NWS lancé avec succès")
                    self._dismiss_startup_dialogs()
                    self._navigate_to_territory_screen()
                    return True

            raise AutomationError("Impossible de se connecter à NWS après lancement")

        except Exception as e:
            self.logger.error(f"Erreur lors du lancement: {e}")
            return False

    def _dismiss_startup_dialogs(self):
        """
        Ferme les dialogues de démarrage (astuces, conseils, etc.).
        Essaie plusieurs méthodes pour s'assurer que les dialogues sont fermés.
        """
        dialog_titles = self.startup_config.get("titles", [])
        if not dialog_titles:
            self.logger.debug("Aucun titre de dialogue configuré, tentative Escape simple")
            # Essayer quand même quelques Escape au cas où
            for _ in range(2):
                pyautogui.press("escape")
                time.sleep(0.2)
            return

        self.logger.debug("Recherche de dialogues de démarrage...")
        time.sleep(self.startup_config.get("wait_time", 2.0))

        # Méthode 1: Recherche par titre avec pywinauto
        if PYWINAUTO_AVAILABLE and self.app:
            for title in dialog_titles:
                try:
                    dialog = self.app.window(title_re=f".*{title}.*", timeout=1)
                    if dialog.exists():
                        self.logger.info(f"Dialogue détecté: {title}")
                        self._close_dialog(dialog)
                        time.sleep(0.5)
                except Exception:
                    pass

        # Méthode 2: Recherche avec pyautogui
        for title in dialog_titles:
            windows = pyautogui.getWindowsWithTitle(title)
            if windows:
                self.logger.info(f"Dialogue détecté (pyautogui): {title}")
                windows[0].activate()
                time.sleep(0.2)
                self._close_dialog_fallback()
                time.sleep(0.5)

        # Méthode 3: Essayer de fermer tout dialogue modal avec Escape
        self.logger.debug("Tentative de fermeture par Escape...")
        for _ in range(3):
            pyautogui.press("escape")
            time.sleep(0.3)

        # Réactiver la fenêtre principale
        self.activate_window()

    def _close_dialog(self, dialog):
        """Ferme un dialogue pywinauto."""
        method = self.startup_config.get("close_method", "escape")

        try:
            if method == "escape":
                dialog.type_keys("{ESC}")
            elif method == "enter":
                dialog.type_keys("{ENTER}")
            elif method == "click_close":
                # Chercher un bouton Fermer/Close/X
                try:
                    close_btn = dialog.child_window(title_re=".*(Fermer|Close|OK|X).*", control_type="Button")
                    if close_btn.exists():
                        close_btn.click()
                        return
                except Exception:
                    pass
                # Fallback: cliquer aux coordonnées configurées
                x, y = self.startup_config.get("close_button", (960, 600))
                pyautogui.click(x, y)
            elif method == "click_ok":
                try:
                    ok_btn = dialog.child_window(title_re=".*(OK|Fermer|Close).*", control_type="Button")
                    if ok_btn.exists():
                        ok_btn.click()
                        return
                except Exception:
                    pass
                pyautogui.press("enter")
            else:
                pyautogui.press("escape")

            self.logger.debug(f"Dialogue fermé avec méthode: {method}")

        except Exception as e:
            self.logger.debug(f"Erreur fermeture dialogue: {e}")
            self._close_dialog_fallback()

    def _close_dialog_fallback(self):
        """Méthode de secours pour fermer un dialogue."""
        method = self.startup_config.get("close_method", "escape")

        if method in ["escape", "click_close"]:
            pyautogui.press("escape")
        elif method in ["enter", "click_ok"]:
            pyautogui.press("enter")
        else:
            pyautogui.press("escape")

    def _navigate_to_territory_screen(self):
        """
        Navigue vers l'écran de saisie des territoires.

        Étapes:
        1. Clic sur le bouton/menu "Territoires"
        2. Clic sur "Liste des territoires"
        """
        self.logger.info("Navigation vers l'écran des territoires...")

        # Étape 1: Cliquer sur "Territoires"
        if "btn_menu_territoires" in self.coords:
            self.logger.debug("Clic sur menu Territoires")
            self.click("btn_menu_territoires")
            time.sleep(0.5)
        else:
            self.logger.warning("Coordonnées 'btn_menu_territoires' non définies")

        # Étape 2: Cliquer sur "Liste des territoires"
        if "btn_liste_territoires" in self.coords:
            self.logger.debug("Clic sur Liste des territoires")
            self.click("btn_liste_territoires")
            time.sleep(0.5)
        else:
            self.logger.warning("Coordonnées 'btn_liste_territoires' non définies")

        self.logger.info("Navigation terminée - écran des territoires")

    def _connect_to_existing(self) -> bool:
        """Tente de se connecter à une instance existante."""
        if not PYWINAUTO_AVAILABLE:
            self.logger.warning("pywinauto non disponible, utilisation de pyautogui seul")
            return self._activate_window_pyautogui()

        try:
            self.app = Application(backend="uia").connect(
                title_re=f".*{self.window_title}.*",
                timeout=2
            )
            self.main_window = self.app.window(title_re=f".*{self.window_title}.*")
            self.main_window.set_focus()
            return True
        except (ElementNotFoundError, PywinautoTimeout):
            return False

    def _activate_window_pyautogui(self) -> bool:
        """Active la fenêtre en utilisant pyautogui."""
        windows = pyautogui.getWindowsWithTitle(self.window_title)
        if windows:
            windows[0].activate()
            time.sleep(0.5)
            return True
        return False

    def activate_window(self):
        """Active et met au premier plan la fenêtre NWS."""
        if self.main_window:
            try:
                self.main_window.set_focus()
                time.sleep(0.2)
            except Exception:
                self._activate_window_pyautogui()
        else:
            self._activate_window_pyautogui()

    def click(self, element_name: str, double: bool = False):
        """
        Clique sur un élément par son nom.

        Args:
            element_name: Nom de l'élément dans le dictionnaire de coordonnées
            double: Si True, effectue un double-clic
        """
        if element_name not in self.coords:
            raise AutomationError(f"Élément inconnu: {element_name}")

        x, y = self.coords[element_name]
        action = "Double-clic" if double else "Clic"
        self.logger.info(f"  → {action} sur [{element_name}] à ({x}, {y})")

        self.activate_window()

        if double:
            pyautogui.doubleClick(x, y)
        else:
            pyautogui.click(x, y)

        time.sleep(self.delays.get("after_click", 0.3))

    def type_text(self, text: str, clear_first: bool = True):
        """
        Saisit du texte dans le champ actif.

        Args:
            text: Texte à saisir
            clear_first: Si True, efface le champ avant la saisie
        """
        if not text:
            return

        preview = text[:30] + "..." if len(text) > 30 else text
        self.logger.info(f"    Saisie: \"{preview}\"")

        if clear_first:
            pyautogui.hotkey("ctrl", "a")
            time.sleep(0.05)

        # Utiliser le presse-papiers pour les caractères spéciaux
        pyperclip.copy(text)
        pyautogui.hotkey("ctrl", "v")

        time.sleep(self.delays.get("after_type", 0.1))

    def select_dropdown_option(self, dropdown_name: str, option_name: str):
        """
        Sélectionne une option dans un dropdown par clic direct.

        Args:
            dropdown_name: Nom du dropdown
            option_name: Nom de l'option à sélectionner
        """
        self.logger.info(f"  Dropdown [{dropdown_name}] → [{option_name}]")
        # Cliquer sur le dropdown pour l'ouvrir
        self.click(dropdown_name)
        time.sleep(0.3)

        # Cliquer sur l'option
        self.click(option_name)

    def select_dropdown_by_typing(self, dropdown_name: str, value: str):
        """
        Sélectionne une option dans un dropdown en tapant le texte.

        Args:
            dropdown_name: Nom du dropdown
            value: Texte à taper pour sélectionner l'option
        """
        if not value:
            return

        # Cliquer sur le dropdown pour l'ouvrir
        self.click(dropdown_name)
        time.sleep(0.3)

        # Taper le texte pour filtrer/sélectionner
        pyperclip.copy(value)
        pyautogui.hotkey("ctrl", "v")
        time.sleep(0.2)

        # Appuyer sur Entrée pour valider la sélection
        pyautogui.press("enter")
        time.sleep(self.delays.get("after_click", 0.3))

        self.logger.debug(f"Dropdown {dropdown_name} sélectionné: {value}")

    def fill_field(self, field_name: str, value: str):
        """
        Remplit un champ de formulaire.

        Args:
            field_name: Nom du champ
            value: Valeur à saisir
        """
        if not value:
            self.logger.info(f"  Champ [{field_name}] ignoré (vide)")
            return

        self.click(field_name)
        self.type_text(value)

    def import_pdf(self, pdf_path: Path) -> bool:
        """
        Importe un fichier PDF/image via la boîte de dialogue Windows.

        Args:
            pdf_path: Chemin vers le fichier

        Returns:
            True si l'import réussit, False sinon
        """
        if not pdf_path.exists():
            self.logger.warning(f"Fichier non trouvé: {pdf_path}")
            return False

        try:
            # Cliquer sur le bouton "Ajouter fichier"
            self.click("btn_import_pdf")

            # Attendre que la boîte de dialogue s'ouvre
            time.sleep(1.5)

            # Dans la boîte de dialogue Windows, le champ "Nom du fichier" a le focus
            # On utilise Ctrl+A pour tout sélectionner puis on colle le chemin
            pyautogui.hotkey("ctrl", "a")
            time.sleep(0.1)

            # Copier le chemin absolu et coller
            absolute_path = str(pdf_path.resolve())
            pyperclip.copy(absolute_path)
            pyautogui.hotkey("ctrl", "v")
            time.sleep(0.3)

            self.logger.debug(f"Chemin collé: {absolute_path}")

            # Appuyer sur Entrée pour valider
            pyautogui.press("enter")
            time.sleep(self.delays.get("after_save", 1.0))

            self.logger.info(f"Fichier importé: {pdf_path.name}")
            return True

        except Exception as e:
            self.logger.error(f"Erreur lors de l'import: {e}")
            return False

    def create_new_territory(self):
        """Clique sur le bouton Nouveau Territoire."""
        self.click("btn_new_territory")
        self.logger.debug("Nouveau territoire créé")

    def process_territory(self, territory: dict, no_save: bool = False) -> bool:
        """
        Traite un territoire complet.

        Args:
            territory: Dictionnaire avec les données du territoire
            no_save: Si True, remplit les champs mais ne sauvegarde pas

        Returns:
            True si le traitement réussit, False sinon
        """
        territory_id = territory.get("numero", "INCONNU")
        self.logger.info(f"")
        self.logger.info(f"{'='*50}")
        self.logger.info(f"TERRITOIRE: {territory_id}")
        self.logger.info(f"{'='*50}")

        try:
            # Activer la fenêtre
            self.activate_window()

            # Créer un nouveau territoire
            self.logger.info("[ÉTAPE 1] Nouveau territoire")
            self.create_new_territory()

            # Remplir les champs (ordre de saisie)
            self.logger.info("[ÉTAPE 2] Catégorie → SAR")
            if "dropdown_categorie" in self.coords and "dropdown_option_sar" in self.coords:
                self.select_dropdown_option("dropdown_categorie", "dropdown_option_sar")

            self.logger.info("[ÉTAPE 3] Numéro")
            self.fill_field("field_numero", territory.get("numero", ""))

            self.logger.info("[ÉTAPE 4] Suffixe")
            self.fill_field("field_suffixe", territory.get("suffixe", ""))

            self.logger.info("[ÉTAPE 5] Type")
            type_value = territory.get("type", "").lower().strip()
            if type_value:
                type_options = {
                    "presentiel": "dropdown_option_presentiel",
                    "en présentiel": "dropdown_option_presentiel",
                    "courrier": "dropdown_option_courrier",
                    "telephone": "dropdown_option_telephone",
                    "téléphone": "dropdown_option_telephone",
                    "entreprise": "dropdown_option_entreprise",
                }
                # Types qui nécessitent confirmation (modal "Êtes-vous sûr")
                types_need_confirm = ["courrier", "telephone", "téléphone", "entreprise"]

                option_id = type_options.get(type_value)
                if option_id:
                    self.select_dropdown_option("dropdown_type", option_id)

                    # Si type autre que "En présentiel", confirmer le modal
                    if type_value in types_need_confirm and "btn_confirm_type" in self.coords:
                        self.logger.info("  → Confirmation modal 'Oui'")
                        time.sleep(0.3)
                        self.click("btn_confirm_type")
                else:
                    self.logger.warning(f"  Type inconnu: {type_value}")
            else:
                self.logger.info("  (pas de type)")

            self.logger.info("[ÉTAPE 6] Ville")
            ville = territory.get("ville", "").upper().strip()
            if ville and "dropdown_ville" in self.coords:
                ville_options = {
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
                option_id = ville_options.get(ville)
                if option_id:
                    self.select_dropdown_option("dropdown_ville", option_id)
                else:
                    self.logger.warning(f"  Ville inconnue: {ville}")
            else:
                self.logger.info("  (pas de ville)")

            self.logger.info("[ÉTAPE 7] Champs texte")
            self.fill_field("field_lien_gps", territory.get("lien_gps", ""))
            self.fill_field("field_notes", territory.get("notes", ""))
            self.fill_field("field_ne_pas_visiter", territory.get("ne_pas_visiter", ""))
            self.fill_field("field_notes_proclamateur", territory.get("notes_proclamateur", ""))

            self.logger.info("[ÉTAPE 8] Onglet Carte")
            if "btn_carte" in self.coords:
                self.click("btn_carte")
                time.sleep(0.5)

            # Importer le fichier (PDF/image)
            self.logger.info("[ÉTAPE 9] Import fichier")
            if not no_save:
                pdf_filename = territory.get("pdf_filename", "")
                if not pdf_filename:
                    # Construire le nom du fichier à partir du numéro
                    pdf_filename = f"{territory_id}.pdf"

                pdf_path = self.pdf_folder / pdf_filename
                self.logger.info(f"  Recherche: {pdf_path}")
                if pdf_path.exists():
                    self.import_pdf(pdf_path)
                else:
                    self.logger.warning(f"  FICHIER NON TROUVÉ: {pdf_path}")

                self.logger.info(f"[OK] Territoire {territory_id} traité avec succès")
            else:
                self.logger.info(f"  (mode --no-save, import ignoré)")
                self.logger.info(f"[OK] Territoire {territory_id} rempli (validation)")

            time.sleep(self.delays.get("between_territories", 0.5))

            return True

        except Exception as e:
            self.logger.error(f"Erreur lors du traitement de {territory_id}: {e}")
            return False

    def get_pdf_path(self, territory: dict) -> Path:
        """
        Détermine le chemin du PDF pour un territoire.

        Args:
            territory: Données du territoire

        Returns:
            Path vers le fichier PDF
        """
        pdf_filename = territory.get("pdf_filename", "")
        if not pdf_filename:
            territory_id = territory.get("numero", "UNKNOWN")
            pdf_filename = f"{territory_id}.pdf"

        return self.pdf_folder / pdf_filename

    def verify_pdf_exists(self, territory: dict) -> tuple[bool, Path]:
        """
        Vérifie si le PDF existe pour un territoire.

        Returns:
            Tuple (existe, chemin)
        """
        pdf_path = self.get_pdf_path(territory)
        return pdf_path.exists(), pdf_path

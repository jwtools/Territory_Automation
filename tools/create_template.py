#!/usr/bin/env python3
"""
Crée un fichier Excel template avec des exemples de données.

Usage:
    python tools/create_template.py
"""

import sys
from pathlib import Path

try:
    import pandas as pd
except ImportError:
    print("ERREUR: pandas n'est pas installé.")
    print("Installez-le avec: pip install pandas openpyxl")
    sys.exit(1)

try:
    import openpyxl
except ImportError:
    print("ERREUR: openpyxl n'est pas installé.")
    print("Installez-le avec: pip install openpyxl")
    sys.exit(1)


def create_template():
    """Crée le fichier template Excel."""
    # Données d'exemple
    # Note: Catégorie est toujours "SAR" (sélectionné automatiquement)
    data = {
        "Numero": ["SAR-1-01", "SAR-1-02", "SAR-1-03", "SAR-2-01", "SAR-2-02"],
        "Suffixe": ["A", "B", "", "A", ""],
        "Type": ["En présentiel", "Courrier", "Téléphone", "Entreprise", "En présentiel"],
        "Ville": ["SARTROUVILLE", "SARTROUVILLE", "MAISONS-LAFFITTE", "MONTESSON", "MESNIL LE ROI"],
        "Lien_GPS": [
            "https://maps.google.com/?q=48.8566,2.3522",
            "https://maps.google.com/?q=48.8584,2.2945",
            "",
            "https://maps.google.com/?q=48.8606,2.3376",
            "",
        ],
        "Notes": [
            "Zone résidentielle",
            "Immeuble avec interphone",
            "Zone commerciale",
            "",
            "Accès difficile le week-end",
        ],
        "Ne_Pas_Visiter": [
            "",
            "Apt 3B - Ne pas déranger",
            "",
            "",
            "Maison bleue au coin",
        ],
        "Notes_Proclamateur": [
            "Prévoir 2h minimum",
            "",
            "Parking gratuit à proximité",
            "",
            "",
        ],
        "PDF_Filename": [
            "",  # Sera déduit de Numero: SAR-1-01.pdf
            "",  # Sera déduit de Numero: SAR-1-02.pdf
            "territoire_special.pdf",  # Nom personnalisé
            "",
            "",
        ],
    }

    df = pd.DataFrame(data)

    # Chemin de sortie
    output_dir = Path(__file__).parent.parent / "data"
    output_dir.mkdir(parents=True, exist_ok=True)
    output_file = output_dir / "territories_template.xlsx"

    # Créer le fichier Excel avec mise en forme
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Territoires")

        # Accéder au workbook pour la mise en forme
        workbook = writer.book
        worksheet = writer.sheets["Territoires"]

        # Ajuster la largeur des colonnes
        column_widths = {
            "A": 15,  # Numero
            "B": 10,  # Suffixe
            "C": 15,  # Type
            "D": 20,  # Ville
            "E": 45,  # Lien_GPS
            "F": 30,  # Notes
            "G": 30,  # Ne_Pas_Visiter
            "H": 30,  # Notes_Proclamateur
            "I": 25,  # PDF_Filename
        }

        for col, width in column_widths.items():
            worksheet.column_dimensions[col].width = width

    print(f"Template créé: {output_file}")
    print()
    print("Structure des colonnes:")
    print("-" * 60)
    print("  Numero            : Numéro du territoire (ex: SAR-1-01)")
    print("  Suffixe           : Suffixe optionnel (ex: A, B)")
    print("  Type              : 'En présentiel', 'Courrier', 'Téléphone' ou 'Entreprise'")
    print("  Ville             : SARTROUVILLE, MAISONS-LAFFITTE, MONTESSON,")
    print("                      MESNIL LE ROI, CARRIERE S/ BOIS ou Aucun")
    print("  Lien_GPS          : URL Google Maps (optionnel)")
    print("  Notes             : Notes générales (optionnel)")
    print("  Ne_Pas_Visiter    : Adresses à éviter (optionnel)")
    print("  Notes_Proclamateur: Notes pour les proclamateurs (optionnel)")
    print("  PDF_Filename      : Nom du PDF (optionnel, sinon Numero.pdf)")
    print()
    print("Note: La catégorie 'SAR' est sélectionnée automatiquement.")
    print()
    print("Nommage des fichiers PDF:")
    print("-" * 60)
    print("  Si PDF_Filename est vide, le script cherchera:")
    print("    - Pour Numero='SAR-1-01' -> SAR-1-01.pdf")
    print("  Placez vos PDFs dans le dossier: data/pdfs/")


if __name__ == "__main__":
    create_template()

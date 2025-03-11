"""
Author : ExEco Environnement
Edition date : 2025/02
Name : 01_package_verification
Group : Biblio_PatNat
"""

import subprocess
import sys

# Liste des packages à installer ou mettre à jour
packages = [
    "python-docx",
    "pandas",
    "pip",
    "openpyxl",
    "requests"
]

# Fonction pour installer ou mettre à jour les packages
def install_or_update(package):
    try:
        # Tentative d'installation ou mise à jour avec pip
        subprocess.check_call([sys.executable, "-m", "pip", "install", "--upgrade", package])
        print(f"{package} a été installé ou mis à jour avec succès.")
    except subprocess.CalledProcessError:
        print(f"Erreur lors de l'installation ou mise à jour de {package}.")

# Boucle à travers la liste des packages
for package in packages:
    install_or_update(package)

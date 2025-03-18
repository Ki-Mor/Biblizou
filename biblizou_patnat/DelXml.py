"""
Auteur : ExEco Environnement - François Botcazou
Date de création : 2025/03
Dernière mise à jour : 2025/03
Version : 1.0
Nom : DelXml.py
Groupe : Biblizou_PatNat
Description : Module pour supprimer tous les xml après traitements.
Dépendances :
    - Python 3.x
    - QGIS (QgsMessageBar, QgsMessageLog)

Utilisation :
    Ce module doit être appelé depuis une extension QGIS. Il prend en entrée un
    dossier contenant des fichiers XML.
"""


import os
from qgis.core import QgsMessageLog
from PyQt5.QtWidgets import QInputDialog, QMessageBox

class DelXml:
    def __init__(self, main_window):
        """Initialisation de la classe."""
        self.main_window = main_window  # Référence à la fenêtre principale contenant la case à cocher delXml

    def delete_files_in_directory(self, folder_path):
        """Supprime les fichiers XML et TXT inutiles dans le dossier spécifié."""
        try:
            xml_files = [f for f in os.listdir(folder_path) if f.endswith('.xml')]

            for file in xml_files :
                file_path = os.path.join(folder_path, file)
                if os.path.exists(file_path):
                    os.remove(file_path)
                    QgsMessageLog.logMessage(f"Fichier supprimé : {file_path}", "Biblizou")
                else:
                    QgsMessageLog.logMessage(f"Fichier non trouvé : {file_path}", "Biblizou")
        except Exception as e:
            QMessageBox.critical(None, "Erreur", f"Erreur lors de la suppression des fichiers : {e}")

    def run(self):
        """Exécute la suppression si la case à cocher est activée."""
        if not self.main_window.delXml.isChecked():
            QgsMessageLog.logMessage("Suppression annulée : case delXml non cochée.", "Biblizou")
            return

        folder_path, ok = QInputDialog.getText(None, "Sélection du dossier", "Entrez le chemin du dossier :")
        if not ok or not os.path.isdir(folder_path):
            QMessageBox.warning(None, "Erreur", "Dossier invalide ou non sélectionné.")
            return

        self.delete_files_in_directory(folder_path)
        QMessageBox.information(None, "Succès", "Suppression des fichiers terminée.")

# Exemple d'utilisation dans l'extension :
def run_module(main_window):
    module = DelXml(main_window)
    module.run()

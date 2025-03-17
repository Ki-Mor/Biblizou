"""
Auteur : ExEco Environnement - François Botcazou
Date de création : 2025/02
Dernière mise à jour : 2025/03
Version : 1.0
Nom : NaturaDwlXml.py
Groupe : Biblizou_PatNat
Description : Module pour télécharger les xml des zonages natura2000 dans un périmètre donné.
Dépendances :
    - Python 3.x
    - QGIS (QgsMessageBar, QgsMessageLog)

Utilisation :
    Ce module doit être appelé depuis une extension QGIS. Il prend en entrée un
    dossier et y télécharge tous les fichiers XML.
"""


import os
import requests
import time
import logging
from qgis.core import (
    QgsProject,
    QgsVectorLayer,
    QgsFeatureRequest,
    QgsMessageLog
)
from qgis.gui import QgsMapLayerComboBox
from PyQt5.QtWidgets import QInputDialog, QMessageBox

class NaturaDwlXml:
    def __init__(self):
        """Initialisation de la classe."""
        self.patrinat_sic = QgsProject.instance().mapLayersByName("Patrinat : SIC")[0]
        self.patrinat_zps = QgsProject.instance().mapLayersByName("Patrinat : ZPS")[0]
        self.id_mnhn_sic = []
        self.id_mnhn_zps = []
        self.ae_eloignee = None


    def select_layer(self):
        """Demande à l'utilisateur de sélectionner une couche vectorielle dans le projet."""
        layers = [layer.name() for layer in QgsProject.instance().mapLayers().values()]
        selected_layer, ok = QInputDialog.getItem(None, "Sélection de la couche",
                                                  "Choisissez une couche de référence :",
                                                  layers, 0, False)
        if ok and selected_layer:
            self.ae_eloignee = QgsProject.instance().mapLayersByName(selected_layer)[0]
        else:
            QMessageBox.warning(None, "Avertissement", "Aucune couche sélectionnée.")

    def selectionner_et_stocker(self, couche_source, liste_stockage):
        """Sélectionne les entités intersectant AE_eloignee et stocke leurs ID."""
        if couche_source and self.ae_eloignee:
            geom_ref = [f.geometry() for f in self.ae_eloignee.getFeatures()]
            couche_source.removeSelection()
            ids_selectionnes = []

            for feature in couche_source.getFeatures():
                if any(feature.geometry().intersects(g) for g in geom_ref):
                    ids_selectionnes.append(feature.id())
                    liste_stockage.append(feature["id_mnhn"])

            if ids_selectionnes:
                couche_source.selectByIds(ids_selectionnes)
                QgsMessageLog.logMessage(f"{len(ids_selectionnes)} entités sélectionnées dans {couche_source.name()}.", "Biblizou")
            else:
                QgsMessageLog.logMessage(f"Aucune entité sélectionnée dans {couche_source.name()}.", "Biblizou")
        else:
            QgsMessageLog.logMessage(f"La couche {couche_source.name()} ou AE_eloignee est introuvable.", "Biblizou")

    def download_file(self, url, save_path, retries=3):
        """Télécharge un fichier XML avec gestion des erreurs."""
        attempt = 0
        while attempt < retries:
            try:
                response = requests.get(url)
                response.raise_for_status()

                with open(save_path, 'wb') as f:
                    f.write(response.content)
                QgsMessageLog.logMessage(f"Fichier téléchargé avec succès : {save_path}", "Biblizou")
                return True
            except requests.exceptions.RequestException as e:
                attempt += 1
                time.sleep(2 ** attempt)
        QgsMessageLog.logMessage(f"Échec du téléchargement après {retries} tentatives : {url}", "Biblizou", level=2)
        return False

    def construct_url_and_download(self, natura_ids, download_folder):
        """Construit les URLs et télécharge les fichiers XML correspondants."""
        if not natura_ids:
            QgsMessageLog.logMessage("Aucun identifiant Natura trouvé.", "Biblizou")
            return

        for natura_id in natura_ids:
            url = f"https://inpn.mnhn.fr/docs/natura2000/fsdxml/{natura_id}.xml"
            save_path = os.path.join(download_folder, f"{natura_id}.xml")
            self.download_file(url, save_path)

    def run(self):
        """Point d'entrée principal du module."""
        self.select_layer()
        if not self.ae_eloignee:
            QMessageBox.warning(None, "Erreur", "Aucune couche de référence sélectionnée.")
            return
        download_folder, ok = QInputDialog.getText(None, "Chemin vers le dossier de travail", "Copier/coller le chemin")
        if not ok:
            QgsMessageLog.logMessage("L'utilisateur a annulé la saisie du chemin.", "Biblizou", level=2)
            return

        if not os.path.isdir(download_folder):
            QMessageBox.warning(None, "Erreur", f"Le dossier {download_folder} n'existe pas.")
            return

        self.selectionner_et_stocker(self.patrinat_sic, self.id_mnhn_sic)
        self.selectionner_et_stocker(self.patrinat_zps, self.id_mnhn_zps)
        self.construct_url_and_download(self.id_mnhn_sic + self.id_mnhn_zps, download_folder)

# Pour exécuter le module dans QGIS
def run_module():
    module = NaturaDwlXml()
    module.run()

run_module()

"""
Author : ExEco Environnement
Edition date : 2025/02
Name : 11_znieff_xml_download_list
Group : Biblio_PatNat
"""

import os
import requests
import time
import logging
from qgis.core import (
    QgsProject,
    QgsVectorLayer,
    QgsFeatureRequest
)

folder_path = QInputDialog.getInt(None,"Chemin vers le dossier de travail", "Copier/coller le chemin")

# Récupérer les couches Patrinat
patrinat_sic = QgsProject.instance().mapLayersByName("Patrinat : SIC")[0]
patrinat_zps = QgsProject.instance().mapLayersByName("Patrinat : ZPS")[0]

# Récupérer la couche de référence
ae_eloignee = QgsProject.instance().mapLayersByName("AE_eloignee")[0]

# Listes pour stocker les id_mnhn des entités sélectionnées
id_mnhn_sic = []
id_mnhn_zps = []

# Fonction de sélection et récupération des id_mnhn
def selectionner_et_stocker(couche_source, liste_stockage):
    if couche_source and ae_eloignee:
        geom_ref = [f.geometry() for f in ae_eloignee.getFeatures()]
        couche_source.removeSelection()
        ids_selectionnes = []

        for feature in couche_source.getFeatures():
            if any(feature.geometry().intersects(g) for g in geom_ref):
                ids_selectionnes.append(feature.id())
                liste_stockage.append(feature["id_mnhn"])

        if ids_selectionnes:
            couche_source.selectByIds(ids_selectionnes)
            print(f"{len(ids_selectionnes)} entités sélectionnées dans {couche_source.name()}.")
        else:
            print(f"Aucune entité sélectionnée dans {couche_source.name()}.")
    else:
        print(f"La couche {couche_source.name()} ou AE_eloignee est introuvable.")

# Appliquer la sélection sur chaque couche Patrinat
selectionner_et_stocker(patrinat_sic, id_mnhn_sic)
selectionner_et_stocker(patrinat_zps, id_mnhn_zps)

# Fusionner les identifiants des deux listes
znieff_ids = list(set(id_mnhn_sic + id_mnhn_zps))

# Afficher les résultats
print("ID MNHN sélectionnés :")
print("SIC :", id_mnhn_sic)
print("ZPS :", id_mnhn_zps)

# Configurer le logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def download_file(url, save_path, retries=3):
    attempt = 0
    while attempt < retries:
        try:
            response = requests.get(url)
            response.raise_for_status()

            with open(save_path, 'wb') as f:
                f.write(response.content)
            logging.info(f"Fichier téléchargé avec succès : {save_path}")
            return True

        except requests.exceptions.HTTPError as e:
            logging.error(f"Erreur HTTP lors du téléchargement de {url}: {e}")
            return False

        except requests.exceptions.RequestException as e:
            attempt += 1
            logging.warning(f"Erreur réseau lors du téléchargement de {url}: {e}. Tentative {attempt} sur {retries}.")
            time.sleep(2 ** attempt)

    logging.error(f"Échec du téléchargement après {retries} tentatives : {url}")
    return False

def construct_url_and_download(znieff_ids, download_folder):
    if not znieff_ids:
        logging.warning("Aucun identifiant ZNIEFF trouvé.")
        return

    successful_downloads = 0

    for znieff_id in znieff_ids:
        url = f"https://inpn.mnhn.fr/docs/ZNIEFF/znieffxml/{znieff_id}.xml"
        save_path = os.path.join(download_folder, f"{znieff_id}.xml")
        if download_file(url, save_path):
            successful_downloads += 1

    if successful_downloads > 0:
        logging.info(f"\n{successful_downloads} fichiers téléchargés avec succès.")
    else:
        logging.warning("\nAucun fichier n'a été téléchargé.")

def main():
    download_folder = folder_path
    construct_url_and_download(znieff_ids, download_folder)

if __name__ == "__main__":
    main()

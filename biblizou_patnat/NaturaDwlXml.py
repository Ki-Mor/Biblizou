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

# Afficher les résultats
print("ID MNHN sélectionnés :")
print("SIC :", id_mnhn_sic)
print("ZPS :", id_mnhn_zps)

def download_file(url, save_path, retries=3):
    attempt = 0
    while attempt < retries:
        try:
            response = requests.get(url)
            response.raise_for_status()  # Vérifie si la requête a réussi (status code 200)

            with open(save_path, 'wb') as f:
                f.write(response.content)
            logging.info(f"Fichier téléchargé avec succès : {save_path}")
            return True  # Le téléchargement a réussi

        except requests.exceptions.HTTPError as e:
            logging.error(f"Erreur HTTP lors du téléchargement de {url}: {e}")
            return False  # Si c'est une erreur HTTP, on retourne False

        except requests.exceptions.RequestException as e:
            attempt += 1
            logging.warning(f"Erreur réseau lors du téléchargement de {url}: {e}. Tentative {attempt} sur {retries}.")
            time.sleep(2 ** attempt)  # Augmentation exponentielle du délai d'attente

    logging.error(f"Échec du téléchargement après {retries} tentatives : {url}")
    return False  # Après épuisement des tentatives, on retourne False

def construct_url_and_download(n2000_ids, download_folder):
    if not n2000_ids:
        logging.warning("Aucun identifiant ZNIEFF trouvé.")
        return

    successful_downloads = 0

    for n2000_id in n2000_ids:
        url = f"https://inpn.mnhn.fr/docs/natura2000/fsdxml/{n2000_id}.xml"
        save_path = os.path.join(download_folder, f"{n2000_id}.xml")
        if download_file(url, save_path):
            successful_downloads += 1

    if successful_downloads > 0:
        logging.info(f"\n{successful_downloads} fichiers téléchargés avec succès.")
    else:
        logging.warning("\nAucun fichier n'a été téléchargé.")

def main():
    logging.info("La fonction main() a été appelée.")
    """
    Point d'entrée principal du script.
    Demande à l'utilisateur de spécifier le chemin du dossier contenant le fichier TXT.
    Le fichier input_xml_n2000_download_list.txt doit être trouvé automatiquement dans ce dossier.
    """
    # Demander à l'utilisateur de spécifier l'emplacement du dossier contenant le fichier TXT
    download_folder, ok = QInputDialog.getText(None, "Chemin vers le dossier de travail", "Copier/coller le chemin")
    if not ok:
        logging.error("L'utilisateur a annulé la saisie du chemin.")
        return

    # Vérifier si le dossier existe
    if not os.path.isdir(download_folder):
        logging.error(f"Le dossier {download_folder} n'existe pas.")
        return


    # Appeler la fonction pour construire les URLs et télécharger les fichiers
    construct_url_and_download(id_mnhn_sic + id_mnhn_zps, download_folder)

main()

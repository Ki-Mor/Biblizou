"""
Author : ExEco Environnement
Edition date : 2025/02
Name : 11_znieff_xml_download_list
Group : Biblio_PatNat
"""

import os
import sys
import requests
import time
import logging

# Configurer le logging pour des informations détaillées
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def obtain_folder_path():
    return sys.argv[1] if len(sys.argv) > 1 else input("Entrez le chemin du dossier contenant le fichier input_xml_znieff_download_list.txt : ")

def download_file(url, save_path, retries=3):
    """
    Télécharge un fichier à partir de l'URL donnée et l'enregistre à l'emplacement spécifié.
    Réessaie en cas d'échec jusqu'à un nombre spécifié de tentatives.
    """
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

def construct_url_and_download(txt_file_path, download_folder):
    """
    Construit les URLs des fichiers à partir des données du fichier TXT et les télécharge.
    """
    try:
        # Lire et nettoyer le fichier TXT contenant les identifiants ZNIEFF
        with open(txt_file_path, 'r', encoding='utf-8') as f:
            znieff_ids = list(filter(None, (id.strip() for id in f.readlines())))

        # Vérifier si des identifiants ont été trouvés
        if not znieff_ids:
            logging.warning(f"Aucun identifiant trouvé dans le fichier {txt_file_path}.")
            return

        successful_downloads = 0

        # Construire les URLs et télécharger les fichiers
        for znieff_id in znieff_ids:
            url = f"https://inpn.mnhn.fr/docs/ZNIEFF/znieffxml/{znieff_id}.xml"
            save_path = os.path.join(download_folder, f"{znieff_id}.xml")
            if download_file(url, save_path):
                successful_downloads += 1

        # Feedback utilisateur : Nombre de fichiers téléchargés
        if successful_downloads > 0:
            logging.info(f"\n{successful_downloads} fichiers téléchargés avec succès.")
        else:
            logging.warning("\nAucun fichier n'a été téléchargé.")

    except FileNotFoundError:
        logging.error(f"Le fichier {txt_file_path} n'a pas été trouvé.")
    except Exception as e:
        logging.error(f"Erreur inattendue : {e}")

def main():
    """
    Point d'entrée principal du script.
    Demande à l'utilisateur de spécifier le chemin du dossier contenant le fichier TXT.
    Le fichier input_xml_znieff_download_list.txt doit être trouvé automatiquement dans ce dossier.
    """
    # Demander à l'utilisateur de spécifier l'emplacement du dossier contenant le fichier TXT
    folder_path = obtain_folder_path()

    # Vérifier si le dossier existe
    if not os.path.isdir(folder_path):
        logging.error(f"Le dossier {folder_path} n'existe pas.")
        return

    # Définir le chemin complet vers le fichier d'entrée (input_xml_znieff_download_list.txt)
    txt_file_path = os.path.join(folder_path, "input_xml_znieff_download_list.txt")

    # Vérifier si le fichier input_xml_znieff_download_list.txt existe dans ce dossier
    if not os.path.isfile(txt_file_path):
        logging.error(f"Le fichier 'input_xml_znieff_download_list.txt' n'a pas été trouvé dans le dossier {folder_path}.")
        return

    # Le dossier de destination pour les fichiers téléchargés est le même que celui du fichier TXT
    download_folder = folder_path

    # Appeler la fonction pour construire les URLs et télécharger les fichiers
    construct_url_and_download(txt_file_path, download_folder)

if __name__ == "__main__":
    main()

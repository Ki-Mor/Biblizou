from typing import Any, List, Tuple
import pandas as pd
import requests
import re
from pathlib import Path
import concurrent.futures
import time
import logging

# Configuration des logs pour suivre l'exécution du script
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Cache global pour stocker les résultats de la recherche Fuzzy afin d'éviter les requêtes répétées
cache_fuzzy_results = {}

# Configuration des options d'affichage de pandas
pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)
pd.set_option('display.width', None)
pd.set_option('display.max_colwidth', None)

# Demander à l'utilisateur le chemin du dossier contenant les fichiers CSV
folder_path = input("Entrez le chemin du dossier contenant les fichiers XML : ")
logging.info(f"Dossier sélectionné : {folder_path}")

# Lister tous les fichiers CSV trouvés dans le dossier donné
csv_files = Path(folder_path).glob("*.csv")

# DataFrame global pour accumuler toutes les données des fichiers CSV
global_df = pd.DataFrame()

# Fonction pour traiter un fichier CSV et extraire les informations nécessaires
def process_csv(file_path: Path) -> pd.DataFrame:
    try:
        logging.info(f"Traitement du fichier : {file_path.name}")
        df = pd.read_csv(file_path, skiprows=range(0, 1), sep=";", on_bad_lines='skip', encoding='utf-8')
        logging.info(f"Les 10 premières lignes du fichier {file_path.name} :\n{df.head(10)}")
    except Exception as e:
        logging.error(f"Erreur lors de la lecture du fichier {file_path.name}: {e}")
        return pd.DataFrame()

    # Extraction des colonnes nécessaires
    NomTaxCBNB = df['NomTaxCBNB'].tolist()
    AnneeDerniereObservation = df['Année_DernièreObservation'].fillna(0).astype(int).tolist()
    CD_Ref = df['CD_Ref'].fillna(0).astype(int).tolist()

    # Extraction du nom de la commune
    file_name = file_path.name
    match = re.search(r'_([^_]+)\.csv$', file_name)
    extracted_commune = match.group(1) if match else "Inconnu"

    data = {
        'NomTaxCBNB': NomTaxCBNB,
        'Année_DernièreObservation': AnneeDerniereObservation,
        'CD_Ref': CD_Ref,
        'Commune': extracted_commune,
        'Obs': 'CBN Brest'
    }

    return pd.DataFrame(data)

# Charger les données depuis tous les fichiers CSV trouvés
for file_path in csv_files:
    df = process_csv(file_path)
    if not df.empty:
        global_df = pd.concat([global_df, df], ignore_index=True)

logging.info(f"Taille du DataFrame global après chargement des fichiers : {global_df.shape}")

# Fonction pour récupérer les statuts des taxons en lots
def fetch_status_data(batch: List[int]) -> List[dict]:
    logging.info(f"Récupération des statuts pour le lot : {batch}")
    url_base = 'https://taxref.mnhn.fr/api/status/search/lines?locationId=INSEEC29241&page=1&size=10000'
    url_complete = url_base + ''.join([f'&taxrefId={elem}' for elem in batch])

    for attempt in range(3):
        try:
            response = requests.get(url_complete, headers={'accept': 'application/hal+json;version=1'})
            if response.status_code == 200:
                logging.info(f"Succès de la requête pour le lot {batch}")
                return response.json().get('_embedded', {}).get('status', [])
            else:
                logging.warning(f"Erreur pour le lot {batch}, tentative {attempt + 1}. Code HTTP: {response.status_code}")
        except requests.RequestException as e:
            logging.error(f"Erreur lors de la requête pour le lot {batch}: {e}")
        time.sleep(2)
    logging.error(f"Échec de la récupération des statuts pour le lot {batch}")
    return []

# Diviser les identifiants taxonomiques en lots et récupérer les données correspondantes
taxref_ids = global_df['CD_Ref'].unique().tolist()
batches = [taxref_ids[i:i + 150] for i in range(0, len(taxref_ids), 150)]
all_status_data = []
for batch in batches:
    logging.info(f"Traitement du lot de CD_Ref : {batch}")
    status_data = fetch_status_data(batch)
    if status_data:
        all_status_data.extend(status_data)

logging.info("Fusion des données de statuts avec le DataFrame principal")

# Construction et fusion des DataFrames pour finaliser le traitement
if all_status_data:
    extracted_data = []
    for item in all_status_data:
        taxon = item.get('taxon', {})
        extracted_data.append({
            'CD_Ref': taxon.get('id', ''),
            'statusCode': item.get('statusCode', ''),
            'statusTypeName': item.get('statusTypeName', '')
        })

    df_status = pd.DataFrame(extracted_data)
    logging.info(f"Nombre de statuts récupérés : {df_status.shape[0]}")
    global_df = global_df.merge(df_status, on='CD_Ref', how='left')
else:
    logging.warning("Aucun statut récupéré. La fusion ne sera pas effectuée.")

logging.info("Affichage du DataFrame global après enrichissement avec les statuts :")
print(global_df.head(20))

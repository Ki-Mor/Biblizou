"""Script Ecalluna en cours de travail"""


from typing import List, Tuple
import pandas as pd
import requests
import re
from pathlib import Path
import concurrent.futures
import time
import logging
from threading import Lock  # Importation du verrou pour synchroniser les mises à jour

# Configuration des logs
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Cache des résultats de la recherche Fuzzy (utilisation d'un cache global)
cache_fuzzy_results = {}
cache_taxref_results = {}
lock = Lock()  # Création d'un verrou pour éviter les accès concurrents au DataFrame
session = requests.Session()  # Session globale pour réutiliser les connexions HTTP
insee = {'Melgven' : 29146, 'Rosporden' : 29241, 'Elliant' : 29049, 'Saint-Yvi' :29272}

# options d'affichage pandas
pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)
pd.set_option('display.width', None)
pd.set_option('display.max_colwidth', None)

# Dossier contenant les fichiers CSV
folder_path = r"\\192.168.1.100\ExEco_Env\Affaires_en_cours\XECO_2403_Carrière Coat Culoden_Rosporden_29\DATA\XECO_PatNat"
# folder_path = input("Entrez le chemin du dossier contenant les fichiers XML : ")

# Vérifier si le dossier existe
if not Path(folder_path).exists():
    raise OSError(f"Le dossier spécifié '{folder_path}' n'existe pas ou n'est pas accessible.")

# Lister tous les fichiers CSV dans le dossier
csv_files = Path(folder_path).glob("*.csv")

# DataFrame global pour accumuler toutes les données
global_df = pd.DataFrame()

# Traiter chaque fichier CSV
def process_csv(file_path: Path) -> pd.DataFrame:
    try:
        # Lire le fichier CSV avec pandas
        df = pd.read_csv(file_path, skiprows=range(0, 1), sep=";", on_bad_lines='skip', encoding='utf-8')
    except Exception as e:
        logging.error(f"Erreur lors de la lecture du fichier {file_path.name}: {e}")
        return pd.DataFrame()  # Retourner un DataFrame vide en cas d'erreur

    # Vérifier la présence des colonnes requises et les gérer
    if 'NomTaxCBNB' not in df.columns:
        logging.warning(f"Colonne 'NomTaxCBNB' absente dans {file_path.name}, remplacement par des valeurs vides.")
        df['NomTaxCBNB'] = ""
    if 'Année_DernièreObservation' not in df.columns:
        logging.warning(f"Colonne 'Année_DernièreObservation' absente dans {file_path.name}, remplacement par NaN.")
        df['Année_DernièreObservation'] = pd.NA
    if 'CD_Ref' not in df.columns:
        logging.warning(f"Colonne 'CD_Ref' absente dans {file_path.name}, remplacement par 0.")
        df['CD_Ref'] = 0

    # Extraire les colonnes spécifiques
    NomTaxCBNB = df['NomTaxCBNB'].tolist()
    AnneeDerniereObservation = pd.to_numeric(df['Année_DernièreObservation'], errors='coerce').astype('Int64')
    CD_Ref = df['CD_Ref'].fillna(0).astype(int).tolist()

    # Extraire le nom de la commune à partir du nom du fichier
    file_name = file_path.name
    match = re.search(r'_([^_]+)\.csv$', file_name)

    extracted_commune = match.group(1) if match else "Inconnu"

    # Ajouter les colonnes supplémentaires à df
    data = {
        'NomTaxCBNB': NomTaxCBNB,
        'Année_DernièreObservation': AnneeDerniereObservation,
        'CD_Ref': CD_Ref,
        'Commune': extracted_commune,
        'Obs': 'CBN Brest'
    }

    return pd.DataFrame(data)

# Charger les données depuis les fichiers CSV
for file_path in csv_files:
    df = process_csv(file_path)
    if not df.empty:
        global_df = pd.concat([global_df, df], ignore_index=True)

# Fonction pour enrichir global_df
def process_row(index, row, cache):
    if pd.notna(row['NomTaxCBNB']) and row['NomTaxCBNB'].strip() != "":
        scientific_name, reference_id = fuzzy_match_taxa_with_cache(row['NomTaxCBNB'])
        with lock:
            global_df.at[index, 'CD_Ref'] = reference_id
        logging.info(f"Modifications pour l'index {index}: {row['NomTaxCBNB']} -> {scientific_name}, {reference_id}")

# Fonction pour corriger global_df
def correct_CD_Ref_data():
    global global_df

    # Appliquer la parallélisation pour remplir le CD_Ref avec un cache global
    with concurrent.futures.ThreadPoolExecutor(max_workers=10) as executor:
        futures = [
            executor.submit(process_row, index, row, cache_fuzzy_results)
            for index, row in global_df.iterrows()
        ]
        for future in concurrent.futures.as_completed(futures):
            future.result()

def fuzzy_match_taxa_with_cache(nom_tax: str) -> Tuple[str, int]:
    if nom_tax in cache_fuzzy_results:
        return cache_fuzzy_results[nom_tax]

    word = nom_tax.split()
    genus, specie = word[0], word[1] if len(word) > 1 else ""
    url = f'https://taxref.mnhn.fr/api/taxa/fuzzyMatch?term={genus}%20{specie}'

    for attempt in range(5):
        try:
            response = session.get(url, headers={'accept': "application/hal+json;version=1"}, timeout=20)
            if response.status_code == 200:
                data = response.json()
                taxa = data.get('_embedded', {}).get('taxa', [])
                if taxa:
                    result = (taxa[0].get("scientificName", ""), int(taxa[0].get("referenceId", 0)))
                    cache_fuzzy_results[nom_tax] = result
                    return result
            return "", 0
        except requests.RequestException as e:
            logging.warning(f"Erreur API pour {nom_tax}, tentative {attempt + 1}: {e}")
            time.sleep(2 ** attempt)
    return "", 0


def get_taxref_data(CD_Ref, cache={}):
    # Vérifier si les données sont déjà en cache
    if CD_Ref in cache:
        return cache[CD_Ref]

    # URL de l'API avec cd_nom comme identifiant
    url = f"https://taxref.mnhn.fr/api/taxa/{CD_Ref}"

    # Définir les headers pour l'API
    headers = {"accept": "application/hal+json;version=1"}

    try:
        # Faire la requête GET
        try:
            response = session.get(url, headers=headers, timeout=30)
            response.raise_for_status()  # Déclenche une exception pour les erreurs HTTP
        except requests.Timeout:
            logging.error(f"Timeout lors de la requête pour {CD_Ref}")
            return {'REGNE': '', 'GROUPE': '', 'NOM_COMPLET': '', 'NOM_VERN': ''}
        except requests.RequestException as e:
            logging.error(f"Erreur API pour {CD_Ref}: {e}")
            return {'REGNE': '', 'GROUPE': '', 'NOM_COMPLET': '', 'NOM_VERN': ''}

        # Vérifier que la réponse est correcte (code 200)
        if response.status_code == 200:
            data = response.json()

            # Vérifier si la réponse contient des informations
            if data:
                # Récupérer les informations pertinentes directement dans la réponse
                result = {
                    'REGNE': data.get('kingdomName', ''),
                    'GROUPE': data.get('vernacularGroup2', ''),
                    'NOM_COMPLET': data.get('fullName', ''),
                    'NOM_VERN': data.get('frenchVernacularName', '')
                }

                # Mettre les données dans le cache
                cache[CD_Ref] = result
                return result
            else:
                print(f"Aucun taxon trouvé pour {CD_Ref}.")
                return {'REGNE': '', 'GROUPE': '', 'NOM_COMPLET': '', 'NOM_VERN': ''}
        else:
            print(f"Erreur API pour {CD_Ref}: {response.status_code}")
            return {'REGNE': '', 'GROUPE': '', 'NOM_COMPLET': '', 'NOM_VERN': ''}
    except Exception as e:
        print(f"Erreur lors de l'interrogation de l'API pour {CD_Ref}: {e}")
        return {'REGNE': '', 'GROUPE': '', 'NOM_COMPLET': '', 'NOM_VERN': ''}

def enrich_taxref_data():
    global global_df
    logging.info("Enrichissement du DataFrame avec les données de TaxRef :")

    # Appliquer la parallélisation pour enrichir global_df
    with concurrent.futures.ThreadPoolExecutor(max_workers=10) as executor:
        futures = [
            executor.submit(enrich_row, index, row)
            for index, row in global_df.iterrows()
            if row['CD_Ref'] != 0 and row['CD_Ref'] not in cache_taxref_results
        ]
        for future in concurrent.futures.as_completed(futures):
            future.result()

def enrich_row(index, row):
    # Récupérer les données de TaxRef pour le CD_Ref donné
    taxref_data = get_taxref_data(row['CD_Ref'], cache_taxref_results)

    # Couper NOM_VERN avant la première virgule
    nom_vern = taxref_data['NOM_VERN']
    if isinstance(nom_vern, str) and ',' in nom_vern:
        taxref_data['NOM_VERN'] = nom_vern.split(',')[0].strip()

    # Mettre à jour les colonnes correspondantes dans global_df
    with lock:
        global_df.at[index, 'REGNE'] = taxref_data['REGNE']
        global_df.at[index, 'GROUPE'] = taxref_data['GROUPE']
        global_df.at[index, 'NOM_COMPLET'] = taxref_data['NOM_COMPLET']
        global_df.at[index, 'NOM_VERN'] = taxref_data['NOM_VERN']

    # Ajouter un message de log pour indiquer que la ligne a été traitée
    logging.info(f"Ligne {index} traitée: CD_Ref={row['CD_Ref']}, REGNE={taxref_data['REGNE']}, GROUPE={taxref_data['GROUPE']}, NOM_COMPLET={taxref_data['NOM_COMPLET']}, NOM_VERN={taxref_data['NOM_VERN']}")


def process_data():
    #Orchestre la correction et l'enrichissement des données.
    logging.info("Début du traitement des données.")
    correct_CD_Ref_data()
    enrich_taxref_data()
    logging.info("Traitement des données terminé.")

process_data()


def fetch_status_data(batch: List[int]) -> List[dict]:
    url_base = 'https://taxref.mnhn.fr/api/status/search/lines?locationId=INSEEC29241&page=1&size=10000'
    url_complete = url_base + ''.join([f'&taxrefId={elem}' for elem in batch])

    for attempt in range(5):  # Tentatives de requêtes
        try:
            response = session.get(url_complete, headers={'accept': 'application/hal+json;version=1'}, timeout=30)
            if response.status_code == 200:
                return response.json().get('_embedded', {}).get('status', [])
            else:
                logging.warning(f"Erreur pour le lot {batch}, tentative {attempt + 1}. Code HTTP: {response.status_code}")
        except requests.RequestException as e:
            logging.error(f"Erreur lors de la requête pour le lot {batch}: {e}")
        time.sleep(2)
    return []

def enrich_status_data():
    global global_df
    logging.info("Enrichissement des statuts de conservation depuis TaxRef")

    # Diviser les CD_Ref en lots de 150
    taxref_ids = global_df['CD_Ref'].unique().tolist()
    batches = [taxref_ids[i:i + 150] for i in range(0, len(taxref_ids), 150)]

    all_status_data = []
    for batch in batches:
        status_data = fetch_status_data(batch)
        if status_data:
            all_status_data.extend(status_data)

    # Création d'un dictionnaire {CD_Ref: Liste de statuts}
    status_dict = {}
    for entry in all_status_data:
        taxref_id = entry.get('taxrefId', 0)
        status = entry.get('statusLabel', '')  # Récupération du libellé du statut
        if taxref_id in status_dict:
            status_dict[taxref_id].append(status)
        else:
            status_dict[taxref_id] = [status]

    # Mise à jour de global_df avec la colonne STATUTS
    with lock:
        global_df['STATUTS'] = global_df['CD_Ref'].map(lambda x: ', '.join(status_dict.get(x, [])))

    logging.info("Enrichissement des statuts terminé.")

# Ajout de la nouvelle fonction au pipeline de traitement
def process_data():
    logging.info("Début du traitement des données.")
    correct_CD_Ref_data()
    enrich_taxref_data()
    enrich_status_data()  # Ajout de l'enrichissement des statuts
    logging.info("Traitement des données terminé.")

process_data()
print(global_df.head(25))


# # Pivot table pour les années par commune
# pivot_df_commune = global_df.pivot_table(
#     index=['CD_Ref', 'NOM_COMPLET'],
#     columns='Commune',
#     values=['Année_DernièreObservation'],
#     aggfunc={'Année_DernièreObservation': 'max'},
#     fill_value=''
# )

# pivot_df_status = global_df.pivot_table(
#     index=['CD_Ref', 'NOM_COMPLET'],
#     columns='statusTypeName',
#     values=['statusCode'],
#     aggfunc='first',
#     fill_value=''
# )

# print(pivot_df_commune.head(25))
# print(pivot_df_status.head(25))

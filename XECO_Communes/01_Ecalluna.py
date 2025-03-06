from typing import Any, List, Tuple
import pandas as pd
import requests
import re
from pathlib import Path
import concurrent.futures
import time
import logging

# Configuration des logs
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Cache des résultats de la recherche Fuzzy (utilisation d'un cache global)
cache_fuzzy_results = {}

# options d'affichage pandas
pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)
pd.set_option('display.width', None)
pd.set_option('display.max_colwidth', None)

# Dossier contenant les fichiers CSV
folder_path = input("Entrez le chemin du dossier contenant les fichiers XML : ")

# Lister tous les fichiers CSV dans le dossier
csv_files = Path(folder_path).glob("*.csv")

# DataFrame global pour accumuler toutes les données
global_df = pd.DataFrame()


# Traiter chaque fichier CSV
def process_csv(file_path: Path) -> pd.DataFrame:
    try:
        # Lire le fichier CSV avec pandas
        df = pd.read_csv(file_path, skiprows=range(0, 1), sep=";", on_bad_lines='skip', encoding='utf-8')
        logging.info(f"Les 10 premières lignes du fichier :\n{df.head(10)}")
    except Exception as e:
        logging.error(f"Erreur lors de la lecture du fichier {file_path.name}: {e}")
        return pd.DataFrame()  # Retourner un DataFrame vide en cas d'erreur

    # Extraire les colonnes spécifiques
    NomTaxCBNB = df['NomTaxCBNB'].tolist()
    AnneeDerniereObservation = df['Année_DernièreObservation'].fillna(0).astype(int).tolist()
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


# Fonction de recherche floue avec cache
def fuzzy_match_taxa_with_cache(nom_tax: str) -> Tuple[str, int]:
    mots = nom_tax.split()
    term1, term2 = mots[0], mots[1] if len(mots) > 1 else ""
    url = f'https://taxref.mnhn.fr/api/taxa/fuzzyMatch?term={term1}%2{term2}'

    # Gestion des erreurs avec tentatives de reconnexion
    for attempt in range(3):  # Nombre de tentatives de requête
        try:
            response = requests.get(url, headers={'accept': 'application/hal+json;version=1'})
            if response.status_code == 200:
                data = response.json()
                taxa = data.get('_embedded', {}).get('taxa', [])
                if taxa:
                    scientific_name = taxa[0].get("scientificName", "")
                    reference_id = int(taxa[0].get("referenceId", 0)) if taxa[0].get("referenceId") else 0
                    return scientific_name, reference_id
                return "", 0
            else:
                logging.warning(
                    f"Requête échouée pour {nom_tax}, tentative {attempt + 1}. Code HTTP: {response.status_code}")
        except requests.RequestException as e:
            logging.error(f"Erreur lors de la requête API pour {nom_tax} : {e}")
        time.sleep(2)  # Attendre 2 secondes avant de réessayer
    return "", 0  # Retourner vide si échec après plusieurs tentatives


# Fonction de traitement des lignes avec parallélisation
def process_row(index, row, cache_fuzzy_results) -> Tuple[int, str, str, int]:
    if row['CD_Ref'] == 0:
        nom_tax = row['NomTaxCBNB']
        if nom_tax in cache_fuzzy_results:
            scientific_name, reference_id = cache_fuzzy_results[nom_tax]
        else:
            scientific_name, reference_id = fuzzy_match_taxa_with_cache(nom_tax)
            cache_fuzzy_results[nom_tax] = (scientific_name, reference_id)

        if reference_id:
            global_df.at[index, 'CD_Ref'] = reference_id
        return index, nom_tax, scientific_name, reference_id
    return None


# Appliquer la parallélisation pour remplir le CD_Ref avec un cache global
with concurrent.futures.ThreadPoolExecutor(max_workers=10) as executor:  # Limiter à 10 threads
    futures = [
        executor.submit(process_row, index, row, cache_fuzzy_results)
        for index, row in global_df.iterrows() if row['CD_Ref'] == 0
    ]

    for future in concurrent.futures.as_completed(futures):
        result = future.result()
        if result:
            index, nom_tax, scientific_name, reference_id = result
            logging.info(f"Modifications pour l'index {index}: {nom_tax} -> {scientific_name}, {reference_id}")


# Fonction pour récupérer les données de statut par lots
def fetch_status_data(batch: List[int]) -> List[dict]:
    url_base = 'https://taxref.mnhn.fr/api/status/search/lines?locationId=INSEEC29241&page=1&size=10000'
    url_complete = url_base + ''.join([f'&taxrefId={elem}' for elem in batch])

    for attempt in range(3):  # Tentatives de requêtes
        try:
            response = requests.get(url_complete, headers={'accept': 'application/hal+json;version=1'})
            if response.status_code == 200:
                return response.json().get('_embedded', {}).get('status', [])
            else:
                logging.warning(
                    f"Erreur pour le lot {batch}, tentative {attempt + 1}. Code HTTP: {response.status_code}")
        except requests.RequestException as e:
            logging.error(f"Erreur lors de la requête pour le lot {batch}: {e}")
        time.sleep(2)
    return []


# Diviser les IDs en lots
taxref_ids = global_df['CD_Ref'].unique().tolist()
batches = [taxref_ids[i:i + 150] for i in range(0, len(taxref_ids), 150)]

# Récupérer les données de statut pour chaque lot
all_status_data = []
for batch in batches:
    status_data = fetch_status_data(batch)
    if status_data:
        all_status_data.extend(status_data)

# Traiter les données extraites
extracted_data = []
for item in all_status_data:
    taxon = item.get('taxon', {})
    extracted_data.append({
        'id': taxon.get('id', ''),
        'scientificName': taxon.get('scientificName', ''),
        'statusTypeName': item.get('statusTypeName', ''),
        'statusCode': item.get('statusCode', '')
    })

# Créer un DataFrame avec les données de statut
df_status = pd.DataFrame(extracted_data)

# Fusionner les données de statut avec le DataFrame global
df_status = df_status.merge(global_df[['CD_Ref', 'Commune', 'Année_DernièreObservation']],
                            left_on='id', right_on='CD_Ref', how='left')

# Pivot table pour les années par commune
pivot_df_commune = df_status.pivot_table(
    index=['id', 'scientificName'],
    columns='Commune',
    values=['Année_DernièreObservation'],
    aggfunc={'Année_DernièreObservation': 'max'},
    fill_value=''
)
# logging.info(f"Table status :\n{pivot_df_commune.head(100)}")

# # Exporter la table pivot_df_commune dans un fichier CSV séparé par des points-virgules
# pivot_commune_path = Path(folder_path) / "test_commune.csv"
# pivot_df_commune.to_csv(pivot_commune_path, sep=";", index=True)  # index=True pour garder les indices dans le fichier
# logging.info(f"Table pivot_commune exportée vers : {pivot_commune_path}")


# Pivot table pour les statuts par type de statut
pivot_df_status = df_status.pivot_table(
    index=['id', 'scientificName'],
    columns='statusTypeName',
    values=['statusCode'],
    aggfunc='first',
    fill_value=''
)


# # Exporter la table pivot_df_commune dans un fichier CSV séparé par des points-virgules
# pivot_df_status_path = Path(folder_path) / "pivot_df_status.csv"
# pivot_df_commune.to_csv(pivot_df_status_path, sep=";", index=True)  # index=True pour garder les indices dans le fichier
# logging.info(f"Table pivot_df_status exportée vers : {pivot_df_status_path}")

# Avant la fusion, trier les DataFrames par les colonnes sur lesquelles vous effectuez la fusion
pivot_df_commune_sorted = pivot_df_commune.reset_index().sort_values(by=['id','scientificName'])
pivot_df_status_sorted = pivot_df_status.reset_index().sort_values(by=['id','scientificName'])

# Fusionner les deux DataFrames triés
final_df = pd.merge(pivot_df_commune_sorted, pivot_df_status_sorted, on=['id', 'scientificName'], how='outer')


# Afficher la table finale fusionnée
# logging.info(f"Table finale fusionnée:\n{final_df.head(25)}")

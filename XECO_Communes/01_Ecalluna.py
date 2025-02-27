from typing import Any
import pandas as pd
import requests
import re
from pathlib import Path
import concurrent.futures
import math

# Cache des résultats de la recherche Fuzzy (utilisation d'un cache global)
cache_fuzzy_results = {}

# options d'affichage pandas
pd.set_option('display.max_columns', None)  # Afficher toutes les colonnes
pd.set_option('display.max_rows', None)  # Afficher toutes les lignes
pd.set_option('display.width', None)  # Pas de limite sur la largeur d'affichage
pd.set_option('display.max_colwidth', None)  # Pas de limite sur la largeur des colonnes

# Dossier contenant les fichiers CSV
folder_path = input("Entrez le chemin du dossier contenant les fichiers XML : ")

# Lister tous les fichiers CSV dans le dossier
csv_files = Path(folder_path).glob("*.csv")

# DataFrame global pour accumuler toutes les données
global_df = pd.DataFrame()

# Traiter chaque fichier CSV
for file_path in csv_files:
    try:
        # Lire le fichier CSV avec pandas
        df = pd.read_csv(file_path, skiprows=range(0, 1), sep=";", on_bad_lines='skip')
    except Exception as e:
        print(f"Erreur lors de la lecture du fichier {file_path.name}: {e}")
        continue

    # Extraire les colonnes spécifiques dans des listes sans 'NomTaxRef'
    NomTaxCBNB = df['NomTaxCBNB'].tolist()
    AnneeDerniereObservation = df['Année_DernièreObservation'].fillna(0).astype(int).tolist()
    CD_Ref = df['CD_Ref'].fillna(0).astype(int).tolist()

    # Extraire le nom de la commune à partir du nom du fichier
    file_name = file_path.name  # Utiliser Path pour extraire le nom du fichier

    # Utiliser une expression régulière pour extraire la commune (après le dernier "_" et avant ".csv")
    match = re.search(r'_([^_]+)\.csv$', file_name)

    if match:
        extracted_commune: str | Any = match.group(1)  # Cela extraira "Rosporden"
        print(f"Nom de la commune pour le fichier {file_name}: {extracted_commune}")
    else:
        print(f"Aucun extrait trouvé pour le fichier {file_name}")
        extracted_commune = "Inconnu"

    observateur = 'CBN Brest'

    # Ajouter les colonnes supplémentaires à df
    data = {
        'NomTaxCBNB': NomTaxCBNB,
        'Année_DernièreObservation': AnneeDerniereObservation,
        'CD_Ref': CD_Ref,
        'Commune': extracted_commune,
        'Obs': observateur
    }

    df = pd.DataFrame(data)

    # Fusionner les données avec le DataFrame global
    global_df = pd.concat([global_df, df], ignore_index=True)


# Fonction pour traiter chaque ligne avec la parallélisation et utiliser un cache
def process_row(index, row, cache_fuzzy_results):
    if row['CD_Ref'] == 0:
        nom_tax = row['NomTaxCBNB']

        # Vérifier si le nom est déjà dans le cache
        if nom_tax in cache_fuzzy_results:
            scientific_name, reference_id = cache_fuzzy_results[nom_tax]
        else:
            scientific_name, reference_id = fuzzy_match_taxa_with_cache(nom_tax)
            # Mémoriser les résultats pour réutilisation future
            cache_fuzzy_results[nom_tax] = (scientific_name, reference_id)

        if row['CD_Ref'] == 0 and reference_id:
            global_df.at[index, 'CD_Ref'] = reference_id
        return index, row['NomTaxCBNB'], scientific_name, reference_id
    return None


# Fonction de match Fuzzy
def fuzzy_match_taxa_with_cache(nom_tax):
    mots = nom_tax.split()
    if len(mots) >= 2:
        term1 = mots[0]
        term2 = mots[1]
    else:
        term1 = mots[0]
        term2 = ""

    # Construire l'URL de la requête API
    url = f'https://taxref.mnhn.fr/api/taxa/fuzzyMatch?term={term1}%2{term2}'

    # Faire la requête API
    response = requests.get(url, headers={'accept': 'application/hal+json;version=1'})

    # Si la requête a réussi
    if response.status_code == 200:
        data = response.json()
        if "_embedded" in data and "taxa" in data["_embedded"]:
            taxa = data["_embedded"]["taxa"]
            if taxa:
                scientific_name = taxa[0].get("scientificName", "")
                reference_id = taxa[0].get("referenceId", "")
                if reference_id is not None:
                    reference_id = int(reference_id)
                return scientific_name, reference_id
    return None, None


# Appliquer la parallélisation pour remplir le CD_Ref avec un cache global
with concurrent.futures.ThreadPoolExecutor() as executor:
    futures = [
        executor.submit(process_row, index, row, cache_fuzzy_results)
        for index, row in global_df.iterrows() if row['CD_Ref'] == 0
    ]

    # Récupérer les résultats et traiter les modifications
    for future in concurrent.futures.as_completed(futures):
        result = future.result()
        if result is not None:
            index, nom_tax, scientific_name, reference_id = result
            print(f"Modifications pour l'index {index}: {nom_tax} -> {scientific_name}, {reference_id}")


'''








INSÉRER ICI LA REQUÊTE TAXA INPN !









'''

# Taille du lot de taxrefId
batch_size = 150  # Vous pouvez ajuster cette taille selon la limite autorisée

# Diviser la liste de taxrefId en lots
taxref_ids = global_df['CD_Ref'].unique().tolist()  # Liste des CD_Ref uniques
batches = [taxref_ids[i:i + batch_size] for i in range(0, len(taxref_ids), batch_size)]

# Fonction pour envoyer les requêtes par lots
def fetch_status_data(batch):
    url_base = 'https://taxref.mnhn.fr/api/status/search/lines?locationId=INSEEC29241&page=1&size=10000'
    url_complete = url_base + ''.join([f'&taxrefId={elem}' for elem in batch])

    headers = {
        'accept': 'application/hal+json;version=1'
    }

    response = requests.get(url_complete, headers=headers)

    if response.status_code == 200:
        data = response.json()
        status_data = data.get('_embedded', {}).get('status', [])
        return status_data
    else:
        print(f"Erreur lors de la récupération des données pour le lot {batch}: {response.status_code}")
        return None

# Faire les requêtes pour chaque lot
all_status_data = []

for batch in batches:
    status_data = fetch_status_data(batch)
    if status_data:
        all_status_data.extend(status_data)

# Traiter les données extraites
extracted_data = []
for item in all_status_data:
    taxon = item.get('taxon', {})
    statusTypeName = item.get('statusTypeName', '')
    statusCode = item.get('statusCode', '')

    taxon_id = taxon.get('id', '')
    scientificName = taxon.get('scientificName', '')

    extracted_data.append({
        'id': taxon_id,
        'scientificName': scientificName,
        'statusTypeName': statusTypeName,
        'statusCode': statusCode
    })

# Créer un DataFrame pandas avec les données extraites
df_status = pd.DataFrame(extracted_data)

# Fusionner les données de statut avec le DataFrame global
df_status = df_status.merge(global_df[['CD_Ref', 'Commune', 'Année_DernièreObservation']],
                            left_on='id', right_on='CD_Ref', how='left')

# Pivot table pour les années par commune
pivot_df_commune = df_status.pivot_table(
    index=['id', 'scientificName'],  # Les indices de la pivot table
    columns='Commune',  # Les colonnes seront les différentes communes
    values=['Année_DernièreObservation'],  # Valeurs que nous voulons afficher
    aggfunc={'Année_DernièreObservation': 'max'},  # Pour chaque combinaison, on prend le max de l'année
    fill_value=''  # Remplacer les NaN par des valeurs vides
)

# Remise en forme pour obtenir une table lisible
pivot_df_commune = pivot_df_commune.reset_index()
print(pivot_df_commune)

# Pivot table pour les status par type de statut
pivot_df_status = df_status.pivot_table(
    index=['id', 'scientificName'],  # Les indices de la pivot table
    columns='statusTypeName',  # Les colonnes seront les différents types de statut
    values=['statusCode'],  # Valeurs que nous voulons afficher
    aggfunc='first',  # Première valeur de statusCode pour chaque groupe
    fill_value=''  # Remplacer les NaN par des valeurs vides
)

# Remise en forme pour obtenir une table lisible
pivot_df_status = pivot_df_status.reset_index()
print(pivot_df_status)

# Fusionner les deux pivot tables sur les colonnes 'id' et 'scientificName'
final_df = pd.merge(pivot_df_commune, pivot_df_status, on=['id', 'scientificName'], how='outer')

# Afficher la table finale fusionnée
print(final_df)


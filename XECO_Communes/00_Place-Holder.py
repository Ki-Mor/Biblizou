from typing import Any
import pandas as pd
import requests
import re
from pathlib import Path
import concurrent.futures

# Cache des résultats de la recherche Fuzzy
cache_fuzzy_results = {}

# options d'affichage pandas
pd.set_option('display.max_columns', None)  # Afficher toutes les colonnes
pd.set_option('display.max_rows', None)  # Afficher toutes les lignes
pd.set_option('display.width', None)  # Pas de limite sur la largeur d'affichage
pd.set_option('display.max_colwidth', None)  # Pas de limite sur la largeur des colonnes

# TODO à implémenter plus tard
# def obtain_folder_path():
#     return sys.argv[1] if len(sys.argv) > 1 else input("Veuillez entrer le chemin du fichier : ")

# Lire le fichier CSV avec pandas
try:
    df = pd.read_csv("C:/Users/celin/OneDrive/Bureau/(17_02_2025)_Rosporden.csv", skiprows=range(0, 1), sep=";",
                     on_bad_lines='skip')
    print(df.head(15))  # Afficher les premières lignes du dataframe pour vérification
except Exception as e:
    print(f"Erreur lors de la lecture du fichier : {e}")

# Extraire les colonnes spécifiques dans des listes
NomTaxCBNB = df['NomTaxCBNB'].tolist()
AnneeDerniereObservation = df['Année_DernièreObservation'].tolist()
NomTaxRef = df['NomTaxRef'].tolist()
CD_Ref = df['CD_Ref'].fillna(0).astype(int).tolist()

# Extraire le nom de la commune à partir du nom du fichier
file_path = "C:/Users/celin/OneDrive/Bureau/(17_02_2025)_Rosporden.csv"
file_name = Path(file_path).name  # Utiliser Path pour extraire le nom du fichier

# Utiliser une expression régulière pour extraire la commune (après le dernier "_" et avant ".csv")
match = re.search(r'_([^_]+)\.csv$', file_name)

if match:
    extracted_commune: str | Any = match.group(1)  # Cela extraira "Rosporden"
    print("Nom de la commune :", extracted_commune)
else:
    print("Aucun extrait trouvé")

data = {
    'NomTaxCBNB': NomTaxCBNB,
    'Année_DernièreObservation': AnneeDerniereObservation,
    'NomTaxRef': NomTaxRef,
    'CD_Ref': CD_Ref,
    'Commune': extracted_commune
}

df = pd.DataFrame(data)


# Fonction pour effectuer une recherche Fuzzy sur l'API avec cache
def fuzzy_match_taxa_with_cache(nom_tax):
    if nom_tax in cache_fuzzy_results:
        return cache_fuzzy_results[nom_tax]
    else:
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

            # Vérifier s'il y a des résultats dans "taxa"
            if "_embedded" in data and "taxa" in data["_embedded"]:
                taxa = data["_embedded"]["taxa"]

                # Prendre le premier résultat (en supposant qu'il est pertinent)
                if taxa:
                    scientific_name = taxa[0].get("scientificName", "")
                    reference_id = taxa[0].get("referenceId", "")

                    # Assurer que reference_id est bien un entier
                    if reference_id is not None:
                        reference_id = int(reference_id)  # Conversion explicite en entier

                    # Mémoriser les résultats pour réutilisation future
                    cache_fuzzy_results[nom_tax] = (scientific_name, reference_id)
                    return scientific_name, reference_id

    return None, None  # Si rien n'est trouvé


# Fonction pour traiter chaque ligne avec la parallélisation
def process_row(index, row):
    if row['NomTaxRef'] == 0 or row['CD_Ref'] == 0:
        nom_tax = row['NomTaxCBNB']
        scientific_name, reference_id = fuzzy_match_taxa_with_cache(nom_tax)

        # Remplacer les valeurs uniquement si elles sont égales à 0
        if row['NomTaxRef'] == 0 and scientific_name:
            df.at[index, 'NomTaxRef'] = scientific_name

        if row['CD_Ref'] == 0 and reference_id:
            df.at[index, 'CD_Ref'] = reference_id
        return index, row['NomTaxCBNB'], scientific_name, reference_id
    return None


# Utilisation de la parallélisation pour effectuer les appels API en parallèle
with concurrent.futures.ThreadPoolExecutor() as executor:
    futures = [
        executor.submit(process_row, index, row)
        for index, row in df.iterrows() if row['NomTaxRef'] == 0 or row['CD_Ref'] == 0
    ]

    # Récupérer les résultats et traiter les modifications
    for future in concurrent.futures.as_completed(futures):
        result = future.result()
        if result is not None:
            index, nom_tax, scientific_name, reference_id = result
            # Si la valeur de 'NomTaxRef' ou 'CD_Ref' a été modifiée, vous pouvez afficher un message ou traiter les résultats ici
            print(f"Modifications pour l'index {index}: {nom_tax} -> {scientific_name}, {reference_id}")

print(df.head(10))

#Compléter avec la BD_Statuts
# URL de base
url_base = 'https://taxref.mnhn.fr/api/status/search/lines?locationId=INSEEC29241&page=1&size=10000'

# Ajouter chaque taxrefId à l'URL
url_complete = url_base + ''.join([f'&taxrefId={elem}' for elem in CD_Ref])

# Afficher l'URL complète générée
print(url_complete)

# En-têtes de la requête
headers = {
    'accept': 'application/hal+json;version=1'
}

# Effectuer la requête GET
response = requests.get(url_complete, headers=headers)

# Vérifier si la requête a réussi (code 200)
if response.status_code == 200:
    # Afficher le contenu de la réponse (en JSON)
    data = response.json()

    # Extraire les données des éléments JSON
    status_data = data.get('_embedded', {}).get('status', [])

    # Créer une liste pour contenir les informations pertinentes
    extracted_data = []

    for item in status_data:
        # Extraire les informations de chaque item
        taxon = item.get('taxon', {})
        statusTypeName = item.get('statusTypeName', '')
        statusCode = item.get('statusCode', '')

        # Récupérer l'id et le nom scientifique
        taxon_id = taxon.get('id', '')
        scientificName = taxon.get('scientificName', '')

        # Ajouter les informations extraites à la liste
        extracted_data.append({
            'id': taxon_id,
            'scientificName': scientificName,
            'statusTypeName': statusTypeName,
            'statusCode': statusCode
        })

    # Créer un DataFrame pandas avec les données extraites
    df = pd.DataFrame(extracted_data)

    # Créer la table pivotée
    pivot_df = df.pivot_table(index=['id', 'scientificName'],
                              columns='statusTypeName',
                              values='statusCode',
                              aggfunc='first',  # 'first' car on s'attend à une seule valeur pour chaque combinaison
                              fill_value='')
# Ajuster pandas pour afficher plus de colonnes et de lignes
# pd.set_option('display.max_columns', None)  # Afficher toutes les colonnes
# pd.set_option('display.max_rows', None)  # Afficher toutes les lignes
# pd.set_option('display.width', None)  # Pas de limite sur la largeur d'affichage
# pd.set_option('display.max_colwidth', None)  # Pas de limite sur la largeur des colonnes

# # Afficher la table pivotée
    print(pivot_df)
else:
    print(f"Erreur : {response.status_code}")

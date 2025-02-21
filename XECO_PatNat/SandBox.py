import requests
import pandas as pd

# URL de l'API
url = 'https://taxref.mnhn.fr/api/status/search/lines?taxrefId=106807&taxrefId=94041&locationId=INSEEC35259&page=1&size=10000'
# En-têtes de la requête
headers = {
    'accept': 'application/hal+json;version=1'
}

# Effectuer la requête GET
response = requests.get(url, headers=headers)

# Vérifier si la requête a réussi (code 200)
if response.status_code == 200:
    # Afficher le contenu de la réponse (en JSON)
    data = response.json()

    # Extraire les données des éléments JSON (en supposant que les données sont sous une clé spécifique, comme 'content')
    # Vous pouvez ajuster selon la structure de la réponse JSON
    content = data.get('content', [])

    # Créer un DataFrame pandas à partir de la liste d'éléments
    df = pd.DataFrame(content)

    # Afficher le dataframe
    print(df)
else:
    print(f"Erreur : {response.status_code}")

# import requests
#
# # URL de l'API
# url = 'https://taxref.mnhn.fr/api/status/search/lines?taxrefId=106807&taxrefId=94041&locationId=INSEEC35259&page=1&size=10000'
#
# # En-têtes de la requête
# headers = {
#     'accept': 'application/hal+json;version=1'
# }
#
# # Effectuer la requête GET
# response = requests.get(url, headers=headers)
#
# # Vérifier si la requête a réussi (code 200)
# if response.status_code == 200:
#     # Récupérer la réponse JSON
#     data = response.json()
#
#     # Extraire et afficher les 'statusName' de chaque élément
#     status_names = [item['statusName'] for item in data['_embedded']['status']]
#
#     # Afficher les 'statusName'
#     print(status_names)
# else:
#     print(f"Erreur : {response.status_code}")

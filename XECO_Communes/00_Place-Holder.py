import requests
import pandas as pd

listesp = [79783,79908,80590,80591,80605,80759,80990,81263,81272,81295,81538,81541,81544,81569,81637,81656,81966,610909,107085,82637,82738,82922,82952,83160,101221,83272,83499,83502,83912,84061,84110,84458,84511,84524,84534,84999,85102,85208,85250,85740,85903,85904,85946,85957,85986,86101,86305,86489,86634,82757,86564,86869,87143,87466,87471,87484,87501,92353,87616,87915,87930,88463,88569,88720,88626,88753,88766,88775,127864,89304,89840,89888,90008,90017,90178,90669,133219,717294,100304,91120,91169,91258,91289,91382,91430,91886,91912,92106,92242,92302,96749,611690,96814,105615,92566,92606,92806,92864,92876,93023,93763,93860,94207,94266,94402,94489,94503,94626,94959,94985,94995,95547,95563,95567,115527,95671,95858,95889,95916,95922,96046,96149,96180,96191,96208,96229,96271,96519,96667,96775,609982,97434,97452,97537,97556,97609,97947,97962,717533,98681,98865,98887,98921,99334,99359,99373,100052,100104,100132,100136,100142,100144,100225,100310,100387,100519,101300,113508,113525,102900,102901,103031,103055,103057,103142,103245,103272,103288,103316,103320,103375,103514,103547,103734,103772,103843,104101,104126,104144,104145,104160,104173,104189,104353,104502,104775,104841,104855,104876,104903,105017,105247,105400,105431,121988,105521,105630,106213,137388,106419,106497,106499,106581,106696,106698,106747,106754,106818,106863,106918,107038,107072,107077,107090,107115,107117,107282,107446,127613,107574,107649,108027,108351,92854,108645,108698,108718,109042,109092,109104,109139,109422,109732,109861,109864,109969,111289,111419,111815,111876,111881,111886,112130,112355,112463,112550,112590,112669,126613,112975,101210,113547,113842,113893,113904,114114,114332,114416,112727,114658,112739,112741,112745,114972,115016,115110,115145,115245,115301,115470,92217,115624,115655,115925,116012,116043,116089,116142,116265,116744,116759,116903,116952,98651,117025,117145,117165,117201,117353,117692,117860,117944,118016,119418,119419,119471,119473,119550,119585,119780,119818,119948,120260,120717,115789,121201,103862,121960,121999,122028,122046,122115,122246,122630,610646,122726,122745,123141,123154,123164,123471,123522,123863,124034,124080,124147,124232,124233,124261,124499,124528,124744,85852,124798,124814,124967,125000,125006,125014,125295,611652,125816,126035,126859,127259,127294,127439,127454,128077,128114,128175,128215,128268,128345,128419,718832,128660,128754,128801,128832,128938,128956,129000,129003,97084,129468,129506,129639,129669,129723,129997]

# URL de base
url_base = 'https://taxref.mnhn.fr/api/status/search/lines?locationId=INSEEC29241&page=1&size=10000'

# Ajouter chaque taxrefId à l'URL
url_complete = url_base + ''.join([f'&taxrefId={elem}' for elem in listesp])

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
    pd.set_option('display.max_columns', None)  # Afficher toutes les colonnes
    pd.set_option('display.max_rows', None)  # Afficher toutes les lignes
    pd.set_option('display.width', None)  # Pas de limite sur la largeur d'affichage
    pd.set_option('display.max_colwidth', None)  # Pas de limite sur la largeur des colonnes
    # Afficher la table pivotée
    print(pivot_df)
else:
    print(f"Erreur : {response.status_code}")

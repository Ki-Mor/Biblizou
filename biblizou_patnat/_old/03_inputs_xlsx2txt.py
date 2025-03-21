import os
import sys
import pandas as pd

def obtain_folder_path():
    return sys.argv[1] if len(sys.argv) > 1 else input("Veuillez entrer le chemin complet du dossier contenant les fichiers Excel : ")

def generate_txt_from_column(excel_file, column_index, txt_file):
    # Lire le fichier Excel sans en-têtes
    df = pd.read_excel(excel_file, header=None)

    # Vérification si le DataFrame est vide
    if df.empty:
        print(f"Le fichier Excel {excel_file} est vide.")
        return False

    # Extraire la colonne spécifique
    column_data = df.iloc[:, column_index]

    # Vérification si la colonne est vide
    if column_data.isnull().all():
        print(f"La colonne {column_index + 1} dans {excel_file} est vide.")
        return False

    # Enregistrer les données dans un fichier texte (chaque valeur sur une ligne)
    with open(txt_file, 'w', encoding='utf-8') as f:
        for value in column_data:
            f.write(str(value) + '\n')

    print(f"Le fichier TXT a été enregistré sous : {txt_file}")
    return True

# Demander le dossier où les fichiers Excel se trouvent
folder_path = obtain_folder_path()

# Vérifier si le dossier existe
if not os.path.isdir(folder_path):
    print("Le dossier spécifié n'existe pas. Veuillez vérifier le chemin.")
else:
    # Fichiers Excel
    n2000_excel_file = os.path.join(folder_path, 'output_qgis_intersect_and_export_n2000.xlsx')
    znieff_excel_file = os.path.join(folder_path, 'output_qgis_intersect_and_export_znieff.xlsx')

    # Vérification que les fichiers existent
    if not os.path.exists(n2000_excel_file):
        print(f"Le fichier {n2000_excel_file} n'existe pas.")
    else:
        # Générer le fichier TXT pour N2000 si les données sont présentes
        output_n2000_txt = os.path.join(folder_path, "input_xml_n2000_download_list.txt")
        if not generate_txt_from_column(n2000_excel_file, 0, output_n2000_txt):  # On prend la première colonne
            print("Pas de zones Natura 2000 détectées")

    if not os.path.exists(znieff_excel_file):
        print(f"Le fichier {znieff_excel_file} n'existe pas.")
    else:
        # Générer le fichier TXT pour ZNIEFF si les données sont présentes
        output_znieff_txt = os.path.join(folder_path, "input_xml_znieff_download_list.txt")
        if not generate_txt_from_column(znieff_excel_file, 0, output_znieff_txt):  # On prend la première colonne
            print("Pas de ZNIEFF détectés")

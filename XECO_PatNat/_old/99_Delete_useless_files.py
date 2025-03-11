import os
import sys

# Fonction pour obtenir le chemin du dossier à partir des arguments de la ligne de commande ou via une entrée utilisateur
def obtain_folder_path():
    return sys.argv[1] if len(sys.argv) > 1 else input("Entrez le chemin du dossier contenant les fichiers XML : ")

# Fonction pour supprimer les fichiers spécifiés dans le répertoire
def delete_files_in_directory(folder_path):
    try:
        # Liste des fichiers à supprimer (fichiers .txt et .xml spécifiques)
        xml_files = [f for f in os.listdir(folder_path) if f.endswith('.xml')]  # Fichiers XML à supprimer
        txt_files = ['input_xml_znieff_download_list.txt', 'input_xml_n2000_download_list.txt']  # Liste des fichiers txt à supprimer

        # Boucle pour supprimer les fichiers XML
        for file in xml_files:
            file_path = os.path.join(folder_path, file)  # Créer le chemin complet du fichier
            print(f"Suppression du fichier XML : {file_path}")  # Affichage pour informer de la suppression
            os.remove(file_path)  # Suppression du fichier

        # Boucle pour supprimer les fichiers TXT spécifiques
        for file in txt_files:
            file_path = os.path.join(folder_path, file)  # Créer le chemin complet du fichier
            if os.path.exists(file_path):  # Vérifier si le fichier existe avant de le supprimer
                print(f"Suppression du fichier TXT : {file_path}")  # Affichage pour informer de la suppression
                os.remove(file_path)  # Suppression du fichier
            else:
                print(f"Le fichier {file} n'existe pas dans le dossier.")  # Message si le fichier n'est pas trouvé

    except Exception as e:
        print(f"Erreur lors de la suppression des fichiers inutiles : {e}")  # Afficher l'erreur en cas de problème

# Programme principal
if __name__ == "__main__":
    folder_path = obtain_folder_path()  # Obtenir le chemin du dossier
    delete_files_in_directory(folder_path)  # Supprimer les fichiers dans ce dossier

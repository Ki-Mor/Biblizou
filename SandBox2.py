"""
Auteur : ExEco Environnement - François Botcazou
Date de création : 2025/02
Dernière mise à jour : 2025/03
Version : 1.0
Nom : NaturaXmlToXlsxEsp.py
Groupe : Biblizou_PatNat
Description : Module pour extraire des données d'espèces directive à partir de fichiers XML et les exporter dans un fichier Excel.
Dépendances :
    - Python 3.x
    - QGIS (QgsMessageLog)
    - xml.etree.ElementTree
    - pandas
    - openpyxl
    - os, datetime

Utilisation :
    Ce module doit être appelé depuis une extension QGIS.
"""

from PyQt5.QtWidgets import QFileDialog
from qgis.core import QgsMessageLog, Qgis
import xml.etree.ElementTree as ElTr
import requests
import pandas as pd
import os
from datetime import datetime
import time

class NaturaXmlToXlsxEsp:
    def __init__(self, iface):
        self.iface = iface

    def run(self):
        folder_path = self.obtain_folder_path()
        if folder_path:
            self.process_xml_files_in_folder(folder_path)

    def obtain_folder_path(self):
        folder_path = QFileDialog.getExistingDirectory(None, "Sélectionner un dossier contenant les fichiers XML")
        if not folder_path:
            self.iface.messageBar().pushMessage("Annulation", "Aucun dossier sélectionné.", level=Qgis.Warning)
            return None
        return folder_path

    def truncate_sheet_name(self, sheet_name):
        return sheet_name[:31]  # Limiter à 31 caractères maximum (restriction d'Excel)

    def get_taxref_data(self, cd_nom, cache={}):
        # Vérifier si les données sont déjà en cache
        if cd_nom in cache:
            return cache[cd_nom]

        # URL de l'API avec cd_nom comme identifiant
        url = f"https://taxref.mnhn.fr/api/taxa/{cd_nom}"

        # Définir les headers pour l'API
        headers = {"accept": "application/hal+json;version=1"}

        try:
            # Faire la requête GET
            response = requests.get(url, headers=headers)

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
                    cache[cd_nom] = result
                    return result
                else:
                    QgsMessageLog.logMessage(f"Aucun taxon trouvé pour {cd_nom}.", "NaturaXmlToXlsxEsp", Qgis.Warning)
                    return {'REGNE': '', 'GROUPE': '', 'NOM_COMPLET': '', 'NOM_VERN': ''}
            else:
                QgsMessageLog.logMessage(f"Erreur API pour {cd_nom}: {response.status_code}", "NaturaXmlToXlsxEsp", Qgis.Warning)
                return {'REGNE': '', 'GROUPE': '', 'NOM_COMPLET': '', 'NOM_VERN': ''}
        except Exception as e:
            QgsMessageLog.logMessage(f"Erreur lors de l'interrogation de l'API pour {cd_nom}: {e}", "NaturaXmlToXlsxEsp", Qgis.Warning)
            return {'REGNE': '', 'GROUPE': '', 'NOM_COMPLET': '', 'NOM_VERN': ''}

    def xml_to_dataframe(self, xml_file, cache, regne_data):
        global taxon_info, cd_nom
        try:
            tree = ElTr.parse(xml_file)
            root = tree.getroot()

            regnes, groupes, cd_noms, noms, nom_complets, nom_vern, sitecode, site_name = [], [], [], [], [], [], "", ""

            start_time: float = time.time()

            # Recherche de l'élément BIOTOP
            for biotop_elem in root.iter('BIOTOP'):
                sitecode_elem = biotop_elem.find('SITECODE')
                site_name_elem = biotop_elem.find('SITE_NAME')

                # Vérification si les éléments sont bien trouvés et non vides
                if sitecode_elem is not None and sitecode_elem.text:
                    sitecode = sitecode_elem.text.strip()
                else:
                    sitecode = None  # ou 'no_sitecode' si tu veux une valeur par défaut

                if site_name_elem is not None and site_name_elem.text:
                    site_name = site_name_elem.text.strip()
                else:
                    site_name = None  # ou 'no_site_name' si tu veux une valeur par défaut

                for species_elem in biotop_elem.iter('SPECIES'):
                    for species_row_elem in species_elem.iter('SPECIES_ROW'):
                        nom_elem = species_row_elem.find('NOM')
                        cd_nom_elem = species_row_elem.find('CD_NOM')

                        if cd_nom_elem is not None:
                            cd_nom = cd_nom_elem.text
                            nom = nom_elem.text
                            taxon_info = self.get_taxref_data(cd_nom, cache)

                            regnes.append(taxon_info['REGNE'])
                            groupes.append(taxon_info['GROUPE'])
                            cd_noms.append(cd_nom_elem.text if cd_nom_elem is not None else "")
                            noms.append(nom)
                            nom_complets.append(taxon_info['NOM_COMPLET'])
                            nom_vern.append(taxon_info['NOM_VERN'])

                        regne = taxon_info['REGNE']  # Récupérer le règne
                        if regne:  # Vérifier que le règne n'est pas vide
                            if regne not in regne_data:
                                regne_data[
                                    regne] = []  # Initialiser une liste pour ce règne si elle n'existe pas encore

                            regne_data[regne].append({
                                "SITECODE - SITE_NAME": f"{sitecode} - {site_name}",
                                "GROUPE": taxon_info['GROUPE'],
                                "CD_NOM": cd_nom,
                                "NOM_COMPLET": taxon_info['NOM_COMPLET'],
                                "NOM_VERN": taxon_info['NOM_VERN']
                            })

            xml_processing_time = time.time() - start_time
            QgsMessageLog.logMessage(f"Fichier {xml_file} traité en {xml_processing_time:.2f} secondes.", "NaturaXmlToXlsxEsp", Qgis.Info)

            # Retourner le DataFrame avec les bonnes données
            data = {
                'REGNE': regnes,
                'GROUPE': groupes,
                'CD_NOM': cd_noms,
                'NOM': noms,
                'NOM_COMPLET': nom_complets,
                'NOM_VERN': nom_vern
            }

            # S'assurer que toutes les listes ont la même longueur
            max_len = max(len(regnes), len(groupes), len(cd_noms), len(noms), len(nom_complets), len(nom_vern))
            regnes += [""] * (max_len - len(regnes))
            groupes += [""] * (max_len - len(groupes))
            cd_noms += [""] * (max_len - len(cd_noms))
            noms += [""] * (max_len - len(noms))
            nom_complets += [""] * (max_len - len(nom_complets))
            nom_vern += [""] * (max_len - len(nom_vern))

            df = pd.DataFrame(data)
            return df, sitecode, site_name

        except ElTr.ParseError as e:
            QgsMessageLog.logMessage(f"Error parsing XML file {xml_file}: {e}", "NaturaXmlToXlsxEsp", Qgis.Critical)
            return pd.DataFrame(columns=['REGNE', 'GROUPE', 'CD_NOM', 'NOM', 'NOM_COMPLET', 'NOM_VERN']), "", ""
        except Exception as e:
            QgsMessageLog.logMessage(f"Unexpected error with file {xml_file}: {e}", "NaturaXmlToXlsxEsp", Qgis.Critical)
            return pd.DataFrame(columns=['REGNE', 'GROUPE', 'CD_NOM', 'NOM', 'NOM_COMPLET', 'NOM_VERN']), "", ""

    def process_xml_files_in_folder(self, folder_path):
        if not os.path.isdir(folder_path):
            QgsMessageLog.logMessage(f"The path {folder_path} is not a valid directory.", "NaturaXmlToXlsxEsp", Qgis.Warning)
            return

        xml_files = [f for f in os.listdir(folder_path) if f.startswith('FR') and f.endswith('.xml') and len(f) == 13]

        if not xml_files:
            QgsMessageLog.logMessage(f"No XML files found in the folder {folder_path}", "NaturaXmlToXlsxEsp", Qgis.Warning)
            return

        cache = {}

        # Initialisation des listes pour les synthèses

        current_time = datetime.now().strftime("%Y%m%d%H%M%S")
        excel_file = os.path.join(folder_path, f'N2000_Synthèse_des_espèces_AnxI-II_{current_time}.xlsx')

        try:
            QgsMessageLog.logMessage(f"\nDébut du traitement des fichiers XML...", "NaturaXmlToXlsxEsp", Qgis.Info)
            start_time = time.time()

            # Création du fichier Excel avec pandas.ExcelWriter() et xlsxwriter comme moteur
            with pd.ExcelWriter(excel_file, engine='xlsxwriter') as writer:
                workbook = writer.book
                header_format = workbook.add_format({'bold': True, 'bg_color': '#009999', 'font_color': '#FFFFFF', 'align': 'center'})
                normal_format = workbook.add_format({'font_size': 9, 'border': 1})
                center_format = workbook.add_format({'font_size': 9, 'border': 1, 'align': 'center'})
                highlight_format = workbook.add_format({'bg_color': '#91d2ff'})

                regne_data = {}  # Stockera les données pour chaque règne
                for idx, xml_file in enumerate(xml_files, start=1):
                    QgsMessageLog.logMessage(f"Traitement du fichier {idx}/{len(xml_files)} : {xml_file}...", "NaturaXmlToXlsxEsp", Qgis.Info)
                    full_path = os.path.join(folder_path, xml_file)
                    df, sitecode, site_name = self.xml_to_dataframe(full_path, cache, regne_data)

                    if not df.empty:
                        sheet_name = self.truncate_sheet_name(f"{sitecode}-{site_name}")  # Assurer un nom valide
                        df.to_excel(writer, sheet_name=sheet_name, index=False)

                        # Ajustement automatique de la largeur des colonnes
                        worksheet = writer.sheets[sheet_name]
                        for i, col in enumerate(df.columns):
                            max_length = max(df[col].astype(str).map(len).max(), len(col)) + 2
                            worksheet.set_column(i, i, max_length)

                        QgsMessageLog.logMessage(f"Feuille ajoutée : {sheet_name}", "NaturaXmlToXlsxEsp", Qgis.Info)

                # Fonction pour écrire les feuilles de synthèse
                def write_summary_sheet(sheet_name, data, writer):
                    QgsMessageLog.logMessage(f"Début de la création de {sheet_name}", "NaturaXmlToXlsxEsp", Qgis.Info)

                    if not data:
                        QgsMessageLog.logMessage(f"Aucune donnée pour {sheet_name}, feuille non créée.",
                                                 "NaturaXmlToXlsxEsp", Qgis.Warning)
                        return

                    # Création du DataFrame à partir des données
                    df_summary = pd.DataFrame(data)
                    QgsMessageLog.logMessage(f"Valeurs None détectées :\n{df_summary.isnull().sum()}",
                                             "NaturaXmlToXlsxEsp", Qgis.Warning)
                    # Vérifier les colonnes disponibles avant de tenter un pivot
                    QgsMessageLog.logMessage(
                        f"Colonnes disponibles dans df_summary ({sheet_name}) : {df_summary.columns.tolist()}",
                        "NaturaXmlToXlsxEsp", Qgis.Info)
                    QgsMessageLog.logMessage(f"Types de df_summary :\n{df_summary.dtypes}", "NaturaXmlToXlsxEsp",
                                             Qgis.Warning)
                    for i, val in enumerate(df_summary["NOM_VERN"]):
                        if val is None:
                            QgsMessageLog.logMessage(f"⚠️ Ligne {i}: Valeur None détectée dans NOM_VERN",
                                                     "NaturaXmlToXlsxEsp", Qgis.Warning)
                        elif pd.isna(val):
                            QgsMessageLog.logMessage(f"⚠️ Ligne {i}: Valeur NaN détectée dans NOM_VERN",
                                                     "NaturaXmlToXlsxEsp", Qgis.Warning)
                        elif not isinstance(val, str):
                            QgsMessageLog.logMessage(f"⚠️ Ligne {i}: Type {type(val)} inattendu dans NOM_VERN → {val}",
                                                     "NaturaXmlToXlsxEsp", Qgis.Warning)

                    # Vérification de l'existence des colonnes nécessaires
                    required_columns = ["SITECODE - SITE_NAME", "GROUPE", "CD_NOM", "NOM_COMPLET", "NOM_VERN"]
                    for col in required_columns:
                        if col not in df_summary.columns:
                            QgsMessageLog.logMessage(f"⚠️ Colonne manquante : {col}", "NaturaXmlToXlsxEsp",
                                                     Qgis.Critical)
                            return
                    QgsMessageLog.logMessage(f"🔍 Aperçu de df_summary :\n{df_summary.head()}", "NaturaXmlToXlsxEsp",
                                             Qgis.Info)

                    # Création du tableau croisé dynamique
                    try:
                        df_summary["SITECODE - SITE_NAME"] = df_summary["SITECODE - SITE_NAME"].astype(str)
                        df_summary["NOM_VERN"] = df_summary["NOM_VERN"].apply(lambda x: "" if pd.isna(x) else x)

                        # Remplacer les NaN/None par une chaîne vide
                        df_summary = df_summary.fillna("")

                        # Convertir toutes les colonnes en chaînes de caractères (évite les erreurs de concaténation)
                        df_summary = df_summary.astype(str)

                        pivot_df = df_summary.pivot_table(
                            index=["GROUPE", "CD_NOM", "NOM_COMPLET", "NOM_VERN"],
                            columns="SITECODE - SITE_NAME",
                            values="CD_NOM",  # Une valeur arbitraire pour éviter la concaténation accidentelle des clés
                            aggfunc=lambda x: "X",
                            fill_value=""
                        )

                        # Réinitialiser les index pour éviter la concaténation automatique
                        pivot_df.reset_index(inplace=True)

                    except Exception as e:
                        QgsMessageLog.logMessage(f"❌ Erreur lors de la création du pivot table ({sheet_name}): {e}",
                                                 "NaturaXmlToXlsxEsp", Qgis.Critical)
                        return

                    # Tronquer le nom de la feuille pour respecter la limite Excel (31 caractères max)
                    sheet_name = sheet_name[:31]
                    pivot_df.to_excel(writer, sheet_name=sheet_name)

                    # Ajustement automatique de la largeur des colonnes
                    worksheet = writer.sheets[sheet_name]
                    for i, col in enumerate(pivot_df.columns.insert(0, "".join(pivot_df.index.names))):
                        max_length = max(pivot_df[col].astype(str).map(len).max(), len(col)) + 2
                        worksheet.set_column(i, i, max_length)

                    QgsMessageLog.logMessage(f"✅ Feuille de synthèse {sheet_name} créée avec succès.",
                                             "NaturaXmlToXlsxEsp", Qgis.Info)

                # Ajout des feuilles de synthèse
                for regne, data in regne_data.items():
                    sheet_name = f"Synthèse {regne}"
                    write_summary_sheet(sheet_name, data, writer)

            processing_time = time.time() - start_time
            QgsMessageLog.logMessage(f"\nTraitement terminé en {processing_time:.2f} secondes.", "NaturaXmlToXlsxEsp", Qgis.Info)

        except Exception as e:
            QgsMessageLog.logMessage(f"Erreur lors du traitement des fichiers XML: {e}", "NaturaXmlToXlsxEsp", Qgis.Critical)

# Pour exécuter le module dans QGIS
def run_module(iface):
        module = NaturaXmlToXlsxEsp(iface)
        module.run()

run_module(iface)
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
import xml.etree.ElementTree as ET
import requests
import pandas as pd
import os
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Side, PatternFill, Font
import time
from collections import defaultdict


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
        return sheet_name[:31]

    def get_taxref_data(self, cd_nom, cache={}):
        if cd_nom in cache:
            return cache[cd_nom]

        url = f"https://taxref.mnhn.fr/api/taxa/{cd_nom}"
        headers = {"accept": "application/hal+json;version=1"}

        try:
            response = requests.get(url, headers=headers)
            if response.status_code == 200:
                data = response.json()
                result = {
                    'REGNE': data.get('kingdomName', ''),
                    'GROUPE': data.get('vernacularGroup2', ''),
                    'NOM_COMPLET': data.get('fullName', ''),
                    'NOM_VERN': data.get('frenchVernacularName', '')
                }
                cache[cd_nom] = result
                return result
            else:
                QgsMessageLog.logMessage(f"Erreur API pour {cd_nom}: {response.status_code}", "Biblizou", Qgis.Warning)
                return {'REGNE': '', 'GROUPE': '', 'NOM_COMPLET': '', 'NOM_VERN': ''}
        except Exception as e:
            QgsMessageLog.logMessage(f"Erreur API: {e}", "Biblizou", Qgis.Critical)
            return {'REGNE': '', 'GROUPE': '', 'NOM_COMPLET': '', 'NOM_VERN': ''}

    def process_xml_files_in_folder(self, folder_path):
        if not os.path.isdir(folder_path):
            self.iface.messageBar().pushMessage("Erreur", f"Le chemin {folder_path} n'est pas un dossier valide.",
                                                level=Qgis.Critical)
            return

        xml_files = [f for f in os.listdir(folder_path) if f.startswith('FR') and f.endswith('.xml') and len(f) == 13]
        if not xml_files:
            self.iface.messageBar().pushMessage("Information", "Aucun fichier XML trouvé dans le dossier.",
                                                level=Qgis.Info)
            return

        cache = {}
        current_time = datetime.now().strftime("%Y%m%d%H%M%S")
        excel_file = os.path.join(folder_path, f'N2000_Synthèse_des_espèces_AnxI-II_{current_time}.xlsx')

        try:
            start_time = time.time()
            with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
                for xml_file in xml_files:
                    full_path = os.path.join(folder_path, xml_file)
                    df, sitecode, site_name = self.xml_to_dataframe(full_path, cache)

                    if not df.empty:
                        sheet_name = self.truncate_sheet_name(f"{sitecode}-{site_name}")
                        df.drop(columns=['NOM'], inplace=True)
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
                        QgsMessageLog.logMessage(f"Fichier traité: {xml_file} ajouté sous {sheet_name}.", "Biblizou",
                                                 Qgis.Info)

            processing_time = time.time() - start_time
            self.iface.messageBar().pushMessage("Succès", f"Traitement terminé en {processing_time:.2f} secondes.",
                                                level=Qgis.Success)
        except Exception as e:
            self.iface.messageBar().pushMessage("Erreur", f"Problème lors du traitement: {e}", level=Qgis.Critical)

    def xml_to_dataframe(self, xml_file, cache):
        try:
            tree = ET.parse(xml_file)
            root = tree.getroot()
            regnes, groupes, cd_noms, noms, nom_complets, nom_vern, sitecode, site_name = [], [], [], [], [], [], "", ""

            for biotop_elem in root.iter('BIOTOP'):
                sitecode_elem = biotop_elem.find('SITECODE')
                site_name_elem = biotop_elem.find('SITE_NAME')
                sitecode = sitecode_elem.text.strip() if sitecode_elem is not None and sitecode_elem.text else ""
                site_name = site_name_elem.text.strip() if site_name_elem is not None and site_name_elem.text else ""

                for species_elem in biotop_elem.iter('SPECIES'):
                    for species_row_elem in species_elem.iter('SPECIES_ROW'):
                        nom_elem = species_row_elem.find('NOM')
                        cd_nom_elem = species_row_elem.find('CD_NOM')
                        if cd_nom_elem is not None:
                            cd_nom = cd_nom_elem.text
                            nom = nom_elem.text if nom_elem is not None else ""
                            taxon_info = self.get_taxref_data(cd_nom, cache)
                            regnes.append(taxon_info['REGNE'])
                            groupes.append(taxon_info['GROUPE'])
                            cd_noms.append(cd_nom)
                            noms.append(nom)
                            nom_complets.append(taxon_info['NOM_COMPLET'])
                            nom_vern.append(taxon_info['NOM_VERN'])

            data = {'REGNE': regnes, 'GROUPE': groupes, 'CD_NOM': cd_noms, 'NOM': noms, 'NOM_COMPLET': nom_complets,
                    'NOM_VERN': nom_vern}
            df = pd.DataFrame(data)
            return df, sitecode, site_name
        except ET.ParseError as e:
            QgsMessageLog.logMessage(f"Erreur parsing XML: {e}", "Biblizou", Qgis.Critical)
            return pd.DataFrame(columns=['REGNE', 'GROUPE', 'CD_NOM', 'NOM', 'NOM_COMPLET', 'NOM_VERN']), "", ""
        except Exception as e:
            QgsMessageLog.logMessage(f"Erreur inattendue: {e}", "Biblizou", Qgis.Critical)
            return pd.DataFrame(columns=['REGNE', 'GROUPE', 'CD_NOM', 'NOM', 'NOM_COMPLET', 'NOM_VERN']), "", ""

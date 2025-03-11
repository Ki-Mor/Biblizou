"""
Auteur : ExEco Environnement - François Botcazou
Date de création : 2025/02
Dernière mise à jour : 2025/03
Version : 1.0
Nom : ZnieffXmlToXlsxEsp.py
Groupe : Biblizou_PatNat
Description : Module pour extraire des données d'espèces déterminantes à partir de fichiers XML et les exporter dans un fichier Excel.
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

from qgis.core import QgsMessageLog, Qgis
from qgis.gui import QgsMessageBar
from PyQt5.QtWidgets import QFileDialog
import xml.etree.ElementTree as ET
import pandas as pd
import os
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Side, PatternFill, Font
from datetime import datetime
import time


class ZnieffXmlToXlsxEsp:
    def __init__(self, iface):
        """
        Module d'extraction des espèces déterminantes à partir de fichiers XML ZNIEFF et exportation vers Excel.
        """
        self.iface = iface

    def log(self, message, level=Qgis.Info):
        QgsMessageLog.logMessage(message, 'Biblizou_PatNat', level)

    def obtain_folder_path(self):
        folder_path = QFileDialog.getExistingDirectory(None, "Sélectionner un dossier contenant les fichiers XML")
        if not folder_path:
            self.log("Aucun dossier sélectionné.", Qgis.Warning)
        return folder_path

    def truncate_sheet_name(self, sheet_name):
        return sheet_name[:31]

    def xml_to_dataframe(self, xml_file):
        try:
            tree = ET.parse(xml_file)
            root = tree.getroot()
            regnes, groupes, cd_noms, nom_complets, nom_vern = [], [], [], [], []
            lb_zn, nm_sffzn = "", ""

            for znieff_elem in root.iter('ZNIEFF'):
                nm_sffzn_elem = znieff_elem.find('NM_SFFZN')
                lb_zn_elem = znieff_elem.find('LB_ZN')

                nm_sffzn = nm_sffzn_elem.text if nm_sffzn_elem is not None else ""
                lb_zn = lb_zn_elem.text if lb_zn_elem is not None else ""

                for espece_row_elem in znieff_elem.iter('ESPECE_ROW'):
                    fg_esp_elem = espece_row_elem.find('FG_ESP')
                    if fg_esp_elem is not None and fg_esp_elem.text == 'D':
                        regnes.append(espece_row_elem.findtext('REGNE', ""))
                        groupes.append(espece_row_elem.findtext('GROUPE', ""))
                        cd_noms.append(espece_row_elem.findtext('CD_NOM', ""))
                        nom_complets.append(espece_row_elem.findtext('NOM_COMPLET', ""))
                        nom_vern.append(espece_row_elem.findtext('NOM_VERN', ""))

            data = {'REGNE': regnes, 'GROUPE': groupes, 'CD_NOM': cd_noms, 'NOM_COMPLET': nom_complets,
                    'NOM_VERN': nom_vern}
            df = pd.DataFrame(data)
            return df, lb_zn, nm_sffzn
        except ET.ParseError as e:
            self.log(f"Erreur de parsing XML : {e}", Qgis.Critical)
            return pd.DataFrame(), "", ""
        except Exception as e:
            self.log(f"Erreur inattendue : {e}", Qgis.Critical)
            return pd.DataFrame(), "", ""

    def process_xml_files_in_folder(self, folder_path):
        if not os.path.isdir(folder_path):
            self.log(f"Le chemin {folder_path} n'est pas un répertoire valide.", Qgis.Warning)
            return

        xml_files = [f for f in os.listdir(folder_path) if f.endswith('.xml')]
        nm_sffzn_data, animalia_data, plantae_data = {}, [], []
        current_time = datetime.now().strftime("%Y%m%d%H%M%S")
        excel_file = os.path.join(folder_path, f'ZNIEFF_synthèse_des_esp_déterminantes_{current_time}.xlsx')

        try:
            with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
                for xml_file in xml_files:
                    df, lb_zn, nm_sffzn = self.xml_to_dataframe(os.path.join(folder_path, xml_file))
                    if not df.empty:
                        nm_sffzn_data.setdefault(nm_sffzn, []).append(df)
                        animalia_data.extend(df[df['REGNE'] == "Animalia"].to_dict('records'))
                        plantae_data.extend(df[df['REGNE'] == "Plantae"].to_dict('records'))

                for nm_sffzn, dfs in nm_sffzn_data.items():
                    combined_df = pd.concat(dfs, ignore_index=True).sort_values(by=['GROUPE', 'NOM_COMPLET'])
                    sheet_name = self.truncate_sheet_name(f"{nm_sffzn} - {lb_zn}")
                    combined_df.to_excel(writer, sheet_name=sheet_name, index=False)

                if animalia_data:
                    pd.DataFrame(animalia_data).to_excel(writer, sheet_name="Synthèse Animalia", index=False)
                if plantae_data:
                    pd.DataFrame(plantae_data).to_excel(writer, sheet_name="Synthèse Plantae", index=False)

            self.log(f"Exportation terminée : {excel_file}")
        except Exception as e:
            self.log(f"Erreur d'écriture dans le fichier Excel : {e}", Qgis.Critical)

    def run(self):
        folder_path = self.obtain_folder_path()
        if folder_path:
            self.process_xml_files_in_folder(folder_path)

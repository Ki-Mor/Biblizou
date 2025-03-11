"""
Auteur : ExEco Environnement - François Botcazou
Date de création : 2025/02
Dernière mise à jour : 2025/03
Version : 1.0
Nom : NaturaXmlToXlsxHab.py
Groupe : Biblizou_PatNat
Description : Module pour extraire des données d'habitats directive à partir de fichiers XML et les exporter dans un fichier Excel.
Dépendances :
    - Python 3.x
    - QGIS (QgsMessageLog, QgsMessageBar)
    - xml.etree.ElementTree
    - pandas
    - openpyxl
    - os, datetime

Utilisation :
    Ce module doit être appelé depuis une extension QGIS.
"""

import xml.etree.ElementTree as ET
import pandas as pd
import os
from openpyxl.styles import Alignment, Border, Side, PatternFill, Font
from openpyxl import load_workbook
from datetime import datetime
from qgis.PyQt.QtWidgets import QFileDialog, QMessageBox
from qgis.utils import iface
from qgis.core import QgsMessageLog, Qgis

class NaturaXmlToXlsxHab:
    def __init__(self):
        self.folder_path = ""

    def run(self):
        """ Exécute le module en demandant un dossier et en traitant les fichiers XML."""
        self.folder_path = QFileDialog.getExistingDirectory(None, "Sélectionner le dossier contenant les fichiers XML")
        if not self.folder_path:
            iface.messageBar().pushMessage("Annulation", "Aucun dossier sélectionné.", level=Qgis.Warning, duration=5)
            return

        self.process_xml_files_in_folder()

    def truncate_sheet_name(self, sheet_name):
        """ Tronque le nom de la feuille à 31 caractères, limite d'Excel. """
        return sheet_name[:31]

    def xml_to_dataframe(self, xml_file):
        """ Analyse un fichier XML et retourne un DataFrame avec les données extraites."""
        try:
            tree = ET.parse(xml_file)
            root = tree.getroot()
            cd_habs, lb_habdh_frs = [], []
            site_name, sitecode = "", ""

            for n2000_elem in root.iter('BIOTOP'):
                sitecode_elem = n2000_elem.find('SITECODE')
                site_name_elem = n2000_elem.find('SITE_NAME')
                if sitecode_elem is not None:
                    sitecode = sitecode_elem.text
                if site_name_elem is not None:
                    site_name = site_name_elem.text
                for typo_info_row_elem in n2000_elem.iter('HABIT1_ROW'):
                    cd_hab_elem = typo_info_row_elem.find('CD_UE')
                    lb_habdh_fr_elem = typo_info_row_elem.find('LB_HABDH_FR')
                    if cd_hab_elem is not None and lb_habdh_fr_elem is not None:
                        cd_habs.append(cd_hab_elem.text)
                        lb_habdh_frs.append(lb_habdh_fr_elem.text)

            df = pd.DataFrame({'CD_UE': cd_habs, 'LB_HABDH_FR': lb_habdh_frs})
            return df, site_name, sitecode
        except ET.ParseError as e:
            QgsMessageLog.logMessage(f"Erreur d'analyse XML : {e}", "Biblizou_PatNat", Qgis.Warning)
            return pd.DataFrame(columns=['CD_UE', 'LB_HABDH_FR']), "", ""
        except Exception as e:
            QgsMessageLog.logMessage(f"Erreur inattendue : {e}", "Biblizou_PatNat", Qgis.Critical)
            return pd.DataFrame(columns=['CD_UE', 'LB_HABDH_FR']), "", ""

    def process_xml_files_in_folder(self):
        """ Traite tous les fichiers XML du dossier et génère un fichier Excel."""
        if not os.path.isdir(self.folder_path):
            iface.messageBar().pushMessage("Erreur", "Chemin de dossier invalide.", level=Qgis.Critical, duration=5)
            return

        xml_files = [f for f in os.listdir(self.folder_path) if f.endswith('.xml') and f.startswith('FR') and len(f) == 13]
        if not xml_files:
            iface.messageBar().pushMessage("Information", "Aucun fichier XML trouvé.", level=Qgis.Info, duration=5)
            return

        unique_cd_ues, unique_lb_habdh_frs = set(), set()
        hab_presence = {}
        current_time = datetime.now().strftime("%Y%m%d%H%M%S")
        excel_file = os.path.join(self.folder_path, f'N2000_Synthèse_habitats_{current_time}.xlsx')

        try:
            with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
                for xml_file in xml_files:
                    full_path = os.path.join(self.folder_path, xml_file)
                    df, site_name, sitecode = self.xml_to_dataframe(full_path)
                    if not df.empty:
                        unique_cd_ues.update(df['CD_UE'].unique())
                        unique_lb_habdh_frs.update(df['LB_HABDH_FR'].unique())
                        sheet_name = self.truncate_sheet_name(f"{sitecode} - {site_name}")
                        hab_presence[sheet_name] = set(df['LB_HABDH_FR'])
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
                summary_data = {'CD_UE': list(unique_cd_ues), 'LB_HABDH_FR': list(unique_lb_habdh_frs)}
                for sheet_name in hab_presence:
                    summary_data[sheet_name] = ['X' if hab in hab_presence[sheet_name] else '' for hab in unique_lb_habdh_frs]
                summary_df = pd.DataFrame(summary_data)
                summary_df.to_excel(writer, sheet_name='Synthèse', index=False)
            iface.messageBar().pushMessage("Succès", f"Fichier Excel généré : {excel_file}", level=Qgis.Success, duration=10)
        except Exception as e:
            iface.messageBar().pushMessage("Erreur", f"Impossible d'écrire le fichier Excel : {e}", level=Qgis.Critical, duration=10)

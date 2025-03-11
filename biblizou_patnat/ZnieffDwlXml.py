from qgis.core import (
    QgsProject,
    QgsVectorLayer,
    QgsFeatureRequest
)

# Récupérer les couches Patrinat
patrinat_sic = QgsProject.instance().mapLayersByName("Patrinat : SIC")[0]
patrinat_zps = QgsProject.instance().mapLayersByName("Patrinat : ZPS")[0]

# Récupérer la couche de référence
ae_eloignee = QgsProject.instance().mapLayersByName("AE_eloignee")[0]

# Listes pour stocker les id_mnhn des entités sélectionnées
id_mnhn_zn2 = []
id_mnhn_zn1 = []


# Fonction de sélection et récupération des id_mnhn
def selectionner_et_stocker(couche_source, liste_stockage):
    if couche_source and ae_eloignee:
        geom_ref = [f.geometry() for f in ae_eloignee.getFeatures()]
        couche_source.removeSelection()
        ids_selectionnes = []

        for feature in couche_source.getFeatures():
            if any(feature.geometry().intersects(g) for g in geom_ref):
                ids_selectionnes.append(feature.id())
                liste_stockage.append(feature["id_mnhn"])

        if ids_selectionnes:
            couche_source.selectByIds(ids_selectionnes)
            print(f"{len(ids_selectionnes)} entités sélectionnées dans {couche_source.name()}.")
        else:
            print(f"Aucune entité sélectionnée dans {couche_source.name()}.")
    else:
        print(f"La couche {couche_source.name()} ou AE_eloignee est introuvable.")


# Appliquer la sélection sur chaque couche Patrinat
selectionner_et_stocker(patrinat_zn2, id_mnhn_zn2)
selectionner_et_stocker(patrinat_zn1, id_mnhn_zn1)

# Afficher les résultats
print("ID MNHN sélectionnés :")
print("ZNIEFF2 :", id_mnhn_zn2)
print("ZNIEFF1 :", id_mnhn_zn1)

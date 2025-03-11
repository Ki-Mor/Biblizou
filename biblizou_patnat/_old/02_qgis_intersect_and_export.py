o

OtherLayer = QgsProject.instance().mapLayersByName("linear")[0]
AreaLayer = QgsProject.instance().mapLayersByName("parcelle_finale")[0]

for feature in intersect_layer.getFeatures():
    sql_expression = f''' fid = {feature.id()} '''
    intersect_layer.setSubsetString(sql_expression)  # This will filter out the features in the `intersect_layer`
    processing.run("native:selectbylocation",
    {'INPUT': input_layer,
    'PREDICATE':[0],
    'INTERSECT': intersect_layer,
    'METHOD':0})




"""
Author : ExEco Environnement
Edition date : 2025/02
Name : 02_qgis_intersect_and_export
Group : Biblio_PatNat
"""

from qgis.core import QgsProcessing
from qgis.core import QgsProcessingAlgorithm
from qgis.core import QgsProcessingMultiStepFeedback
from qgis.core import QgsProcessingParameterVectorLayer
from qgis.core import QgsProcessingParameterFileDestination
from qgis.core import QgsProcessingFeatureSourceDefinition
import processing


class Modle(QgsProcessingAlgorithm):

    def initAlgorithm(self, config=None):
        self.addParameter(QgsProcessingParameterVectorLayer('aire_etude_eloignee', 'Aire_etude_eloignee', types=[QgsProcessing.TypeVectorPolygon], defaultValue=None))
        self.addParameter(QgsProcessingParameterFileDestination('Output_qgis_intersect_and_export_znieff', 'output_qgis_intersect_and_export_znieff', fileFilter='GeoPackage (*.gpkg *.GPKG);;ESRI Shapefile (*.shp *.SHP);;AutoCAD DXF (*.dxf *.DXF);;FlatGeobuf (*.fgb *.FGB);;Geoconcept (*.gxt *.txt *.GXT *.TXT);;Geography Markup Language [GML] (*.gml *.GML);;GeoJSON - Newline Delimited (*.geojsonl *.geojsons *.json *.GEOJSONL *.GEOJSONS *.JSON);;GeoJSON (*.geojson *.GEOJSON);;GeoRSS (*.xml *.XML);;GPS eXchange Format [GPX] (*.gpx *.GPX);;INTERLIS 1 (*.itf *.xml *.ili *.ITF *.XML *.ILI);;INTERLIS 2 (*.xtf *.xml *.ili *.XTF *.XML *.ILI);;Keyhole Markup Language [KML] (*.kml *.KML);;Mapinfo TAB (*.tab *.TAB);;Microstation DGN (*.dgn *.DGN);;PostgreSQL SQL dump (*.sql *.SQL);;S-57 Base file (*.000 *.000);;SQLite (*.sqlite *.SQLITE);;Tableur MS Office Open XML [XLSX] (*.xlsx *.XLSX);;Tableur Open Document  [ODS] (*.ods *.ODS);;Valeurs séparées par une virgule [CSV] (*.csv *.CSV)', createByDefault=True, defaultValue=''))
        self.addParameter(QgsProcessingParameterFileDestination('Output_qgis_intersect_and_export_n2000', 'output_qgis_intersect_and_export_n2000', fileFilter='GeoPackage (*.gpkg *.GPKG);;ESRI Shapefile (*.shp *.SHP);;AutoCAD DXF (*.dxf *.DXF);;FlatGeobuf (*.fgb *.FGB);;Geoconcept (*.gxt *.txt *.GXT *.TXT);;Geography Markup Language [GML] (*.gml *.GML);;GeoJSON - Newline Delimited (*.geojsonl *.geojsons *.json *.GEOJSONL *.GEOJSONS *.JSON);;GeoJSON (*.geojson *.GEOJSON);;GeoRSS (*.xml *.XML);;GPS eXchange Format [GPX] (*.gpx *.GPX);;INTERLIS 1 (*.itf *.xml *.ili *.ITF *.XML *.ILI);;INTERLIS 2 (*.xtf *.xml *.ili *.XTF *.XML *.ILI);;Keyhole Markup Language [KML] (*.kml *.KML);;Mapinfo TAB (*.tab *.TAB);;Microstation DGN (*.dgn *.DGN);;PostgreSQL SQL dump (*.sql *.SQL);;S-57 Base file (*.000 *.000);;SQLite (*.sqlite *.SQLITE);;Tableur MS Office Open XML [XLSX] (*.xlsx *.XLSX);;Tableur Open Document  [ODS] (*.ods *.ODS);;Valeurs séparées par une virgule [CSV] (*.csv *.CSV)', createByDefault=True, defaultValue=''))

    def processAlgorithm(self, parameters, context, model_feedback):
        # Use a multi-step feedback, so that individual child algorithm progress reports are adjusted for the
        # overall progress through the model
        feedback = QgsProcessingMultiStepFeedback(8, model_feedback)
        results = {}
        outputs = {}

        # Intersection (ZNIEFF1)
        alg_params = {
            'INPUT': QgsProcessingFeatureSourceDefinition('ms_Znieff1_a13a77d0_aa69_4cf1_9f6f_0aead1c6c104', selectedFeaturesOnly=False, featureLimit=-1, flags=QgsProcessingFeatureSourceDefinition.FlagOverrideDefaultGeometryCheck, geometryCheck=QgsFeatureRequest.GeometrySkipInvalid),
            'INPUT_FIELDS': ['ID_MNHN'],
            'OVERLAY': parameters['aire_etude_eloignee'],
            'OVERLAY_FIELDS': ['ISFAKE'],
            'OVERLAY_FIELDS_PREFIX': '',
            'OUTPUT': QgsProcessing.TEMPORARY_OUTPUT
        }
        outputs['IntersectionZnieff1'] = processing.run('native:intersection', alg_params, context=context, feedback=feedback, is_child_algorithm=True)

        feedback.setCurrentStep(1)
        if feedback.isCanceled():
            return {}

        # Intersection (ZNIEFF2)
        alg_params = {
            'INPUT': QgsProcessingFeatureSourceDefinition('ms_Znieff2_3e15138b_a40e_43a9_8dfd_6f55f9dfb10e', selectedFeaturesOnly=False, featureLimit=-1, flags=QgsProcessingFeatureSourceDefinition.FlagOverrideDefaultGeometryCheck, geometryCheck=QgsFeatureRequest.GeometrySkipInvalid),
            'INPUT_FIELDS': ['ID_MNHN'],
            'OVERLAY': parameters['aire_etude_eloignee'],
            'OVERLAY_FIELDS': ['ISFAKE'],
            'OVERLAY_FIELDS_PREFIX': '',
            'OUTPUT': QgsProcessing.TEMPORARY_OUTPUT
        }
        outputs['IntersectionZnieff2'] = processing.run('native:intersection', alg_params, context=context, feedback=feedback, is_child_algorithm=True)

        feedback.setCurrentStep(2)
        if feedback.isCanceled():
            return {}

        # Intersection (SIC)
        alg_params = {
            'INPUT': QgsProcessingFeatureSourceDefinition('ms_Sites_d_importance_communautaire_f5e5a100_42b2_4950_807b_34740bf1045c', selectedFeaturesOnly=False, featureLimit=-1, flags=QgsProcessingFeatureSourceDefinition.FlagOverrideDefaultGeometryCheck, geometryCheck=QgsFeatureRequest.GeometrySkipInvalid),
            'INPUT_FIELDS': ['SITECODE'],
            'OVERLAY': parameters['aire_etude_eloignee'],
            'OVERLAY_FIELDS': ['ISFAKE'],
            'OVERLAY_FIELDS_PREFIX': '',
            'OUTPUT': QgsProcessing.TEMPORARY_OUTPUT
        }
        outputs['IntersectionSic'] = processing.run('native:intersection', alg_params, context=context, feedback=feedback, is_child_algorithm=True)

        feedback.setCurrentStep(3)
        if feedback.isCanceled():
            return {}

        # Intersection (ZPS)
        alg_params = {
            'INPUT': QgsProcessingFeatureSourceDefinition('ms_Zones_de_protection_speciale_33a7516d_44d9_4b47_ae09_86a3c261c6e6', selectedFeaturesOnly=False, featureLimit=-1, flags=QgsProcessingFeatureSourceDefinition.FlagOverrideDefaultGeometryCheck, geometryCheck=QgsFeatureRequest.GeometrySkipInvalid),
            'INPUT_FIELDS': ['SITECODE'],
            'OVERLAY': parameters['aire_etude_eloignee'],
            'OVERLAY_FIELDS': ['ISFAKE'],
            'OVERLAY_FIELDS_PREFIX': '',
            'OUTPUT': QgsProcessing.TEMPORARY_OUTPUT
        }
        outputs['IntersectionZps'] = processing.run('native:intersection', alg_params, context=context, feedback=feedback, is_child_algorithm=True)

        feedback.setCurrentStep(4)
        if feedback.isCanceled():
            return {}

        # Fusionner des couches vecteur (N2000)
        alg_params = {
            'CRS': None,
            'LAYERS': [outputs['IntersectionSic']['OUTPUT'],outputs['IntersectionZps']['OUTPUT']],
            'OUTPUT': QgsProcessing.TEMPORARY_OUTPUT
        }
        outputs['FusionnerDesCouchesVecteurN2000'] = processing.run('native:mergevectorlayers', alg_params, context=context, feedback=feedback, is_child_algorithm=True)

        feedback.setCurrentStep(5)
        if feedback.isCanceled():
            return {}

        # Fusionner des couches vecteur (ZNIEFF)
        alg_params = {
            'CRS': None,
            'LAYERS': [outputs['IntersectionZnieff1']['OUTPUT'],outputs['IntersectionZnieff2']['OUTPUT']],
            'OUTPUT': QgsProcessing.TEMPORARY_OUTPUT
        }
        outputs['FusionnerDesCouchesVecteurZnieff'] = processing.run('native:mergevectorlayers', alg_params, context=context, feedback=feedback, is_child_algorithm=True)

        feedback.setCurrentStep(6)
        if feedback.isCanceled():
            return {}

        # Sauvegarder les entités vectorielles dans un fichier (N2000)
        alg_params = {
            'DATASOURCE_OPTIONS': '',
            'INPUT': outputs['FusionnerDesCouchesVecteurN2000']['OUTPUT'],
            'LAYER_NAME': '',
            'LAYER_OPTIONS': '',
            'OUTPUT': parameters['Output_qgis_intersect_and_export_n2000']
        }
        outputs['SauvegarderLesEntitsVectoriellesDansUnFichierN2000'] = processing.run('native:savefeatures', alg_params, context=context, feedback=feedback, is_child_algorithm=True)
        results['Output_qgis_intersect_and_export_n2000'] = outputs['SauvegarderLesEntitsVectoriellesDansUnFichierN2000']['OUTPUT']

        feedback.setCurrentStep(7)
        if feedback.isCanceled():
            return {}

        # Sauvegarder les entités vectorielles dans un fichier (ZNIEFF)
        alg_params = {
            'DATASOURCE_OPTIONS': '',
            'INPUT': outputs['FusionnerDesCouchesVecteurZnieff']['OUTPUT'],
            'LAYER_NAME': '',
            'LAYER_OPTIONS': '',
            'OUTPUT': parameters['Output_qgis_intersect_and_export_znieff']
        }
        outputs['SauvegarderLesEntitsVectoriellesDansUnFichierZnieff'] = processing.run('native:savefeatures', alg_params, context=context, feedback=feedback, is_child_algorithm=True)
        results['Output_qgis_intersect_and_export_znieff'] = outputs['SauvegarderLesEntitsVectoriellesDansUnFichierZnieff']['OUTPUT']
        return results

    def name(self):
        return 'Modèle'

    def displayName(self):
        return 'Modèle'

    def group(self):
        return ''

    def groupId(self):
        return ''

    def createInstance(self):
        return Modle()

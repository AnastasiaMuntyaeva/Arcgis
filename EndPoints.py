# -*- coding: utf-8 -*-
# Создание точек начала и конца в ArcGIS PRO

# 25.11.2024
# Мунтяева А.Е.

import arcpy


class Toolbox(object):
    def __init__(self):
        """Define the toolbox (the name of the toolbox is the name of the
        .pyt file)."""
        self.label = "EndPoints"
        self.alias = "Создает закрепы осей"

        # List of tool classes associated with this toolbox
        self.tools = [Tool]


class Tool(object):
    def __init__(self):
        """Define the tool (tool name is the name of the class)."""
        self.label = "EndPointsTool"
        self.description = ""
        self.canRunInBackground = False

    def getParameterInfo(self):
        """Define parameter definitions"""
        params = [
                arcpy.Parameter(
                displayName="Слой осей",
                name="in_features",
                datatype="GPFeatureLayer",
                parameterType="Required",
                direction="Input"),

                arcpy.Parameter(
                displayName="Папка для сохранения полученных точек в формате KML",
                name="KML",
                datatype="DEFolder",
                parameterType="Required",
                direction="Input")
               ]
        return params

    def isLicensed(self):
        """Set whether tool is licensed to execute."""
        return True

    def updateParameters(self, parameters):
        """Modify the values and properties of parameters before internal
        validation is performed.  This method is called whenever a parameter
        has been changed."""
        return

    def updateMessages(self, parameters):
        """Modify the messages created by internal validation for each tool
        parameter.  This method is called after internal validation."""
        return

    def execute(self, parameters, messages):

        Axis = parameters[0].valueASText  # Слой оси дороги

        KML = parameters[1].valueASText  #папка для кмл

        workspace = arcpy.Describe(Axis).path

        EndPoints_fc = workspace + "\\EndPoints"

        arcpy.env.overwriteOutput = True

        arcpy.management.FeatureVerticesToPoints(in_features=Axis, out_feature_class=EndPoints_fc,
                                                                                                point_location="START")

        arcpy.AddField_management(EndPoints_fc, "PointType", "TEXT", field_length=500, field_alias='Тип закрепа')

        arcpy.CalculateField_management(EndPoints_fc, "PointType","'Начало'", "PYTHON3")

        EndPoints_temp = workspace + "\\EndPoints_temp"

        arcpy.management.FeatureVerticesToPoints(in_features=Axis, out_feature_class=EndPoints_temp,
                                                                                                   point_location="END")

        arcpy.AddField_management(EndPoints_temp, "PointType", "TEXT", field_length=500, field_alias='Тип закрепа')

        arcpy.CalculateField_management(EndPoints_temp, "PointType", "'Конец'", "PYTHON3")

        arcpy.management.Append(inputs=EndPoints_temp, target=EndPoints_fc, schema_type='NO_TEST')

        arcpy.management.Delete(EndPoints_temp)

        arcpy.management.DeleteField(EndPoints_fc, ["Name","PointType"], 'KEEP_FIELDS')

        arcpy.AddField_management(EndPoints_fc, 'Longitude', 'DOUBLE', field_alias='Долгота')

        arcpy.AddField_management(EndPoints_fc, 'Latitude', 'DOUBLE', field_alias='Широта')

        geo_cs = arcpy.SpatialReference(4326) # WGS 84

        arcpy.management.CalculateGeometryAttributes(EndPoints_fc, [['Longitude', 'POINT_X'],  ['Latitude','POINT_Y']],
                                                                                               coordinate_system=geo_cs)

        arcpy.AlterAliasName(EndPoints_fc, 'Закрепы')

        aprx = arcpy.mp.ArcGISProject("CURRENT")

        aprxMap = aprx.listMaps("Map")[0]

        aprxMap.addDataFromPath(EndPoints_fc)

        lyr = aprxMap.listLayers('Закрепы')[0]

        sym = lyr.symbology

        sym.updateRenderer("UniqueValueRenderer")

        sym.renderer.fields = ["PointType"]  # Правильное имя свойства `fields`

        for grp in sym.renderer.groups:
            for itm in grp.items:
                myVal = itm.values[0][0]

                if myVal == "Начало":
                    itm.symbol.applySymbolFromGallery("Triangle 3")  # Настраиваем на `itm.symbol`
                    itm.symbol.color = {'RGB': [0, 255, 0, 100]}  # Зеленый цвет
                    itm.symbol.outlineColor = {'CMYK': [0, 0, 0, 100, 100]}
                    itm.symbol.size = 14
                    itm.label = str(myVal)  # Устанавливаем метку

                elif myVal == "Конец":
                    itm.symbol.applySymbolFromGallery("Triangle 3")  # Настраиваем на `itm.symbol`
                    itm.symbol.color = {'RGB': [255, 0, 0, 100]}  # Красный цвет
                    itm.symbol.outlineColor = {'CMYK': [0, 0, 0, 100, 100]}
                    itm.symbol.size = 14
                    itm.label = str(myVal)  # Устанавливаем метку
        lyr.symbology = sym
        arcpy.conversion.LayerToKML(lyr, (KML + '\\' + 'Закрепы'+'.kmz'))

        return

    def postExecute(self, parameters):
        """This method takes place after outputs are processed and
        added to the display."""
        return

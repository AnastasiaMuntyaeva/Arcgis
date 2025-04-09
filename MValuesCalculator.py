# -*- coding: utf-8 -*-

# Расчет абсолютного местоположения точки относительно начала дороги в ArcGIS PRO

# 07.04.2025
# Мунтяева А.Е.

import arcpy


class Toolbox:
    def __init__(self):
        self.label = "MValuesCalculatorNew"
        self.alias = "Пересчет значений координаты M"
        self.tools = [Tool]


class Tool:
    def __init__(self):
        self.label = "MValuesCalculator"
        self.description = ""
        self.canRunInBackground = False

    def getParameterInfo(self):
        params = [
            arcpy.Parameter(displayName="Оси для пересчета", name="in_features", datatype="GPFeatureLayer",
                            parameterType="Required", direction="Input"),
            arcpy.Parameter(displayName="Координатная система для пересчета", name="cs", datatype="GPCoordinateSystem",
                            parameterType="Required", direction="Input"),
            arcpy.Parameter(displayName="Поле с абсолютными началами в метрах", name="first_m", datatype="GPString",
                            parameterType="Required", direction="Input"),
            arcpy.Parameter(displayName="Поле для заполнения абсолютных концов в метрах (Необязательно)", name="last_m",
                            datatype="GPString", parameterType="Optional", direction="Input"),
            arcpy.Parameter(displayName="Поле для заполнения протяженности в метрах (Необязательно)", name="length",
                            datatype="GPString", parameterType="Optional", direction="Input")
        ]
        return params

    def isLicensed(self):
        return True

    def updateParameters(self, parameters):
        in_layer = parameters[0].valueAsText
        if in_layer:
            field_names = [field.name for field in arcpy.ListFields(in_layer)]
            parameters[2].filter.type = 'ValueList'
            parameters[2].filter.list = field_names
            parameters[3].filter.type = 'ValueList'
            parameters[3].filter.list = field_names
            parameters[4].filter.type = 'ValueList'
            parameters[4].filter.list = field_names
        return

    def updateMessages(self, parameters):
        return

    def process_polyline(self, polyline, absolute_start):
        new_parts = []
        total_cumulative_distance = 0.0

        for part in polyline:
            new_vertices = []
            cumulative_distance = 0.0
            prev_point = None

            for i, pnt in enumerate(part):
                if i > 0:
                    prev_geom = arcpy.PointGeometry(prev_point, polyline.spatialReference)
                    curr_geom = arcpy.PointGeometry(pnt, polyline.spatialReference)
                    cumulative_distance += prev_geom.distanceTo(curr_geom)
                else:
                    cumulative_distance = 0.0

                new_pnt = arcpy.Point(pnt.X, pnt.Y, None, cumulative_distance + absolute_start)
                new_pnt.M = cumulative_distance + absolute_start
                new_vertices.append(new_pnt)
                prev_point = pnt

            new_parts.append(new_vertices)
            total_cumulative_distance += cumulative_distance

        new_parts_array = arcpy.Array([arcpy.Array(part) for part in new_parts])
        new_polyline = arcpy.Polyline(new_parts_array, polyline.spatialReference, False, True)

        return new_polyline, total_cumulative_distance

    def execute(self, parameters, messages):
        arcpy.env.overwriteOutput = True
        axis = parameters[0].valueASText
        cs_name = parameters[1].value
        cs = arcpy.SpatialReference()
        cs.loadFromString(cs_name)
        first_m_field = parameters[2].valueASText
        last_m_field = parameters[3].valueASText
        length_field = parameters[4].valueASText

        workspace = arcpy.Describe(axis).path
        axis_name = arcpy.Describe(axis).featureClass.name
        axis_path = workspace + "\\" + axis_name

        id_list = []
        with arcpy.da.SearchCursor(axis, ["OBJECTID"]) as cursor:
            for row in cursor:
                id_list.append(row[0])

        arcpy.conversion.FeatureClassToFeatureClass(axis, workspace, "axis_temp2")
        temp_lyr = workspace + "\\axis_temp2"
        lyr = workspace + "\\axis_temp"
        arcpy.management.Project(in_dataset=temp_lyr, out_dataset=lyr, out_coor_system=cs)

        field_combinations = [
            (["SHAPE@", first_m_field, last_m_field, length_field], last_m_field and length_field),
            (["SHAPE@", first_m_field, last_m_field], last_m_field and not length_field),
            (["SHAPE@", first_m_field, length_field], length_field and not last_m_field),
            (["SHAPE@", first_m_field], not last_m_field and not length_field)
        ]

        for fields, condition in field_combinations:
            if condition:
                with arcpy.da.UpdateCursor(lyr, fields) as cursor:
                    for row in cursor:
                        polyline, total_distance = self.process_polyline(row[0], row[1])
                        row[0] = polyline

                        if len(fields) > 2:
                            row[2] = total_distance + row[1]  # last_m_field
                        if len(fields) > 3:
                            row[3] = total_distance  # length_field

                        cursor.updateRow(row)

        with arcpy.da.UpdateCursor(axis, ["OBJECTID"]) as cursor:
            for row in cursor:
                if row[0] in id_list:
                    cursor.deleteRow()

        arcpy.management.Append(lyr, axis_path)
        arcpy.management.Delete([lyr, temp_lyr])

        return

    def postExecute(self, parameters):
        return

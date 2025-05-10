# -*- coding: utf-8 -*-

# Разбивка оси по км.

# 19.12.2024
# Мунтяева А.Е.

import arcpy


class Toolbox(object):
    def __init__(self):
        """Define the toolbox (the name of the toolbox is the name of the
        .pyt file)."""
        self.label = "KMRoadSplitter"
        self.alias = "KMRoadSplitter"

        # List of tool classes associated with this toolbox
        self.tools = [Tool]


class Tool(object):
    def __init__(self):
        """Define the tool (tool name is the name of the class)."""
        self.label = "KMRoadSplitter"
        self.description = ""
        self.canRunInBackground = False

    def getParameterInfo(self):
        """Define parameter definitions"""

        params =[
                arcpy.Parameter(
                displayName="Ось",
                name="in_features_axis",
                datatype="GPFeatureLayer",
                parameterType="Required",
                direction="Input"),

                arcpy.Parameter(
                displayName="Километровые столбы",
                name="in_features_km",
                datatype="GPFeatureLayer",
                parameterType="Required",
                direction="Input"),

                arcpy.Parameter(
                displayName="Координатная система для пересчета",
                name="cs",
                datatype="GPCoordinateSystem",
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

        axis = parameters[0].valueASText  # Слой оси дороги
        km = parameters[1].valueASText  # км столбы
        cs_name = parameters[2].value # координатная система
        cs=arcpy.SpatialReference()
        cs.loadFromString(cs_name)

        arcpy.env.overwriteOutput = True

        workspace = arcpy.Describe(axis).path
        axis_name = arcpy.Describe(axis).featureClass.name
        ROUTES = workspace + "\\" + axis_name # дорога в wgs исходная

        km_name = arcpy.Describe(km).featureClass.name
        POSTSKMS= workspace + "\\"+ km_name # километровые столбы в wgs исходные


        geo_cs = arcpy.SpatialReference(4326) # WGS 84

        # считаем поля Широта, Долгота

        arcpy.management.CalculateGeometryAttributes(POSTSKMS, [['LON', 'POINT_X'],
                                                                ['LAT','POINT_Y']],
                                                                coordinate_system=geo_cs)

        # Создаем список значений KM_START и OID@
        KM_START_LIST = []
        with arcpy.da.SearchCursor(POSTSKMS, ["KM_START", "OID@"]) as cursor:
            for row in cursor:
                if row[0] is not None:  # Проверяем, что KM_START не пустой
                    KM_START_LIST.append((row[0], row[1]))

        if not KM_START_LIST:
            raise RuntimeError("Список KM_START пуст. Проверьте данные в POSTSKMS.")

        # Сортируем по KM_START
        KM_START_LIST.sort(key=lambda x: x[0])

        # Обновляем KM_END
        with arcpy.da.UpdateCursor(POSTSKMS, ["KM_START", "KM_END", "OID@"]) as cursor:
            for row in cursor:
                for i, (km_start, oid) in enumerate(KM_START_LIST):
                    if row[2] == oid:
                        if i < len(KM_START_LIST) - 1:
                            row[1] = KM_START_LIST[i + 1][0]  # Следующий KM_START
                        else:
                            row[1] = None  # Последний элемент
                        cursor.updateRow(row)
                        break


        # создаем перепроецированные классы в заданной ск
        ROUTES_UTM = workspace + "\\ROUTES_UTM"
        POSTSKMS_UTM = workspace + "\\POSTSKMS_UTM"

        arcpy.management.Project(in_dataset=ROUTES, out_dataset=ROUTES_UTM, out_coor_system=cs)
        arcpy.management.Project(in_dataset=POSTSKMS, out_dataset=POSTSKMS_UTM, out_coor_system=cs)

        # делаем копию слоя км столбов для проецирования на линию
        POSTSKMS_PROJECTED = workspace + "\\POSTSKMS_PROJECTED"
        arcpy.management.CopyFeatures(POSTSKMS_UTM, POSTSKMS_PROJECTED)

        # проецируем столбы на линию
        arcpy.edit.Snap(POSTSKMS_PROJECTED, [[ROUTES_UTM, "EDGE", "100 Meters"]])

        # разделяем линию в полученных точках
        KM_ROUTES_UTM = workspace + "\\KM_ROUTES_UTM"
        arcpy.management.SplitLineAtPoint(ROUTES_UTM, POSTSKMS_PROJECTED, KM_ROUTES_UTM, "100 Meters")

        POSTSKMS_table = workspace + "\\POSTSKMS_table"

        #создаем табличку со значениями отступа и измерения на маршруте
        arcpy.lr.LocateFeaturesAlongRoutes(in_features=POSTSKMS_UTM, in_routes=ROUTES_UTM,
                                           route_id_field="Name", radius_or_tolerance="100 Meters",
                                           out_table=POSTSKMS_table,
                                           out_event_properties="RID; Point; MEAS",
                                           route_locations="FIRST", distance_field="DISTANCE",
                                           zero_length_events="ZERO", in_fields="FIELDS",
                                           m_direction_offsetting="M_DIRECTON")




        # Создаём словарь из таблицы POSTSKMS_table
        postskms_dict = {}
        with arcpy.da.SearchCursor(POSTSKMS_table, ["KM_START", "MEAS", "DISTANCE"]) as cursor:
            for row in cursor:
                km_start = row[0]
                meas = row[1]
                distance = row[2]
                if km_start not in postskms_dict:
                    postskms_dict[km_start] = (meas, distance)

        # Используем UpdateCursor для обновления значений MEASURE_ROUTE и OFFSET
        with arcpy.da.UpdateCursor(POSTSKMS, ["KM_START", "MEASURE_ROUTE", "OFFSET"]) as cursor:
            for row in cursor:
                km_start = row[0]

                # Проверяем наличие данных в словаре
                if km_start in postskms_dict:
                    meas, distance = postskms_dict[km_start]

                    # Заполняем MEASURE_ROUTE и OFFSET
                    row[1] = round(meas)  # MEASURE_ROUTE
                    row[2] = round(abs(distance), 2)  # OFFSET

                    # Обновляем строку
                    cursor.updateRow(row)



# находим начало оси
        with arcpy.da.SearchCursor(ROUTES, ["StartKM", "M_ABS_START"]) as cursor:
            for row in cursor:
                StartKM=row[0]
                M_ABS_START= row[1]



        rows = []
        with arcpy.da.SearchCursor(POSTSKMS,  ["KM_START", "MEASURE_ROUTE", "KM_DISTANCE"]) as cursor:
            for row in cursor:
                rows.append(list(row))


        # Обновляем таблицу с вычислением разницы

        with arcpy.da.UpdateCursor(POSTSKMS,  ["KM_START", "MEASURE_ROUTE", "KM_DISTANCE"]) as update_cursor:
            for i, row in enumerate(update_cursor):
                if i < len(rows) - 2:  # Проверяем, чтобы не выйти за пределы
                    current_measure = row[1]  # MEASURE_ROUTE текущей строки
                    next_measure = rows[i + 1][1]  # MEASURE_ROUTE следующей строки
                    km_distance = abs(current_measure - next_measure)  # Разница по модулю
                else:
                    km_distance = None  # Для последней строки

                # Обновляем значение KM_DISTANCE в текущей строке
                row[2] = km_distance
                update_cursor.updateRow(row)


        # Проверяем StartKM и корректируем список
        if StartKM != KM_START_LIST[0][0]:  # Проверка относительно первого элемента списка
            KM_START_LIST.insert(0, (StartKM, 0))

        KM_START_LIST_WITH_INDEX = [(index + 1, item[0], item[1]) for index, item in enumerate(KM_START_LIST)]

#для первой М нужно учесть км+м 

        with arcpy.da.UpdateCursor(KM_ROUTES_UTM, ["SHAPE@", "StartKM", "ORIG_SEQ"]) as cursor:
            for row in cursor:
                polyline = row[0]
                ORIG_SEQ = row[2]  # Связь с километровым столбом через поле ORIG_SEQ

                km_start_entry = next((km for km in KM_START_LIST_WITH_INDEX if km[0] == ORIG_SEQ), None)

                if km_start_entry[1]==StartKM:
                    current_m = M_ABS_START  # Начальное значение M для первого участка
                    updated_paths = arcpy.Array()
                else:
                    km_start = km_start_entry[1]
                    current_m = km_start * 1000  # Начальное значение M для участка
                    updated_paths = arcpy.Array()

                    for part in polyline:
                        updated_part = arcpy.Array()

                        for i, point in enumerate(part):
                            if i == 0:
                                # Первая точка участка
                                updated_point = arcpy.Point(point.X, point.Y, None, current_m)
                            else:
                                # Остальные точки
                                prev_point = updated_part[-1]
                                temp_line = arcpy.Polyline(
                                    arcpy.Array([prev_point, point]),
                                    polyline.spatialReference
                                )
                                segment_length = temp_line.length

                                # Обновляем M-координату
                                current_m += segment_length
                                updated_point = arcpy.Point(point.X, point.Y, None, current_m)

                            updated_part.add(updated_point)

                        updated_paths.add(updated_part)

                    # Задаём новую геометрию с пересчитанными координатами M
                    row[0] = arcpy.Polyline(updated_paths, polyline.spatialReference, True)
                    row[1] = km_start  # Обновляем поле StartKM
                    cursor.updateRow(row)


        arcpy.AddField_management(KM_ROUTES_UTM, 'KM', 'LONG', field_alias='KM')
        arcpy.management.CalculateField(KM_ROUTES_UTM, 'KM', "!StartKM!")

        arcpy.management.DeleteField(KM_ROUTES_UTM, ['KM'], 'KEEP_FIELDS')

        KM_BUFFERS_UTM= workspace + "\\KM_BUFFERS_UTM"
        arcpy.analysis.Buffer(in_features=KM_ROUTES_UTM, out_feature_class=KM_BUFFERS_UTM,
                             buffer_distance_or_field="50 Meters", line_side="FULL", line_end_type="FLAT",
                             dissolve_option="NONE", dissolve_field=[], method="PLANAR")

        arcpy.edit.Snap(POSTSKMS_PROJECTED, [[ROUTES_UTM, "EDGE", "100 Meters"]])

        arcpy.management.Delete(POSTSKMS_table)

        arcpy.AlterAliasName(KM_BUFFERS_UTM, "Буферы километровых участков")
        arcpy.AlterAliasName(KM_ROUTES_UTM, "Километровые участки")
        arcpy.AlterAliasName(POSTSKMS, "Километровые столбы")
        arcpy.AlterAliasName(POSTSKMS_PROJECTED, "Километровые столбы спроецированые")
        arcpy.AlterAliasName(POSTSKMS_UTM, "Километровые столбы UTM")
        arcpy.AlterAliasName(ROUTES, "Ось дороги")
        arcpy.AlterAliasName(ROUTES_UTM, "Ось дороги UTM")

        aprx = arcpy.mp.ArcGISProject("CURRENT")
        aprxMap = aprx.listMaps("Map")[0]

        aprxMap.addDataFromPath(KM_BUFFERS_UTM)
        aprxMap.addDataFromPath(KM_ROUTES_UTM)
        aprxMap.addDataFromPath(POSTSKMS_PROJECTED)


        return

    def postExecute(self, parameters):
        """This method takes place after outputs are processed and
        added to the display."""
        return

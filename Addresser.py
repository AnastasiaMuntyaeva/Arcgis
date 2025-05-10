# -*- coding: utf-8 -*-
# Адресация объектов УДС в ArcGIS PRO

# 23.12.2024
# Мунтяева А.Е.


import arcpy
import pandas as pd
import os
import time

class Toolbox(object):
    def __init__(self):
        """Define the toolbox (the name of the toolbox is the name of the
        .pyt file)."""
        self.label = "Toolbox"
        self.alias = "toolbox"

        # List of tool classes associated with this toolbox
        self.tools = [Tool]


class Tool(object):
    def __init__(self):
        """Define the tool (tool name is the name of the class)."""
        self.label = "Tool"
        self.description = ""
        self.canRunInBackground = False


    def getParameterInfo(self):
        """Define parameter definitions"""
        params =[
                arcpy.Parameter(
                displayName="Объекты для адрессации",
                name="in_features",
                datatype="GPFeatureLayer",
                parameterType="Required",
                direction="Input"),

                arcpy.Parameter(
                displayName="Ось дороги",
                name="routes",
                datatype="GPFeatureLayer",
                parameterType="Required",
                direction="Input"),


                arcpy.Parameter(
                displayName="Папка для сохранения выходных файлов",
                name="output_folder",
                datatype="DEFolder",
                parameterType="Required",
                direction="Input"),
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

        # Функция для создания уникального имени файла
        def generate_unique_filename(base_filename, output_folder):
            """Генерирует уникальное имя файла, добавляя порядковый номер, если файл существует"""
            count = 1
            file_name = "{}".format(base_filename)
            while os.path.exists(os.path.join(output_folder, file_name)):
                file_name = "{}_{}".format(count, base_filename)
                count += 1
            return os.path.join(output_folder, file_name)

        # Указываем входные параметры
        OBJECTS = parameters[0].valueASText
        ROUTES = parameters[1].valueASText
        OUTPUT_FOLDER = parameters[2].valueASText

        # Разрешаем перезапись
        arcpy.env.overwriteOutput = True

        # Указываем рабочее пространство
        workspace = arcpy.Describe(ROUTES).path
        routes_name = arcpy.Describe(ROUTES).featureClass.name
        KM_ROUTES_UTM = workspace + "\\KM_ROUTES_UTM"
        ROUTES_UTM = workspace + "\\ROUTES_UTM"
        workspace = arcpy.Describe(ROUTES).path
        desc = arcpy.Describe(OBJECTS)

        # Если объект является точкой
        if desc.shapeType == "Point":
            OBJECTS_TEMP = workspace + "\\OBJECTS_TEMP"
            OBJECTS_TEMP1 = workspace + "\\OBJECTS_TEMP1"

            # Генерируем уникальные имена файлов для Excel
            Measure_xlsx = generate_unique_filename("Measure.xlsx", OUTPUT_FOLDER)
            km_m_xlsx = generate_unique_filename("km_m.xlsx", OUTPUT_FOLDER)

            # Проецируем объекты на оси
            arcpy.lr.LocateFeaturesAlongRoutes(in_features = OBJECTS, in_routes = ROUTES_UTM,
                                           route_id_field = "Name", radius_or_tolerance = "800 Meters",
                                           out_table = OBJECTS_TEMP,
                                           out_event_properties = "KM; Point; MEASURE",
                                           route_locations = "FIRST", distance_field = "DISTANCE",
                                           zero_length_events = "ZERO", in_fields = "FIELDS",
                                           m_direction_offsetting = "M_DIRECTON")

            arcpy.lr.LocateFeaturesAlongRoutes(in_features = OBJECTS, in_routes = KM_ROUTES_UTM,
                                           route_id_field="KM", radius_or_tolerance="800 Meters",
                                           out_table=OBJECTS_TEMP1,
                                           out_event_properties="KM; Point; MEASURE",
                                           route_locations="FIRST", distance_field="DISTANCE",
                                           zero_length_events="ZERO", in_fields="FIELDS",
                                           m_direction_offsetting="M_DIRECTON")

            # Экспортируем слой в таблицы Excel
            arcpy.conversion.TableToExcel(OBJECTS_TEMP, Measure_xlsx, Use_field_alias_as_column_header = "NAME",
                                                                    Use_domain_and_subtype_description = "CODE")

            arcpy.conversion.TableToExcel(OBJECTS_TEMP1, km_m_xlsx, Use_field_alias_as_column_header = "NAME",
                                                                  Use_domain_and_subtype_description = "CODE")

            # Читаем данные из Excel
            try:
                measure_data = pd.read_excel(Measure_xlsx)
                km_data = pd.read_excel(km_m_xlsx)

                # Выводим имена столбцов для проверки
                arcpy.AddMessage("Столбцы в Measure.xlsx: {}".format(measure_data.columns.tolist()))
                arcpy.AddMessage("Столбцы в km_m.xlsx: {}".format(km_data.columns.tolist()))

                # Проверяем наличие необходимых столбцов в данных Excel
                measure_required = ['UUID', 'MEASURE']
                km_required = ['UUID', 'KM', 'MEASURE']  # Добавляем поля 'KM' и 'MEASURE' для создания адреса

                measure_missing = [col for col in measure_required if col not in measure_data.columns]
                km_missing = [col for col in km_required if col not in km_data.columns]

                if measure_missing:
                    arcpy.AddError("Отсутствуют необходимые столбцы в Measure.xlsx: {}".format(", ".join(measure_missing)))
                    raise Exception("Отсутствуют необходимые столбцы в Measure.xlsx.")

                if km_missing:
                    arcpy.AddError("Отсутствуют необходимые столбцы в km_m.xlsx: {}".format(", ".join(km_missing)))
                    raise Exception("Отсутствуют необходимые столбцы в km_m.xlsx.")

                # Проверка типов данных
                try:
                    # Проверка столбца 'KM'
                    km_data['KM'] = pd.to_numeric(km_data['KM'], errors='raise')
                except ValueError:
                    arcpy.AddError("Столбец 'KM' в km_m.xlsx должен содержать числовые значения.")
                    raise Exception("Столбец 'KM' в km_m.xlsx должен содержать числовые значения.")

                try:
                    # Проверка столбца 'MEASURE'
                    measure_data['MEASURE'] = pd.to_numeric(measure_data['MEASURE'], errors='raise')
                except ValueError:
                    arcpy.AddError("Столбец 'MEASURE' в Measure.xlsx должен содержать числовые значения.")
                    raise Exception("Столбец 'MEASURE' в Measure.xlsx должен содержать числовые значения.")

                # Округляем значения 'MEASURE' в km_data и формируем адреса 'км+м'
                km_data['KM'] = pd.to_numeric(km_data['KM'], errors='coerce')
                km_data['MEASURE'] = pd.to_numeric(km_data['MEASURE'], errors='coerce')

                km_data['MEASURE'] = km_data['MEASURE'] - km_data['KM'] * 1000

                km_data['MEASURE'] = km_data['MEASURE'].apply(lambda x: "{:03d}".format(int(round(float(x)))))
                km_data['Adress'] = km_data['KM'].astype(str) + "+" + km_data['MEASURE'].astype(str)
                arcpy.AddMessage("Округление и добавление столбца с адресом 'км+м' в km_data")

            except Exception as e:
                arcpy.AddError("Ошибка при чтении Excel файлов: {}".format(str(e)))
                raise  # Прерываем выполнение при ошибке


            # Обновление значений в исходном слое 
            with arcpy.da.UpdateCursor(OBJECTS, ['UUID', 'MCoord', 'Adress']) as cursor:
                for row in cursor:
                    uuid_value = row[0]  # Получаем UUID текущей строки
                    # Фильтрация по 'UUID'
                    measure_row = measure_data[measure_data['UUID'] == uuid_value]  # Используем 'UUID' из Measure.xlsx
                    km_row = km_data[km_data['UUID'] == uuid_value]  # Используем 'UUID' из km_m.xlsx

                    if not measure_row.empty:
                        # Округляем значение 'MEASURE' и форматируем его как трехзначное число
                        try:
                            measure_value = measure_row['MEASURE'].values[0]  # Получаем значение MEASURE
                            rounded_measure = int(round(float(measure_value)))  # Округляем до ближайшего целого
                            formatted_measure = "{:03d}".format(rounded_measure)  # Форматируем как трехзначное число
                            arcpy.AddMessage(
                                "Обновление 'MCoord' для UUID: {} с отформатированным значением: {}".format(uuid_value,
                                                                                                     formatted_measure))
                            row[1] = formatted_measure  # Обновляем столбец 'MCoord'
                        except ValueError as ve:
                            arcpy.AddWarning("Ошибка преобразования для UUID {}: {}".format(uuid_value, str(ve)))
                    else:
                        arcpy.AddMessage("Нет данных для 'MCoord' для UUID: {}".format(uuid_value))

                    if not km_row.empty:
                        arcpy.AddMessage("Обновление 'Adress' для UUID: {}".format(uuid_value))
                        row[2] = km_row['Adress'].values[0]  # Обновляем столбец 'Adress' с новыми данными 'км+м'
                    else:
                        arcpy.AddMessage("Нет данных для 'Adress' для UUID: {}".format(uuid_value))

                    cursor.updateRow(row)

            # Удаляем временные слои
            arcpy.management.Delete(OBJECTS_TEMP)
            arcpy.management.Delete(OBJECTS_TEMP1)

            # Сообщаем об окончании процесса
            arcpy.AddMessage("Процесс завершён.")

        # Если объект является линией
        elif desc.shapeType == "Polyline":
            fields = arcpy.ListFields(OBJECTS)

            # Если у линии нужно получить адрес начала и конца
            if 'MCoordS' in [field.name for field in fields]:

                OBJECTS_TEMP = workspace + "\\OBJECTS_TEMP" # Точки начал линии
                OBJECTS_TEMP1 = workspace + "\\OBJECTS_TEMP1" # Точки конца линии
                OBJECTS_TEMP2 = workspace + "\\OBJECTS_TEMP2" # Таблица MEASURE начал линии
                OBJECTS_TEMP3 = workspace + "\\OBJECTS_TEMP3" # Таблица KM+M начал линии
                OBJECTS_TEMP4 = workspace + "\\OBJECTS_TEMP4" # Таблица MEASURE конца линии
                OBJECTS_TEMP5 = workspace + "\\OBJECTS_TEMP5" # Таблица KM+M конца линии

                # Генерируем уникальные имена файлов для Excel
                Line_start_Measure = generate_unique_filename("Line_start_Measure.xlsx", OUTPUT_FOLDER)
                Line_end_Measure = generate_unique_filename("Line_end_Measure.xlsx", OUTPUT_FOLDER)
                Line_start_Km_m = generate_unique_filename("Line_start_Km_m.xlsx", OUTPUT_FOLDER)
                Line_end_Km_m = generate_unique_filename("Line_end_Km_m.xlsx", OUTPUT_FOLDER)

                # Создает точки начала и конца
                arcpy.management.FeatureVerticesToPoints(in_features=OBJECTS, out_feature_class=OBJECTS_TEMP,
                                                                                                    point_location="START")
                arcpy.management.FeatureVerticesToPoints(in_features=OBJECTS, out_feature_class=OBJECTS_TEMP1,
                                                                                                       point_location="END")

                # Проецируем объекты на оси
                arcpy.lr.LocateFeaturesAlongRoutes(in_features = OBJECTS_TEMP, in_routes = ROUTES_UTM,
                                               route_id_field = "Name", radius_or_tolerance = "800 Meters",
                                               out_table = OBJECTS_TEMP2,
                                               out_event_properties = "KM; Point; MEASURE",
                                               route_locations = "FIRST", distance_field = "DISTANCE",
                                               zero_length_events = "ZERO", in_fields = "FIELDS",
                                               m_direction_offsetting = "M_DIRECTON")

                arcpy.lr.LocateFeaturesAlongRoutes(in_features = OBJECTS_TEMP, in_routes = KM_ROUTES_UTM,
                                               route_id_field="KM", radius_or_tolerance="800 Meters",
                                               out_table=OBJECTS_TEMP3,
                                               out_event_properties="KM; Point; MEASURE",
                                               route_locations="FIRST", distance_field="DISTANCE",
                                               zero_length_events="ZERO", in_fields="FIELDS",
                                               m_direction_offsetting="M_DIRECTON")

                arcpy.lr.LocateFeaturesAlongRoutes(in_features = OBJECTS_TEMP1, in_routes = ROUTES_UTM,
                                               route_id_field = "Name", radius_or_tolerance = "800 Meters",
                                               out_table = OBJECTS_TEMP4,
                                               out_event_properties = "KM; Point; MEASURE",
                                               route_locations = "FIRST", distance_field = "DISTANCE",
                                               zero_length_events = "ZERO", in_fields = "FIELDS",
                                               m_direction_offsetting = "M_DIRECTON")

                arcpy.lr.LocateFeaturesAlongRoutes(in_features = OBJECTS_TEMP1, in_routes = KM_ROUTES_UTM,
                                               route_id_field="KM", radius_or_tolerance="800 Meters",
                                               out_table=OBJECTS_TEMP5,
                                               out_event_properties="KM; Point; MEASURE",
                                               route_locations="FIRST", distance_field="DISTANCE",
                                               zero_length_events="ZERO", in_fields="FIELDS",
                                               m_direction_offsetting="M_DIRECTON")

                # Экспортируем слой в таблицы Excel
                arcpy.conversion.TableToExcel(OBJECTS_TEMP2, Line_start_Measure, Use_field_alias_as_column_header = "NAME",
                                                                                Use_domain_and_subtype_description = "CODE")
                arcpy.conversion.TableToExcel(OBJECTS_TEMP3, Line_start_Km_m, Use_field_alias_as_column_header = "NAME",
                                                                                Use_domain_and_subtype_description = "CODE")
                arcpy.conversion.TableToExcel(OBJECTS_TEMP4, Line_end_Measure, Use_field_alias_as_column_header = "NAME",
                                                                                Use_domain_and_subtype_description = "CODE")
                arcpy.conversion.TableToExcel(OBJECTS_TEMP5, Line_end_Km_m, Use_field_alias_as_column_header = "NAME",
                                                                                Use_domain_and_subtype_description = "CODE")

                # Удаляем временные слои
                arcpy.management.Delete(OBJECTS_TEMP)
                arcpy.management.Delete(OBJECTS_TEMP1)
                arcpy.management.Delete(OBJECTS_TEMP2)
                arcpy.management.Delete(OBJECTS_TEMP3)
                arcpy.management.Delete(OBJECTS_TEMP4)
                arcpy.management.Delete(OBJECTS_TEMP5)

                # Чтение данных из Excel
                try:
                    measure_start_data = pd.read_excel(Line_start_Measure)
                    measure_end_data = pd.read_excel(Line_end_Measure)
                    km_start_data = pd.read_excel(Line_start_Km_m)
                    km_end_data = pd.read_excel(Line_end_Km_m)

                    # Проверка наличия необходимых столбцов
                    required_measure_columns = ['UUID', 'MEASURE']
                    for col in required_measure_columns:
                        if col not in measure_start_data.columns:
                            raise Exception("Отсутствует столбец {} в Line_start_Measure.xlsx".format(col))
                        if col not in measure_end_data.columns:
                            raise Exception("Отсутствует столбец {} в Line_end_Measure.xlsx".format(col))

                    required_km_columns = ['UUID', 'KM']
                    for col in required_km_columns:
                        if col not in km_start_data.columns:
                            raise Exception("Отсутствует столбец {} в Line_start_Km_m.xlsx".format(col))
                        if col not in km_end_data.columns:
                            raise Exception("Отсутствует столбец {} в Line_end_Km_m.xlsx".format(col))

                    # Округление и добавление столбца адреса
                    km_start_data['MEASURE'] = km_start_data['MEASURE'] - km_start_data['KM'] * 1000
                    km_start_data['MEASURE'] = km_start_data['MEASURE'].apply(lambda x: "{:03d}".format(int(round(float(x))))) if 'MEASURE' in km_start_data.columns else None
                    km_start_data['Adress'] = km_start_data['KM'].astype(str) + "+" + km_start_data['MEASURE']

                    km_end_data['MEASURE'] = km_end_data['MEASURE'] - km_end_data['KM'] * 1000
                    km_end_data['MEASURE'] = km_end_data['MEASURE'].apply(lambda x: "{:03d}".format(int(round(float(x))))) if 'MEASURE' in km_end_data.columns else None
                    km_end_data['Adress'] = km_end_data['KM'].astype(str) + "+" + km_end_data['MEASURE']
                    arcpy.AddMessage("Округление и добавление столбца с адресом 'км+м' завершено")

                    # Округление значений MEASURE в measure_start_data и measure_end_data
                    measure_start_data['MEASURE'] = measure_start_data['MEASURE'].apply(
                        lambda x: round(x)) if 'MEASURE' in measure_start_data.columns else None
                    measure_end_data['MEASURE'] = measure_end_data['MEASURE'].apply(
                        lambda x: round(x)) if 'MEASURE' in measure_end_data.columns else None

                except Exception as e:
                    arcpy.AddError("Ошибка при чтении Excel файлов: {}".format(str(e)))
                    raise

                # Обновление значений в исходном слое
                with arcpy.da.UpdateCursor(OBJECTS, ['UUID', 'MCoordS', 'MCoordE', 'AdressS', 'AdressE']) as cursor:
                    for row in cursor:
                        uuid_value = row[0]

                        # Обновление MCoordS
                        measure_start_row = measure_start_data[measure_start_data['UUID'] == uuid_value]
                        if not measure_start_row.empty:
                            row[1] = measure_start_row['MEASURE'].values[0]
                            arcpy.AddMessage("Обновление MCoordS для UUID {} с значением {}".format(uuid_value, row[1]))

                        # Обновление MCoordE
                        measure_end_row = measure_end_data[measure_end_data['UUID'] == uuid_value]
                        if not measure_end_row.empty:
                            row[2] = measure_end_row['MEASURE'].values[0]
                            arcpy.AddMessage("Обновление MCoordE для UUID {} с значением {}".format(uuid_value, row[2]))

                        # Обновление AdressS
                        if 'Adress' in km_start_data.columns:
                            km_start_row = km_start_data[km_start_data['UUID'] == uuid_value]
                            if not km_start_row.empty:
                                row[3] = km_start_row['Adress'].values[0]
                                arcpy.AddMessage("Обновление AdressS для UUID {} с значением {}".format(uuid_value, row[3]))

                        # Обновление AdressE
                        if 'Adress' in km_end_data.columns:
                            km_end_row = km_end_data[km_end_data['UUID'] == uuid_value]
                            if not km_end_row.empty:
                                row[4] = km_end_row['Adress'].values[0]
                                arcpy.AddMessage("Обновление AdressE для UUID {} с значением {}".format(uuid_value, row[4]))

                        cursor.updateRow(row)

            # Если у линии нужно получить адрес центра
            else:
                OBJECTS_TEMP = workspace + "\\OBJECTS_TEMP" # Точки центра линии
                OBJECTS_TEMP1 = workspace + "\\OBJECTS_TEMP1" # Таблица MEASURE
                OBJECTS_TEMP2 = workspace + "\\OBJECTS_TEMP2" # Таблица Km

                # Генерируем уникальные имена файлов для Excel
                One_points_line_measure = generate_unique_filename("One_points_line_measure.xlsx", OUTPUT_FOLDER)
                One_points_line_km_m = generate_unique_filename("One_points_line_km_m.xlsx", OUTPUT_FOLDER)

                # Создаем точки середины линии
                arcpy.management.FeatureVerticesToPoints(in_features=OBJECTS, out_feature_class=OBJECTS_TEMP,
                                                                                         point_location="MID")

                # Проецируем объекты на оси
                arcpy.lr.LocateFeaturesAlongRoutes(in_features = OBJECTS_TEMP, in_routes = ROUTES_UTM,
                                               route_id_field = "Name", radius_or_tolerance = "800 Meters",
                                               out_table = OBJECTS_TEMP1,
                                               out_event_properties = "KM; Point; MEASURE",
                                               route_locations = "FIRST", distance_field = "DISTANCE",
                                               zero_length_events = "ZERO", in_fields = "FIELDS",
                                               m_direction_offsetting = "M_DIRECTON")

                arcpy.lr.LocateFeaturesAlongRoutes(in_features = OBJECTS_TEMP, in_routes = KM_ROUTES_UTM,
                                                   route_id_field="KM", radius_or_tolerance="800 Meters",
                                                   out_table=OBJECTS_TEMP2,
                                                   out_event_properties="KM; Point; MEASURE",
                                                   route_locations="FIRST", distance_field="DISTANCE",
                                                   zero_length_events="ZERO", in_fields="FIELDS",
                                                   m_direction_offsetting="M_DIRECTON")

                # Экспорт таблиц в Excel
                arcpy.conversion.TableToExcel(OBJECTS_TEMP1, One_points_line_measure)
                arcpy.conversion.TableToExcel(OBJECTS_TEMP2, One_points_line_km_m)

                # Считываем данные из созданных таблиц
                try:
                    measure_data = pd.read_excel(One_points_line_measure)
                    km_data = pd.read_excel(One_points_line_km_m)

                    # Выводим имена столбцов для проверки
                    arcpy.AddMessage("Столбцы в Measure.xlsx: {}".format(measure_data.columns.tolist()))
                    arcpy.AddMessage("Столбцы в km_m.xlsx: {}".format(km_data.columns.tolist()))

                    # Проверяем наличие необходимых столбцов в данных Excel
                    measure_required = ['UUID', 'MEASURE']
                    km_required = ['UUID', 'KM', 'MEASURE']  # Добавляем поля 'KM' и 'MEASURE' для создания адреса

                    measure_missing = [col for col in measure_required if col not in measure_data.columns]
                    km_missing = [col for col in km_required if col not in km_data.columns]

                    if measure_missing:
                        arcpy.AddError("Отсутствуют необходимые столбцы в Measure.xlsx: {}".format(", ".join(measure_missing)))
                    if km_missing:
                        arcpy.AddError("Отсутствуют необходимые столбцы в km_m.xlsx: {}".format(", ".join(km_missing)))


                    # Округляем значения 'MEASURE' в km_data и формируем адреса 'км+м'
                    km_data['MEASURE'] = km_data['MEASURE'] - km_data['KM'] * 1000
                    km_data['MEASURE'] = km_data['MEASURE'].apply(lambda x: "{:03d}".format(int(round(float(x)))))
                    km_data['Adress'] = km_data['KM'].astype(str) + "+" + km_data['MEASURE']
                    arcpy.AddMessage("Округление и добавление столбца с адресом 'км+м' в km_data")

                except Exception as e:
                    arcpy.AddError("Ошибка при чтении Excel файлов: {}".format(str(e)))
                    raise  # Завершаем выполнение, если произошла ошибка

                # Обновление значений в исходном слое
                with arcpy.da.UpdateCursor(OBJECTS, ['UUID', 'MCoord', 'Adress']) as cursor:
                    for row in cursor:
                        uuid_value = row[0]  # Получаем UUID текущей строки

                        # Фильтрация по 'UUID'
                        measure_row = measure_data[measure_data['UUID'] == uuid_value]
                        km_row = km_data[km_data['UUID'] == uuid_value]

                        # Проверяем, что строка не пуста и что значение 'MEASURE' не является NaN
                        if not measure_row.empty and not pd.isnull(measure_row['MEASURE'].values[0]):
                            try:
                                measure_value = measure_row['MEASURE'].values[0]
                                rounded_measure = round(float(measure_value))
                                formatted_measure = "{:03d}".format(int(rounded_measure))
                                arcpy.AddMessage("Обновление 'MCoord' для UUID: {} с отформатированным значением: {}".format(uuid_value, formatted_measure))
                                row[1] = formatted_measure  # Обновляем MCoord
                            except ValueError as ve:
                                arcpy.AddWarning("Ошибка преобразования для UUID {}: {}".format(uuid_value, ve))
                        else:
                            arcpy.AddMessage("Нет данных для 'MCoord' для UUID: {}".format(uuid_value))

                        # Проверяем, что строка не пуста и что значение 'Adress' не является NaN
                        if not km_row.empty and not pd.isnull(km_row['Adress'].values[0]):
                            arcpy.AddMessage("Обновление 'Adress' для UUID: {}".format(uuid_value))
                            row[2] = km_row['Adress'].values[0]  # Обновляем Adress
                        else:
                            arcpy.AddMessage("Нет данных для 'Adress' для UUID: {}".format(uuid_value))

                        # Сохраняем изменения
                        cursor.updateRow(row)

                        # Удаляем временные слои
                        arcpy.management.Delete(OBJECTS_TEMP)
                        arcpy.management.Delete(OBJECTS_TEMP1)
                        arcpy.management.Delete(OBJECTS_TEMP2)



        # Если объект является полигоном
        elif desc.shapeType == "Polygon":
            fields = arcpy.ListFields(OBJECTS)

            # Если у полигона нужно получить адрес начала и конца
            if 'MCoordS' in [field.name for field in fields]:
                OBJECTS_TEMP = workspace + "\\OBJECTS_TEMP" # Точки вершин полигона
                OBJECTS_TEMP1 = workspace + "\\OBJECTS_TEMP1" # Таблица MEASURE
                OBJECTS_TEMP2 = workspace + "\\OBJECTS_TEMP2" # Таблица Km

                # Генерируем уникальные имена файлов для Excel
                Allpoints_polygon_measure = generate_unique_filename("Allpoints_polygon_measure.xlsx", OUTPUT_FOLDER)
                Allpoints_polygon_km_m = generate_unique_filename("AllOnepoints_polygon_km_m.xlsx", OUTPUT_FOLDER)

                arcpy.management.FeatureVerticesToPoints(in_features=OBJECTS, out_feature_class=OBJECTS_TEMP,
                                                                                         point_location="ALL")
                # Проецируем объекты на оси
                arcpy.lr.LocateFeaturesAlongRoutes(in_features = OBJECTS_TEMP, in_routes = ROUTES_UTM,
                                               route_id_field = "Name", radius_or_tolerance = "800 Meters",
                                               out_table = OBJECTS_TEMP1,
                                               out_event_properties = "KM; Point; MEASURE",
                                               route_locations = "FIRST", distance_field = "DISTANCE",
                                               zero_length_events = "ZERO", in_fields = "FIELDS",
                                               m_direction_offsetting = "M_DIRECTON")

                arcpy.lr.LocateFeaturesAlongRoutes(in_features = OBJECTS_TEMP, in_routes = KM_ROUTES_UTM,
                                                   route_id_field="KM", radius_or_tolerance="800 Meters",
                                                   out_table=OBJECTS_TEMP2,
                                                   out_event_properties="KM; Point; MEASURE",
                                                   route_locations="FIRST", distance_field="DISTANCE",
                                                   zero_length_events="ZERO", in_fields="FIELDS",
                                                   m_direction_offsetting="M_DIRECTON")

                # Экспорт таблиц в Excel
                arcpy.conversion.TableToExcel(OBJECTS_TEMP1, Allpoints_polygon_measure)
                arcpy.conversion.TableToExcel(OBJECTS_TEMP2, Allpoints_polygon_km_m)

                # Считываем данные из созданной таблицы Allpoints_polygon_measure
                try:
                    arcpy.AddMessage("Чтение данных для MCoordS и MCoordE...")
                    measure_df = pd.read_excel(Allpoints_polygon_measure)

                    # Проверка наличия необходимых столбцов
                    required_columns = ['MEASURE', 'UUID']
                    if not all(col in measure_df.columns for col in required_columns):
                        arcpy.AddError("Ошибка: Не все необходимые столбцы присутствуют в таблице. Ожидаются: {}".format(required_columns))
                        raise Exception("Ошибка: Не все необходимые столбцы присутствуют в таблице.")

                    # Извлечение уникальных значений UUID
                    unique_uuids = measure_df['UUID'].unique()

                    # Создаем новый DataFrame для хранения результатов
                    results_df = pd.DataFrame(unique_uuids, columns=['UUID'])

                    # Вычисление минимальных и максимальных значений MEASURE для каждого UUID
                    results_df['min_MEASURE'] = results_df['UUID'].apply(
                        lambda x: "{:03d}".format(int(round(measure_df.loc[measure_df['UUID'] == x, 'MEASURE'].min())))
                        if not measure_df.loc[measure_df['UUID'] == x, 'MEASURE'].empty else None
                    )

                    results_df['max_MEASURE'] = results_df['UUID'].apply(
                        lambda x: "{:03d}".format(int(round(measure_df.loc[measure_df['UUID'] == x, 'MEASURE'].max())))
                        if not measure_df.loc[measure_df['UUID'] == x, 'MEASURE'].empty else None
                    )

                except Exception as e:
                    arcpy.AddError("Ошибка при обработке данных: {}".format(str(e)))
                    raise

                arcpy.management.ClearWorkspaceCache()
                time.sleep(1)

                # Обновляем исходный слой Polygon
                arcpy.AddMessage("Обновление значений MCoordS и MCoordE в исходном слое...")
                with arcpy.da.UpdateCursor(OBJECTS, ['UUID', 'MCoordS', 'MCoordE']) as cursor:
                    for row in cursor:
                        uuid_value = row[0]  # Читаем UUID объекта

                        # Обновляем MCoordS (начальная координата)
                        min_measure_row = results_df[results_df['UUID'] == uuid_value]
                        if not min_measure_row.empty:
                            row[1] = min_measure_row['min_MEASURE'].values[0]  # Берем минимальное значение
                            arcpy.AddMessage("Обновление MCoordS для UUID {} с значением {}".format(uuid_value, row[1]))

                        # Обновляем MCoordE (конечная координата)
                        max_measure_row = results_df[results_df['UUID'] == uuid_value]
                        if not max_measure_row.empty:
                            row[2] = max_measure_row['max_MEASURE'].values[0]  # Берем максимальное значение
                            arcpy.AddMessage("Обновление MCoordE для UUID {} с значением {}".format(uuid_value, row[2]))

                        cursor.updateRow(row)

                # Чтение данных из созданной таблицы Allpoints_polygon_km_m
                arcpy.AddMessage("Чтение данных для AdressS и AdressE...")
                try:
                    km_df = pd.read_excel(Allpoints_polygon_km_m)

                    # Проверка наличия необходимых столбцов
                    required_columns = ['MEASURE', 'UUID', 'KM']
                    if not all(col in km_df.columns for col in required_columns):
                        arcpy.AddError("Ошибка: Не все необходимые столбцы присутствуют в таблице. Ожидаются: {}".format(required_columns))
                        raise Exception("Ошибка: Не все необходимые столбцы присутствуют в таблице.")

                    # Фильтрация по столбцу UUID
                    filtered_km_df = km_df[km_df['UUID'].notnull()]

                    # Группировка по UUID и расчет минимальных и максимальных значений KM
                    grouped_km_df = filtered_km_df.groupby('UUID').agg({
                        'KM': ['min', 'max']
                    }).reset_index()

                    # Преобразуем многоуровневые имена столбцов
                    grouped_km_df.columns = ['UUID', 'min_KM', 'max_KM']

                    # Для каждого UUID получаем min_measure, основываясь на min_KM
                    grouped_km_df['min_measure'] = grouped_km_df.apply(lambda row: filtered_km_df[
                        (filtered_km_df['UUID'] == row['UUID']) &
                        (filtered_km_df['KM'] == row['min_KM'])
                    ]['MEASURE'].min(), axis=1)

                    # Для каждого UUID получаем max_measure, основываясь на max_KM
                    grouped_km_df['max_measure'] = grouped_km_df.apply(lambda row: filtered_km_df[
                        (filtered_km_df['UUID'] == row['UUID']) &
                        (filtered_km_df['KM'] == row['max_KM'])
                    ]['MEASURE'].max(), axis=1)

                    # Округление min_measure и max_measure до целых чисел
                    grouped_km_df['min_measure'] = grouped_km_df['min_measure'].round().astype(int)
                    grouped_km_df['max_measure'] = grouped_km_df['max_measure'].round().astype(int)

                    # Преобразование min_measure и max_measure в формат с ведущими нулями (до трёх знаков)
                    grouped_km_df['min_measure'] = grouped_km_df['min_measure'] - grouped_km_df['min_KM'] * 1000
                    grouped_km_df['max_measure'] = grouped_km_df['max_measure'] - grouped_km_df['max_KM'] * 1000
                    grouped_km_df['min_measure'] = grouped_km_df['min_measure'].apply(lambda x: str(x).zfill(3))
                    grouped_km_df['max_measure'] = grouped_km_df['max_measure'].apply(lambda x: str(x).zfill(3))

                    # Формулы для создания адреса км+м
                    grouped_km_df['address_km_m_start'] = grouped_km_df['min_KM'].astype(str) + "+" + grouped_km_df['min_measure']
                    grouped_km_df['address_km_m_end'] = grouped_km_df['max_KM'].astype(str) + "+" + grouped_km_df['max_measure']

                    arcpy.AddMessage("Данные успешно обработаны.")

                except Exception as e:
                    arcpy.AddError("Ошибка при обработке данных: {}".format(str(e)))
                    raise

                # Обновляем исходный слой Polygon с новыми значениями
                arcpy.AddMessage("Обновление значений AdressS и AdressE в исходном слое...")
                with arcpy.da.UpdateCursor(OBJECTS, ['UUID', 'AdressS', 'AdressE']) as cursor:
                    for row in cursor:
                        uuid_value = row[0]  # Читаем UUID объекта

                        # Обновляем AdressS (начальная адресация)
                        min_measure_row = grouped_km_df[grouped_km_df['UUID'] == uuid_value]
                        if not min_measure_row.empty:
                            row[1] = min_measure_row['address_km_m_start'].values[0]  # Берем начальный адрес
                            arcpy.AddMessage("Обновление AdressS для UUID {} с значением {}".format(uuid_value, row[1]))

                        # Обновляем AdressE (конечная адресация)
                        max_measure_row = grouped_km_df[grouped_km_df['UUID'] == uuid_value]
                        if not max_measure_row.empty:
                            row[2] = max_measure_row['address_km_m_end'].values[0]  # Берем конечный адрес
                            arcpy.AddMessage("Обновление AdressE для UUID {} с значением {}".format(uuid_value, row[2]))

                        cursor.updateRow(row)

                arcpy.AddMessage("Обновление значений завершено.")

                # Удаление временных таблиц после использования
                arcpy.Delete_management(OBJECTS_TEMP)
                arcpy.Delete_management(OBJECTS_TEMP1)
                arcpy.Delete_management(OBJECTS_TEMP2)

            # Если у полигона нужно получить адрес центра
            else:
                OBJECTS_TEMP = workspace + "\\OBJECTS_TEMP" # Точки центра полигона
                OBJECTS_TEMP1 = workspace + "\\OBJECTS_TEMP1" # Таблица MEASURE
                OBJECTS_TEMP2 = workspace + "\\OBJECTS_TEMP2" # Таблица Km


                # Генерируем уникальные имена файлов для Excel
                Onepoints_polygon_measure = generate_unique_filename("Onepoints_polygon_measure.xlsx", OUTPUT_FOLDER)
                Onepoints_polygon_km_m = generate_unique_filename("Onepoints_polygon_km_m.xlsx", OUTPUT_FOLDER)

                # Создаем столбцы с координатами
                arcpy.AddField_management(OBJECTS, "Longitude_TEMP", "DOUBLE")
                arcpy.AddField_management(OBJECTS, "Latitude_TEMP", "DOUBLE")

                # Указываем систему координата
                geo_cs = arcpy.SpatialReference(4326)

                # Рассчитываем координаты
                arcpy.management.CalculateGeometryAttributes(OBJECTS, [['Longitude_TEMP', 'CENTROID_X'],
                                                                        ['Latitude_TEMP','CENTROID_Y']],
                                                                               coordinate_system=geo_cs)

                # Создаем точечный слой из таблицы
                Temp = arcpy.conversion.TableToTable(in_rows=OBJECTS, out_path=workspace, out_name="Temp")
                arcpy.management.XYTableToPoint(Temp, OBJECTS_TEMP, 'Longitude_TEMP', 'Latitude_TEMP')

                # Проецируем объекты на оси
                arcpy.lr.LocateFeaturesAlongRoutes(in_features = OBJECTS_TEMP, in_routes = ROUTES_UTM,
                                               route_id_field = "Name", radius_or_tolerance = "800 Meters",
                                               out_table = OBJECTS_TEMP1,
                                               out_event_properties = "KM; Point; MEASURE",
                                               route_locations = "FIRST", distance_field = "DISTANCE",
                                               zero_length_events = "ZERO", in_fields = "FIELDS",
                                               m_direction_offsetting = "M_DIRECTON")

                arcpy.lr.LocateFeaturesAlongRoutes(in_features = OBJECTS_TEMP, in_routes = KM_ROUTES_UTM,
                                               route_id_field="KM", radius_or_tolerance="800 Meters",
                                               out_table=OBJECTS_TEMP2,
                                               out_event_properties="KM; Point; MEASURE",
                                               route_locations="FIRST", distance_field="DISTANCE",
                                               zero_length_events="ZERO", in_fields="FIELDS",
                                               m_direction_offsetting="M_DIRECTON")

                # Экспорт таблиц в Excel
                arcpy.conversion.TableToExcel(OBJECTS_TEMP1, Onepoints_polygon_measure,
                                             Use_field_alias_as_column_header = "NAME",
                                           Use_domain_and_subtype_description = "CODE")
                arcpy.conversion.TableToExcel(OBJECTS_TEMP2, Onepoints_polygon_km_m,
                                          Use_field_alias_as_column_header = "NAME",
                                         Use_domain_and_subtype_description = "CODE")


                # Проверяем, созданы ли файлы
                if os.path.exists(Onepoints_polygon_measure) and os.path.exists(Onepoints_polygon_km_m):
                    arcpy.AddMessage("Оба файла Excel успешно созданы.")
                else:
                    arcpy.AddError("Не удалось создать один или оба Excel файла.")

                # Считываем данные из созданных таблиц
                try:
                    measure_data = pd.read_excel(Onepoints_polygon_measure)
                    km_data = pd.read_excel(Onepoints_polygon_km_m)

                    # Выводим имена столбцов для проверки
                    arcpy.AddMessage("Столбцы в Measure.xlsx: {}".format(measure_data.columns.tolist()))
                    arcpy.AddMessage("Столбцы в km_m.xlsx: {}".format(km_data.columns.tolist()))

                    # Проверяем наличие необходимых столбцов в данных Excel
                    measure_required = ['UUID', 'MEASURE']
                    km_required = ['UUID', 'KM', 'MEASURE']  # Добавляем поля 'KM' и 'MEASURE' для создания адреса

                    measure_missing = [col for col in measure_required if col not in measure_data.columns]
                    km_missing = [col for col in km_required if col not in km_data.columns]

                    if measure_missing:
                        arcpy.AddError("Отсутствуют необходимые столбцы в Measure.xlsx: {}".format(", ".join(measure_missing)))
                    if km_missing:
                        arcpy.AddError("Отсутствуют необходимые столбцы в km_m.xlsx: {}".format(", ".join(km_missing)))


                    # Округляем значения 'MEASURE' в km_data и формируем адреса 'км+м'
                    km_data['MEASURE'] = km_data['MEASURE'] - km_data['KM'] * 1000
                    km_data['MEASURE'] = km_data['MEASURE'].apply(lambda x: "{:03d}".format(int(round(float(x)))))
                    km_data['Adress'] = km_data['KM'].astype(str) + "+" + km_data['MEASURE']
                    arcpy.AddMessage("Округление и добавление столбца с адресом 'км+м' в km_data")

                except Exception as e:
                    arcpy.AddError("Ошибка при чтении Excel файлов: {}".format(str(e)))
                    raise  # Завершаем выполнение, если произошла ошибка

                # Обновление значений в исходном слое
                with arcpy.da.UpdateCursor(OBJECTS, ['UUID', 'MCoord', 'Adress']) as cursor:
                    for row in cursor:
                        uuid_value = row[0]  # Получаем UUID текущей строки

                        # Фильтрация по 'UUID'
                        measure_row = measure_data[measure_data['UUID'] == uuid_value]
                        km_row = km_data[km_data['UUID'] == uuid_value]

                        # Проверяем, что строка не пуста и что значение 'MEASURE' не является NaN
                        if not measure_row.empty and not pd.isnull(measure_row['MEASURE'].values[0]):
                            try:
                                measure_value = measure_row['MEASURE'].values[0]
                                rounded_measure = round(float(measure_value))
                                formatted_measure = "{:03d}".format(int(rounded_measure))
                                arcpy.AddMessage("Обновление 'MCoord' для UUID: {} с отформатированным значением: {}".format(uuid_value, formatted_measure))
                                row[1] = formatted_measure  # Обновляем MCoord
                            except ValueError as ve:
                                arcpy.AddWarning("Ошибка преобразования для UUID {}: {}".format(uuid_value, ve))
                        else:
                            arcpy.AddMessage("Нет данных для 'MCoord' для UUID: {}".format(uuid_value))

                        # Проверяем, что строка не пуста и что значение 'Adress' не является NaN
                        if not km_row.empty and not pd.isnull(km_row['Adress'].values[0]):
                            arcpy.AddMessage("Обновление 'Adress' для UUID: {}".format(uuid_value))
                            row[2] = km_row['Adress'].values[0]  # Обновляем Adress
                        else:
                            arcpy.AddMessage("Нет данных для 'Adress' для UUID: {}".format(uuid_value))

                        # Сохраняем изменения
                        cursor.updateRow(row)

                        # Удаляем временные слои
                        arcpy.management.Delete(OBJECTS_TEMP)
                        arcpy.management.Delete(OBJECTS_TEMP1)
                        arcpy.management.Delete(OBJECTS_TEMP2)
                        arcpy.management.Delete(Temp)

        return

    def postExecute(self, parameters):
        """This method takes place after outputs are processed and
        added to the display."""
        return

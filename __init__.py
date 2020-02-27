# coding: utf-8
"""
Base para desarrollo de modulos externos.
Para obtener el modulo/Funcion que se esta llamando:
     GetParams("module")

Para obtener las variables enviadas desde formulario/comando Rocketbot:
    var = GetParams(variable)
    Las "variable" se define en forms del archivo package.json

Para modificar la variable de Rocketbot:
    SetVar(Variable_Rocketbot, "dato")

Para obtener una variable de Rocketbot:
    var = GetVar(Variable_Rocketbot)

Para obtener la Opcion seleccionada:
    opcion = GetParams("option")


Para instalar librerias se debe ingresar por terminal a la carpeta "libs"

    pip install <package> -t .

"""
# Changing the data types of all strings in the module at once
from __future__ import unicode_literals
import os
import sys

base_path = tmp_global_obj["basepath"]
cur_path = base_path + 'modules' + os.sep + 'AdvancedExcel' + os.sep + 'libs' + os.sep
sys.path.append(cur_path)

global constants

constants = {"xlRowField": 1, "xlColumnField": 2, "xlPageField": 3}

module = GetParams("module")

if module == "pivotTable":

    sheet = GetParams("sheet")
    pivotTableName = GetParams("table")
    data = GetParams("data")
    option = GetParams("option_")

    excel = GetGlobals("excel")
    xls = excel.file_[excel.actual_id]

    try:
        wb = xls['workbook']
        sht = wb.sheets[sheet].select()

        pivotTable = wb.api.ActiveSheet.PivotTables(pivotTableName)
        cubeField = pivotTable.CubeFields(data)
        if option != "data":
            cubeField = pivotTable.CubeFields(data)
            cubeField.Orientation = constants[option]
            cubeField.Position = 1

        else:
            title = data.split(".")[-1]
            title = title.replace("[", "").replace("]", "")
            pivotTable.AddDataField(cubeField, title)

    except Exception as e:
        PrintException()
        raise e

if module == "updatePivot":

    sheet = GetParams("sheet")
    pivotTableName = GetParams("table")
    excel = GetGlobals("excel")

    try:
        xls = excel.file_[excel.actual_id]
        wb = xls['workbook']
        sht = wb.sheets[sheet].select()

        wb.api.ActiveSheet.PivotTables(pivotTableName).PivotCache().refresh()
    except Exception as e:
        PrintException()
        raise e

if module == "filter":

    sheet = GetParams("sheet")
    pivotTableName = GetParams("table")
    data = GetParams("data")
    more = GetParams("more")

    excel = GetGlobals("excel")
    xls = excel.file_[excel.actual_id]
    try:
        wb = xls['workbook']
        sht = wb.sheets[sheet].select()

        if data and data.count("]") != 1:
            father = data.split('.')[0]
        else:
            father = data

        pivotTable = wb.api.ActiveSheet.PivotTables(pivotTableName)
        cubeField = pivotTable.CubeFields(father)

        cubeField.Orientation = constants["xlPageField"]
        cubeField.Position = 1
        if not more:
            more = data.split(".")[0][1:-1]
            more = "[{data}].[All {data}]".format(data=more)

        print(more)

        pivotTable.PivotFields(data).CurrentPageName = more

    except Exception as e:
        PrintException()
        raise e

if module == "open":

    sheet = GetParams("sheet")
    pivotTableName = GetParams("table")
    data = GetParams("data")
    child = GetParams("child")

    excel = GetGlobals("excel")
    xls = excel.file_[excel.actual_id]
    try:
        wb = xls['workbook']
        sht = wb.sheets[sheet].select()

        if data.count("]") != 1:
            father = data.split('.')[0]
        else:
            father = data

        pivotTable = wb.api.ActiveSheet.PivotTables(pivotTableName)
        cubeField = pivotTable.CubeFields(father)

        pivotField = pivotTable.PivotFields(data).PivotItems(child).DrilledDown = True

    except Exception as e:
        PrintException()
        raise e

if module == "listFields":
    sheet = GetParams("sheet")
    pivotTableName = GetParams("table")
    result = GetParams("result")

    excel = GetGlobals("excel")
    xls = excel.file_[excel.actual_id]
    try:
        wb = xls['workbook']
        sht = wb.sheets[sheet].select()

        pivotTable = wb.api.ActiveSheet.PivotTables(pivotTableName)

        cubeFields = pivotTable.CubeFields

        fields = [field.Name for field in cubeFields]

        SetVar(result, fields)

    except Exception as e:
        PrintException()
        raise e

if module == "listSubField":
    sheet = GetParams("sheet")
    pivotTableName = GetParams("table")
    field = GetParams("data")
    result = GetParams("result")

    excel = GetGlobals("excel")
    xls = excel.file_[excel.actual_id]
    try:
        wb = xls['workbook']
        sht = wb.sheets[sheet].select()

        pivotTable = wb.api.ActiveSheet.PivotTables(pivotTableName)
        print(pivotTable.Name)
        cubeFields = pivotTable.CubeFields("[{}]".format(field))

        
        fields = [e.Name for e in cubeFields.PivotFields]

        SetVar(result, fields)

    except Exception as e:
        PrintException()
        raise e
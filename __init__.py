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
cur_path = base_path + 'modules' + os.sep + 'CubeExcel' + os.sep + 'libs' + os.sep

cur_path_x64 = os.path.join(cur_path, 'Windows' + os.sep +  'x64' + os.sep)
cur_path_x86 = os.path.join(cur_path, 'Windows' + os.sep +  'x86' + os.sep)

if sys.maxsize > 2**32 and cur_path_x64 not in sys.path:
        sys.path.append(cur_path_x64)
if sys.maxsize <= 2**32 and cur_path_x86 not in sys.path:
        sys.path.append(cur_path_x86)

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
            cubeField.Orientation = constants[option]
            cubeField.Position = 1

        else:
            title = data.split(".")[-1]
            title = title.replace("[", "").replace("]", "")
            pivotTable.AddDataField(cubeField, title)

    except Exception as e:
        PrintException()
        raise e
    
if module == "removeField":

    sheet = GetParams("sheet")
    pivotTableName = GetParams("table")
    data = GetParams("data")

    excel = GetGlobals("excel")
    xls = excel.file_[excel.actual_id]

    try:
        wb = xls['workbook']
        sht = wb.sheets[sheet].select()

        pivotTable = wb.api.ActiveSheet.PivotTables(pivotTableName)
        cubeField = pivotTable.CubeFields(data)
        cubeField.Orientation = 0 #xlHidden

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
        try:
            cubeField = pivotTable.CubeFields(father)
        except:
            cubeField = pivotTable.CubeFields(data)

        cubeField.Orientation = constants["xlPageField"]
        cubeField.Position = 1
        if not more:
            more = data.split(".")[0][1:-1]
            more = "[{data}].[All {data}]".format(data=more)

        try:
            pivotTable.PivotFields(data).CurrentPageName = more
        except:
            pass
    except Exception as e:
        PrintException()
        raise e

if module == "applyFilter":
    sheet = GetParams("sheet")
    pivotTableName = GetParams("table")
    field = GetParams("data")
    filter = GetParams("filter")
    numAsText = GetParams("numAsText")

    if not filter:
        filter = ['']
    
    excel = GetGlobals("excel")
    xls = excel.file_[excel.actual_id]
    try:
        wb = xls['workbook']
        sht = wb.sheets[sheet].select()
        pivotTable = wb.api.ActiveSheet.PivotTables(pivotTableName)
        pivotTable.PivotCache().refresh()
        
        pivotField = pivotTable.PivotFields(field)     
        
        global parent_
        parent_ = ".".join(pivotField.Name.split('.')[:-1])
        cubeField = pivotTable.CubeFields(parent_)
        cubeField.EnableMultiplePageItems = True
        
        if filter.startswith("["):
            filter = eval(filter)
        else:
            if "," in filter:
                filter = filter.split(",")
            else:
                filter = [filter]
        
        filter_parent = [parent_ == ".".join(f.split('.')[:-1]) for f in filter]
        if all(filter_parent) or filter == ['']:
            pivotField.VisibleItemsList = filter
        else:
            filter_aux = []
            for f in filter:
                if f.isnumeric():
                    if numAsText and eval(numAsText):
                        pass
                    elif len(f) == 1:
                        f = str(f) + "."
                    else:
                        point = str(float("0."+str(f)[1:])).replace("0.", "")
                        
                        f = str(f)[0] + "." + point + "E" + str(len(f)-1)
                filter_aux.append(f)
            
            filter_ = [parent_+f".&[{value}]" for value in filter_aux]
            
            pivotField.VisibleItemsList = filter_

    except Exception as e:
        PrintException()
        raise e

if module == "clearFilter":
    sheet = GetParams("sheet")
    pivotTableName = GetParams("table")
    field = GetParams("data")   

    excel = GetGlobals("excel")
    xls = excel.file_[excel.actual_id]
    try:
        wb = xls['workbook']
        sht = wb.sheets[sheet].select()
        pivotTable = wb.api.ActiveSheet.PivotTables(pivotTableName)
        
        pivotField = pivotTable.PivotFields(field)
        pivotField.VisibleItemsList = [""]
        
    except Exception as e:
        PrintException()
        raise e

if module == "clearAllFilters":
    sheet = GetParams("sheet")
    pivotTableName = GetParams("table")

    excel = GetGlobals("excel")
    xls = excel.file_[excel.actual_id]
    try:
        wb = xls['workbook']
        sht = wb.sheets[sheet].select()
        pivotTable = wb.api.ActiveSheet.PivotTables(pivotTableName)
        pivotTable.ClearAllFilters()
        
    except Exception as e:
        PrintException()
        raise e

if module == "getVisibleItems":
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
        
        pivotField = pivotTable.PivotFields(field)      
        visible = list(pivotField.VisibleItemsList)
        
        SetVar(result, visible)
        
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
        try:
            cubeFields = pivotTable.CubeFields("[{}]".format(field))
        except:
            cubeFields = pivotTable.CubeFields(field)
            
        fields = [e.Name for e in cubeFields.PivotFields]

        SetVar(result, fields)

    except Exception as e:
        PrintException()
        raise e
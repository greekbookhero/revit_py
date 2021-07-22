# -*- coding: utf-8 -*-
__doc__ = 'Внутренняя отделка'

__title__ = 'Внутренняя \n отделка'
__author__ = 'Bekzhan Zharkynbek'
__highlight__ = 'new'

import clr
import math
import collections
clr.AddReferenceByPartialName('PresentationCore')
clr.AddReferenceByPartialName("PresentationFramework")
clr.AddReferenceByPartialName('System')
clr.AddReferenceByPartialName('System.Windows.Forms')

from Autodesk.Revit import DB
from Autodesk.Revit import UI

from rpw.ui.forms import *
import xlrd
import xlwt

_eps = 1.0e-9
# creates variables for selected elements in global scope
# e1, e2, ...
# print('\n'.join(sys.path))
max_elements = 5
gdict = globals()
uidoc = __revit__.ActiveUIDocument
if uidoc:
    doc = __revit__.ActiveUIDocument.Document
    selection = [doc.GetElement(x) for x in uidoc.Selection.GetElementIds()]
    for idx, el in enumerate(selection):
        if idx < max_elements:
            gdict['e{}'.format(idx+1)] = el
        else:
            break


def alert(msg):
    TaskDialog.Show('RPS', msg)

def quit():
    __window__.Close()

if len(selection) > 0:
    el = selection[0]

def filterTypeRule(name):
    comments_param_id = DB.ElementId(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
    comments_param_prov = DB.ParameterValueProvider(comments_param_id)
    param_cont = DB.FilterStringContains()
    comments_value_rule = DB.FilterStringRule(comments_param_prov, param_cont, name, False)
    comment_filter = DB.ElementParameterFilter(comments_value_rule)
    return comment_filter

def MMToFeet(p_value):
    return p_value / 304.8

def excel_upload(sender, event):
    global level_id
    level = doc.GetElement(level_id)
    elf = DB.ElementLevelFilter(level_id)
    rooms = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Rooms).WherePasses(elf).ToElements()
    kj_name = doc.Title.Replace("AR","KJ").Replace("_отсоединено","")
    dict = {}
    for j in rooms:
        s = j.GetBoundarySegments(DB.SpatialElementBoundaryOptions())
        try:
            for i in s[0]:
                if i.ElementId.IntegerValue != -1:
                    if "(перегородки)" in doc.GetElement(i.ElementId).Name or "(внутренние)" in doc.GetElement(i.ElementId).Name or "(наружные)" in doc.GetElement(i.ElementId).Name or "(утеплитель_внутренний)" in doc.GetElement(i.ElementId).Name :
                        if dict.get(j.LookupParameter("BI_индекс_помещения").AsString()[0]+" "+j.Parameter[DB.BuiltInParameter.ROOM_NAME].AsString()):
                            if  doc.GetElement(i.ElementId).Name not in dict[j.LookupParameter("BI_индекс_помещения").AsString()+" "+j.Parameter[DB.BuiltInParameter.ROOM_NAME].AsString()]:
                                dict[j.LookupParameter("BI_индекс_помещения").AsString()[0]+" "+j.Parameter[DB.BuiltInParameter.ROOM_NAME].AsString()].append(doc.GetElement(i.ElementId).Name)
                        else:
                            dict[j.LookupParameter("BI_индекс_помещения").AsString()[0]+" "+j.Parameter[DB.BuiltInParameter.ROOM_NAME].AsString()] = [doc.GetElement(i.ElementId).Name]
                    elif kj_name in doc.GetElement(i.ElementId).Name:
                        if dict.get(j.LookupParameter("BI_индекс_помещения").AsString()[0]+" "+j.Parameter[DB.BuiltInParameter.ROOM_NAME].AsString()):
                            if  "КЖ" not in dict[j.Parameter[DB.BuiltInParameter.ROOM_NAME].AsString()]:
                                dict[j.LookupParameter("BI_индекс_помещения").AsString()[0]+" "+j.Parameter[DB.BuiltInParameter.ROOM_NAME].AsString()].append("КЖ")
                        else:
                            dict[j.LookupParameter("BI_индекс_помещения").AsString()[0]+" "+j.Parameter[DB.BuiltInParameter.ROOM_NAME].AsString()] = ["КЖ"]
        except Exception as e:
            pass
    path = select_folder()
    wb = xlwt.Workbook()
    sheet = wb.add_sheet('Типы')
    cnt = 0
    for k,v in dict.iteritems():
        sheet.write(cnt,0,k)
        sheet.col(0).width = 10000
        cnt += 1
        sheet.write(cnt,0,"")
        for i in range(len(v)):
            sheet.write(cnt-1,1,v[i])
            sheet.write(cnt-1,2,"(отделка_по_кирпичу)квартира_t=15")
            sheet.col(1).width = 15000
            sheet.col(2).width = 10000
            cnt +=1
    wb.save(path+"\\Отделка_помещения_"+str(level.Name)+".xls")

def excel_download(sender, event):
    global level_id
    dict = {}
    list_of_otdelka = []
    list_of_wall = []
    path = dialogWindow()
    try:
        book = xlrd.open_workbook(path)
    except Exception as e:
        Alert("Выберите файл", title='Ошибка', header='Нужно выбрать файл', exit=True)
    sh = book.sheet_by_index(0)
    sh = list(sh)
    type = ""
    special_rooms = ["ПУИ","Шлюз","С/У","Тамбур-шлюз"]
    kj_name = doc.Title.Replace("AR","KJ").Replace("_отсоединено","")
    plane = doc.ActiveView
    walls_types = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Walls).WhereElementIsElementType().ToElements()
    level = doc.GetElement(level_id)
    elf = DB.ElementLevelFilter(level_id)
    rooms = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Rooms).WherePasses(elf).ToElements()
    for i in range(len(sh)):
        if sh[i][0].value != "":
            j = i
            while sh[j][1].value != "":
                if dict.get(sh[i][0].value):
                    dict[sh[i][0].value][sh[j][1].value] = sh[j][2].value
                else:
                    dict[sh[i][0].value] = {}
                    dict[sh[i][0].value][sh[j][1].value] = sh[j][2].value
                if j+1 ==len(sh):
                    break
                else:
                    j+=1
    with DB.Transaction(doc, "создание отделки") as trans:
        trans.Start()
        for i in rooms:
            bb = i.get_BoundingBox(None)
            solid = GetElementSolid(i)
            roofs = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Floors).WherePasses(DB.ElementIntersectsSolidFilter(solid)).ToElements()
            height = 1000
            if len(roofs) > 0:
                for g in roofs:
                    temp = g.LookupParameter("Смещение от уровня").AsDouble() - g.LookupParameter("Толщина").AsDouble()
                    if temp < height:
                        height = temp
            else:
                height = i.LookupParameter("Полная высота").AsDouble()
            s = i.GetBoundarySegments(DB.SpatialElementBoundaryOptions())
            cp = (bb.Max+bb.Min)/2
            for j in s[0]:
                if j.ElementId.IntegerValue != -1:
                    if "(перегородки)" in doc.GetElement(j.ElementId).Name or "(внутренние)" in doc.GetElement(j.ElementId).Name or "(наружные)" in doc.GetElement(j.ElementId).Name or "(утеплитель_внутренний)" in doc.GetElement(j.ElementId).Name :
                        if dict.get(i.LookupParameter("BI_индекс_помещения").AsString()[0]+" "+i.Parameter[DB.BuiltInParameter.ROOM_NAME].AsString()):
                            if dict.get(i.LookupParameter("BI_индекс_помещения").AsString()[0]+" "+i.Parameter[DB.BuiltInParameter.ROOM_NAME].AsString()).get(doc.GetElement(j.ElementId).Name):
                                type = dict.get(i.LookupParameter("BI_индекс_помещения").AsString()[0]+" "+i.Parameter[DB.BuiltInParameter.ROOM_NAME].AsString()).get(doc.GetElement(j.ElementId).Name)
                        if not type:
                            continue
                        type_element = next((f for f in walls_types if type in f.Parameter[DB.BuiltInParameter.SYMBOL_NAME_PARAM].AsString() ), None)
                        type_width = type_element.Width/2
                        type_id = type_element.Id
                        curve = j.GetCurve()
                        level_id = i.LevelId
                        pl = DB.Plane.CreateByThreePoints(curve.GetEndPoint(0), curve.GetEndPoint(1), curve.GetEndPoint(1).Add(DB.XYZ.BasisZ))
                        lp = ProjectOnto(pl, cp)
                        dir = (cp - lp).Normalize()
                        trf = DB.Transform.CreateTranslation(dir * type_width)
                        curve = curve.CreateTransformed(trf)
                        slope = DB.Wall.Create(doc, curve, type_id, level_id, height, 0, False, False)
                        list_of_wall.append((slope, doc.GetElement(j.ElementId)))
                    elif kj_name in doc.GetElement(j.ElementId).Name:
                        if dict.get(i.LookupParameter("BI_индекс_помещения").AsString()[0]+" "+i.Parameter[DB.BuiltInParameter.ROOM_NAME].AsString()):
                            if dict.get(i.LookupParameter("BI_индекс_помещения").AsString()[0]+" "+i.Parameter[DB.BuiltInParameter.ROOM_NAME].AsString()).get("КЖ"):
                                type = dict.get(i.LookupParameter("BI_индекс_помещения").AsString()[0]+" "+i.Parameter[DB.BuiltInParameter.ROOM_NAME].AsString()).get("КЖ")
                        if not type:
                            continue
                        type_element = next((f for f in walls_types if type in f.Parameter[DB.BuiltInParameter.SYMBOL_NAME_PARAM].AsString() ), None)
                        type_width = type_element.Width/2
                        type_id = type_element.Id
                        curve = j.GetCurve()
                        level_id = i.LevelId
                        pl = DB.Plane.CreateByThreePoints(curve.GetEndPoint(0), curve.GetEndPoint(1), curve.GetEndPoint(1).Add(DB.XYZ.BasisZ))
                        lp = ProjectOnto(pl, cp)
                        dir = (cp - lp).Normalize()
                        trf = DB.Transform.CreateTranslation(dir * type_width)
                        curve = curve.CreateTransformed(trf)
                        slope = DB.Wall.Create(doc, curve, type_id, level_id, height, 0, False, False)
                else:
                    if dict.get(i.LookupParameter("BI_индекс_помещения").AsString()[0]+" "+i.Parameter[DB.BuiltInParameter.ROOM_NAME].AsString()):
                        if dict.get(i.LookupParameter("BI_индекс_помещения").AsString()[0]+" "+i.Parameter[DB.BuiltInParameter.ROOM_NAME].AsString()).get("КЖ"):
                            type = dict.get(i.LookupParameter("BI_индекс_помещения").AsString()[0]+" "+i.Parameter[DB.BuiltInParameter.ROOM_NAME].AsString()).get("КЖ")
                    if not type:
                        continue
                    type_element = next((f for f in walls_types if type in f.Parameter[DB.BuiltInParameter.SYMBOL_NAME_PARAM].AsString() ), None)
                    type_width = type_element.Width/2
                    type_id = type_element.Id
                    curve = j.GetCurve()
                    level_id = i.LevelId
                    pl = DB.Plane.CreateByThreePoints(curve.GetEndPoint(0), curve.GetEndPoint(1), curve.GetEndPoint(1).Add(DB.XYZ.BasisZ))
                    lp = ProjectOnto(pl, cp)
                    dir = (cp - lp).Normalize()
                    trf = DB.Transform.CreateTranslation(dir * type_width)
                    curve = curve.CreateTransformed(trf)
                    slope = DB.Wall.Create(doc, curve, type_id, level_id, height, 0, False, False)
        trans.Commit()
    with DB.Transaction(doc, "соединение отделки") as trans:
        trans.Start()
        for i in list_of_wall:
            DB.JoinGeometryUtils.JoinGeometry(doc, i[0], i[1])
        trans.Commit()




def dialogWindow():
    folderpath = select_file('Файл Excel (*.xls)|*.xls', 'Выберите файл excel для потолков')
    return folderpath

def GetElementSolid(p_elem):
    elemSolid = None
    geomElem = p_elem.get_Geometry(DB.Options())
    for geomObj in geomElem:
        if type(geomObj) is DB.Solid:
            elemSolid = geomObj
            if elemSolid:
                break
    return elemSolid

def IsZero(a, tolerance = _eps):
    return tolerance > abs(a)

def IsEqual(a, b, tolerance = _eps):
    return IsZero(b - a, tolerance)

def SignedDistanceTo(plane, p):
    if not IsEqual(plane.Normal.GetLength(), 1):
        msg = 'Ошбика нормали плоскости!'
        Alert(msg, title='Ошибка', header='Ошибка плоскости', exit=True)
    return plane.Normal.DotProduct(p - plane.Origin)

def ProjectOnto(plane, p):
    d = SignedDistanceTo(plane, p)
    q = p - d * plane.Normal
    if IsZero(SignedDistanceTo(plane, q)):
        return q
    else:
        msg = 'Точка лежит вне плоскости!'
        Alert(msg, title='Ошибка', header='Ошибка проецирования', exit=True)
# for i in s[0]:
#  	if i.ElementId.IntegerValue != -1:
#  		print(doc.GetElement(i.ElementId).Name+":"+str(i.GetCurve().GetEndPoint(0)))
levels = DB.FilteredElementCollector(doc).OfCategory(
            DB.BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements()
levels_dict = {}
for l in levels:
    levels_dict[l.Name] = l.Id
level_id = SelectFromList('Выберите уровень', levels_dict)
button_1 = Button('Выгрузить Excel', on_click=excel_upload)
button_2 = Button('Запустить excel', on_click=excel_download)
components = [button_1,button_2]
form = FlexForm(" ", components)
form.show()

# pt = type_width * slope.Orientation
# DB.ElementTransformUtils.MoveElement(doc, slope.Id, pt)
# try:
#     DB.JoinGeometryUtils.JoinGeometry(doc, doc.GetElement(j.ElementId), slope)
# except:
#     print(j.ElementId)

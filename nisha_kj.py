# -*- coding: utf-8 -*-
__doc__ = 'Создание ниш для КЖ'

__title__ = 'Создание\n ниш для КЖ'
__author__ = 'Zharkynbek Bekzhan'
__highlight__ = 'new'

import clr
clr.AddReferenceByPartialName('PresentationCore')
clr.AddReferenceByPartialName("PresentationFramework")
clr.AddReferenceByPartialName('System')
clr.AddReferenceByPartialName('System.Windows.Forms')
from Autodesk.Revit import DB
from Autodesk.Revit import UI

from rpw.ui.forms import *

uidoc = __revit__.ActiveUIDocument

_eps = 1.0e-9
# creates variables for selected elements in global scope
# e1, e2,
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

def MMToFeet(p_value):
    return p_value / 304.8

def IsZero(a, tolerance = _eps):
    return tolerance > abs(a)


def IsEqual(a, b, tolerance = _eps):
    return IsZero(b - a, tolerance)

def Compare_double(a, b, tolerance=_eps):
    return 0 if IsEqual(a, b, tolerance) else (-1 if a < b else 1)

def alert(msg):
    TaskDialog.Show('RPS', msg)

def quit():
    __window__.Close()

if len(selection) > 0:
    el = selection[0]

def ActivateSymbol(p_symbol):
    with DB.Transaction(p_symbol.Document, "activateSymbol") as t:
        t.Start()
        p_symbol.Activate()
        t.Commit()

def em():
    with DB.Transaction(doc, 'создание ниш') as trans:
        trans.Start()
        type_param_id = DB.ElementId(DB.BuiltInParameter.ELEM_TYPE_PARAM)
        type_param_prov = DB.ParameterValueProvider(type_param_id)
        param_ev = DB.FilterStringContains()
        type1_value_rule = DB.FilterStringRule(type_param_prov, param_ev, 'двухместная', False)
        type2_value_rule = DB.FilterStringRule(type_param_prov, param_ev, '(плита_перекрытия)', False)
        plate_filter = DB.ElementParameterFilter(type2_value_rule,False)
        inverse_filter = DB.ElementParameterFilter(type1_value_rule,True)
        two_filter = DB.ElementParameterFilter(type1_value_rule,False)
        list_of_points_one = []
        list_of_points_two = []
        list_of_points_light = []
        list_socket = []
        list_light = []
        linkInstances = DB.FilteredElementCollector(doc).OfClass(DB.RevitLinkInstance)
        walls = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Walls).WhereElementIsNotElementType().ToElements()
        plate = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Floors).WhereElementIsNotElementType().WherePasses(plate_filter).ToElements()
        for linkInstance in linkInstances:
            linkDoc = linkInstance.GetLinkDocument()
            if linkDoc:
                if linkDoc.Title.Contains("EM"):
                    em = linkDoc
                    transform = linkInstance.GetTotalTransform()
        one_type_sockets = DB.FilteredElementCollector(em).OfCategory(DB.BuiltInCategory.OST_ElectricalFixtures).WhereElementIsNotElementType().WherePasses(inverse_filter).ToElements()
        light = DB.FilteredElementCollector(em).OfCategory(DB.BuiltInCategory.OST_LightingDevices).WhereElementIsNotElementType().ToElements()
        if not socket_symbol.IsActive:
            socket_symbol.Activate()
        two_type_sockets = DB.FilteredElementCollector(em).OfCategory(DB.BuiltInCategory.OST_ElectricalFixtures).WhereElementIsNotElementType().WherePasses(two_filter).ToElements()
        for i in light:
            list_of_points_light.append(transform.OfPoint(i.Location.Point))
        for i in one_type_sockets:
            list_of_points_one.append(transform.OfPoint(i.Location.Point))
        for i in two_type_sockets:
            list_of_points_two.append(transform.OfPoint(i.Location.Point))
        for i in walls:
            bb = i.get_BoundingBox(None)
            solid = GetSolid(i)
            cur = i.Location.Curve
            for j in list_of_points_one:
                if bb.Max.Z> j.Z and bb.Min.Z<j.Z:
                    d = i.Location.Curve.Distance(DB.XYZ(j.X, j.Y, i.Location.Curve.GetEndPoint(0).Z))
                    if d <= i.Width :
                        lp1 = cur.Project(j).XYZPoint
                        r = (DB.XYZ(j.X, j.Y, 0) - DB.XYZ(lp1.X, lp1.Y, 0)).Normalize()
                        for g in solid.Faces:
                            if str(g.FaceNormal) == str(r):
                                storona = g
                        try:
                            nisha_one = doc.Create.NewFamilyInstance(storona, j, i.Location.Curve.Direction, socket_symbol)
                            list_socket.append(nisha_one.Id)
                        except Exception as e:
                            pass
            for j in list_of_points_two:
                if bb.Max.Z> j.Z and bb.Min.Z<j.Z:
                    d = i.Location.Curve.Distance(DB.XYZ(j.X, j.Y, i.Location.Curve.GetEndPoint(0).Z))
                    if d <= i.Width :
                        lp1 = cur.Project(j).XYZPoint
                        r = (DB.XYZ(j.X, j.Y, 0) - DB.XYZ(lp1.X, lp1.Y, 0)).Normalize()
                        for g in solid.Faces:
                            if str(g.FaceNormal) == str(r):
                                storona = g.Reference
                        try:
                            nisha_two = doc.Create.NewFamilyInstance(storona, j, i.Location.Curve.Direction, socket_symbol)
                            nisha_two_2 = doc.Create.NewFamilyInstance(storona, j.Add(i.Location.Curve.Direction * MMToFeet(75)), i.Location.Curve.Direction, socket_symbol)
                            list_socket.append(nisha_two.Id)
                            list_socket.append(nisha_two_2.Id)
                        except Exception as e:
                            pass
        for i in plate:
            bb = i.get_BoundingBox(None)
            solid = GetSolid(i)
            width = i.LookupParameter("Толщина").AsDouble()
            for j in list_of_points_light:
                if bb.Max.Z > j.Z and bb.Min.Z-width <= j.Z:
                    lp1 = cur.Project(j).XYZPoint
                    for g in solid.Faces:
                        if "PlanarFace" in str(type(g)):
                            if IsEqual(g.FaceNormal.Z, -1.0):
                                storona = g.Reference
                                break
                    nisha_light = doc.Create.NewFamilyInstance(storona, j, DB.XYZ(0, -1, 0), socket_symbol)
                    list_light.append(nisha_light.Id)
        trans.Commit()
    with DB.Transaction(doc, 'поправка ниш') as trans:
        trans.Start()
        for i in list_socket:
            nisha_socket = doc.GetElement(i)
            nisha_socket.LookupParameter("BI_размер_радиус").Set(diam/2)
            nisha_socket.LookupParameter("BI_размер_глубина").Set(depth)
        for i in list_light:
            light_nisha = doc.GetElement(i)
            light_nisha.LookupParameter("BI_размер_радиус").Set(light_diam/2)
            light_nisha.LookupParameter("BI_размер_глубина").Set(light_depth)
        trans.Commit()



def resize_recess(sender, event):
    global depth
    global diam
    global light_depth
    global light_diam
    if nisha_diam.value.isdecimal() and nisha_depth.value.isdecimal() and light_diam.value.isdecimal() and light_depth.value.isdecimal():
        diam = MMToFeet(int(nisha_diam.value))
        depth = MMToFeet(int(nisha_depth.value))
        light_diam = MMToFeet(int(light_diam.value))
        light_depth = MMToFeet(int(light_depth.value))
        em()
    else:
        msg = 'Убедитесь, что параметры содержат только цифры'
        Alert(msg, title='Ошибка', header='Ошибка ввода')
        return


def GetSolid(p_elem):
    elemSolid = None
    geomOptions = DB.Options()
    geomOptions.ComputeReferences = True
    geomElem = p_elem.get_Geometry(geomOptions)
    for geomObj in geomElem:
        if type(geomObj) is DB.Solid:
            elemSolid = geomObj
            if elemSolid:
                break
    return elemSolid

global socket_symbol
OpeningsFamily = DB.FilteredElementCollector(doc).OfClass(clr.GetClrType(DB.Family))
socket_family = next((f for f in OpeningsFamily if f.Name == 'Ниша для розеток'), None)
try:
    for i in socket_family.GetFamilySymbolIds():
        socket_symbol = doc.GetElement(i)
except Exception as e:
    Alert('Вам нужно импортировать семейство с названием "Ниша для розеток"', title="Ошибка", header="Нету семейства", exit=True)
nisha_diam = TextBox('Диаметр', Text="")
nisha_depth = TextBox('Глубина', Text="")
light_diam = TextBox('Диаметр', Text="")
light_depth = TextBox('Глубина', Text="")
button_ok = Button('Ok', on_click=resize_recess)
components = [Label("Введите данные для розетки"),Label("Введите сюда диаметр"), nisha_diam,Label("Введите сюда глубину"), nisha_depth, Label(""), Label("Введите данные для люстры"), Label("Введите сюда диаметр"), light_diam,Label("Введите сюда глубину"), light_depth, button_ok]
form = FlexForm("", components)
form.show()

440386

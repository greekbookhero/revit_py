# -*- coding: utf-8 -*-

__doc__ = 'Маркировка отметок'
__title__ = 'Маркировка'
__author__ = 'Bekzhan Zharkynbek'
__highlight__ = 'new'

import clr
import math
from rpw.ui.forms import *

clr.AddReferenceByPartialName('PresentationCore')
clr.AddReferenceByPartialName("PresentationFramework")
clr.AddReferenceByPartialName('System')
clr.AddReferenceByPartialName('System.Windows.Forms')
from System.Collections.Generic import List
from Autodesk.Revit import DB
from Autodesk.Revit import UI

_eps = 1.0e-9
max_elements = 5
gdict = globals()
uidoc = __revit__.ActiveUIDocument
if uidoc:
    doc = __revit__.ActiveUIDocument.Document
    selection = [doc.GetElement(x) for x in uidoc.Selection.GetElementIds()]
    for idx, el in enumerate(selection):
        if idx < max_elements:
            gdict['e{}'.format(idx + 1)] = el
        else:
            break

plane = doc.ActiveView

def alert(msg):
    TaskDialog.Show('RPS', msg)

def quit():
    window.Close()


if len(selection) > 0:
    el = selection[0]

def MMToFeet(p_value):
    return p_value / 304.8

def filterTypeRule(name):
    comments_param_id = DB.ElementId(DB.BuiltInParameter.ELEM_TYPE_PARAM)
    comments_param_prov = DB.ParameterValueProvider(comments_param_id)
    param_cont = DB.FilterStringContains()
    comments_value_rule = DB.FilterStringRule(comments_param_prov, param_cont, name, False)
    comment_filter = DB.ElementParameterFilter(comments_value_rule)
    return comment_filter

def line_inter(point, solid, x, y, z):
    line = DB.Line.CreateBound(point.Add(DB.XYZ(0,0,0)), point.Add(DB.XYZ(x, y, z)))
    line_inter = solid.IntersectWithCurve(line, DB.SolidCurveIntersectionOptions())
    return line_inter

def isVerticalRoom(bb):
    if (bb.Max.X-bb.Min.X < bb.Max.Y-bb.Min.Y):
        return True
    else:
        return False

def name_tags(room, p_doc, name_tag, plane):
    i = room
    sk = plane.SketchPlane.GetPlane()
    room_solid = GetElementSolid(i)
    bb = i.get_BoundingBox(None)
    bc = (bb.Max + bb.Min)/2
    room_point = i.Location.Point
    if not i.IsPointInRoom(bc):
        y_line = line_inter(bc, room_solid, 0 , 1000, MMToFeet(150))
        down_x_line = line_inter(bc, room_solid, -1000, 0, MMToFeet(150))
        if y_line.SegmentCount != 0:
            y_line = y_line.GetCurveSegment(0).GetEndPoint(0)
        else:
            y_line = bc
        if down_x_line.SegmentCount != 0:
            down_x_line = down_x_line.GetCurveSegment(0).GetEndPoint(0)
        else:
            down_x_line = bc
        if bc.DistanceTo(y_line) > bc.DistanceTo(down_x_line):
            pt_line = down_x_line
        else:
            pt_line = y_line
        bc = DB.XYZ(pt_line.X-2, pt_line.Y, pt_line.Z)
    if "балкон" in i.Parameter[DB.BuiltInParameter.ROOM_NAME].AsString().lower() or "лоджия" in i.Parameter[DB.BuiltInParameter.ROOM_NAME].AsString().lower():
        if isVerticalRoom(bb):
            bc = DB.XYZ(bc.X, bb.Max.Y-1.5, bc.Z)
        else:
            bc = bc
    elif(i.Area > 55):
        bc = bc.Add(DB.XYZ(0, 2, 0))
    else:
        bc = bc.Add(DB.XYZ(0, 1, 0))
    UV = ProjectInto(sk,bc)
    tag_name = p_doc.Create.NewRoomTag(DB.LinkElementId(i.Id), UV, plane.Id)
    tag_name.ChangeTypeId(name_tag)

def lounge_tags(room, p_doc, lounge_tag, plane):
    i = room
    sk = plane.SketchPlane.GetPlane()
    room_solid = GetElementSolid(i)
    bb = i.get_BoundingBox(None)
    bc = (bb.Max + bb.Min)/2
    room_point = i.Location.Point
    if not i.IsPointInRoom(bc):
        y_line = line_inter(bc, room_solid, 0 , 1000, MMToFeet(150))
        down_x_line = line_inter(bc, room_solid, -1000, 0, MMToFeet(150))
        if y_line.SegmentCount != 0:
            y_line = y_line.GetCurveSegment(0).GetEndPoint(0)
        else:
            y_line = bc
        if down_x_line.SegmentCount != 0:
            down_x_line = down_x_line.GetCurveSegment(0).GetEndPoint(0)
        else:
            down_x_line = bc
        if bc.DistanceTo(y_line) > bc.DistanceTo(down_x_line):
            pt_line = down_x_line
        else:
            pt_line = y_line
        bc = DB.XYZ(pt_line.X-2, pt_line.Y, pt_line.Z)
    if isVerticalRoom(bb):
        bc = bc.Add(DB.XYZ(0, -5, 0))
        UV = ProjectInto(sk,bc.Add(DB.XYZ(-1.5, 10, 0)))
        tag_name = p_doc.Create.NewRoomTag(DB.LinkElementId(i.Id), UV, plane.Id)
        tag_name.ChangeTypeId(lounge_tag)
    else:
        bc = bc.Add(DB.XYZ(0, -3, 0))
        UV = ProjectInto(sk,bc.Add(DB.XYZ(-1.5, 6.2, 0)))
        tag_name = p_doc.Create.NewRoomTag(DB.LinkElementId(i.Id), UV, plane.Id)
        tag_name.ChangeTypeId(lounge_tag)

def area_tags(room, p_doc, tag, plane):
    i = room
    sk = plane.SketchPlane.GetPlane()
    if i.LookupParameter("Имя").AsString() == "ВК" or i.LookupParameter("Имя").AsString() == "ЭЛ" or i.LookupParameter("Имя").AsString() == "ОВ":
        pass
    else:
        room_solid = GetElementSolid(i)
        bb = i.get_BoundingBox(None)
        bl = DB.XYZ(bb.Max.X, bb.Min.Y, bb.Min.Z)
        if not i.IsPointInRoom(bl):
            y_line = line_inter(bl, room_solid, 0 , 1000, MMToFeet(150))
            down_x_line = line_inter(bl, room_solid, -1000, 0, MMToFeet(150))
            if y_line.SegmentCount != 0:
                y_line = y_line.GetCurveSegment(0).GetEndPoint(0)
            else:
                y_line = bl
            if down_x_line.SegmentCount != 0:
                down_x_line = down_x_line.GetCurveSegment(0).GetEndPoint(0)
            else:
                down_x_line = bl
            if bl.DistanceTo(y_line) > bl.DistanceTo(down_x_line):
                pt_line = down_x_line
            else:
                pt_line = y_line
            bl = DB.XYZ(pt_line.X, pt_line.Y, pt_line.Z)
        if "лоджия" in i.LookupParameter("Имя").AsString().lower() or "балкон" in i.LookupParameter("Имя").AsString().lower() :
            bl = bl.Add( DB.XYZ(-1.7, 1.9, 0))
        elif i.Area > 60:
            bl = bl.Add( DB.XYZ(-2.9, 3.0, 0))
        else:
            bl = bl.Add( DB.XYZ(-1.5, 1, 0))
        UV2 = ProjectInto( sk, bl)
        tag_volume = p_doc.Create.NewRoomTag( DB.LinkElementId(i.Id), UV2, plane.Id)
        tag_volume.ChangeTypeId(tag)

def room_tags(p_doc, plane):
    global bc
    try:
        sk = plane.SketchPlane.GetPlane()
    except Exception as e:
        Alert('На виде есть неразмещенные комнаты', title="Ошибка", header="Неразмещенные комнаты")
    OpeningsFamily = DB.FilteredElementCollector(doc).OfClass(clr.GetClrType(DB.Family))
    floor_family = next((f for f in OpeningsFamily if f.Name == 'ECO_(марка)полы'), None)
    lounge_family = next((f for f in OpeningsFamily if f.Name == 'ECO_(марка)квартирография'), None)
    area_family = next((f for f in OpeningsFamily if f.Name == 'ECO_(марка)помещения'), None)
    name_family = next((f for f in OpeningsFamily if f.Name == 'ECO_(марка)помещения'), None)
    loggia_family = next((f for f in OpeningsFamily if f.Name == 'ECO_(марка)площадь_с_коэффициентом'), None)
    floors = DB.FilteredElementCollector(doc,plane.Id).OfCategory(DB.BuiltInCategory.OST_Floors).WherePasses(filterTypeRule("(полы)")).ToElementIds()
    floor_tags = DB.FilteredElementCollector(doc,plane.Id).OfCategory(DB.BuiltInCategory.OST_FloorTags).ToElements()
    try:
        for i in floor_family.GetFamilySymbolIds():
            floor_tag = i
    except Exception as e:
        Alert('Вам нужно импортировать семейство с названием "ECO_(марка)полы"', title="Ошибка", header="Нету семейства")
        return
    try:
        for i in lounge_family.GetFamilySymbolIds():
            lounge_tag = i
    except Exception as e:
        Alert('Вам нужно импортировать семейство с названием "ECO_(марка)площадь_с_коэффициентом""', title="Ошибка", header="Нету семейства")
        return
    try:
        for i in name_family.GetFamilySymbolIds():
            j = doc.GetElement(i)
            if j.LookupParameter("Имя"):
                if j.LookupParameter("Имя").AsValueString() == "Да" and j.LookupParameter("Площадь").AsValueString() == "Нет" and j.LookupParameter("Номер").AsValueString() == "Нет":
                    name_tag = i
    except Exception as e:
        Alert('Вам нужно импортировать семейство с названием "ECO_(марка)помещения"', title="Ошибка", header="Нету семейства")
        return
    try:
        for i in area_family.GetFamilySymbolIds():
            j = doc.GetElement(i)
            if j.LookupParameter("Площадь"):
                if j.LookupParameter("Имя").AsValueString() == "Нет" and j.LookupParameter("Площадь").AsValueString() == "Да" and j.LookupParameter("Номер").AsValueString() == "Нет":
                    area_tag = i
    except Exception as e:
         Alert('Вам нужно импортировать семейство с названием "ECO_(марка)помещения"', title="Ошибка", header="Нету семейства")
         return
    try:
        for i in loggia_family.GetFamilySymbolIds():
            loggia_tag = i
    except Exception as e:
         Alert('Вам нужно импортировать семейство с названием "ECO_(марка)площадь_с_коэффициентом"', title="Ошибка", header="Нету семейства")
         return
    rooms = DB.FilteredElementCollector(p_doc, plane.Id).OfCategory(DB.BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElementIds()
    lounge_rooms = [f for f in rooms if "гостиная" in doc.GetElement(f).Parameter[DB.BuiltInParameter.ROOM_NAME].AsString().lower()]
    loggia_rooms = [f for f in rooms if "балкон" in doc.GetElement(f).Parameter[DB.BuiltInParameter.ROOM_NAME].AsString().lower() or "лоджия" in doc.GetElement(f).Parameter[DB.BuiltInParameter.ROOM_NAME].AsString().lower()]
    rooms_without_loggia = list(set(rooms)-set(loggia_rooms))
    room_tags = DB.FilteredElementCollector(doc, plane.Id).OfCategory(DB.BuiltInCategory.OST_RoomTags).WhereElementIsNotElementType().ToElements()
    room_area_tags = [f.Room.Id for f in room_tags if f.GetTypeId() == area_tag]
    room_name_tags = [f.Room.Id for f in room_tags if f.GetTypeId() == name_tag]
    room_loggia_tags = [f.Room.Id for f in room_tags if f.GetTypeId() == loggia_tag]
    room_lounge_tags = [f.Room.Id for f in room_tags if f.GetTypeId() == lounge_tag]
    room_floor_tags = [f.TaggedLocalElementId for f in floor_tags ]
    rooms_without_area_tag = [doc.GetElement(room) for room in rooms_without_loggia if room not in room_area_tags]
    rooms_without_name_tag = [doc.GetElement(room) for room in rooms if room not in room_name_tags]
    rooms_without_loggia_tag = [doc.GetElement(room) for room in loggia_rooms if room not in room_loggia_tags]
    rooms_without_lounge_tag = [doc.GetElement(room) for room in lounge_rooms if room not in room_lounge_tags]
    rooms_without_floor_tag = [floor for floor in floors if floor not in room_floor_tags]
    trans = DB.Transaction(p_doc, "room tags")
    trans.Start()
    for i in rooms_without_lounge_tag:
        lounge_tags(i, p_doc, lounge_tag, plane)
    for i in rooms_without_name_tag:
        name_tags(i, p_doc, name_tag, plane)
    for i in rooms_without_area_tag:
        area_tags(i, p_doc, area_tag, plane)
    for i in rooms_without_loggia_tag:
        area_tags(i, p_doc, loggia_tag, plane)
    for j in rooms_without_floor_tag:
        j = p_doc.GetElement(j)
        room_solid = GetElementSolid(j)
        bb = j.get_BoundingBox(None)
        br = DB.XYZ(bb.Min.X, bb.Min.Y, bb.Min.Z)
        bc = (bb.Max+bb.Min)/2
        var = line_inter(bc, room_solid, 0, 1000, 0)
        if var.SegmentCount > 0:
            poi = var.GetCurveSegment(0).GetEndPoint(0)
            if Compare( br, poi) == 0:
                pass
            else:
                bc = poi
            bc.Add(DB.XYZ(0, 2, 0))
        else:
            var = line_inter(bc, room_solid, 0, -1000, 0)
            if var.SegmentCount > 0:
                poi = var.GetCurveSegment(0).GetEndPoint(0)
                if Compare(br, poi) == 0:
                    pass
                else:
                    bc = poi
            bc.Add(DB.XYZ(0, -2, 0))
        get_ref = DB.Reference(j)
        hor_tag = DB.TagOrientation.Horizontal
        if "балкон" in j.Parameter[DB.BuiltInParameter.ELEM_TYPE_PARAM].AsValueString().lower() or "лоджия" in j.Parameter[DB.BuiltInParameter.ELEM_TYPE_PARAM].AsValueString().lower() :
            if isVerticalRoom(bb):
                new_floor_tag = DB.IndependentTag.Create(p_doc, floor_tag, p_doc.ActiveView.Id, get_ref, False, hor_tag, bc)
            else:
                new_floor_tag = DB.IndependentTag.Create(p_doc, floor_tag, p_doc.ActiveView.Id, get_ref, False, hor_tag, br.Add(DB.XYZ(2, 1.3, 0)))
        else:
            new_floor_tag = DB.IndependentTag.Create(p_doc, floor_tag, p_doc.ActiveView.Id, get_ref, False, hor_tag, bc.Add(DB.XYZ(0, -1, 0)))
    trans.Commit()

def GetElementSolid(p_elem):
    elemSolid = None
    geomElem = p_elem.get_Geometry(DB.Options())
    for geomObj in geomElem:
        if type(geomObj) is DB.Solid:
            elemSolid = geomObj
            if elemSolid:
                break
    return elemSolid

def glass(p_doc, plane):
    OpeningsFamily = DB.FilteredElementCollector(doc).OfClass(clr.GetClrType(DB.Family))
    glass_family = next((f for f in OpeningsFamily if f.Name == 'ECO_(марка)стена'), None)
    try:
        for i in glass_family.GetFamilySymbolIds():
            j = doc.GetElement(i)
            if j.LookupParameter("Марка"):
                if j.LookupParameter("Марка").AsValueString() == "Да":
                    glass_tag = i
    except Exception as e:
        Alert('Вам нужно импортировать семейство с названием "ECO_(марка)стена"', title="Ошибка", header="Нету семейства")
        return
    glass = DB.FilteredElementCollector(doc,plane.Id).OfCategory(DB.BuiltInCategory.OST_Walls).WherePasses(filterTypeRule("(витражи)")).ToElementIds()
    glass_tags = DB.FilteredElementCollector(doc, plane.Id).OfCategory(DB.BuiltInCategory.OST_WallTags).WhereElementIsNotElementType().ToElements()
    glass_exist = [dt.TaggedLocalElementId for dt in glass_tags]
    glass_without_tags = [g for g in glass if g not in glass_exist]
    trans = DB.Transaction(p_doc,"витражи")
    trans.Start()
    for i in glass_without_tags:
        i = p_doc.GetElement(i)
        ref = DB.Reference(i)
        cur = i.Location.Curve
        lp = (cur.GetEndPoint(0) + cur.GetEndPoint(1)) / 2
        if IsEqual(abs(i.Orientation.X), 1):
            hor_tag = DB.TagOrientation.Vertical
        else:
            hor_tag = DB.TagOrientation.Horizontal
        new_ws_tag = DB.IndependentTag.Create(p_doc, glass_tag, p_doc.ActiveView.Id, ref, False, hor_tag, lp.Add(MMToFeet(300) * i.Orientation))
    trans.Commit()

def window(p_doc, plane):
    OpeningsFamily = DB.FilteredElementCollector(doc).OfClass(clr.GetClrType(DB.Family))
    window = DB.FilteredElementCollector(doc,plane.Id).WhereElementIsNotElementType().WherePasses(filterTypeRule("(окно)")).ToElementIds()
    window_tags = DB.FilteredElementCollector(doc, plane.Id).OfCategory(DB.BuiltInCategory.OST_WindowTags).WherePasses(filterTypeRule("окно")).WhereElementIsNotElementType().ToElements()
    windowsill = DB.FilteredElementCollector(p_doc,plane.Id).OfCategory(DB.BuiltInCategory.OST_Windows).WherePasses(filterTypeRule("подоконник")).ToElementIds()
    windowsill_tags = DB.FilteredElementCollector(doc, plane.Id).OfCategory(DB.BuiltInCategory.OST_WindowTags).WherePasses(filterTypeRule("подоконник")).WhereElementIsNotElementType().ToElements()
    window_exist = [dt.TaggedLocalElementId for dt in window_tags]
    windowsill_exist = [dt.TaggedLocalElementId for dt in windowsill_tags]
    window_without_tags = [w for w in window if w not in window_exist]
    windowsill_without_tags = [ws for ws in windowsill if ws not in windowsill_exist]
    el_family = next((f for f in OpeningsFamily if f.Name == 'ECO_(марка)окно'), None)
    try:
        for i in el_family.GetFamilySymbolIds():
            j = doc.GetElement(i)
            if j.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString() == '(марка)окно':
                window_symbol = j
            elif j.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString() == '(марка)подоконник':
                windowsill_symbol = j
    except Exception as e:
        Alert('Вам нужно импортировать семейство с названием "ECO_(марка)окно"', title="Ошибка", header="Нету семейства")
        return
    trans = DB.Transaction(doc, "windowsill")
    trans.Start()
    for i in windowsill_without_tags:
        windowsill = p_doc.GetElement(i)
        space = windowsill.LookupParameter("Ширина").AsDouble()/2
        bb = windowsill.get_BoundingBox(None)
        mark_pos = (bb.Max+bb.Min)/2
        ref = DB.Reference(windowsill)
        if IsEqual(abs(windowsill.FacingOrientation.X), 1):
            hor_tag = DB.TagOrientation.Vertical
        else:
            hor_tag = DB.TagOrientation.Horizontal
        new_ws_tag = DB.IndependentTag.Create(p_doc, windowsill_symbol.Id, p_doc.ActiveView.Id, ref, False, hor_tag, mark_pos.Subtract((space+MMToFeet(200)) * windowsill.FacingOrientation))
    trans.Commit()
    trans = DB.Transaction(doc, "window")
    trans.Start()
    for i in window_without_tags:
        window = p_doc.GetElement(i)
        ref = DB.Reference(window)
        for i in window.GetSubComponentIds():
            if "(откос)наружн" in p_doc.GetElement(i).Name:
                otkos = p_doc.GetElement(i)
                bb = otkos.get_BoundingBox(None)
                cp = (bb.Max + bb.Min)/2
                space = otkos.LookupParameter("BI_глубина_откоса").AsDouble()/2
        if IsEqual(abs(window.FacingOrientation.X), 1):
            hor_tag = DB.TagOrientation.Vertical
        else:
            hor_tag = DB.TagOrientation.Horizontal
        new_w_tag = DB.IndependentTag.Create(p_doc, window_symbol.Id, p_doc.ActiveView.Id, ref, False, hor_tag, cp.Add((space+MMToFeet(220)) * window.FacingOrientation))
    trans.Commit()

def door(p_doc, plane):
    OpeningsFamily = DB.FilteredElementCollector(doc).OfClass(clr.GetClrType(DB.Family))
    doors = DB.FilteredElementCollector(p_doc, plane.Id).OfCategory( DB.BuiltInCategory.OST_Doors).WherePasses(filterTypeRule("(дверь)")).ToElementIds()
    door_family = next((f for f in OpeningsFamily if f.Name == 'ECO_(марка)дверь'), None)
    door_tags = DB.FilteredElementCollector(doc, plane.Id).OfCategory(DB.BuiltInCategory.OST_DoorTags).WhereElementIsNotElementType().ToElements()
    doors_exist = [dt.TaggedLocalElementId for dt in door_tags]
    door_without_tags = list(set(doors)-set(doors_exist))
    try:
        for i in door_family.GetFamilySymbolIds():
            door_sym = i
    except Exception as e:
        Alert('Вам нужно импортировать семейство с названием "Марка.двери1"', title="Ошибка", header="Нету семейства")
        return
    trans = DB.Transaction(doc, "door tag")
    trans.Start()
    for i in door_without_tags:
        door = doc.GetElement(i)
        ref = DB.Reference(door)
        lp = door.Location.Point
        hor_tag = DB.TagOrientation.Horizontal
        new_dr_tag = DB.IndependentTag.Create(p_doc, door_sym, p_doc.ActiveView.Id, ref, False, hor_tag, lp)
        if( new_dr_tag.TagText == "ЭЛ" or new_dr_tag.TagText == "ВК" or new_dr_tag.TagText == "ОВ"):
            p_doc.Delete(new_dr_tag.Id)
    trans.Commit()

if len(selection) > 0:
	el = selection[0]

def IsZero(a, tolerance = _eps):
    return tolerance > abs(a)


def IsEqual(a, b, tolerance = _eps):
    return IsZero(b - a, tolerance)

def Compare_double(a, b, tolerance=_eps):
    return 0 if IsEqual(a, b, tolerance) else (-1 if a < b else 1)

def Compare(p, q, tolerance=_eps):
    d = Compare_double(p.X, q.X, tolerance)
    if d == 0:
        d = Compare_double(p.Y, q.Y, tolerance)
        if d == 0:
            d = Compare_double(p.Z, q.Z, tolerance)
    return d

def ProjectInto(plane, p):
    q = ProjectOnto(plane, p)
    d = q - plane.Origin
    u = d.DotProduct(plane.XVec)
    v = d.DotProduct(plane.YVec)
    return DB.UV(u, v)

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

def AngleToDegree(angle):
    return angle * (180.0 / math.pi)

def AngleToRadian(angle):
    return math.pi * angle / 180.0

def CalcPosAngleInRadians(radian):
    angle = AngleToDegree(math.acos(radian))
    if angle > 90:
        angle = 180 - angle
    return AngleToRadian(angle)

def GetHostElementHorAngle(hostElement):
    curve = hostElement.Location.Point
    direction = (curve.GetEndPoint(1) - curve.GetEndPoint(0)).Normalize()
    return CalcPosAngleInRadians(direction.X), curve.GetEndPoint(0), curve.GetEndPoint(1)

def tags():
    if "План отделочных работ" not in plane.Name:
        Alert('Активируйте окно "План отделочных работ"', title="Ошибка", header="Плане отделочных работ", exit=True)
    else:
        room_tags(doc, plane)
        window(doc, plane)
        door(doc, plane)
        glass(doc, plane)

tags()

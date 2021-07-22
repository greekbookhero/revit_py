# -*- coding: utf-8 -*-
__doc__ = "Нанесение аннотаций на виде в разрезе. \n \n 1. Проставление вертикальных размеров, вывод высотных отметок наружных проемов. \n 2. Проставление горизонтальных размеров координационных осей. \n 3. Вывод высотных отметок уровней. \n 4. Проставление вертикальных размеров по высоте помещений. \n 5. Проставление вертикальных размеров  внутренних проемов.\n \n Для нанесения одного из выбранных аннотаций разместите точку в нужное местоположение на виде."

__title__ = 'Drawing\nSection Ann'
__author__ = 'Bekzhan Zharkynbek'
__highlight__ = 'new'

import clr
import math
import System

clr.AddReferenceByPartialName('PresentationCore')
clr.AddReferenceByPartialName("PresentationFramework")
clr.AddReferenceByPartialName('System')
clr.AddReferenceByPartialName('System.Windows.Forms')

from Autodesk.Revit import DB
from Autodesk.Revit import UI
from System.Collections.Generic import List
from Autodesk.Revit.UI.Selection import ObjectType
from rpw.ui.forms import *

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
structuralType = DB.Structure.StructuralType.NonStructural
_eps = 1.0e-9

def IsZero(a, tolerance = _eps):
    return tolerance > abs(a)

def IsEqual(a, b, tolerance = _eps):
    return IsZero(b - a, tolerance)

def IsPerp(a, b):
    d = a.AngleTo(b)
    return IsEqual(d, math.pi/2)

def MMToFeet(p_value):
    return p_value / 304.8

def filterTypeRule(name):
    comments_param_id = DB.ElementId(DB.BuiltInParameter.ELEM_TYPE_PARAM)
    comments_param_prov = DB.ParameterValueProvider(comments_param_id)
    param_cont = DB.FilterStringContains()
    comments_value_rule = DB.FilterStringRule(comments_param_prov, param_cont, name, False)
    comment_filter = DB.ElementParameterFilter(comments_value_rule)
    return comment_filter

def GetWallSolid(p_elem):
    geomOptions = DB.Options()
    geomOptions.ComputeReferences = True
    geomElem = p_elem.get_Geometry(geomOptions)
    for g in geomElem:
     return g

def GetSolids(p_elem):
    ss = []
    geomOptions = DB.Options()
    geomOptions.ComputeReferences = True
    geomOptions.IncludeNonVisibleObjects = True
    geomElem = p_elem.get_Geometry(geomOptions)
    for g in geomElem:
     ge = g.GetSymbolGeometry()
     for geomObj in ge:
         if type(geomObj) is DB.Solid:
             ss.append(geomObj)
    return ss

def doors(sender, event):
    with DB.Transaction(doc, 'линейные размеры для окон справа') as tran:
        tran.Start()
        Alert('Выберите точку', title="Выбор точки" )
        sk = DB.Plane.CreateByNormalAndOrigin(doc.ActiveView.ViewDirection, doc.ActiveView.Origin)
        new_sk = DB.SketchPlane.Create(doc, sk)
        doc.ActiveView.SketchPlane = new_sk
        av = doc.ActiveView
        levels = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements()
        levels_sorted = sorted(levels, key=lambda c: c.Elevation)
        f_level = levels_sorted[0]
        l_level = levels_sorted[-1]
        try:
            picked_point5 = uidoc.Selection.PickPoint()
        except Exception as e:
            Alert("Не выбрана точка", title="Ошибка")
            return
        start_pt2 = DB.XYZ(picked_point5.X, picked_point5.Y, f_level.Elevation)
        end_pt2 = DB.XYZ(picked_point5.X, picked_point5.Y, l_level.Elevation)
        door_line = DB.Line.CreateBound(start_pt2, end_pt2)
        outline = DB.Outline(start_pt2, end_pt2)
        out_inter = DB.BoundingBoxIntersectsFilter(outline)
        roof = DB.FilteredElementCollector(doc, doc.ActiveView.Id).OfCategory(DB.BuiltInCategory.OST_Floors).WherePasses(filterTypeRule("потолки")).WhereElementIsNotElementType().ToElements()
        doors = DB.FilteredElementCollector(doc, av.Id).OfCategory(DB.BuiltInCategory.OST_Doors).WhereElementIsNotElementType().WherePasses(filterTypeRule("(откос)дверь")).WherePasses(out_inter).ToElements()
        door_ref = DB.ReferenceArray()
        if len(doors) == 0:
            Alert("На линии не найдено откоса дверей", title="Ошибка")
            return
        for l in levels_sorted:
            if [l_r for l_r in roof if l_r.LevelId == l.Id]:
                level_roof = [l_r for l_r in roof if l_r.LevelId == l.Id][0]
                solid_roof = GetWallSolid(level_roof)
                for f in solid_roof.Faces:
                    if IsEqual(f.FaceNormal.Z, 1):
                        door_ref.Append(f.Reference)
                        break
        for d in doors:
            solid = GetSolids(d)
            for f in solid[0].Faces:
                if IsEqual(f.FaceNormal.Z, 1):
                    if IsEqual(f.YVector.Y, 1):
                        door_ref.Append(f.Reference)
                if IsEqual(f.FaceNormal.Z, -1):
                    if IsEqual(f.YVector.X, 1):
                        door_ref.Append(f.Reference)
        Alert('Выберите точку для разметки дверей ', title="Выбор точки" )
        picked_point7 = uidoc.Selection.PickPoint()
        line_dim = DB.Line.CreateUnbound(picked_point7, DB.XYZ.BasisZ)
        doc.Create.NewDimension(av, line_dim, door_ref)
        tran.Commit()

def windows(sender, event):
    cnt = 1
    with DB.Transaction(doc, 'линейные размеры для окон справа') as tran:
        tran.Start()
        Alert('Выберите точку', title="Выбор точки" )
        sk = DB.Plane.CreateByNormalAndOrigin(doc.ActiveView.ViewDirection, doc.ActiveView.Origin)
        new_sk = DB.SketchPlane.Create(doc, sk)
        doc.ActiveView.SketchPlane = new_sk
        av = doc.ActiveView
        dir = av.ViewDirection
        origin = av.Origin
        xy_view = abs(origin.DotProduct(dir))
        Spot_Families = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_SpotElevSymbols).WhereElementIsElementType()
        spot_down = [f.Id for f in Spot_Families if "Вниз" in f.Parameter[DB.BuiltInParameter.SYMBOL_NAME_PARAM].AsString() ][0]
        grids = DB.FilteredElementCollector(doc, doc.ActiveView.Id).OfCategory(DB.BuiltInCategory.OST_Grids).WhereElementIsNotElementType().ToElements()
        view_ratio = int(doc.ActiveView.LookupParameter("Масштаб вида").AsValueString().Split(":")[1])
        sorted_grids = sorted(grids, key=lambda c: c.Curve.GetEndPoint(0).DotProduct(av.RightDirection))
        max = sorted_grids[-1]
        min = sorted_grids[0]
        cp = (max.Curve.GetEndPoint(0)+min.Curve.GetEndPoint(0))/2
        windows = DB.FilteredElementCollector(doc, av.Id).OfCategory(DB.BuiltInCategory.OST_Windows).WherePasses(DB.LogicalOrFilter(filterTypeRule("(проем)"),filterTypeRule("(окно)"))).WhereElementIsNotElementType().ToElements()
        doors = DB.FilteredElementCollector(doc, av.Id).OfCategory(DB.BuiltInCategory.OST_Doors).WherePasses(DB.LogicalOrFilter(filterTypeRule("(проем)"),filterTypeRule("(дверь)"))).WhereElementIsNotElementType().ToElements()
        try:
            picked_point3 = uidoc.Selection.PickPoint()
        except Exception as e:
            Alert("Не выбрана точка", title="Ошибка")
            return
        window_prep = [w for w in windows if IsPerp(w.FacingOrientation,dir) and abs(w.Location.Point.DotProduct(dir)) - w.LookupParameter("Ширина").AsDouble()/2 <= xy_view <= abs(w.Location.Point.DotProduct(dir)) + w.LookupParameter("Ширина").AsDouble()/2]
        door_prep = [d for d in doors if IsPerp(d.FacingOrientation,dir)and abs(d.Location.Point.DotProduct(dir)) - d.LookupParameter("Ширина").AsDouble()/2 <= xy_view <= abs(d.Location.Point.DotProduct(dir)) + d.LookupParameter("Ширина").AsDouble()/2]
        levels = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements()
        levels_sorted = sorted(levels, key=lambda c: c.Elevation)
        right = []
        left = []
        window_ref = DB.ReferenceArray()
        for l in levels_sorted:
            right_window = None
            left_window = None
            window_level = [w_p_l for w_p_l in window_prep if w_p_l.LevelId == l.Id]
            door_level = [d_p_l for d_p_l in door_prep if d_p_l.LevelId == l.Id]
            windows_sorted = sorted(window_level, key=lambda c: c.Location.Point.DotProduct(av.RightDirection))
            doors_sorted = sorted(door_level, key=lambda c: c.Location.Point.DotProduct(av.RightDirection))
            if len(windows_sorted) != 0:
                if cp.DotProduct(av.RightDirection) < windows_sorted[-1].Location.Point.DotProduct(av.RightDirection):
                    right_window = windows_sorted[-1]
                if cp.DotProduct(av.RightDirection) > windows_sorted[0].Location.Point.DotProduct(av.RightDirection) :
                    left_window = windows_sorted[0]
                if len(doors_sorted) != 0:
                    if max.Curve.GetEndPoint(0).DotProduct(av.RightDirection)-MMToFeet(2000) <= doors_sorted[-1].Location.Point.DotProduct(av.RightDirection) <= max.Curve.GetEndPoint(0).DotProduct(av.RightDirection) + MMToFeet(2000) :
                        if right_window:
                            if right_window.Location.Point.DotProduct(av.RightDirection) > doors_sorted[-1].Location.Point.DotProduct(av.RightDirection):
                                right.append(right_window)
                            else:
                                right.append(doors_sorted[-1])
                        else:
                            right.append(doors_sorted[-1])
                    else:
                        if right_window:
                            right.append(right_window)
                    if min.Curve.GetEndPoint(0).DotProduct(av.RightDirection)-MMToFeet(2000) >= doors_sorted[0].Location.Point.DotProduct(av.RightDirection) >= min.Curve.GetEndPoint(0).DotProduct(av.RightDirection) + MMToFeet(2000) :
                        if left_window:
                            if left_window.Location.Point.DotProduct(av.RightDirection) < doors_sorted[0].Location.Point.DotProduct(av.RightDirection):
                                left.append(left_window)
                            else:
                                left.append(doors_sorted[0])
                        else:
                            left.append(doors_sorted[0])
                    else:
                        if left_window:
                            left.append(left_window)
            else:
                if len(doors_sorted) != 0:
                    if max.Curve.GetEndPoint(0).DotProduct(av.RightDirection)-MMToFeet(2000) <= doors_sorted[-1].Location.Point.DotProduct(av.RightDirection) <= max.Curve.GetEndPoint(0).DotProduct(av.RightDirection) + MMToFeet(2000) :
                        right.append(doors_sorted[-1])
                    if min.Curve.GetEndPoint(0).DotProduct(av.RightDirection)-MMToFeet(2000) >= doors_sorted[0].Location.Point.DotProduct(av.RightDirection) >= min.Curve.GetEndPoint(0).DotProduct(av.RightDirection) + MMToFeet(2000) :
                        left.append(doors_sorted[0])
        if picked_point3.DotProduct(av.RightDirection) > cp.DotProduct(av.RightDirection):
            if len(right) == 0 :
                Alert("Не выбрана точка", title="Ошибка")
                return
            for r in right:
                wall = r.Host
                height = r.LookupParameter("Высота").AsDouble()
                up_and_down = []
                solid = GetWallSolid(wall)
                for j in solid.Faces:
                    if not IsEqual(j.FaceNormal.Z, 1) and not IsEqual(j.FaceNormal.Z, -1) and (r.Id in r.Host.GetGeneratingElementIds(j)):
                        side = j
                        break
                for e in side.EdgeLoops[0]:
                    if e.AsCurve().GetEndPoint(0).Z == e.AsCurve().GetEndPoint(1).Z :
                        up_and_down.append(e)
                        window_ref.Append(e.Reference)
                sorted_up_and_down = sorted(up_and_down, key=lambda c: c.AsCurve().GetEndPoint(0).Z)
                pt_min = DB.XYZ(picked_point3.X, picked_point3.Y, r.Location.Point.Z)
                pt_min = pt_min.Add(MMToFeet(15 * view_ratio) * av.RightDirection)
                pt_max = DB.XYZ(picked_point3.X, picked_point3.Y, r.Location.Point.Z + height)
                pt_max = pt_max.Add(MMToFeet(15 * view_ratio) * av.RightDirection)
                elev_up = doc.Create.NewSpotElevation(av, sorted_up_and_down[0].Reference, pt_min, pt_min.Add(1*av.RightDirection), pt_min, pt_min, True)
                elev_down = doc.Create.NewSpotElevation(av, sorted_up_and_down[1].Reference, pt_max, pt_max.Add(1*av.RightDirection), pt_max, pt_max, True)
                if cnt == 1:
                    spot_down = next((f for f in elev_up.GetValidTypes() if doc.GetElement(f).Parameter[DB.BuiltInParameter.SYMBOL_NAME_PARAM].AsString() == 'Вниз'), None)
                    spot_up = next((f for f in elev_up.GetValidTypes() if doc.GetElement(f).Parameter[DB.BuiltInParameter.SYMBOL_NAME_PARAM].AsString() == 'Вверх'), None)
                cnt +=1
                elev_up.ChangeTypeId(spot_up)
                elev_down.ChangeTypeId(spot_down)
                if view_ratio == 200 and round((pt_max.Z-pt_min.Z)*304.8) < 2200:
                    elev_down.ChangeTypeId(spot_up)
            window_line = DB.Line.CreateUnbound( picked_point3, DB.XYZ.BasisZ)
            dim = doc.Create.NewDimension(av,window_line, window_ref)
            if view_ratio == 200 :
                for i in dim.Segments:
                    if round(i.Value*304.8) <= 999 :
                         i.TextPosition = i.TextPosition.Add(4 * doc.ActiveView.RightDirection).Add(DB.XYZ(0,0,-3))
        else:
            if len(left) == 0 :
                Alert("Не выбрана точка", title="Ошибка")
                return
            for l in left:
                wall = l.Host
                height = l.LookupParameter("Высота").AsDouble()
                up_and_down = []
                solid = GetWallSolid(wall)
                for j in solid.Faces:
                    if not IsEqual(j.FaceNormal.Z, 1) and not IsEqual(j.FaceNormal.Z, -1) and (l.Id in l.Host.GetGeneratingElementIds(j)):
                        side = j
                        break
                for e in side.EdgeLoops[0]:
                    if e.AsCurve().GetEndPoint(0).Z == e.AsCurve().GetEndPoint(1).Z :
                        up_and_down.append(e)
                        window_ref.Append(e.Reference)
                sorted_up_and_down = sorted(up_and_down, key=lambda c: c.AsCurve().GetEndPoint(0).Z)
                pt_min = DB.XYZ(picked_point3.X, picked_point3.Y, l.Location.Point.Z)
                pt_min = pt_min.Subtract(MMToFeet(15 * view_ratio) * av.RightDirection)
                pt_max = DB.XYZ(picked_point3.X, picked_point3.Y, l.Location.Point.Z + height)
                pt_max = pt_max.Subtract(MMToFeet(15 * view_ratio) * av.RightDirection)
                elev_up = doc.Create.NewSpotElevation(av, sorted_up_and_down[0].Reference, pt_min, pt_min.Add(-1*av.RightDirection), pt_min, pt_min, True)
                elev_down = doc.Create.NewSpotElevation(av, sorted_up_and_down[1].Reference, pt_max, pt_max.Add(-1*av.RightDirection), pt_max, pt_max, True)
                if cnt == 1:
                    spot_down = next((f for f in elev_up.GetValidTypes() if doc.GetElement(f).Parameter[DB.BuiltInParameter.SYMBOL_NAME_PARAM].AsString() == 'Вниз'), None)
                    spot_up = next((f for f in elev_up.GetValidTypes() if doc.GetElement(f).Parameter[DB.BuiltInParameter.SYMBOL_NAME_PARAM].AsString() == 'Вверх'), None)
                cnt +=1
                elev_up.ChangeTypeId(spot_up)
                elev_down.ChangeTypeId(spot_down)
                if view_ratio == 200 and round((pt_max.Z-pt_min.Z)*304.8) < 2200:
                    elev_down.ChangeTypeId(spot_up)
            window_line = DB.Line.CreateUnbound( picked_point3, DB.XYZ.BasisZ)
            dim = doc.Create.NewDimension(av,window_line, window_ref)
            if view_ratio == 200 :
                for i in dim.Segments:
                    if round(i.Value*304.8) <= 999 :
                         i.TextPosition = i.TextPosition.Add(4 * doc.ActiveView.RightDirection).Add(DB.XYZ(0,0,-3))
        tran.Commit()

def floor_dimension(sender, event):
    grids = DB.FilteredElementCollector(doc, doc.ActiveView.Id).OfCategory(DB.BuiltInCategory.OST_Grids).WhereElementIsNotElementType().ToElements()
    with DB.Transaction(doc, 'линейные размеры для окон слева') as tran:
        tran.Start()
        sk = DB.Plane.CreateByNormalAndOrigin(doc.ActiveView.ViewDirection, doc.ActiveView.Origin)
        new_sk = DB.SketchPlane.Create(doc, sk)
        doc.ActiveView.SketchPlane = new_sk
        av = doc.ActiveView
        dir = av.ViewDirection
        sorted_grids = sorted(grids, key=lambda c: c.Curve.GetEndPoint(0).DotProduct(av.RightDirection))
        max = sorted_grids[-1]
        min = sorted_grids[0]
        cp = (max.Curve.GetEndPoint(0)+min.Curve.GetEndPoint(0))/2
        levels = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements()
        levels_sorted = sorted(levels, key=lambda c: c.Elevation)
        f_level = levels_sorted[0]
        l_level = levels_sorted[-1]
        Alert('Выберите точку', title="Выбор точки" )
        try:
            picked_point = uidoc.Selection.PickPoint()
        except Exception as e:
            Alert("Не выбрана точка", title="Ошибка")
            return
        start_pt = DB.XYZ(picked_point.X, picked_point.Y, f_level.Elevation)
        end_pt = DB.XYZ(picked_point.X, picked_point.Y, l_level.Elevation)
        outline = DB.Outline(start_pt, end_pt)
        out_inter = DB.BoundingBoxIntersectsFilter(outline)
        floors = DB.FilteredElementCollector(doc, av.Id).OfCategory(DB.BuiltInCategory.OST_Floors).WhereElementIsNotElementType().WherePasses(filterTypeRule("полы")).WherePasses(out_inter).ToElements()
        for i in floors:
            go = DB.Options()
            go.ComputeReferences = True
            go.IncludeNonVisibleObjects = True
            go.View = doc.ActiveView
            ge = i.get_Geometry(go)
            for g in ge:
                for f in g.Faces:
                    if IsEqual(f.FaceNormal.Z,1):
                        for j in f.EdgeLoops[0]:
                            if IsZero(dir.X):## TODO: need to change this
                                if IsEqual(j.AsCurve().Direction.X, 1) or IsEqual(j.AsCurve().Direction.X, -1):
                                    r = j.Reference
                                    break
                            else:
                                if IsEqual(j.AsCurve().Direction.Y, 1) or IsEqual(j.AsCurve().Direction.Y, -1):
                                    r = j.Reference
                                    break
                        bx = i.BoundingBox[doc.ActiveView].Max
                        pt = DB.XYZ(picked_point.X, picked_point.Y, bx.Z)
                        pt_dir = abs(pt.DotProduct(av.RightDirection))
                        if pt_dir > abs(cp.DotProduct(av.RightDirection)):
                            doc.Create.NewSpotElevation(av, r, pt, pt.Add(1*av.RightDirection), pt, pt, True)
                            break
                        else:
                            doc.Create.NewSpotElevation(av, r, pt, pt.Add(-1*av.RightDirection), pt, pt, True)
                            break
        tran.Commit()

def level_room_dimension(sender, event):
    with DB.Transaction(doc, 'линейные размеры для уровня этажа') as tran:
        tran.Start()
        Alert('Выберите точку', title="Выбор точки")
        av = doc.ActiveView
        sk = DB.Plane.CreateByNormalAndOrigin(doc.ActiveView.ViewDirection, doc.ActiveView.Origin)
        new_sk = DB.SketchPlane.Create(doc, sk)
        doc.ActiveView.SketchPlane = new_sk
        try:
            picked_point4 = uidoc.Selection.PickPoint()
        except Exception as e:
             Alert("Не выбрана точка", title="Ошибка")
             return
        line3 = DB.Line.CreateUnbound(picked_point4, DB.XYZ.BasisZ)
        line4 = DB.Line.CreateUnbound(picked_point4.Add(2 * av.RightDirection), DB.XYZ.BasisZ)
        levels = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements()
        levels_sorted = sorted(levels, key=lambda c: c.Elevation)
        floor = DB.FilteredElementCollector(doc, doc.ActiveView.Id).OfCategory(DB.BuiltInCategory.OST_Floors).WherePasses(filterTypeRule("пол")).WhereElementIsNotElementType().ToElements()
        roof = DB.FilteredElementCollector(doc, doc.ActiveView.Id).OfCategory(DB.BuiltInCategory.OST_Floors).WherePasses(filterTypeRule("потолки")).WhereElementIsNotElementType().ToElements()
        level_ref_arr = DB.ReferenceArray()
        for i in levels_sorted:
            if [l_f for l_f in floor if l_f.LevelId == i.Id]:
                level_floor = [l_f for l_f in floor if l_f.LevelId == i.Id][0]
                solid_floor = GetWallSolid(level_floor)
                for face in solid_floor.Faces:
                    if IsEqual(face.FaceNormal.Z, 1):
                        level_ref_arr.Append(face.Reference)
                        break
            if [l_r for l_r in roof if l_r.LevelId == i.Id]:
                level_roof = [l_r for l_r in roof if l_r.LevelId == i.Id][0]
                solid_roof = GetWallSolid(level_roof)
                for f in solid_roof.Faces:
                    if IsEqual(f.FaceNormal.Z, 1):
                        level_ref_arr.Append(f.Reference)
                        break
        floor_dim = doc.Create.NewDimension(av, line3, level_ref_arr)
        segments = floor_dim.Segments
        for i in segments:
            if round(i.Value*304.8) <= 1000 :
                i.TextPosition = i.TextPosition.Add(2 * doc.ActiveView.RightDirection).Add(DB.XYZ(0,0,3))
        tran.Commit()

def grid_dimension(sender,event):
    x = []
    y = []
    with DB.Transaction(doc, 'создание линейной разметки на оси') as tran:
        tran.Start()
        sk = DB.Plane.CreateByNormalAndOrigin(doc.ActiveView.ViewDirection, doc.ActiveView.Origin)
        new_sk = DB.SketchPlane.Create(doc, sk)
        doc.ActiveView.SketchPlane = new_sk
        av = doc.ActiveView
        dir = doc.ActiveView.ViewDirection
        ref_arr = DB.ReferenceArray()
        under_ref_arr = DB.ReferenceArray()
        grids = DB.FilteredElementCollector(doc, doc.ActiveView.Id).OfCategory(DB.BuiltInCategory.OST_Grids).WhereElementIsNotElementType().ToElements()
        for i in grids:
            point = i.Curve.GetEndPoint(0)
            if round(point.X) not in x:
                x.append(round(point.X))
            if round(point.Y) not in y:
                y.append(round(point.Y))
            ref_arr.Append(DB.Reference(i))
        sorted_grids = sorted(grids, key=lambda c: c.Curve.GetEndPoint(0).DotProduct(av.RightDirection))
        max = sorted_grids[-1]
        min = sorted_grids[0]
        cp = (max.Curve.GetEndPoint(0)+min.Curve.GetEndPoint(0))/2
        under_ref_arr.Append(DB.Reference(min))
        under_ref_arr.Append(DB.Reference(max))
        bb = min.BoundingBox[doc.ActiveView].Min
        bb2 = max.BoundingBox[doc.ActiveView].Min
        bb2 = DB.XYZ(bb2.X, bb2.Y, bb.Z)
        line2 = DB.Line.CreateBound(bb.Add(DB.XYZ(0, 0, 2)), bb2.Add(DB.XYZ(0, 0, 2)))
        line = DB.Line.CreateBound(bb.Add(DB.XYZ(0, 0, 5)), bb2.Add(DB.XYZ(0, 0, 5)))
        new_dim = doc.Create.NewDimension(av, line, ref_arr)
        new_dim2 = doc.Create.NewDimension(av, line2, under_ref_arr)
        tran.Commit()

button_right_window = Button('1. Размеры и отметки наружных проемов', on_click = windows )
button_grid = Button('2. Размеры координационных осей', on_click=grid_dimension)
button_floor = Button('3. Высотные отметки уровней', on_click=floor_dimension)
button_level = Button('4. Высота помещений', on_click=level_room_dimension)
button_door = Button('5. Размеры внутренних проемов', on_click=doors)
components = [ button_right_window, button_grid, button_floor, button_level, button_door]
form = FlexForm('Разметка', components)
form.WindowStartupLocation = System.Windows.WindowStartupLocation.Manual
form.show()

# -*- coding: utf-8 -*-
__doc__ = 'Образмеривание арматурных стержней на текущем виде'

__title__ = 'Образмеривание\nарматуры по середине'
__author__ = 'Bekzhan'
__highlight__ = 'new'


import clr
import math
from itertools import groupby
from System.Collections.Generic import List
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
clr.AddReferenceByPartialName('PresentationCore')
clr.AddReferenceByPartialName("PresentationFramework")
clr.AddReferenceByPartialName('System')


from Autodesk.Revit import DB
from rpw.ui.forms import *
from System.Collections.Generic import List

_eps = 1.0e-9
uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document
activeView = doc.ActiveView

def AngleToDegree(angle):
    return angle * (180.0 / math.pi)

def AngleToRadian(angle):
    return math.pi * angle / 180.0

def CalcPosAngleInRadians(radian):
    angle = AngleToDegree(math.acos(radian))
    return AngleToRadian(angle)

def MMToFeet(p_value):
    return p_value / 304.8

def IsZero(a, tolerance = _eps):
    return tolerance > abs(a)

def IsEqual(a, b, tolerance = _eps):
    return IsZero(b - a, tolerance)

def IsPerp(a, b):
    d = a.AngleTo(b)
    return IsEqual(d, math.pi/2)

def IsEqualVec(a, b, tolerance = _eps):
    return IsEqual(a.X, b.X, tolerance) and IsEqual(a.Y, b.Y, tolerance) and IsEqual(a.Z, b.Z, tolerance)

def GetRefForBar(bar, cur):
    options = DB.Options()
    options.View = activeView
    options.ComputeReferences = True
    options.IncludeNonVisibleObjects = True
    geom_ri = bar.get_Geometry(options)
    try:
        for line in geom_ri:
            print(type(line))
            if type(line) is DB.Line:
                if IsEqualVec(line.GetEndPoint(0), cur.GetEndPoint(0)) and IsEqualVec(line.GetEndPoint(1), cur.GetEndPoint(1)):
                    ref_line = line
                    break
        return ref_line
    except Exception as e:
        print(bar.Id)

def IsZero(a, tolerance = _eps):
    return tolerance > abs(a)
# OST_RebarTags Типоразмер стержня, второстепенное направление, верхняя грань
t = DB.Transaction(doc,'t')
t.Start()
OpeningsFamily = DB.FilteredElementCollector(doc).OfClass(clr.GetClrType(DB.Family))
for f in OpeningsFamily:
    if f.Name == '(Марки) Арматура':
        mark_family = f
for r_f in mark_family.GetFamilySymbolIds():
    if doc.GetElement(r_f).Parameter[DB.BuiltInParameter.SYMBOL_NAME_PARAM].AsString() == 'Марка' :
        mark_tag = doc.GetElement(r_f)
rebar_family = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MultiReferenceAnnotations).WhereElementIsNotElementType().FirstElement()
for r_f in rebar_family.GetValidTypes():
    if doc.GetElement(r_f).Parameter[DB.BuiltInParameter.SYMBOL_NAME_PARAM].AsString() == 'Зона армирования Бекжан' :
        rebar_tag = doc.GetElement(r_f)
mrao = DB.MultiReferenceAnnotationOptions(rebar_tag)
hor_tag = DB.TagOrientation.Horizontal
mrao.DimensionPlaneNormal = doc.ActiveView.ViewDirection
rebar = DB.FilteredElementCollector(doc, doc.ActiveView.Id).OfCategory(DB.BuiltInCategory.OST_AreaRein)
for i in rebar:
    riss = i.GetRebarInSystemIds()
    elem_riss = [doc.GetElement(elem) for elem in riss]
    for j in i.GetRebarInSystemIds():
        rcol = List[DB.ElementId(-1).GetType()]()
        if doc.GetElement(j).NumberOfBarPositions > 1:
            rcol.Add(j)
            bb = doc.GetElement(j).BoundingBox[doc.ActiveView]
            pt = (bb.Max+bb.Min)/2
            doc.GetElement(j).SetPresentationMode(doc.ActiveView, DB.Structure.RebarPresentationMode.FirstLast)
            cur_mr = elem_riss[0].GetCenterlineCurves(False,False,False)[0]
            cp0 = cur_mr.GetEndPoint(0)
            cp1 = cur_mr.GetEndPoint(1)
            cp = (cp0 + cp1) / 2
            cur_dir_pos = doc.ActiveView.ViewDirection
            bb_el = elem_riss[0].BoundingBox[doc.ActiveView]
            # trf_st = doc.GetElement(j).GetBarPositionTransform(0)
            # trf_end = doc.GetElement(j).GetBarPositionTransform(doc.GetElement(j).NumberOfBarPositions-1)
            # cur_cp = cur.CreateTransformed(trf_cp)
            if IsEqual(abs(elem_riss[0].Normal.X),1) or IsEqual(abs(elem_riss[0].Normal.Z),1):
                a = CalcPosAngleInRadians(elem_riss[0].Normal.X)
            elif IsEqual(abs(elem_riss[0].Normal.Y),1) :
                a = CalcPosAngleInRadians(elem_riss[0].Normal.Y)
            if (a > math.pi/2 and a <= math.pi) or (a > 3 * math.pi/2 and a < 2 * math.pi):
                mcp = pt
                cp_cp_l = mcp +MMToFeet(300) * doc.ActiveView.RightDirection
                mrao.DimensionLineOrigin = mcp
                mrao.TagHeadPosition = cp_cp_l
            else:
                mcp = pt
                cp_cp_l = mcp +MMToFeet(300) * doc.ActiveView.RightDirection
                mrao.DimensionLineOrigin = mcp
                mrao.TagHeadPosition = cp_cp_l
            mrao.SetElementsToDimension(rcol)
            mrao.DimensionLineDirection = elem_riss[0].Normal
            mra = DB.MultiReferenceAnnotation.Create(doc, activeView.Id, mrao)
        else:
            get_ref = DB.Reference(doc.GetElement(j))
            DB.IndependentTag.Create(doc, mark_tag.Id, doc.ActiveView.Id, get_ref, True, hor_tag, (doc.GetElement(j).GetCenterlineCurves(False,False,False)[0].GetEndPoint(0)/2))
t.Commit()

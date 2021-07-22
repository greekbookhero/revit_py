# -*- coding: utf-8 -*-
__doc__ = 'Маркировка плиты'

__title__ = 'Маркировка плиты'
__author__ = 'Zharkynbek Bekzhan'
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
from System.Collections.Generic import List
from Autodesk.Revit.UI.Selection import ObjectType
from rpw.ui.forms import *

el = None
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
structuralType = DB.Structure.StructuralType.NonStructural
_eps = 1.0e-9
max_elements = 5
gdict = globals()
if uidoc:
    doc = __revit__.ActiveUIDocument.Document
    selection = [doc.GetElement(x) for x in uidoc.Selection.GetElementIds()]
    for idx, el in enumerate(selection):
        if idx < max_elements:
            gdict['e{}'.format(idx+1)] = el
        else:
            break

def IsZero(a, tolerance = _eps):
    return tolerance > abs(a)

def filterTypeRule(name):
    comments_param_id = DB.ElementId(DB.BuiltInParameter.REBAR_ELEM_HOST_MARK)
    comments_param_prov = DB.ParameterValueProvider(comments_param_id)
    param_cont = DB.FilterStringContains()
    comments_value_rule = DB.FilterStringRule(comments_param_prov, param_cont, name, False)
    comment_filter = DB.ElementParameterFilter(comments_value_rule)
    return comment_filter

def rebar_classification(el):
    with DB.Transaction(doc, 'Маркировка плиты') as tran:
        tran.Start()
        order = {}
        order[0] = 'плита фон верх У'
        order[1] = 'плита фон верх Х'
        order[2] = 'плита фон низ У'
        order[3] = 'плита фон низ Х'
        order[4] = 'плита_У_верх'
        order[5] = 'плита_Х_верх'
        order[6] = 'плита_У_низ'
        order[7] = 'плита_Х_низ'
        level = int(doc.GetElement(el.LevelId).Name.Split(' ')[0])
        value = el.LookupParameter('Марка').AsString()
        if value:
            rebar = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Rebar).WherePasses(filterTypeRule(value)).WhereElementIsNotElementType().ToElements()
        else:
            Alert('У выбранной плиты нет Марки',title='Ошибка')
            return
        filter = {}
        sorted_filter = {}
        for j in rebar:
            if level >= 1:
                j.LookupParameter('BI_этаж').Set(str(level))
            else:
                under_level = 0 + level -1
                j.LookupParameter('BI_этаж').Set(str(under_level))
            try:
                item_diam = doc.GetElement(j.GetTypeId()).LookupParameter("Диаметр стержня").AsDouble()*304.8
                j.LookupParameter('BI_диаметр_арматуры').Set(item_diam/304.8)
            except Exception as e:
                form = j.LookupParameter('Форма').AsValueString()
                Alert('Удалите BI_диаметр_арматуры в семействе '+form,title='Ошибка')
                return
            try:
                item_length = j.LookupParameter('Длина стержня').AsDouble()*304.8
                j.LookupParameter('BI_длина').Set(item_length/304.8)
            except Exception as e:
                form = j.LookupParameter('Форма').AsValueString()
                Alert('Удалите BI_длина в семействе'+form,title='Ошибка')
                print(j.Name+str(j.Id))
                return
            j.LookupParameter('BI_марка_конструкции').Set(value)
            try:
                j.LookupParameter('BI_масса').Set((math.pi*(item_diam**2)/4*item_length*0.00000785))
            except Exception as e:
                form = j.LookupParameter('Форма').AsValueString()
                Alert('Удалите BI_масса в семействе'+form,title='Ошибка')
                return
            filter_var = doc.GetElement(j.GetTypeId()).LookupParameter("BI_фильтр_арматуры")
            if filter_var.AsString() not in filter and filter_var.AsString() != None :
                filter[filter_var.AsString()] = [(j)]
            if filter_var.AsString() in filter and filter_var.AsString() != None:
                filter[filter_var.AsString()].append((j))
        for k,v in filter.iteritems():
            dict = {}
            for item in v:
                value_str = str(int(round(doc.GetElement(item.GetTypeId()).LookupParameter("Диаметр стержня").AsDouble()*304.8)))+" "+'%9d' % (int(round(item.LookupParameter("Длина стержня").AsDouble()*304.8)),)
                if dict.get(value_str):
                    dict[value_str].append((item))
                else:
                    dict[value_str] = [(item)]
            sort_dict = collections.OrderedDict(sorted(dict.items()))
            sorted_filter[k] = sort_dict
        cnt = 1
        for i in range(len(order)):
            if sorted_filter.get(order[i]):
                for k,v in sorted_filter[order[i]].iteritems():
                    for i_v in v:
                        i_v.LookupParameter('Марка').Set(str(cnt))
                    cnt += 1
        tran.Commit()



if el:
    rebar_classification(el)
else:
    Alert('Не выбрана плита перекрытия',title='Ошибка')

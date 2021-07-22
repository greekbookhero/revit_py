# -*- coding: utf-8 -*-
__doc__ = 'Маркировка'

__title__ = 'Заполнение  \n проемов'
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
    comments_param_id = DB.ElementId(DB.BuiltInParameter.ELEM_TYPE_PARAM)
    comments_param_prov = DB.ParameterValueProvider(comments_param_id)
    param_cont = DB.FilterStringContains()
    comments_value_rule = DB.FilterStringRule(comments_param_prov, param_cont, name, False)
    comment_filter = DB.ElementParameterFilter(comments_value_rule)
    return comment_filter

def glass_classification(p_doc, levels, glass, pans, mulls, window_glass, door_glass):
    types_of_glass = collections.OrderedDict()
    for i in levels:
        level = i
        types_of_glass_level = {}
        glass_level = [g for g in glass if g.LevelId == level.Id]
        for i in glass_level:
            pan_temp_list = []
            mulls_temp_list = []
            window_temp_list = []
            door_temp_list = []
            pans_glass = [p for p in pans if p.Host.Id == i.Id]
            mulls_glass = [m for m in mulls if m.Host.Id == i.Id]
            window_glass_level = [wg for wg in window_glass if wg.Host.Id == i.Id]
            door_glass_level = [dg for dg in door_glass if dg.Host.Id == i.Id]
            glass_str = str(int(round(i.LookupParameter("Неприсоединенная высота").AsDouble()*304.8)))+" "+str(int(i.LookupParameter("Длина").AsDouble()*304.8))+" "+str(i.GetTypeId()) + " " + str(len(mulls_glass))
            for j in pans_glass:
                pan_temp_str = str(int(j.LookupParameter("Высота").AsDouble()*304.8))+" "+str(int(j.LookupParameter("Ширина").AsDouble()*304.8))+" "+str(j.GetTypeId())
                if pan_temp_str not in pan_temp_list:
                    pan_temp_list.append(pan_temp_str)
            pan_temp_list.sort()
            pan_temp_list.insert(0,str(len(pans_glass)))
            for j in door_glass_level:
                door_temp_str = str(int(j.LookupParameter("Высота").AsDouble()*304.8))+" "+str(int(j.LookupParameter("Ширина").AsDouble()*304.8))+" "+str(j.GetTypeId())
                if door_temp_str not in door_temp_list:
                    door_temp_list.append(door_temp_str)
            door_temp_list.sort()
            door_temp_list.insert(0,str(len(door_glass)))
            for j in window_glass_level:
                window_temp_str = str(int(j.LookupParameter("Высота").AsDouble()*304.8))+" "+str(int(j.LookupParameter("Ширина").AsDouble()*304.8))+" "+str(j.GetTypeId())
                if window_temp_str not in window_temp_list:
                    window_temp_list.append(window_temp_str)
            window_temp_list.sort()
            window_temp_list.insert(0,str(len(window_temp_list)))
            final_str = glass_str + " ".join(pan_temp_list)+" ".join(mulls_temp_list)+" ".join(door_temp_list)+" ".join(window_temp_list)
            if types_of_glass_level.get(final_str):
                types_of_glass_level[final_str].append((i, pans_glass, mulls_glass, window_glass_level , door_glass_level))
            else:
                types_of_glass_level[final_str] = [(i, pans_glass, mulls_glass, window_glass_level , door_glass_level)]
        level_class = collections.OrderedDict(sorted(types_of_glass_level.items()))
        for k, v in level_class.iteritems():
            if types_of_glass.get(k):
                for i in v:
                    types_of_glass[k].append(i)
            else:
                types_of_glass[k] = v
    return types_of_glass

def glass_mark(mark_str, glass, p_doc):
    cnt = 1
    for i in glass.values():
        for j in i:
            j[0].LookupParameter('Марка').Set(mark_str+str(cnt))
            level_id = j[0].LevelId
            level = doc.GetElement(level_id)
            if level.Name.split(" ")[0].isdigit() :
                level_str = level.Name.split(" ")[0]+" эт."
                level_param = j[0].LookupParameter(level_str)
                if level_param:
                    level_param.Set(1)
                else:
                    level_str2 = level.Name.split(" ")[0]+" эт"
                    level_param = j[0].LookupParameter(level_str2)
                    level_param.Set(1)
            else:
                level_str = level.Name.split(" ")[0]
                level_param = j[0].LookupParameter(level_str)
                level_param.Set(1)
            for p in j[1]:
                p.LookupParameter('Марка').Set(mark_str+str(cnt))
            for m in j[2]:
                m.LookupParameter('Марка').Set(mark_str+str(cnt))
            for wg in j[3]:
                wg.LookupParameter('Марка').Set(mark_str+str(cnt))
                level_id = wg.LevelId
                level = doc.GetElement(level_id)
                if level.Name.split(" ")[0].isdigit() :
                    level_str = level.Name.split(" ")[0]+" эт."
                    level_param = wg.LookupParameter(level_str)
                    if level_param:
                        level_param.Set(1)
                    else:
                        level_str2 = level.Name.split(" ")[0]+" эт"
                        level_param = wg.LookupParameter(level_str2)
                        level_param.Set(1)
                else:
                    level_str = level.Name.split(" ")[0]
                    level_param = wg.LookupParameter(level_str)
                    level_param.Set(1)
            for dg in j[4]:
                dg.LookupParameter('Марка').Set(mark_str+str(cnt))
                level_id = dg.LevelId
                level = doc.GetElement(level_id)
                if level.Name.split(" ")[0].isdigit() :
                    level_str = level.Name.split(" ")[0]+" эт."
                    level_param = dg.LookupParameter(level_str)
                    if level_param:
                        level_param.Set(1)
                    else:
                        level_str2 = level.Name.split(" ")[0]+" эт"
                        level_param = dg.LookupParameter(level_str2)
                        level_param.Set(1)
                else:
                    level_str = level.Name.split(" ")[0]
                    level_param = dg.LookupParameter(level_str)
                    level_param.Set(1)
        cnt += 1

def window(p_doc):
    types_of_window_door = []
    types_of_window = []
    window_door = DB.FilteredElementCollector(p_doc).OfCategory(DB.BuiltInCategory.OST_Windows).WherePasses(filterTypeRule("(оконно-дверной блок)")).ToElements()
    window = DB.FilteredElementCollector(p_doc).OfCategory(DB.BuiltInCategory.OST_Windows).WherePasses(filterTypeRule("(окно)")).ToElements()
    for i in window_door:
        level_id = i.LevelId
        level = doc.GetElement(level_id)
        if level.Name.split(" ")[0].isdigit() :
            level_str = level.Name.split(" ")[0]+" эт."
            level_param = i.LookupParameter(level_str)
            if level_param:
                level_param.Set(1)
            else:
                level_str2 = level.Name.split(" ")[0]+" эт"
                level_param = i.LookupParameter(level_str2)
                level_param.Set(1)
        else:
            level_str = level.Name.split(" ")[0]
            level_param = i.LookupParameter(level_str)
            level_param.Set(1)
        temp = str(i.LookupParameter("Высота").AsDouble())+ " " + str(i.LookupParameter("Ширина").AsDouble())+" "+str(i.GetTypeId())
        if temp not in types_of_window_door:
            types_of_window_door.append(temp)
    types_of_window_door.sort()
    for i in window_door:
        temp = str(i.LookupParameter("Высота").AsDouble())+ " " + str(i.LookupParameter("Ширина").AsDouble())
        for j in range(len(types_of_window_door)):
            if temp == types_of_window_door[j]:
                mark = i.LookupParameter("Марка")
                mark.Set("ОДБ-"+str(j+1))
    for i in window:
        level_id = i.LevelId
        level = doc.GetElement(level_id)
        if level.Name.split(" ")[0].isdigit() :
            level_str = level.Name.split(" ")[0]+" эт."
            level_param = i.LookupParameter(level_str)
            if level_param:
                level_param.Set(1)
            else:
                level_str2 = level.Name.split(" ")[0]+" эт"
                level_param = i.LookupParameter(level_str2)
                level_param.Set(1)
        else:
            level_str = level.Name.split(" ")[0]
            level_param = i.LookupParameter(level_str)
            level_param.Set(1)
        temp = str(i.LookupParameter("Высота").AsDouble())+ " " + str(i.LookupParameter("Ширина").AsDouble())+" "+str(i.GetTypeId())
        if temp not in types_of_window:
            types_of_window.append(temp)
    types_of_window.sort()
    for i in window:
        window_el = p_doc.GetElement(i.GetTypeId())
        temp = str(i.LookupParameter("Высота").AsDouble())+ " " + str(i.LookupParameter("Ширина").AsDouble())+" "+str(i.GetTypeId())
        for j in range(len(types_of_window)):
            if temp == types_of_window[j]:
                if p_doc.GetElement(i.GetTypeId()).LookupParameter("Маркировка типоразмера").AsString() != "ОК-"+str(j+1):
                    mark = p_doc.GetElement(i.GetTypeId()).LookupParameter("Маркировка типоразмера")
                    mark.Set("ОК-"+str(j+1))
                mark = i.LookupParameter("Марка")
                mark.Set("ОК-"+str(j+1))

def window_button(sender, event):
    trans = DB.Transaction(doc,"Окна")
    trans.Start()
    window(doc)
    trans.Commit()

def windowsill(p_doc):
    types_of_windowsill = []
    windowsill = DB.FilteredElementCollector(p_doc).OfCategory(DB.BuiltInCategory.OST_Windows).WherePasses(filterTypeRule("(подоконник)")).ToElements()
    for i in windowsill:
        level_id = i.LevelId
        level = doc.GetElement(level_id)
        if level.Name.split(" ")[0].isdigit() :
            level_str = level.Name.split(" ")[0]+" эт."
            level_param = i.LookupParameter(level_str)
            if level_param:
                level_param.Set(1)
            else:
                level_str2 = level.Name.split(" ")[0]+" эт"
                level_param = i.LookupParameter(level_str2)
                level_param.Set(1)
        else:
            level_str = level.Name.split(" ")[0]
            level_param = i.LookupParameter(level_str)
            level_param.Set(1)
        temp = str(i.LookupParameter("BI_длина").AsDouble())+ " " + str(i.LookupParameter("Ширина").AsDouble())
        if temp not in types_of_windowsill:
            types_of_windowsill.append(temp)
    types_of_windowsill.sort()
    for i in windowsill:
        temp = str(i.LookupParameter("BI_длина").AsDouble())+ " " + str(i.LookupParameter("Ширина").AsDouble())
        for j in range(len(types_of_windowsill)):
            if temp == types_of_windowsill[j]:
                mark = i.LookupParameter("Марка")
                mark.Set("ПД-"+str(j+1))

def windowsill_button(sender, event):
    trans = DB.Transaction(doc,"Подоконник")
    trans.Start()
    windowsill(doc)
    trans.Commit()

def door(p_doc):
    types_of_doors_level = []
    types_of_doors = []
    doors = DB.FilteredElementCollector(p_doc).OfCategory(DB.BuiltInCategory.OST_Doors).WherePasses(filterTypeRule("(дверь)")).ToElements()
    levels = DB.FilteredElementCollector(p_doc).OfCategory(DB.BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements()
    levels_sorted = sorted(levels, key=lambda c: c.Elevation)
    for i in levels_sorted:
        level = i
        level_doors = DB.FilteredElementCollector(p_doc).OfCategory(DB.BuiltInCategory.OST_Doors).WherePasses(filterTypeRule("(дверь)")).WherePasses(DB.ElementLevelFilter(level.Id)).ToElements()
        for j in level_doors:
            if p_doc.GetElement(i.GetTypeId()).LookupParameter("Маркировка типоразмера"):
                if p_doc.GetElement(i.GetTypeId()).LookupParameter("Маркировка типоразмера").AsString() != "ЭЛ":
                    temp = str(j.LookupParameter("Высота").AsDouble())+ " " + str(j.LookupParameter("Ширина").AsDouble())+" "+str(j.GetTypeId())
                    if temp not in types_of_doors_level:
                        types_of_doors_level.append(temp)
            else:
                temp = str(j.LookupParameter("Высота").AsDouble())+ " " + str(j.LookupParameter("Ширина").AsDouble())+" "+str(j.GetTypeId())
                if temp not in types_of_doors_level:
                    types_of_doors_level.append(temp)
        types_of_doors_level.sort()
        types_of_doors = types_of_doors + [i for i in types_of_doors_level if i not in types_of_doors]
        for i in doors:
            level_id = i.LevelId
            level = doc.GetElement(level_id)
            if level.Name.split(" ")[0].isdigit() :
                level_str = level.Name.split(" ")[0]+" эт."
                level_param = i.LookupParameter(level_str)
                if level_param:
                    level_param.Set(1)
                else:
                    level_str2 = level.Name.split(" ")[0]+" эт"
                    level_param = i.LookupParameter(level_str2)
                    level_param.Set(1)
            else:
                level_str = level.Name.split(" ")[0]
                level_param = i.LookupParameter(level_str)
                level_param.Set(1)
        for i in doors:
            temp = str(i.LookupParameter("Высота").AsDouble())+ " " + str(i.LookupParameter("Ширина").AsDouble())+" "+str(i.GetTypeId())
            for j in range(len(types_of_doors)):
                if temp == types_of_doors[j]:
                    if p_doc.GetElement(i.GetTypeId()).LookupParameter("Маркировка типоразмера").AsString() !="ЭЛ":
                        mark = p_doc.GetElement(i.GetTypeId()).LookupParameter("Маркировка типоразмера")
                        mark.Set("Д"+str(j+1))
                    mark = i.LookupParameter("Марка")
                    mark.Set("Д"+str(j+1))

def door_button(sender, event):
    trans = DB.Transaction(doc,"дверь")
    trans.Start()
    door(doc)
    trans.Commit()

def vitrazhi(p_doc):
    glass_inner = DB.FilteredElementCollector(p_doc).OfCategory(DB.BuiltInCategory.OST_Walls).WherePasses(filterTypeRule("(витражи)внутренние")).ToElements()
    glass_outter = DB.FilteredElementCollector(p_doc).OfCategory(DB.BuiltInCategory.OST_Walls).WherePasses(filterTypeRule("(витражи)наружные")).ToElements()
    door_glass = DB.FilteredElementCollector(p_doc).OfCategory(DB.BuiltInCategory.OST_Doors).WherePasses(filterTypeRule("(панель_витража)дверь")).ToElements()
    window_glass = DB.FilteredElementCollector(p_doc).OfCategory(DB.BuiltInCategory.OST_Windows).WherePasses(filterTypeRule("(панель_витража)окно")).ToElements()
    levels = DB.FilteredElementCollector(p_doc).OfCategory(DB.BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements()
    pans = DB.FilteredElementCollector(p_doc).OfCategory(DB.BuiltInCategory.OST_CurtainWallPanels).WhereElementIsNotElementType().ToElements()
    mulls = DB.FilteredElementCollector(p_doc).OfCategory(DB.BuiltInCategory.OST_CurtainWallMullions).WhereElementIsNotElementType().ToElements()
    levels_sorted = sorted(levels, key=lambda c: c.Elevation)
    glass_inner_classification = glass_classification(p_doc, levels_sorted, glass_inner, pans, mulls, window_glass, door_glass)
    glass_outter_classification = glass_classification(p_doc, levels_sorted, glass_outter, pans, mulls, window_glass, door_glass)
    glass_mark("ВВ-",glass_inner_classification, p_doc)
    glass_mark("ВН-",glass_outter_classification, p_doc)

def vitrazhi_button(sender, event):
    trans = DB.Transaction(doc,"Витраж")
    trans.Start()
    vitrazhi(doc)
    trans.Commit()

button_1 = Button('Окна', on_click=window_button)
button_2 = Button('Подоконни', on_click=windowsill_button)
button_3 = Button('Двери', on_click=door_button)
button_4 = Button('Витражи', on_click=vitrazhi_button)
components = [Label("Кнопки"),button_1,button_2,button_3,button_4]
form = FlexForm("Заполнение проемов", components)
form.show()

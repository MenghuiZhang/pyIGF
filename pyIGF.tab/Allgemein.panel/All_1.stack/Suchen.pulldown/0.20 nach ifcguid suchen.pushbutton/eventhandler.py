# coding: utf8
from Autodesk.Revit.UI import IExternalEventHandler,ExternalEvent,TaskDialog
import Autodesk.Revit.DB as DB
from System.Windows import Visibility 

def IFC_Filter(ifc):
    param_equality=DB.FilterStringEquals()
    param_id = DB.ElementId(DB.BuiltInParameter.IFC_GUID)
    param_prov=DB.ParameterValueProvider(param_id)
    param_value_rule=DB.FilterStringRule(param_prov,param_equality,ifc,True)
    param_filter = DB.ElementParameterFilter(param_value_rule)
    return param_filter


class ANZEIGEN(IExternalEventHandler):
    def __init__(self):
        self.ifcguid = None
        self.ifcfilter = IFC_Filter
        
    def Execute(self,app):
        uidoc = app.ActiveUIDocument
        doc = uidoc.Document
        if not self.ifcguid:
            TaskDialog.Show('Fehler','keine IFC-GUID') 
            return
        try:
            elemid = DB.FilteredElementCollector(doc).WherePasses(self.ifcfilter(self.ifcguid)).WhereElementIsNotElementType().ToElementIds()[0]
            elem = doc.GetElement(elemid)
        except:
            TaskDialog.Show('Fehler','ungültige IFC-GUID') 
            return
        if not elem:
            TaskDialog.Show('Fehler','falsche IFC-GUID') 
            return
            
        sel = uidoc.Selection.GetElementIds()
        sel.Clear()
        sel.Add(elem.Id)
        uidoc.Selection.SetElementIds(sel)
        uidoc.ShowElements(elem)


    def GetName(self):
        return "Element anzeigen"

class SELECT(IExternalEventHandler):
    def __init__(self):
        self.ifcguid = None
        self.ifcfilter = IFC_Filter
      
        
    def Execute(self,app):
        uidoc = app.ActiveUIDocument
        doc = uidoc.Document
        if not self.ifcguid:
            TaskDialog.Show('Fehler','keine IFC-GUID') 
            return
        try:
            elemid = DB.FilteredElementCollector(doc).WherePasses(self.ifcfilter(self.ifcguid)).WhereElementIsNotElementType().ToElementIds()[0]
            elem = doc.GetElement(elemid)
        except:
            TaskDialog.Show('Fehler','ungültige IFC-GUID') 
            return
        if not elem:
            TaskDialog.Show('Fehler','falsche IFC-GUID') 
            return
            
        sel = uidoc.Selection.GetElementIds()
        sel.Clear()
        sel.Add(elem.Id)
        uidoc.Selection.SetElementIds(sel)

    def GetName(self):
        return "Element auswählen"
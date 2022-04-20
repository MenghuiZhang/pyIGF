# coding: utf8
from Autodesk.Revit.UI import IExternalEventHandler,ExternalEvent,TaskDialog
import Autodesk.Revit.DB as DB
from pyrevit import revit
import os


class FamilyLoadOptions(DB.IFamilyLoadOptions):
    def OnFamilyFound(self,familyInUse, overwriteParameterValues = False): return True
    def OnSharedFamilyFound(self,familyInUse, source, overwriteParameterValues = False): return True

class Familien(object):
    def __init__(self,elem,name):
        self.elem = elem
        self.name = name

def Get_IS():
    _dict = {}
    Liste = []
    FamilySymbols = DB.FilteredElementCollector(revit.doc).OfCategory(DB.BuiltInCategory.OST_DuctFitting).WhereElementIsElementType().ToElements()
    for el in FamilySymbols:
        famname = el.FamilyName
        if famname not in _dict.keys():
            _dict[famname] = el.Family
    for el in sorted(_dict.keys()):
        Liste.append(Familien(_dict[el],el))
    return Liste

LISTE_IS = Get_IS()

class CHANGEFAMILY(IExternalEventHandler):
    def __init__(self):
        self.Desktop = os.path.expanduser("~\Desktop")
        
        self.Kombi = None
        self.Familien_Name = 'DistributionBox_'
        self.Vorlage = None
        self.Path = None
        self.modell = True
        self.ansicht = False
        self.auswahl = False
        self.elems = None
        self.startnum = 100

        
    def Execute(self,app):
        if not self.Kombi or self.Kombi == '':
            TaskDialog.Show('Fehler','Bitte zuerst Kombirahmen-Familie auswählen!')
            return  
        if not self.Vorlage:
            TaskDialog.Show('Fehler','Bitte zuerst DistributionBox-Familie auswählen!')
            return  
        uidoc = app.ActiveUIDocument
        doc = uidoc.Document
        param_equality = DB.FilterStringContains()
        Fam_name_id = DB.ElementId(DB.BuiltInParameter.ELEM_FAMILY_PARAM)
        Fam_name_prov=DB.ParameterValueProvider(Fam_name_id)
        Fam_name_value_rule=DB.FilterStringRule(Fam_name_prov,param_equality,self.Kombi,True)
        Filter = DB.ElementParameterFilter(Fam_name_value_rule)
        if self.modell:
            self.elems = DB.FilteredElementCollector(doc).WherePasses(Filter).WhereElementIsNotElementType().ToElementIds()
        elif self.ansicht:
            self.elems = DB.FilteredElementCollector(doc,uidoc.ActiveView.Id).WherePasses(Filter).WhereElementIsNotElementType().ToElementIds()
        else:pass
        if self.elems.Count == 0:
            TaskDialog.Show('Info.','Keine Komnirahmen gefunden! Haben Sie Button (Auswahl Aktualisieren) gedruckt?')
            return
        
        
        Familien_Liste = []
        for n,item in enumerate(self.elems):
            self.startnum += 1
            self.Path = self.Desktop+'\\' +self.Familien_Name+str(self.startnum)+'.rfa'
            famdoc = doc.EditFamily(self.Vorlage)
            famdoc.SaveAs(self.Path)
            try:
                family = famdoc.LoadFamily(doc, FamilyLoadOptions())
                Familien_Liste.append(family)
            except:pass
            famdoc.Close(False)
            os.remove(self.Path)
        t = DB.Transaction(doc,'change Bauteiltyp')
        t.Start()
        for n,item in enumerate(self.elems):

            try:
                famid = list(Familien_Liste[n].GetFamilySymbolIds())[0]
                doc.GetElement(item).ChangeTypeId(famid)
                famtyp = doc.GetElement(famid)
                h = int(round(float(doc.GetElement(item).LookupParameter('MC Height Instance').AsValueString()),0))
                l = int(round(float(doc.GetElement(item).LookupParameter('MC Length Instance').AsValueString()),0))
                w = int(round(float(doc.GetElement(item).LookupParameter('MC Width Instance').AsValueString()),0))
                famtyp.Name = str(2*l) +'x'+str(w)+'x'+str(h)
            except:pass

        t.Commit()
        t.Dispose()
        TaskDialog.Show('Fertig','Erledigt!')

    def GetName(self):
        return "Bauteil austauschen"

class Aktualisieren(IExternalEventHandler):
    def __init__(self):
        self.Auswahl = []
        self.Muster = None
        
    def Execute(self,app):    
        if not self.Muster or self.Muster == '':
            TaskDialog.Show('Fehler','Bitte zuerst Kombirahmen-Familie auswählen!')
            return
        uidoc = app.ActiveUIDocument
        doc = uidoc.Document
        Auswahl = uidoc.Selection.GetElementIds()
        Liste = []
        for el in Auswahl:
            elem = doc.GetElement(el)
            try:
                if elem.Symbol.FamilyName == self.Muster:
                    Liste.append(el)
            except:pass
        Auswahl.Clear()
        self.Auswahl = Liste
        if len(Liste) == 0:
            uidoc.Selection.SetElementIds(Auswahl)
            TaskDialog.Show('Fertig','keine gefunden!')
            return
        else:
            for el in Liste:
                Auswahl.Add(el)
                uidoc.Selection.SetElementIds(Auswahl)
          
        TaskDialog.Show('Fertig','Auswahl aktualisiert!')

    def GetName(self):
        return "Auswahl aktualsieren"



class NUMMERIEREN(IExternalEventHandler):
    def __init__(self):
        self.Desktop = os.path.expanduser("~\Desktop")
        self.Familien_Name = 'DistributionBox_'
        self.Vorlage = None
        self.Path = None
        self.elems = None
        self.startnum = 100
        self.zeit = None

        
    def Execute(self,app):
        uidoc = app.ActiveUIDocument
        doc = uidoc.Document
        param_equality = DB.FilterStringContains()
        Fam_name_id = DB.ElementId(DB.BuiltInParameter.ELEM_FAMILY_PARAM)
        Fam_name_prov=DB.ParameterValueProvider(Fam_name_id)
        Fam_name_value_rule=DB.FilterStringRule(Fam_name_prov,param_equality,self.Familien_Name,True)
        Filter = DB.ElementParameterFilter(Fam_name_value_rule)
        self.elems = DB.FilteredElementCollector(doc).WherePasses(Filter).WhereElementIsNotElementType().ToElements()
        kopiert = []
        vorhanden = []
        for el in self.elems:
            famname = el.Symbol.FamilyName
            if not famname in vorhanden:
                vorhanden.append(famname)
                continue
            else:kopiert.append(el)

        Familien_Liste = []
        for item in enumerate(kopiert):
            self.startnum += 1
            self.Path = self.Desktop+'\\' +self.Familien_Name+str(self.startnum)+'.rfa'
            famdoc = doc.EditFamily(self.Vorlage)
            famdoc.SaveAs(self.Path)
            try:
                family = famdoc.LoadFamily(doc, FamilyLoadOptions())
                Familien_Liste.append(family)
            except:pass
            famdoc.Close(False)
            os.remove(self.Path)
        t = DB.Transaction(doc,'neu nummerieren')
        t.Start()
        for n,item in enumerate(kopiert):
            try:
                famid = list(Familien_Liste[n].GetFamilySymbolIds())[0]
                item.ChangeTypeId(famid)
                famtyp = doc.GetElement(famid)
                h = int(round(float(item.LookupParameter('MC Height Instance').AsValueString()),0))
                l = int(round(float(item.LookupParameter('MC Length Instance').AsValueString()),0))
                w = int(round(float(item.LookupParameter('MC Width Instance').AsValueString()),0))
                famtyp.Name = str(2*l) +'x'+str(w)+'x'+str(h)
            except:pass

        t.Commit()
        t.Dispose()
        bericht = open(self.Desktop+"\\DistributionBox_Nummerierung.txt","a")
        bericht.write("\n[{}]".format(self.zeit))
        elemid = ''
        for el in kopiert:
            elemid += el.Id.ToString() + ', '
        bericht.write("\n{}".format(elemid))
        bericht.close()

        TaskDialog.Show('Fertig','Erledigt! neue nummerierten Elements in Datei Desktop\DistributionBox_Nummerierung.txt exportiert.')

    def GetName(self):
        return "DistributionBox nummerieren"

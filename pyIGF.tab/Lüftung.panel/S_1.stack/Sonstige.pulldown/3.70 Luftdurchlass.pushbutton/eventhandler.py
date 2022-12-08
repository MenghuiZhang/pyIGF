# coding: utf8
from Autodesk.Revit.UI import IExternalEventHandler,ExternalEvent,TaskDialog
import Autodesk.Revit.DB as DB
from pyrevit import forms
import time


class CHANGEPARAM(IExternalEventHandler):
    def __init__(self):
        self.class_GUI = None

    def Execute(self,uiapp):
        if self.class_GUI.In.SelectedIndex == -1:
            TaskDialog.Show('Fehler','Keinen Input-Parameter ausgewählt!')
            return  
        if self.class_GUI.Aus.SelectedIndex == -1:
            TaskDialog.Show('Fehler','Keinen Output-Parameter ausgewählt!')
            return  


        uidoc = uiapp.ActiveUIDocument
        doc = uidoc.Document

        inparam = self.class_GUI.In.SelectedItem.ToString()
        outparam = self.class_GUI.Aus.SelectedItem.ToString()

        if inparam == outparam:
            TaskDialog.Show('Fehler','Der Input-Parameter ist gleich mit den Output-Parameter!')
            return 

        if self.class_GUI.modell.IsChecked:
            elems = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_DuctTerminal).WhereElementIsNotElementType().ToElementIds()
        elif self.class_GUI.ansicht.IsChecked:
            elems = DB.FilteredElementCollector(doc,uidoc.ActiveView.Id).OfCategory(DB.BuiltInCategory.OST_DuctTerminal).WhereElementIsNotElementType().ToElementIds()
        elif self.class_GUI.auswahl.IsChecked:
            Auswahl = uidoc.Selection.GetElementIds()
            Liste = []
            for el in Auswahl:
                elem = doc.GetElement(el)
                try:
                    if elem.Category.Id.ToString() == '-2008013':
                        Liste.append(el)
                except:pass
            elems = Liste
        if elems.Count == 0:
            TaskDialog.Show('Info.','Keinen Luftauslass gefunden!')
            return 
        
        t = DB.Transaction(doc,'Volumenstrom übertragen')
        t.Start()
        self.class_GUI.progress.Maximum = elems.Count
        self.class_GUI.progress.Minimum = 0
        n = 0
        step = 10
        for elid in elems:
            n+=1
            if n % step == 0:
                self.class_GUI.title_p.Text = str(n) + ' / '+ str(elems.Count)
                print(n)
            self.class_GUI.progress.Value = n

            el = doc.GetElement(elid)
            min_vol = el.LookupParameter('IGF_RLT_AuslassVolumenstromMin')
            max_vol = el.LookupParameter('IGF_RLT_AuslassVolumenstromMax')
            nacht_vol = el.LookupParameter('IGF_RLT_AuslassVolumenstromNacht')
            tiefenacht_vol = el.LookupParameter('IGF_RLT_AuslassVolumenstromTiefeNacht')
            vol = el.get_Parameter(DB.BuiltInParameter.RBS_DUCT_FLOW_PARAM)

            if inparam == 'IGF_RLT_AuslassVolumenstromMin':
                inputparam = min_vol
            elif inparam == 'IGF_RLT_AuslassVolumenstromMax':
                inputparam = max_vol
            elif inparam == 'IGF_RLT_AuslassVolumenstromNacht':
                inputparam = nacht_vol
            elif inparam == 'IGF_RLT_AuslassVolumenstromTiefeNacht':
                inputparam = tiefenacht_vol
            elif inparam == 'Volumenstrom':
                inputparam = vol

            if outparam == 'IGF_RLT_AuslassVolumenstromMin':
                outputparam = min_vol
            elif outparam == 'IGF_RLT_AuslassVolumenstromMax':
                outputparam = max_vol
            elif outparam == 'IGF_RLT_AuslassVolumenstromNacht':
                outputparam = nacht_vol
            elif outparam == 'IGF_RLT_AuslassVolumenstromTiefeNacht':
                outputparam = tiefenacht_vol
            elif outparam == 'Volumenstrom':
                outputparam = vol
            outputparam.Set(inputparam.AsDouble())
        t.Commit()
        t.Dispose()

    def GetName(self):
        return "Volumenstrom übertragen"
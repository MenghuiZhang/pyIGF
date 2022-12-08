# coding: utf8
from Autodesk.Revit.UI import IExternalEventHandler,ExternalEvent,TaskDialog,TaskDialogCommonButtons,TaskDialogResult 
import Autodesk.Revit.DB as DB
from pyrevit import revit
import os
from System.Collections.Generic import List
import System

# class FailureHandler(DB.IFailuresPreprocessor):
#     def __init__(self):
#         self.ErrorMessage = []
#         self.ErrorSeverity = []
 
#     def PreprocessFailures(self,failuresAccessor):
#         failureMessages = failuresAccessor.GetFailureMessages()
#         for failureMessageAccessor in failureMessages:
#             Id = failureMessageAccessor.GetFailureDefinitionId()
#             self.ErrorMessage.append(failuresAccessor.GetAttemptedResolutionTypes(failureMessageAccessor))
#             self.ErrorSeverity.append(Id)
#             # if Id == DB.BuiltInFailures.SystemsFailures.FamilyDoesntMatchSystemPropertiesWasDisconnected:
#             #     failureMessageAccessor.SetCurrentResolutionType(DB.FailureResolutionType.Others)
#             #     self.ErrorMessage = failuresAccessor.GetAttemptedResolutionTypes(failureMessageAccessor)
#                 # failureMessageAccessor.SetCurrentResolutionType(DB.FailureResolutionType.FixElements)
#                 # print(failuresAccessor.GetAttemptedResolutionTypes(failureMessageAccessor))
#                 # failuresAccessor.failureMessageAccessor(failureMessageAccessor)
                
#                 # failuresAccessor.DeleteWarning(failureMessageAccessor)

#         return DB.FailureProcessingResult.ProceedWithCommit


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
        self.Desktop = 'C:\\temp'
        self.Familien_Name = 'DistributionBox_'
        self.Path = None
        self.class_GUI = None

        
    def Execute(self,uiapp):
        
        if self.class_GUI.kombi.SelectedIndex == -1: 
            TaskDialog.Show('Fehler','Bitte zuerst Kombirahmen-Familie auswählen!')
            return  
        if self.class_GUI.distribution.SelectedIndex == -1:
            TaskDialog.Show('Fehler','Bitte Familie "DistributionBox_IGF" laden!')
            return  
        uidoc = uiapp.ActiveUIDocument
        # app = uiapp.Application
        doc = uidoc.Document
        param_equality = DB.FilterStringEquals()
        Fam_name_id = DB.ElementId(DB.BuiltInParameter.ELEM_FAMILY_PARAM)
        Fam_name_prov=DB.ParameterValueProvider(Fam_name_id)
        Fam_name_value_rule=DB.FilterStringRule(Fam_name_prov,param_equality,self.class_GUI.kombi.SelectedItem.name,True)
        Filter = DB.ElementParameterFilter(Fam_name_value_rule)
        if self.class_GUI.modell.IsChecked:
            elems = DB.FilteredElementCollector(doc).WherePasses(Filter).WhereElementIsNotElementType().ToElementIds()
        elif self.class_GUI.ansicht.IsChecked:
            elems = DB.FilteredElementCollector(doc,uidoc.ActiveView.Id).WherePasses(Filter).WhereElementIsNotElementType().ToElementIds()
        elif self.class_GUI.auswahl.IsChecked:
            Auswahl = uidoc.Selection.GetElementIds()
            Liste = []
            for el in Auswahl:
                elem = doc.GetElement(el)
                try:
                    if elem.Symbol.FamilyName == self.class_GUI.kombi.SelectedItem.name:
                        Liste.append(el)
                except:pass
            elems = Liste
        if elems.Count == 0:
            TaskDialog.Show('Info.','Keine Komnirahmen gefunden!')
            return
        else:
            task = TaskDialog.Show('Abfrage','Komnirahmen: '+str(elems.Count)+'\nBauteil wechseln?',TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No)
            if task == TaskDialogResult.No:
                return   
        
        self.class_GUI.austausch.Visibility = self.class_GUI.hidden
        self.class_GUI.progress_aus.Visibility = self.class_GUI.visible
        self.class_GUI.title_p_aus.Visibility = self.class_GUI.visible

        try:    
        
            Familien_Liste = []
            famdoc = doc.EditFamily(self.class_GUI.distribution.SelectedItem.elem)

            self.class_GUI.maxvalue = elems.Count
            self.class_GUI.minvalue = 0
            self.class_GUI.value = 1

            for n,item in enumerate(elems):
                self.class_GUI.PB_text = str(n+1) + ' / '+ str(elems.Count) + ' DistributionBox laden'
                self.class_GUI.value = n+1
                self.class_GUI.progress_aus.Dispatcher.Invoke(System.Action(self._update_pbar),
                                    System.Windows.Threading.DispatcherPriority.Background)
                self.class_GUI.startnum += 1
                self.Path = self.Desktop+'\\' +self.Familien_Name+str(self.class_GUI.startnum)+'.rfa'
                famdoc.SaveAs(self.Path)
                try:
                    family = famdoc.LoadFamily(doc, FamilyLoadOptions())
                    Familien_Liste.append(family)
                except:pass
                os.remove(self.Path)
            famdoc.Close(False)
            elems_neu = []
            self.class_GUI.value = 1

            # app.FailuresProcessing += System.EventHandler[DB.Events.FailuresProcessingEventArgs]()

            t = DB.Transaction(doc,'Bauteiltyp wechseln')
            # opt = t.GetFailureHandlingOptions()
            # hand = FailureHandler()
            
            # opt.SetFailuresPreprocessor(hand)
            # t.SetFailureHandlingOptions(opt)
            t.Start()
            for n,item in enumerate(elems):
                self.class_GUI.PB_text = str(n+1) + ' / '+ str(elems.Count) + ' Bauteile wechseln'
                self.class_GUI.value = n+1
                self.class_GUI.progress_aus.Dispatcher.Invoke(System.Action(self._update_pbar),
                                    System.Windows.Threading.DispatcherPriority.Background)
                famid = list(Familien_Liste[n].GetFamilySymbolIds())[0]
                elem_item = doc.GetElement(item)
                elem_item.Pinned = False
                try:
                    elem_item.ChangeTypeId(famid)
                    doc.Regenerate()

                    elems_neu.append(elem_item)
                    try:
                        famtyp = doc.GetElement(famid)
                        h0 = int(round(float(elem_item.LookupParameter('MC Height Instance').AsValueString()),0))
                        w0 = int(round(float(elem_item.LookupParameter('MC Width Instance').AsValueString()),0))
                        elem_item.LookupParameter('MC Width Instance').AsValueString()
                        elem_item.LookupParameter('MC Height Instance').SetValueString(str(h0+40))
                        elem_item.LookupParameter('MC Width Instance').SetValueString(str(w0+40))
                        h = h0 + 40 
                        l = int(round(float(elem_item.LookupParameter('MC Length Instance').AsValueString()),0))
                        w = w0 + 40
                        famtyp.Name = str(2*l) +'x'+str(w)+'x'+str(h)
                    except Exception as e:print(e)
                except Exception as e:
                    print(e)
                    print('Austausch von Bauteil {} gescheitert!'.format(item.ToString()))

            t.Commit()
            t.Dispose()

            self.class_GUI.maxvalue = elems_neu.Count
            self.class_GUI.value = 1
            t = DB.Transaction(doc,' Bauteile Spiegeln')

            for n,item in enumerate(elems_neu):
                self.class_GUI.PB_text = str(n+1) + ' / '+ str(elems_neu.Count) + ' Bauteile Spiegeln'
                self.class_GUI.value = n+1
                self.class_GUI.progress_aus.Dispatcher.Invoke(System.Action(self._update_pbar),
                                    System.Windows.Threading.DispatcherPriority.Background)

                if item.Pinned:
                    item.Pinned = False
                if item.Mirrored == False:
                    continue
                
                t.Start()
                transform = item.GetTransform()
                plane_XZ = DB.Plane.CreateByOriginAndBasis(transform.Origin,transform.BasisX,transform.BasisZ)
                plane_XY = DB.Plane.CreateByOriginAndBasis(transform.Origin,transform.BasisX,transform.BasisY)
                plane_YZ = DB.Plane.CreateByOriginAndBasis(transform.Origin,transform.BasisY,transform.BasisZ)
                try:
                    DB.ElementTransformUtils.MirrorElements(doc, List[DB.ElementId]([item.Id]), plane_XZ,False)
                    doc.Regenerate()
                    if item.Mirrored != False:
                        t.RollBack()
                    else:
                        t.Commit()
                        continue
                    t.Start()
                    DB.ElementTransformUtils.MirrorElements(doc, List[DB.ElementId]([item.Id]), plane_XY,False)
                    doc.Regenerate()
                    if item.Mirrored != False:
                        t.RollBack()
                    else:
                        t.Commit()
                        continue
                    t.Start()
                    DB.ElementTransformUtils.MirrorElements(doc, List[DB.ElementId]([item.Id]), plane_YZ,False)
                    doc.Regenerate()
                    if item.Mirrored != False:
                        t.RollBack()
                        print("Spiegeln für DistributionBox: {} gescheitert".fromat(item.Id.ToString()))     
                    else:
                        t.Commit()
                        continue
                except:
                    t.RollBack()
                    print("Spiegeln für DistributionBox: {} gescheitert".fromat(item.Id.ToString()))     
            
            self.class_GUI.austausch.Visibility = self.class_GUI.visible
            self.class_GUI.progress_aus.Visibility = self.class_GUI.hidden
            self.class_GUI.title_p_aus.Visibility = self.class_GUI.hidden

            TaskDialog.Show('Fertig','Erledigt!')

        except Exception as e:print(e)


    def GetName(self):
        return "Bauteil austauschen"
    
    def _update_pbar(self):
        self.class_GUI.progress_aus.Maximum = self.class_GUI.maxvalue
        self.class_GUI.progress_aus.Value = self.class_GUI.value
        self.class_GUI.title_p_aus.Text = self.class_GUI.PB_text

class NUMMERIEREN(IExternalEventHandler):
    def __init__(self):
        # self.Desktop = os.path.expanduser("~\Desktop")
        self.Desktop = 'C:\\temp'
        self.Familien_Name = 'DistributionBox_'
        self.class_GUI = None
        self.Path = None
        self.zeit = None

        
    def Execute(self,app):
        uidoc = app.ActiveUIDocument
        doc = uidoc.Document
        param_equality = DB.FilterStringContains()
        Fam_name_id = DB.ElementId(DB.BuiltInParameter.ELEM_FAMILY_PARAM)
        Fam_name_prov=DB.ParameterValueProvider(Fam_name_id)
        Fam_name_value_rule=DB.FilterStringRule(Fam_name_prov,param_equality,self.Familien_Name,True)
        Filter = DB.ElementParameterFilter(Fam_name_value_rule)
        elems = DB.FilteredElementCollector(doc).WherePasses(Filter).WhereElementIsNotElementType().ToElements()
        kopiert = []
        vorhanden = []
        for el in elems:
            famname = el.Symbol.FamilyName
            if not famname in vorhanden: 
                vorhanden.append(famname)
                continue
            else:kopiert.append(el)
        
        self.class_GUI.maxvalue = kopiert.Count
        self.class_GUI.minvalue = 0
        self.class_GUI.value = 1
        
        Familien_Liste = []
        for n,item in enumerate(kopiert):
            self.class_GUI.PB_text = str(n+1) + ' / '+ str(kopiert.Count) + ' DistributionBox laden'
            self.class_GUI.value = n+1
            self.class_GUI.progress_aus.Dispatcher.Invoke(System.Action(self._update_pbar),
                                System.Windows.Threading.DispatcherPriority.Background)

            self.class_GUI.startnum += 1
            self.Path = self.Desktop+'\\' +self.Familien_Name+str(self.class_GUI.startnum)+'.rfa'
            famdoc = doc.EditFamily(item.Symbol.Family)
            famdoc.SaveAs(self.Path)
            try:
                family = famdoc.LoadFamily(doc, FamilyLoadOptions())
                Familien_Liste.append(family)
            except:pass
            famdoc.Close(False)
            os.remove(self.Path)
        
        self.class_GUI.value = 1

        t = DB.Transaction(doc,'neu nummerieren')
        t.Start()
        for n,item in enumerate(kopiert):

            self.class_GUI.PB_text = str(n+1) + ' / '+ str(kopiert.Count) + ' DistributionBox wechseln'
            self.class_GUI.value = n+1
            self.class_GUI.progress_aus.Dispatcher.Invoke(System.Action(self._update_pbar),
                                System.Windows.Threading.DispatcherPriority.Background)


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
        self.neu_nummer.Visibility = self.visible
        self.progress_neu.Visibility = self.hidden
        self.title_p_neu.Visibility = self.hidden

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
    
    def _update_pbar(self):
        self.class_GUI.progress_aus.Maximum = self.class_GUI.maxvalue
        self.class_GUI.progress_aus.Value = self.class_GUI.value
        self.class_GUI.title_p_aus.Text = self.class_GUI.PB_text

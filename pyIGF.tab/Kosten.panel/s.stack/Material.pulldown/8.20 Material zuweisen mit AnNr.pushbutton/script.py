# coding: utf8
from System import Guid
from pyrevit import revit
from pyrevit import script, forms
from Autodesk.Revit.DB import *
from System.Collections.ObjectModel import ObservableCollection
from pyrevit.forms import WPFWindow
from System.Windows import FontWeights, FontStyles
from System.Windows.Media import Brushes, BrushConverter

__title__ = "8.20 Material zuweisen(mit Anlagennummer)"
__doc__ = """CAx Materialkz --->>> IGF_X_Material_Text
CAx Materialkz: 7e758303-a7ae-470f-ab85-738065c2824e
IGF_X_Material_Text: 1c657b13-cb97-47f6-aaf6-8fea74223c3c
 """
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()

uidoc = revit.uidoc
doc = revit.doc
active_view = uidoc.ActiveView

system_Luft = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctSystem).WhereElementIsNotElementType()
system_Luft_dict = {}
# Rohr System
rohrsys = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipingSystem).WhereElementIsNotElementType()
system_Rohr_dict = {}



def coll2dict(coll,dict):
    for el in coll:
      #  name = el.get_Parameter(BuiltInParameter.RBS_SYSTEM_NAME_PARAM).AsString()
        type = el.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString()
        if type in dict.Keys:
            dict[type].append(el.Id)
        else:
            dict[type] = [el.Id]

coll2dict(system_Luft,system_Luft_dict)
coll2dict(rohrsys,system_Rohr_dict)
system_Luft.Dispose()
rohrsys.Dispose()


class System(object):
    def __init__(self):
        self.checked = False
        self.SystemName = ''
        self.TypName = ''

    @property
    def TypName(self):
        return self._TypName
    @TypName.setter
    def TypName(self, value):
        self._TypName = value
    @property
    def checked(self):
        return self._checked
    @checked.setter
    def checked(self, value):
        self._checked = value
    @property
    def ElementId(self):
        return self._ElementId
    @ElementId.setter
    def ElementId(self, value):
        self._ElementId = value

Liste_Luft = ObservableCollection[System]()
Liste_Rohr = ObservableCollection[System]()
Liste_Alle = ObservableCollection[System]()

for key in system_Luft_dict.Keys:
    temp_system = System()
    temp_system.TypName = key
    temp_system.ElementId = system_Luft_dict[key]
    Liste_Luft.Add(temp_system)
    Liste_Alle.Add(temp_system)

for key in system_Rohr_dict.Keys:
    temp_system = System()
    temp_system.TypName = key
    temp_system.ElementId = system_Rohr_dict[key]
    Liste_Rohr.Add(temp_system)
    Liste_Alle.Add(temp_system)

# GUI Systemauswahl
class Systemauswahl(WPFWindow):
    def __init__(self, xaml_file_name,liste_Rohr,liste_Luft,liste_Alle):
        self.liste_Rohr = liste_Rohr
        self.liste_Luft = liste_Luft
        self.liste_Alle = liste_Alle
        WPFWindow.__init__(self, xaml_file_name)
        self.tempcoll = ObservableCollection[System]()
        self.altdatagrid = None

        try:
            self.dataGrid.ItemsSource = self.liste_Alle
            self.altdatagrid = self.liste_Alle
            self.backAll()
            self.click(self.alle)
        except Exception as e:
            logger.error(e)

        self.suche.TextChanged += self.search_txt_changed

    def click(self,button):
        button.Background = BrushConverter().ConvertFromString("#FF707070")
        button.FontWeight = FontWeights.Bold
        button.FontStyle = FontStyles.Italic
    def back(self,button):
        button.Background  = Brushes.White
        button.FontWeight = FontWeights.Normal
        button.FontStyle = FontStyles.Normal
    def backAll(self):
        self.back(self.luft)
        self.back(self.rohr)
        self.back(self.alle)
    
    def rohre(self,sender,args):
        self.backAll()
        self.click(self.rohr)
        self.dataGrid.ItemsSource = self.liste_Rohr
        self.altdatagrid = self.liste_Rohr
        self.dataGrid.Items.Refresh()
    def luftu(self,sender,args):
        self.backAll()
        self.click(self.luft)
        self.dataGrid.ItemsSource = self.liste_Luft
        self.altdatagrid = self.liste_Luft
        self.dataGrid.Items.Refresh()
    def allen(self,sender,args):
        self.backAll()
        self.click(self.alle)
        self.dataGrid.ItemsSource = self.liste_Alle
        self.altdatagrid = self.liste_Alle
        self.dataGrid.Items.Refresh()
    
    def search_txt_changed(self, sender, args):
        """Handle text change in search box."""
        self.tempcoll.Clear()
        text_typ = self.suche.Text.upper()
        if text_typ in ['',None]:
            self.dataGrid.ItemsSource = self.altdatagrid

        else:
            if text_typ == None:
                text_typ = ''
            for item in self.altdatagrid:
                if item.TypName.upper().find(text_typ) != -1:
                    self.tempcoll.Add(item)
            self.dataGrid.ItemsSource = self.tempcoll
        self.dataGrid.Items.Refresh()

    def checkall(self,sender,args):
        for item in self.dataGrid.Items:
            item.checked = True
        self.dataGrid.Items.Refresh()

    def uncheckall(self,sender,args):
        for item in self.dataGrid.Items:
            item.checked = False
        self.dataGrid.Items.Refresh()

    def toggleall(self,sender,args):
        for item in self.dataGrid.Items:
            value = item.checked
            item.checked = not value
        self.dataGrid.Items.Refresh()

    def auswahl(self,sender,args):
        self.Close()
    
Systemwindows = Systemauswahl("System.xaml",Liste_Rohr,Liste_Luft,Liste_Alle)
Systemwindows.ShowDialog()
Cate = forms.SelectFromList.show(['Luftkanäle','Luftkanalzubehör','Luftkanalformteile','Rohre','Rohrzubehör','Rohrformteile'], multiselect=True,button_name='Select Category')
SystemListe_luft = {}
SystemListe_rohr = {}
def colltodict(coll,dict):
    for el in coll:
        if el.checked == True:
            for it in el.ElementId:
                elem = doc.GetElement(it)
                systype = elem.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString()
                if not systype in dict.Keys:
                    dict[systype] = [elem]
                else:
                    dict[systype].append(elem)

colltodict(Liste_Luft,SystemListe_luft)
colltodict(Liste_Rohr,SystemListe_rohr)

def Datenschreiben(SystemListe):
    for systyp in SystemListe.Keys:
        sysliste = SystemListe[systyp]
        
        for sys_ele in sysliste:
            elements = None
            try:
                elements = sys_ele.DuctNetwork
            except:
                elements = sys_ele.PipingNetwork
            name = sys_ele.Name
            Leiung = []
            Formteile = []
            Zubehör = []
            
            try:
                anlagenr = sys_ele.LookupParameter('TGA_RLT_AnlagenNr').AsValueString()
            except:
                anlagenr = ''
            for elem in elements:
                if elem.Category.Id.ToString() in ['-2008000','-2008044',]:
                    Leiung.append(elem)
                elif elem.Category.Id.ToString()  in ['-2008010','-2008049']:
                    Formteile.append(elem)
                elif elem.Category.Id.ToString()  in ['-2008016','-2008055']:
                    Zubehör.append(elem)
            if 'Luftkanäle' in Cate or 'Rohre' in Cate:
                title = '{value}/{max_value} Leitungen in System ' + systyp+': '+name
                with forms.ProgressBar(title=title,cancellable=True, step=10) as pb1:
                    for n_1,elem in enumerate(Leiung):
                        if pb1.cancelled:
                            t.RollBack()
                            script.exit()
                        pb1.update_progress(n_1+1, len(Leiung))
                        try:
                            typeid = elem.GetTypeId()
                            alt = doc.GetElement(typeid).get_Parameter(Guid('7e758303-a7ae-470f-ab85-738065c2824e')).AsString()
                            if anlagenr:
                                wert = anlagenr+'-'+alt
                            else:
                                wert = alt
                            elem.get_Parameter(Guid('1c657b13-cb97-47f6-aaf6-8fea74223c3c')).Set(str(wert))
                        except Exception as e:
                            print(30*'-')
                            logger.error(doc.GetElement(elem.GetTypeId()).FamilyName)
                            logger.error(e)
            if 'Luftkanalformteile' in Cate or 'Rohrformteile' in Cate:
                title = '{value}/{max_value} Formteile in System ' + systyp+': '+name
                with forms.ProgressBar(title=title,cancellable=True, step=10) as pb2:
                    for n_1,elem in enumerate(Formteile):
                        if pb2.cancelled:
                            t.RollBack()
                            script.exit()
                        pb2.update_progress(n_1+1, len(Formteile))
                        try:
                            typeid = elem.GetTypeId()
                            alt = doc.GetElement(typeid).get_Parameter(Guid('7e758303-a7ae-470f-ab85-738065c2824e')).AsString()
                            if anlagenr:
                                wert = anlagenr+'-'+alt
                            else:
                                wert = alt
                            elem.get_Parameter(Guid('1c657b13-cb97-47f6-aaf6-8fea74223c3c')).Set(str(wert))
                        except Exception as e:
                            print(30*'-')
                            logger.error(doc.GetElement(elem.GetTypeId()).FamilyName)
                            logger.error(e)
            if 'Rohrzubehör' in Cate or 'Luftkanalzubehör' in Cate:
                title = '{value}/{max_value} Zubehör in System ' + systyp+': '+name
                with forms.ProgressBar(title=title,cancellable=True, step=10) as pb3:
                    for n_1,elem in enumerate(Zubehör):
                        if pb3.cancelled:
                            t.RollBack()
                            script.exit()
                        pb3.update_progress(n_1+1, len(Zubehör))
                        wert = None
                        conns = elem.MEPModel.ConnectorManager.Connectors
                        for conn in conns:
                            refs = conn.AllRefs
                            for ref in refs:
                                try:
                                    if ref.Owner.Category.Id.ToString() in ['-2008000','-2008044']:
                                        typeid = ref.Owner.GetTypeId()
                                        alt = doc.GetElement(typeid).get_Parameter(Guid('7e758303-a7ae-470f-ab85-738065c2824e')).AsString()
                                        if anlagenr:
                                            wert = anlagenr+'-'+alt
                                        else:
                                            wert = alt
                                        elem.get_Parameter(Guid('1c657b13-cb97-47f6-aaf6-8fea74223c3c')).Set(str(wert))
                                        break
                                
                                except Exception as e:
                                    pass
                            if wert:
                                break


if any(SystemListe_luft.keys()) or any(SystemListe_rohr.keys()):
    if forms.alert("Daten übernehmen?", ok=False, yes=True, no=True):
        t = Transaction(doc,'Material zuweisen')
        t.Start()
        Datenschreiben(SystemListe_luft)
        Datenschreiben(SystemListe_rohr)
        t.Commit()

# coding: utf8
from IGF_log import getlog
from rpw import revit, DB, UI
from pyrevit import script, forms
from IGF_lib import get_value
from System.Collections.ObjectModel import ObservableCollection
from System.Windows import FontWeights, FontStyles
from System.Windows.Media import Brushes, BrushConverter

__title__ = "2.02 Daten von ULK in MEP schreiben"
__doc__ = """Temperatur und Leistung von Umluftkühler in MEP Raum schreiben
--------------------------
ULK Category: HLS-Bauteile
--------------------------
Parameter in MEP-Räume: 
--------------------------
IGF_K_ULK-VL_Temp: Vorlauftemperatur, 
ermittelt von System. 
Typparameter in System: 'Temperatur von Medium',  'Abkürzung': ('K VL')

IGF_K_ULK-RL_Temp: Rücklauftemperatur, 
ermittelt von System. 
Typparameter in System: 'Temperatur von Medium',  'Abkürzung': ('K RL')

IGF_K_ULK_Leistung: ULK-Kühlleistung, 
ermittelt von ULK-Bauteile

IGF_K_ULK_Wassermenge: [L/s] 
berechnungsformel: IGF_K_ULK_Leistung/4200/(IGF_K_ULK-RL_Temp-IGF_K_ULK-VL_Temp)
--------------------------

[Version: 1.1]
[2022.10.31]
"""

__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()

uidoc = revit.uidoc
doc = revit.doc
name = doc.ProjectInformation.Name
number = doc.ProjectInformation.Number
config = script.get_config(name+number+'ULK-Familie')

try:getlog(__title__)
except:pass

bauteile = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MechanicalEquipment).WhereElementIsNotElementType().ToElementIds()

HLS_Dict = {}
HLS_Para_Dict = {}
for elid in bauteile:
    elem = doc.GetElement(elid)
    Family = elem.Symbol.FamilyName
    if not Family in HLS_Dict.keys():
        HLS_Dict[Family] = [elem]
        Paraliste = []
        for para in elem.Parameters:
            if not para.Definition.Name in Paraliste:
                try:Paraliste.append(para.Definition.Name)
                except:pass
        HLS_Para_Dict[Family]=Paraliste
    else:
        HLS_Dict[Family].append(elem)

class MechanicalEquipment(object):
    def __init__(self,name,elems):
        self.Name = name
        self.Selectedindex = -1
        self.Paras = []
        self.checked = False
        self.elems = elems

Liste_ULK = ObservableCollection[MechanicalEquipment]()

for name in sorted(HLS_Dict.keys()):
    temp = MechanicalEquipment(name,HLS_Dict[name])
    params = HLS_Para_Dict[name][:]
    temp.Paras = sorted(params)
    Liste_ULK.Add(temp)

class ULK_Familie(forms.WPFWindow):
    def __init__(self):
        self.Liste_ULK = Liste_ULK
        forms.WPFWindow.__init__(self, "window.xaml")
        self.tempcoll = ObservableCollection[MechanicalEquipment]()
        self.altdatagrid = Liste_ULK
        self.dataGrid.ItemsSource = Liste_ULK
        self.read_config()
        self.start = False
    
    def click(self,button):
        button.Background = BrushConverter().ConvertFromString("#FF707070")
        button.FontWeight = FontWeights.Bold
        button.FontStyle = FontStyles.Italic

    def back(self,button):
        button.Background  = Brushes.White
        button.FontWeight = FontWeights.Normal
        button.FontStyle = FontStyles.Normal

    
    def read_config(self):
        try:
            ULK_dict = config.ULK_dict
            for item in self.Liste_ULK:
                if item.Name in ULK_dict.keys():
                    try:item.checked = True
                    except:pass
                    try:item.Selectedindex = item.Paras.index(ULK_dict[item.Name])
                    except:item.Selectedindex = -1
            self.dataGrid.Items.Refresh()
         
        except:
            pass        
        
    def write_config(self):
        dict_ = {}
        for item in self.Liste_ULK:
            if item.checked:
                try:
                    if item.Selectedindex != -1:
                        dict_[item.Name] = item.Paras[item.Selectedindex]
                except:
                    pass

        config.ULK_dict = dict_
        script.save_config()

    def ok(self,sender,args):
        self.write_config()
        for el in self.Liste_ULK:
            if el.checked and el.Selectedindex == -1:
                UI.TaskDialog.Show('Fehler','Bitte Kühlleitung-Parameter auswählen!')
                return
        self.start = True
        self.Close()

    def hide(self,sender,args):
        self.tempcoll.Clear()
        for item in self.dataGrid.Items:
            if item.checked:
                self.tempcoll.Add(item)
        self.dataGrid.ItemsSource = self.tempcoll
        self.click(self.aus)
        self.back(self.ein)

    def show(self,sender,args):
        self.dataGrid.ItemsSource = self.altdatagrid
        self.click(self.ein)
        self.back(self.aus)

    def close(self,sender,args):
        self.Close()
        script.exit()
    
ULK_Familie_Auswahl = ULK_Familie()
try:
    ULK_Familie_Auswahl.ShowDialog()
except Exception as e:
    logger.error(e)
    ULK_Familie_Auswahl.Close()
    script.exit()

if ULK_Familie_Auswahl.start == False:
    script.exit()

Familie = []
for el in Liste_ULK:
    if el.checked:
        Familie.append(el)

if len(Familie) == 0:
    UI.TaskDialog.Show('Fehler','Keine ULK-Familie ausgewählt!')
    script.exit()

class Umluftkuehler:
    def __init__(self, elem, param):
        self.elem = elem
        self.param = param
        if self.Muster_Pruefen():
            return
        try:
            self.RaumNr = self.elem.Space[doc.GetElement(self.elem.CreatedPhaseId)].Number
            self.Raum = self.elem.Space[doc.GetElement(self.elem.CreatedPhaseId)].Id.ToString()
        except:
            self.RaumNr = ''
            self.Raum = ''
            logger.error('kein MEP-Raum für Element {} gefunden'.format(self.elem.Id.ToString()))
        try:
            self.leistung = get_value(self.elem.LookupParameter(self.param))
        except:
            self.leistung = 0
            logger.error("Parameter {} konnte nicht gefunden werden".format(self.param))
        self.VL_Temp = ''
        self.RL_Temp = ''
        self.Temperaturermittelung()
    
    def Muster_Pruefen(self):
        '''prüft, ob der Bauteil sich in Musterbereich befindet.'''
        try:
            bb = self.elem.LookupParameter('Bearbeitungsbereich').AsValueString()
            if bb == 'KG4xx_Musterbereich':return True
            else:return False
        except:return False
    
    def Temperaturermittelung(self):
        conns = self.elem.MEPModel.ConnectorManager.Connectors
        for conn in conns:
            if conn.IsConnected:
                system  = conn.MEPSystem
                if system:
                    typ = doc.GetElement(system.get_Parameter(DB.BuiltInParameter.ELEM_TYPE_PARAM).AsElementId())
                    try:
                        abk = typ.LookupParameter('Abkürzung').AsString()
                        temp = typ.LookupParameter('Temperatur von Medium').AsDouble() - 273.15
                        if abk == 'K VL':
                            self.VL_Temp = temp
                        elif abk == 'K RL':
                            self.RL_Temp = temp
                    except:
                        pass


MEP_Dict = {}

with forms.ProgressBar(title='{value}/{max_value} Umluftkühler', cancellable=True, step=10) as pb:
    for n, ulk in enumerate(Familie):
        pb.title='{value}/{max_value} Exemplare von ' + ulk.Name + ' --- ' + str(n+1) + ' / '+ str(len(Familie)) + 'Familien'
        for n1, elem in enumerate(ulk.elems):
            if pb.cancelled:
                script.exit()
            pb.update_progress(n1 + 1, len(ulk.elems))
            umluft = Umluftkuehler(elem,ulk.Paras[ulk.Selectedindex])
            if not umluft.Muster_Pruefen():
                if umluft.Raum:
                    if umluft.Raum not in MEP_Dict.keys():
                        MEP_Dict[umluft.Raum] = [umluft]
                    else:
                        MEP_Dict[umluft.Raum].append(umluft)
                

class MEPRaum:
    def __init__(self, elemid, liste_ulk):
        self.elem = doc.GetElement(DB.ElementId(int(elemid)))
        self.ulk_liste = liste_ulk
        self.leistung = 0
        self.VL_Temp = 0
        self.RL_Temp = 0
        self.Wasser = 0
        if len(self.ulk_liste) > 0:
            self.leistungberechnen()
            self.Tempberechnen()
            self.Wassermengenberechnen()

    def leistungberechnen(self):
        for item in self.ulk_liste:
            self.leistung += item.leistung

    def Tempberechnen(self):
        liste_VL = []
        liste_RL = []
        for item in self.ulk_liste:
            if item.VL_Temp not in liste_VL:
                liste_VL.append(item.VL_Temp)
            if item.RL_Temp not in liste_RL:
                liste_RL.append(item.RL_Temp)
        if len(liste_VL) == 1:
            self.VL_Temp = liste_VL[0]
        else:
            logger.error(50*'-')
            logger.error('Vorlauftemperatur von ULK in MEPRaum {}, {} nicht gleich'.format(self.elem.Number,self.elem.Id.ToString()))
            for el in self.ulk_liste:
                logger.error('Vorlauftemperatur: {} °C, ULK-Id: {}'.format(el.VL_Temp,el.element.Id.ToString()))
            liste_VL_neu = [float(a) for a in liste_VL]
            self.VL_Temp = min(liste_VL_neu)
            logger.error('!!! Achtung: ULK-Vorlauftemperatur in MEPRaum {}, {} wird auf {} °C eingesetzt'.format(self.elem.Number,self.elem.Id.ToString(),self.VL_Temp))
            logger.error(50*'-')

        if len(liste_RL) == 1:
            self.RL_Temp = liste_RL[0]
        else:
            logger.error(50*'-')
            logger.error('Rücklauftemperatur von ULK in MEPRaum {}, {} nicht gleich'.format(self.elem.Number,self.elem.Id.ToString()))
            for el in self.ulk_liste:
                logger.error('Rücklauftemperatur: {} °C, ULK-Id: {}'.format(el.RL_Temp,el.element.Id.ToString()))
            liste_RL_neu = [float(a) for a in liste_RL]
            self.RL_Temp = max(liste_RL_neu)
            logger.error('!!! Achtung: ULK-Rücklauftemperatur in MEPRaum {}, {} wird auf {} °C eingesetzt'.format(self.elem.Number,self.elem.Id.ToString(),self.RL_Temp))
            logger.error(50*'-')

    def Wassermengenberechnen(self):
        if self.VL_Temp == 0 or self.RL_Temp == 0:
            return
        else:
            try:self.Wasser = self.leistung / 4200.0 / abs(self.RL_Temp - self.VL_Temp)
            except Exception as e:
                print(self.RL_Temp,self.VL_Temp)
                self.Wasser = 0
    
    def werte_schreiben(self):
        """Schreibt die berechneten Werte zurück in das Modell."""
        def wert_schreiben(param_name, wert):
            if not wert is None:
                param = self.elem.LookupParameter(param_name)
                if param:
                    if param.StorageType.ToString() == 'Double':
                        param.SetValueString(str(wert))
                    else:
                        param.Set(wert)

        
        wert_schreiben("IGF_K_ULK-VL_Temp", self.VL_Temp)
        wert_schreiben("IGF_K_ULK-RL_Temp", self.RL_Temp)
        wert_schreiben("IGF_K_ULK_Leistung", self.leistung)
        wert_schreiben("IGF_K_ULK_Wassermenge", self.Wasser)
    

mep_liste = []

MEP_Raum_Liste = [e.ToString() for e in DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MEPSpaces).ToElementIds()]

with forms.ProgressBar(title='{value}/{max_value} MEP-Räume',cancellable=True, step=10) as pb1:
    for n, mep in enumerate(MEP_Raum_Liste):
        if pb1.cancelled:
            script.exit()
        pb1.update_progress(n + 1, len(MEP_Raum_Liste))
        if mep in MEP_Dict.keys():
            mepraum = MEPRaum(mep,MEP_Dict[mep])
        else:
            mepraum = MEPRaum(mep,[])
        mep_liste.append(mepraum)


# Werte zuückschreiben + Abfrage
if forms.alert("Berechnete Werte in MEP-Räume schreiben?", ok=False, yes=True, no=True):
    with forms.ProgressBar(title="{value}/{max_value} MEP-Räume",cancellable=True, step=10) as pb2:
        t = DB.Transaction(doc)
        t.Start('Daten von ULK in MEP-Räume schreiben')
        for n,mepraum in enumerate(mep_liste):
            if pb2.cancelled:
                t.RollBack()
                script.exit()
            pb2.update_progress(n+1, len(mep_liste))
            mepraum.werte_schreiben()
        t.Commit()
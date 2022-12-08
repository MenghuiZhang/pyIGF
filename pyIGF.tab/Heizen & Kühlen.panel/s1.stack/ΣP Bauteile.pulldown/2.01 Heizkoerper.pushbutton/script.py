# coding: utf8
from IGF_log import getlog
from rpw import revit, DB, UI
from pyrevit import script, forms
from IGF_lib import get_value
from System.Collections.ObjectModel import ObservableCollection
from System.Windows import FontWeights, FontStyles
from System.Windows.Media import Brushes, BrushConverter

__title__ = "2.01 Daten von HK in MEP schreiben"
__doc__ = """Temperatur und Volumenstrom bzw. Heizleistung von Heizkörper in MEP Raum schreiben

--------------------------
Heizkörper Category: HLS-Bauteile
--------------------------
Parameter in MEP-Räume: 
--------------------------
IGF_H_HK-VL_Temp: Vorlauftemperatur, 
ermittelt von System. 
Typparameter in System: 
'Temperatur von Medium',  
'Abkürzung': 'H VL'

IGF_H_HK-RL_Temp: Rücklauftemperatur, 
ermittelt von System. 
Typparameter in System: 
'Temperatur von Medium',  
'Abkürzung': 'H RL'

IGF_H_HK_Leistung: HK-Heizleistung, 
ermittelt von HK-Bauteile

IGF_H_HK_Wassermenge: [L/s] 
ermittelt von HK-Bauteile
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
config = script.get_config(name+number+'HK-Familie')

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
        self.Selectedindex1 = -1
        self.Paras = []
        self.checked = False
        self.elems = elems

Liste_HK = ObservableCollection[MechanicalEquipment]()

for name in sorted(HLS_Dict.keys()):
    temp = MechanicalEquipment(name,HLS_Dict[name])
    params = HLS_Para_Dict[name][:]
    temp.Paras = sorted(params)
    Liste_HK.Add(temp)

class HK_Familie(forms.WPFWindow):
    def __init__(self):
        self.Liste_HK = Liste_HK
        forms.WPFWindow.__init__(self, "window.xaml")
        self.tempcoll = ObservableCollection[MechanicalEquipment]()
        self.altdatagrid = Liste_HK
        self.dataGrid.ItemsSource = Liste_HK
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
            HK_dict = config.HK_dict
            for item in self.Liste_HK:
                if item.Name in HK_dict.keys():
                    try:item.checked = True
                    except:pass
                    try:item.Selectedindex = item.Paras.index(HK_dict[item.Name][0])
                    except:item.Selectedindex = -1
                    try:item.Selectedindex1 = item.Paras.index(HK_dict[item.Name][1])
                    except:item.Selectedindex1 = -1
            self.dataGrid.Items.Refresh()
         
        except:
            pass        
        
    def write_config(self):
        dict_ = {}
        for item in self.Liste_HK:
            if item.checked:
                try:
                    if item.Selectedindex != -1 and item.Selectedindex1 != -1:
                        dict_[item.Name] = [item.Paras[item.Selectedindex],item.Paras[item.Selectedindex1]]
                except:
                    pass

        config.HK_dict = dict_
        script.save_config()

    def ok(self,sender,args):
        self.write_config()
        for el in self.Liste_HK:
            if el.checked and (el.Selectedindex == -1 or el.Selectedindex1 == -1):
                UI.TaskDialog.Show('Fehler','Bitte Parameter auswählen!')
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
    
HK_Familie_Auswahl = HK_Familie()
try:
    HK_Familie_Auswahl.ShowDialog()
except Exception as e:
    logger.error(e)
    HK_Familie_Auswahl.Close()
    script.exit()

if HK_Familie_Auswahl.start == False:
    script.exit()

    
Familie = []
for el in Liste_HK:
    if el.checked:
        Familie.append(el)

if len(Familie) == 0:
    UI.TaskDialog.Show('Fehler','Keine HK-Familie ausgewählt!')
    script.exit()


class Heizkoerper:
    def __init__(self, elem, Liste):
        self.elem = elem
        self.vol = Liste[0]
        self.HL = Liste[1]
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
            self.leistung = get_value(self.elem.LookupParameter(self.HL))
        except:
            self.leistung = 0
            logger.error("Parameter {} konnte nicht gefunden werden".format(self.HL))
        try:
            self.VOL = round(get_value(self.elem.LookupParameter(self.vol)),2)
        except:
            self.VOL = 0
            logger.error("Parameter {} konnte nicht gefunden werden".format(self.vol))

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
                        temp = typ.LookupParameter('Temperatur von Medium').AsDouble() -273.15
                        if abk == 'H VL':
                            self.VL_Temp = temp
                        elif abk == 'H RL':
                            self.RL_Temp = temp
                    except:
                        pass

MEP_Dict = {}

with forms.ProgressBar(title='{value}/{max_value} Heizkörper', cancellable=True, step=10) as pb:
    for n, hk in enumerate(Familie):
        pb.title='{value}/{max_value} Exemplare von ' + hk.Name + ' --- ' + str(n+1) + ' / '+ str(len(Familie)) + 'Familien'
        for n1, elem in enumerate(hk.elems):
            if pb.cancelled:
                script.exit()
            pb.update_progress(n1 + 1, len(hk.elems))
            heiz = Heizkoerper(elem,[hk.Paras[hk.Selectedindex],hk.Paras[hk.Selectedindex1]])
            if not heiz.Muster_Pruefen():
                if heiz.Raum:
                    if heiz.Raum not in MEP_Dict.keys():
                        MEP_Dict[heiz.Raum] = [heiz]
                    else:
                        MEP_Dict[heiz.Raum].append(heiz)


class MEPRaum:
    def __init__(self, elemid, liste_hk):
        self.elem = doc.GetElement(DB.ElementId(int(elemid)))
        self.hk_liste = liste_hk
        self.leistung = 0
        self.VL_Temp = 0
        self.RL_Temp = 0
        self.Wasser = 0
        if len(self.hk_liste) > 0:
            self.berechnen()


    def berechnen(self):
        liste_VL = []
        liste_RL = []
        for item in self.hk_liste:
            self.leistung += item.leistung
            self.Wasser += item.VOL
            if item.VL_Temp not in liste_VL:
                liste_VL.append(item.VL_Temp)
            if item.RL_Temp not in liste_RL:
                liste_RL.append(item.RL_Temp)
        if len(liste_VL) == 1:
            self.VL_Temp = liste_VL[0]
        else:
            logger.error(50*'-')
            logger.error('Vorlauftemperatur von Heizkörper in MEPRaum {}, {} nicht gleich'.format(self.elem.Number,self.elem.Id.ToString()))
            for el in self.hk_liste:
                logger.error('Vorlauftemperatur: {} °C, Heizkörper-Id: {}'.format(el.VL_Temp,el.elem.Id.ToString()))
            liste_VL_neu = [float(a) for a in liste_VL]
            self.VL_Temp = max(liste_VL_neu)
            logger.error('!!! Achtung: Heizkörper-Vorlauftemperatur in MEPRaum {}, {} wird auf {} °C eingesetzt'.format(self.elem.Number,self.elem.Id.ToString(),self.VL_Temp))
            logger.error(50*'-')

        if len(liste_RL) == 1:
            self.RL_Temp = liste_RL[0]
        else:
            logger.error(50*'-')
            logger.error('Rücklauftemperatur von hk in MEPRaum {}, {} nicht gleich'.format(self.elem.Number,self.elem.Id.ToString()))
            for el in self.hk_liste:
                logger.error('Rücklauftemperatur: {} °C, Heizkörper-Id: {}'.format(el.RL_Temp,el.elem.Id.ToString()))
            liste_RL_neu = [float(a) for a in liste_RL]
            self.RL_Temp = min(liste_RL_neu)
            logger.error('!!! Achtung: Heizkörper-Rücklauftemperatur in MEPRaum {}, {} wird auf {} °C eingesetzt'.format(self.elem.Number,self.elem.Id.ToString(),self.RL_Temp))
            logger.error(50*'-')   
    
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

        
        wert_schreiben("IGF_H_HK-VL_Temp", self.VL_Temp)
        wert_schreiben("IGF_H_HK-RL_Temp", self.RL_Temp)
        wert_schreiben("IGF_H_HK_Leistung", self.leistung)
        wert_schreiben("IGF_H_HK_Wassermenge", self.Wasser)
    

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

# Werte zurückschreiben + Abfrage
if forms.alert("Berechnete Werte in MEP-Räume schreiben?", ok=False, yes=True, no=True):
    with forms.ProgressBar(title="{value}/{max_value} MEP-Räume",cancellable=True, step=10) as pb2:
        t = DB.Transaction(doc)
        t.Start('Daten von Heizkörper in MEP-Räume schreiben')
        for n,mepraum in enumerate(mep_liste):
            if pb2.cancelled:
                t.RollBack()
                script.exit()
            pb2.update_progress(n+1, len(mep_liste))
            mepraum.werte_schreiben()
        t.Commit()
# coding: utf8
from IGF_log import getlog
from rpw import revit, DB, UI
from pyrevit import script, forms
from IGF_lib import get_value
from System.Collections.ObjectModel import ObservableCollection
from System.Windows import FontWeights, FontStyles
from System.Windows.Media import Brushes, BrushConverter



__title__ = "2.00  Daten von DeS in MEP schreiben"
__doc__ = """Temperatur und Volumenstrom bzw. Heiz- und Kühlleistung der Deckensegel in MEP Raum schreiben

-------------------------
DeS Category: HLS-Bauteile
-------------------------
Parameter in DeS: manuell definiert
-------------------------
Parameter in MEP-Räume:
-------------------------
IGF_H_DeS-VL_Win_Temp: VL_Temp Heizung

IGF_H_DeS-RL_Win_Temp: RL_Temp Heizung

IGF_K_DeS-VL_Som_Temp: VL_Temp Kühlung

IGF_K_DeS-RL_Som_Temp: RL_Temp Kühlung

IGF_H_DeS_Winter: Volumenstrom Heizung

IGF_K_DeS_Sommer: Volumenstrom Kühlung

IGF_H_DeS_Leistung: Leistung Heizung

IGF_K_DeS_Leistung: Leistung Kühlung
-------------------------

[Version: 1.1]
[2022.10.32]


"""
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()

uidoc = revit.uidoc
doc = revit.doc
name = doc.ProjectInformation.Name
number = doc.ProjectInformation.Number
config = script.get_config(name+number+'DeS-Familie')

try:
    getlog(__title__)
except:
    pass


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
        self.Selectedindex2 = -1
        self.Selectedindex3 = -1
        self.Selectedindex4 = -1
        self.Selectedindex5 = -1
        self.Selectedindex6 = -1
        self.Paras = []
        self.checked = False
        self.elems = elems

Liste_DeS = ObservableCollection[MechanicalEquipment]()

for name in sorted(HLS_Dict.keys()):
    temp = MechanicalEquipment(name,HLS_Dict[name])
    params = HLS_Para_Dict[name][:]
    temp.Paras = sorted(params)
    Liste_DeS.Add(temp)

class DeS_Familie(forms.WPFWindow):
    def __init__(self):
        self.Liste_DeS = Liste_DeS
        forms.WPFWindow.__init__(self, "window.xaml")
        self.tempcoll = ObservableCollection[MechanicalEquipment]()
        self.altdatagrid = Liste_DeS
        self.dataGrid.ItemsSource = Liste_DeS
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
            DeS_dict = config.DeS_dict
            for item in self.Liste_DeS:
                if item.Name in DeS_dict.keys():
                    try:item.checked = True
                    except:pass
                    try:item.Selectedindex = item.Paras.index(DeS_dict[item.Name][0])
                    except:item.Selectedindex = -1
                    try:item.Selectedindex = item.Paras.index(DeS_dict[item.Name][1])
                    except:item.Selectedindex1 = -1
                    try:item.Selectedindex = item.Paras.index(DeS_dict[item.Name][2])
                    except:item.Selectedindex2 = -1
                    try:item.Selectedindex = item.Paras.index(DeS_dict[item.Name][3])
                    except:item.Selectedindex3 = -1
                    try:item.Selectedindex = item.Paras.index(DeS_dict[item.Name][4])
                    except:item.Selectedindex4 = -1
                    try:item.Selectedindex = item.Paras.index(DeS_dict[item.Name][5])
                    except:item.Selectedindex5 = -1
                    try:item.Selectedindex = item.Paras.index(DeS_dict[item.Name][6])
                    except:item.Selectedindex6 = -1

            self.dataGrid.Items.Refresh()
         
        except:
            pass
            
    def write_config(self):
        dict_ = {}
        for item in self.Liste_DeS:
            if item.checked:
                try:
                    if item.Selectedindex != -1 and item.Selectedindex1 != -1 and\
                    item.Selectedindex2 != -1 and item.Selectedindex3 != -1 and\
                        item.Selectedindex4 != -1 and item.Selectedindex5 != -1 and item.Selectedindex6 != -1:
                        dict_[item.Name] = [item.Paras[item.Selectedindex],item.Paras[item.Selectedindex1],
                        item.Paras[item.Selectedindex2],item.Paras[item.Selectedindex3],item.Paras[item.Selectedindex4],
                        item.Paras[item.Selectedindex5],item.Paras[item.Selectedindex6]]
                except:
                    pass

        config.DeS_dict = dict_
        script.save_config()

    def ok(self,sender,args):
        self.write_config()
        for el in self.Liste_DeS:
            if el.checked and (el.Selectedindex == -1 or el.Selectedindex1 == -1 or\
            	el.Selectedindex2 == -1 or el.Selectedindex3 == -1 or el.Selectedindex4 == -1\
                     or el.Selectedindex5 == -1 or el.Selectedindex6 == -1):
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
        config.Hide = True
        self.click(self.aus)
        self.back(self.ein)

    def show(self,sender,args):
        self.dataGrid.ItemsSource = self.altdatagrid
        config.Hide = False
        self.click(self.ein)
        self.back(self.aus)

    def close(self,sender,args):
        self.Close()
        script.exit()
    
DeS_Familie_Auswahl = DeS_Familie()
try:
    DeS_Familie_Auswahl.ShowDialog()
except Exception as e:
    DeS_Familie_Auswahl.Close()
    script.exit()

if DeS_Familie_Auswahl.start == False:
    script.exit()

Familie = []
for el in Liste_DeS:
    if el.checked:
        Familie.append(el)

if len(Familie) == 0:
    UI.TaskDialog.Show('Fehler','Keine DeS-Familie ausgewählt!')
    script.exit()

class Deckensegel:
    def __init__(self, elem, Liste):
        self.elem = elem
        self.vol = Liste[0]
        self.KL = Liste[1]
        self.VK = Liste[2]
        self.RK = Liste[3]
        self.HL = Liste[4]
        self.VH = Liste[5]
        self.RH = Liste[6]
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
            self.leistung_H = get_value(self.elem.LookupParameter(self.HL))
        except:
            self.leistung_H = 0
            logger.error("Parameter {} konnte nicht gefunden werden".format(self.HL))
        try:
            self.leistung_K = get_value(self.elem.LookupParameter(self.KL))
        except:
            self.leistung_K = 0
            logger.error("Parameter {} konnte nicht gefunden werden".format(self.KL))
        try:
            self.VLH = get_value(self.elem.LookupParameter(self.VH))
        except:
            self.VLH = 0
            logger.error("Parameter {} konnte nicht gefunden werden".format(self.VH))
        try:
            self.RLH = get_value(self.elem.LookupParameter(self.RH))
        except:
            self.RLH = 0
            logger.error("Parameter {} konnte nicht gefunden werden".format(self.RH))
        try:
            self.VLK = get_value(self.elem.LookupParameter(self.VK))
        except:
            self.VLK = 0
            logger.error("Parameter {} konnte nicht gefunden werden".format(self.VK))
        try:
            self.RLK = get_value(self.elem.LookupParameter(self.RK))
        except:
            self.RLK = 0
            logger.error("Parameter {} konnte nicht gefunden werden".format(self.RK))
        try:
            self.VOL = get_value(self.elem.LookupParameter(self.vol))
        except:
            self.VOL = 0
            logger.error("Parameter {} konnte nicht gefunden werden".format(self.vol))
        
    def Muster_Pruefen(self):
        '''prüft, ob der Bauteil sich in Musterbereich befindet.'''
        try:
            bb = self.elem.LookupParameter('Bearbeitungsbereich').AsValueString()
            if bb == 'KG4xx_Musterbereich':return True
            else:return False
        except:return False
    
MEP_Dict = {}

with forms.ProgressBar(title='{value}/{max_value} Deckensegel', cancellable=True, step=10) as pb:
    for n, des in enumerate(Familie):
        pb.title='{value}/{max_value} Exemplare von ' + des.Name + ' --- ' + str(n+1) + ' / '+ str(len(Familie)) + 'Familien'
        for n1, elem in enumerate(des.elems):
            if pb.cancelled:
                script.exit()
            pb.update_progress(n1 + 1, len(des.elems))
            decken = Deckensegel(elem, [des.Paras[des.Selectedindex],des.Paras[des.Selectedindex1],
                                        des.Paras[des.Selectedindex2],des.Paras[des.Selectedindex3],
                                        des.Paras[des.Selectedindex4],des.Paras[des.Selectedindex5],
                                        des.Paras[des.Selectedindex6]])
            if not decken.Muster_Pruefen():
                if decken.Raum:
                    if decken.Raum not in MEP_Dict.keys():
                        MEP_Dict[decken.Raum] = [decken]
                    else:
                        MEP_Dict[decken.Raum].append(decken)
                
class MEPRaum:
    def __init__(self, elemid, liste_des):
        self.elem = doc.GetElement(DB.ElementId(int(elemid)))
        self.des_liste = liste_des
        self.leistungH = 0
        self.leistungK = 0
        self.VLH_temp = 0
        self.RLH_temp = 0
        self.VLK_temp = 0
        self.RLK_temp = 0
        self.Vol = 0
        if len(self.des_liste) > 0:
            self.berechnen()


    def berechnen(self):
        liste_VLH = []
        liste_RLH = []
        liste_VLK = []
        liste_RLK = []
        for item in self.des_liste:
            self.leistungH += item.leistung_H
            self.leistungK += item.leistung_K
            self.Vol += item.VOL
            if item.VLH not in liste_VLH:
                liste_VLH.append(item.VLH)
            if item.RLH not in liste_RLH:
                liste_RLH.append(item.RLH)
            if item.VLK not in liste_VLK:
                liste_VLK.append(item.VLK)
            if item.RLK not in liste_RLK:
                liste_RLK.append(item.RLK)
        if len(liste_VLH) == 1:
            self.VLH_temp = liste_VLH[0]
        else:
            logger.error(50*'-')
            logger.error('Vorlauftemperatur Heizen von DeS in MEPRaum {}, {} nicht gleich'.format(self.elem.Number,self.elem.Id.ToString()))
            for el in self.des_liste:
                logger.error('Vorlauftemperatur Heizen: {} °C, DeS-Id: {}'.format(el.VLH,el.elem.Id.ToString()))
            liste_VLH_neu = [float(a) for a in liste_VLH]
            self.VLH_temp = max(liste_VLH_neu)
            logger.error('!!! Achtung: DeS-Vorlauftemperatur Heizen in MEPRaum {}, {} wird auf {} °C eingesetzt'.format(self.elem.Number,self.elem.Id.ToString(),self.VLH_temp))
            logger.error(50*'-')

        if len(liste_RLH) == 1:
            self.RLH_temp = liste_RLH[0]
        else:
            logger.error(50*'-')
            logger.error('Rücklauftemperatur Heizen von DeS in MEPRaum {}, {} nicht gleich'.format(self.elem.Number,self.elem.Id.ToString()))
            for el in self.des_liste:
                logger.error('Rücklauftemperatur Heizen: {} °C, DeS-Id: {}'.format(el.RLH,el.elem.Id.ToString()))
            liste_RLH_neu = [float(a) for a in liste_RLH]
            self.RLH_temp = min(liste_RLH_neu)
            logger.error('!!! Achtung: DeS-Vorlauftemperatur Heizen in MEPRaum {}, {} wird auf {} °C eingesetzt'.format(self.elem.Number,self.elem.Id.ToString(),self.RLH_temp))
            logger.error(50*'-')

        if len(liste_VLK) == 1:
            self.VLK_temp = liste_VLK[0]
        else:
            logger.error(50*'-')
            logger.error('Vorlauftemperatur Kälte von DeS in MEPRaum {}, {} nicht gleich'.format(self.elem.Number,self.elem.Id.ToString()))
            for el in self.des_liste:
                logger.error('Vorlauftemperatur Kälte: {} °C, DeS-Id: {}'.format(el.VLK,el.elem.Id.ToString()))
            liste_VLK_neu = [float(a) for a in liste_VLK]
            self.VLK_temp = min(liste_VLK_neu)
            logger.error('!!! Achtung: DeS-Vorlauftemperatur Kälte in MEPRaum {}, {} wird auf {} °C eingesetzt'.format(self.elem.Number,self.elem.Id.ToString(),self.VLK_temp))
            logger.error(50*'-')

        if len(liste_RLK) == 1:
            self.RLK_temp = liste_RLK[0]
        else:
            logger.error(50*'-')
            logger.error('Rücklauftemperatur Kälte von DeS in MEPRaum {}, {} nicht gleich'.format(self.elem.Number,self.element.Id.ToString()))
            for el in self.des_liste:
                logger.error('Rücklauftemperatur Kälte: {} °C, DeS-Id: {}'.format(el.RLK,el.elem.Id.ToString()))
            liste_RLK_neu = [float(a) for a in liste_RLK]
            self.RLK_temp = max(liste_RLK_neu)
            logger.error('!!! Achtung: DeS-Rücklauftemperatur Kälte in MEPRaum {}, {} wird auf {} °C eingesetzt'.format(self.element.Number,self.element.Id.ToString(),self.RLK_temp))
            logger.error(50*'-')

    
    def werte_schreiben(self):
        """Schreibt die berechneten Werte zurück in das Modell."""
        def wert_schreiben(param_name, wert):
            if not wert is None:
                param = self.element.LookupParameter(param_name)
                if param:
                    if param.StorageType.ToString() == 'Double':
                        param.SetValueString(str(wert))
                    else:
                        param.Set(wert)
        
        wert_schreiben("IGF_H_DeS-VL_Win_Temp", self.VLH_temp)
        wert_schreiben("IGF_H_DeS-RL_Win_Temp", self.RLH_temp)
        wert_schreiben("IGF_K_DeS-VL_Som_Temp", self.VLK_temp)
        wert_schreiben("IGF_K_DeS-RL_Som_Temp", self.RLK_temp)
        wert_schreiben("IGF_H_DeS_Winter", self.Vol)
        wert_schreiben("IGF_K_DeS_Sommer", self.Vol)
        wert_schreiben("IGF_H_DeS_Leistung", self.leistungH)
        wert_schreiben("IGF_K_DeS_Leistung", self.leistungK)

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
        t.Start('Daten von DeS in MEP-Räume schreiben')
        for n,mepraum in enumerate(mep_liste):
            if pb2.cancelled:
                t.RollBack()
                script.exit()
            pb2.update_progress(n+1, len(mep_liste))
            mepraum.werte_schreiben()
        t.Commit()
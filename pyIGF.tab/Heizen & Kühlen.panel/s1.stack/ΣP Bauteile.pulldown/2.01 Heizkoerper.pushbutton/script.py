# coding: utf8
import sys
sys.path.append(r'R:\pyRevit\xx_Skripte\libs\IGF_libs')
from IGF_log import getlog
from rpw import revit, DB
from pyrevit import script, forms
from IGF_libKopie import get_value
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
[2021.11.18]
"""
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()

uidoc = revit.uidoc
doc = revit.doc
name = doc.ProjectInformation.Name
number = doc.ProjectInformation.Number
config = script.get_config(name+number+'HK-Familie')

try:
    getlog(__title__)
except:
    pass

bauteile_collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MechanicalEquipment).WhereElementIsNotElementType()
bauteile = bauteile_collector.ToElementIds()
bauteile_collector.Dispose()

phase = doc.Phases[0]

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
                try:
                    Paraliste.append(para.Definition.Name)
                except:
                    pass
        HLS_Para_Dict[Family]=Paraliste
    else:
        HLS_Dict[Family].append(elem)

class MechanicalEquipment(object):
    def __init__(self,name):
        self.Name = name
        self.Vol_Id = -1
        self.HL_Id = -1
        self.Paras = []
        self.checked = False

    @property
    def Name(self):
        return self._Name
    @Name.setter
    def Name(self, value):
        self._Name = value
    @property
    def HL_Id(self):
        return self._HL_Id
    @HL_Id.setter
    def HL_Id(self, value):
        self._HL_Id = value
    @property
    def Vol_Id(self):
        return self._Vol_Id
    @Vol_Id.setter
    def Vol_Id(self, value):
        self._Vol_Id = value
    @property
    def Paras(self):
        return self._Paras
    @Paras.setter
    def Paras(self, value):
        self._Paras = value
    @property
    def Para_dict(self):
        return self._Para_dict
    @Para_dict.setter
    def Para_dict(self, value):
        self._Para_dict = value
    @property
    def checked(self):
        return self._checked
    @checked.setter
    def checked(self, value):
        self._checked = value
    @property
    def Para_id_dict(self):
        return self._Para_id_dict
    @Para_id_dict.setter
    def Para_id_dict(self, value):
        self._Para_id_dict = value

class select(object):
    def __init__(self,n,param):
        self.selectId = n
        self.ParaName = param

    @property
    def ParaName(self):
        return self._ParaName
    @ParaName.setter
    def ParaName(self, value):
        self._ParaName = value
    @property
    def selectId(self):
        return self._selectId
    @selectId.setter
    def selectId(self, value):
        self._selectId = value


Liste_HK = ObservableCollection[MechanicalEquipment]()

HKName = HLS_Para_Dict.keys()[:]
HKName.sort()

for elem in HKName:
    temp = MechanicalEquipment(elem)
    params = HLS_Para_Dict[elem][:]
    params.sort()
    para_Liste = []
    para_dict = {}
    para_id_dict = {}
    for n,para in enumerate(params):
        temp_para = select(n,para)
        para_dict[n] = para
        para_id_dict[para] = n
        para_Liste.append(temp_para)
    temp.Paras = para_Liste
    temp.Para_dict = para_dict
    temp.Para_id_dict = para_id_dict
    Liste_HK.Add(temp)

class HK_Familie(forms.WPFWindow):
    def __init__(self, xaml_file_name,Liste_HK):
        self.Liste_HK = Liste_HK
        forms.WPFWindow.__init__(self, xaml_file_name)
        self.tempcoll = ObservableCollection[MechanicalEquipment]()
        self.altdatagrid = Liste_HK
        self.dataGrid.ItemsSource = Liste_HK
        self.read_config()
    
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
                    try:
                        item.checked = HK_dict[item.Name][0]
                    except:
                        pass
                    try:
                        item.Vol_Id = item.Para_id_dict[HK_dict[item.Name][1]]
                    except:
                        pass
                    
                    try:
                        item.HL_Id = item.Para_id_dict[HK_dict[item.Name][2]]
                    except:
                        pass
            self.dataGrid.Items.Refresh()
         
        except:
            pass

        try:
            if config.Hide:
                self.click(self.aus)
                for item in self.dataGrid.Items:
                    if item.checked:
                        self.tempcoll.Add(item)
                self.dataGrid.ItemsSource = self.tempcoll
                self.dataGrid.Items.Refresh()
            else:
                self.click(self.ein)
                   
        except:
            pass
        
        
    def write_config(self):
        dict_ = {}
        for item in self.Liste_HK:
            if item.checked:
                try:
                    dict_[item.Name] = [item.checked,
                    item.Para_dict[item.Vol_Id],
                    item.Para_dict[item.HL_Id]
                    ]
                except:
                    pass

        config.HK_dict = dict_
        script.save_config()

    def ok(self,sender,args):
        self.write_config()
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
    
HK_Familie_Auswahl = HK_Familie("window.xaml",Liste_HK)
try:
    HK_Familie_Auswahl.ShowDialog()
except Exception as e:
    HK_Familie_Auswahl.Close()
    script.exit()

Familie_Bearbeiten = {}
for el in Liste_HK:
    if el.checked:
        for elem in HLS_Dict[el.Name]:
            Familie_Bearbeiten[elem] = [
                el.Para_dict[el.Vol_Id],
                el.Para_dict[el.HL_Id]]

if len(Familie_Bearbeiten.keys()) == 0:
    script.exit()


class Heizkoerper:
    def __init__(self, elem, Liste):
        self.element = elem
        self.vol = Liste[0]
        self.HL = Liste[1]
        try:
            self.RaumNr = self.element.Space[phase].Number
            self.Raum = self.element.Space[phase].Id.ToString()
        except:
            self.RaumNr = ''
            self.Raum = ''
            logger.error('kein MEP-Raum für Element {} gefunden'.format(self.element.Id.ToString()))
        try:
            self.leistung = get_value(self.element.LookupParameter(self.HL))
        except:
            self.leistung = 0
            logger.error("Parameter {} konnte nicht gefunden werden".format(self.HL))
        try:
            self.VOL = get_value(self.element.LookupParameter(self.vol))
        except:
            self.VOL = 0
            logger.error("Parameter {} konnte nicht gefunden werden".format(self.vol))

        self.VL_Temp = ''
        self.RL_Temp = ''
        self.Temperaturermittelung()
        self.Pruefen()
    
    def Temperaturermittelung(self):
        conns = self.element.MEPModel.ConnectorManager.Connectors
        for conn in conns:
            if conn.IsConnected:
                system  = conn.MEPSystem
                if system:
                    typ = doc.GetElement(system.get_Parameter(DB.BuiltInParameter.ELEM_TYPE_PARAM).AsElementId())
                    try:
                        abk = typ.LookupParameter('Abkürzung').AsString()
                        temp = get_value(typ.LookupParameter('Temperatur von Medium'))
                        if abk == 'H VL':
                            self.VL_Temp = temp
                        elif abk == 'H RL':
                            self.RL_Temp = temp
                    except:
                        pass
    def Pruefen(self):
        if not self.VL_Temp:
           self.VL_Temp = 0
        if not self.RL_Temp:
           self.RL_Temp = 100


MEP_Dict = {}
hkfamilieliste = Familie_Bearbeiten.keys()[:]

with forms.ProgressBar(title='{value}/{max_value} Heizkörper', cancellable=True, step=10) as pb:
    for n, hk in enumerate(hkfamilieliste):
        if pb.cancelled:
            script.exit()
        pb.update_progress(n + 1, len(hkfamilieliste))
        heizkoerper = Heizkoerper(hk,Familie_Bearbeiten[hk])
        if heizkoerper.Raum:
            if heizkoerper.Raum not in MEP_Dict.keys():
                MEP_Dict[heizkoerper.Raum] = [heizkoerper]
            else:
                MEP_Dict[heizkoerper.Raum].append(heizkoerper)

class MEPRaum:
    def __init__(self, elemid, liste_hk):
        self.element = doc.GetElement(DB.ElementId(int(elemid)))
        self.hk_liste = liste_hk
        self.leistung = 0
        self.VL_temp = -200
        self.RL_temp = -200
        self.Wasser = 0
        self.berechnen()


    def berechnen(self):
        liste_VL = []
        liste_RL = []
        for item in self.hk_liste:
            self.leistung += item.leistung
            self.Wasser -= item.VOL
            if item.VL_Temp not in liste_VL:
                liste_VL.append(item.VL_Temp)
            if item.RL_Temp not in liste_RL:
                liste_RL.append(item.RL_Temp)
        if len(liste_VL) == 1:
            self.VL_Temp = liste_VL[0]
        else:
            logger.error(50*'-')
            logger.error('Vorlauftemperatur von Heizkörper in MEPRaum {}, {} nicht gleich'.format(self.element.Number,self.element.Id.ToString()))
            for el in self.hk_liste:
                logger.error('Vorlauftemperatur: {} °C, Heizkörper-Id: {}'.format(el.VL_Temp,el.element.Id.ToString()))
            liste_VL_neu = [float(a) for a in liste_VL]
            self.VL_Temp = max(liste_VL_neu)
            logger.error('!!! Achtung: Heizkörper-Vorlauftemperatur in MEPRaum {}, {} wird auf {} °C eingesetzt'.format(self.element.Number,self.element.Id.ToString(),self.VL_Temp))
            logger.error(50*'-')

        if len(liste_RL) == 1:
            self.RL_Temp = liste_RL[0]
        else:
            logger.error(50*'-')
            logger.error('Rücklauftemperatur von hk in MEPRaum {}, {} nicht gleich'.format(self.element.Number,self.element.Id.ToString()))
            for el in self.hk_liste:
                logger.error('Rücklauftemperatur: {} °C, Heizkörper-Id: {}'.format(el.RL_Temp,el.element.Id.ToString()))
            liste_RL_neu = [float(a) for a in liste_RL]
            self.RL_Temp = min(liste_RL_neu)
            logger.error('!!! Achtung: Heizkörper-Rücklauftemperatur in MEPRaum {}, {} wird auf {} °C eingesetzt'.format(self.element.Number,self.element.Id.ToString(),self.RL_Temp))
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

        
        wert_schreiben("IGF_H_HK-VL_Temp", self.VL_Temp)
        wert_schreiben("IGF_H_HK-RL_Temp", self.RL_Temp)
        wert_schreiben("IGF_H_HK_Leistung", self.leistung)
        wert_schreiben("IGF_H_HK_Wassermenge", self.Wasser)
    

mep_liste = []
MEP_Liste = MEP_Dict.keys()[:]
with forms.ProgressBar(title='{value}/{max_value} MEP-Räume',cancellable=True, step=10) as pb1:
    for n, mep in enumerate(MEP_Liste):
        if pb1.cancelled:
            script.exit()
        pb1.update_progress(n + 1, len(MEP_Liste))
        mepraum = MEPRaum(mep,MEP_Dict[mep])
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
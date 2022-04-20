# coding: utf8
import sys
sys.path.append(r'R:\pyRevit\xx_Skripte\libs\IGF_libs')
from rpw import revit,DB
from pyrevit import script, forms
from IGF_log import getlog
from IGF_lib import get_value
from pyrevit.forms import WPFWindow
from System.Collections.ObjectModel import ObservableCollection


__title__ = "3.20 Ebenen sortieren"
__doc__ = """
Ebenen sortieren für Anlagen- und Schächteberechnung

[2021.11.26]
Version: 2.0
"""
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()

uidoc = revit.uidoc
doc = revit.doc

name = doc.ProjectInformation.Name
number = doc.ProjectInformation.Number

config = script.get_config(name+number+'Ebenensort')

try:
    getlog(__title__)
except:
    pass

# MEP Räume aus aktueller Projekt
spaces_collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MEPSpaces).WhereElementIsNotElementType()
spaces = spaces_collector.ToElementIds()
spaces_collector.Dispose()


class Ebenen(object):
    def __init__(self,Ebene):
        self.Ebene = Ebene
        self.Abk = ""
        self.Nr = ""

    @property
    def Abk(self):
        return self._Abk
    @Abk.setter
    def Abk(self, value):
        self._Abk = value
    @property
    def Nr(self):
        return self._Nr
    @Nr.setter
    def Nr(self, value):
        self._Nr = value

class MEPRaum:
    def __init__(self,elemid):
        self.elemid = elemid
        self.elem = doc.GetElement(self.elemid)
        self.ebene = self.elem.Level.Name
        self.abkuerzung = ""
        self.nr = ""

    @property
    def abkuerzung(self):
        return self._abkuerzung
    @abkuerzung.setter
    def abkuerzung(self, value):
        self._abkuerzung = value
    @property
    def nr(self):
        return self._nr
    @nr.setter
    def nr(self, value):
        self._nr = value
    
    def werte_schreiben(self):
        def wert_schreiben(param_name,wert):
            param = self.elem.LookupParameter(param_name)
            if param.StorageType.ToString() == 'Double':
                param.SetValueString(str(wert))
            else:
                param.Set(wert)
        wert_schreiben("IGF_RLT_Verteilung_EbenenSortieren",int(self.nr))
        wert_schreiben("IGF_RLT_Verteilung_EbenenName",str(self.abkuerzung))

Level_Liste = []
mepraum_liste = []

with forms.ProgressBar(title="{value}/{max_value} MEP Räume", cancellable=True, step=int(len(spaces)/100)+1) as pb:
    for n, spaceid in enumerate(spaces):
        if pb.cancelled:
            script.exit()
        pb.update_progress(n + 1, len(spaces))
        mepraum = MEPRaum(spaceid)
        if mepraum.ebene not in Level_Liste:
            Level_Liste.append(mepraum.ebene)
        mepraum_liste.append(mepraum)

sorted(Level_Liste)

Liste_Levels = ObservableCollection[Ebenen]()
for level in Level_Liste:
    Liste_Levels.Add(Ebenen(level))


class EbenenSort(WPFWindow):
    def __init__(self, xaml_file_name,liste):
        self.Liste = liste
        WPFWindow.__init__(self, xaml_file_name)
        self.LV_Ebene.ItemsSource = self.Liste
        self.read_config()
    
    def read_config(self):
        _dict = {}
        try:
            for item in config.liste:
                _dict[item[0]] = [item[1],item[2]]
            for item1 in self.Liste:
                if item1.Ebene in _dict.keys():
                    item1.Abk = _dict[item1.Ebene][0]
                    item1.Nr = _dict[item1.Ebene][1]
            self.LV_Ebene.Items.Refresh()
        except:
            config.liste = []
        
    def write_config(self):
        liste = []
        for el in self.Liste:
            liste.append([el.Ebene,el.Abk,el.Nr])
        config.liste = liste
        script.save_config()

    def ok(self,sender,args):
        self.write_config()
        self.Close()
    def abbrechen(self,sender,args):
        self.Close()
        script.exit()

ebenensort = EbenenSort('window.xaml',Liste_Levels)
try:
    ebenensort.ShowDialog()
except Exception as e:
    ebenensort.Close()
    logger.error(e)
    script.exit()

Ebenen_dict = {}

for item in Liste_Levels:
    Ebenen_dict[item.Ebene] = [item.Abk,item.Nr]

mepraum.abkuerzung
if forms.alert('Ebenesabkürzung und Ebenensortierungsnummer in MEP Räume schreiben?', ok=False, yes=True, no=True):
    t = DB.Transaction(doc,'Ebenen sortieren')
    t.Start()
    with forms.ProgressBar(title='{value}/{max_value} MEP Räume', cancellable=True, step=int(len(mepraum_liste)/200)+10) as pb:
        for n,mepraum in enumerate(mepraum_liste):
            if pb.cancelled:
                t.RollBack()
                script.exit()
            pb.update_progress(n + 1, len(mepraum_liste))
            try:
                mepraum.abkuerzung = Ebenen_dict[mepraum.ebene][0]
                mepraum.nr = Ebenen_dict[mepraum.ebene][1]
                mepraum.werte_schreiben()
            except Exception as e:
                logger.error(e)
    t.Commit()
    t.Dispose()
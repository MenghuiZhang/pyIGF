# coding: utf8
from IGF_log import getlog,getloglocal
from rpw import revit
from pyrevit import script, forms
from eventhandler import CHANGEFAMILY,Aktualisieren,Familien,LISTE_IS,ExternalEvent,NUMMERIEREN
from time import strftime, localtime
import os


__title__ = "3.60 Kombirahmen -> MagiCAD DistributionBox"
__doc__ = """

Kombirahmen --> MagiCAD DistributionBox
[2022.04.13]
Version: 1.1
"""
__authors__ = "Menghui Zhang"

try:
    getlog(__title__)
except:
    pass

try:
    getloglocal(__title__)
except:
    pass

logger = script.get_logger()

uidoc = revit.uidoc
doc = revit.doc


class AktuelleBerechnung(forms.WPFWindow):
    def __init__(self):
        self.changefamily = CHANGEFAMILY()
        self.aktualisieren = Aktualisieren()
        self.nummerieren = NUMMERIEREN()
        self.strftime = strftime
        self.localtime = localtime
        self.changefamilyEvent = ExternalEvent.Create(self.changefamily)
        self.aktualisierenEvent = ExternalEvent.Create(self.aktualisieren)
        self.nummerierenEvent = ExternalEvent.Create(self.nummerieren)
        
        self.Elems_Selected = []

        self.classFamilien = Familien
        self.liste_is = LISTE_IS                
        forms.WPFWindow.__init__(self,'window.xaml')
        self.kombi.ItemsSource = self.liste_is
        self.distribution.ItemsSource = self.liste_is
        self.startnum = 0
        self.get_startnum()
        self.get_vorlage()
        self.set_icon(os.path.join(os.path.dirname(__file__), 'icon.png'))
    
    def get_startnum(self):
        for el in self.liste_is:
            if el.name.find('DistributionBox_') != -1:
                try:
                    num = int(el.name[el.name.find('_')+1:])
                    if num > self.startnum:
                        self.startnum = num
                except:pass
        self.changefamily.startnum = self.startnum
        self.nummerieren.startnum = self.startnum
    def get_zeit(self):
        return self.strftime("%d.%m.%Y-%H:%M", self.localtime())

    
    def get_vorlage(self):
        for el in self.liste_is:
            if el.name == 'DistributionBox_IGF':
                self.distribution.SelectedItem = el
                self.changefamily.Vorlage = el.elem
                self.nummerieren.Vorlage = el.elem
                break

    def update(self, sender, args):
        self.aktualisierenEvent.Raise()

    def changetomodell(self, sender, args):
        self.changefamily.modell = True
        self.changefamily.ansicht = False
        self.changefamily.auswahl = False
        self.changefamily.elems = []
        self.aktu.IsEnabled = False

    def changetoansicht(self, sender, args):
        self.changefamily.modell = False
        self.changefamily.ansicht = True
        self.changefamily.auswahl = False
        self.changefamily.elems = []
        self.aktu.IsEnabled = False

    def changetoselect(self, sender, args):
        self.changefamily.modell = False
        self.changefamily.ansicht = False
        self.changefamily.auswahl = True
        self.changefamily.elems = self.Elems_Selected
        self.aktu.IsEnabled = True

    def sourcechanged(self, sender, args):
        self.changefamily.Kombi = self.kombi.SelectedItem.name
        self.aktualisieren.Muster = self.kombi.SelectedItem.name
    

    def neunummerieren(self, sender, args):
        if self.nummerieren.startnum < self.changefamily.startnum:
            self.nummerieren.startnum = self.changefamily.startnum
        self.nummerieren.zeit = self.get_zeit()
        self.nummerierenEvent.Raise()  

    
    def start(self, sender, args):
        if self.changefamily.auswahl:
            self.changefamily.elems = self.aktualisieren.Auswahl
        if self.nummerieren.startnum > self.changefamily.startnum:self.changefamily.startnum=self.nummerieren.startnum
        self.changefamilyEvent.Raise()   
       
wind = AktuelleBerechnung()
wind.Show()

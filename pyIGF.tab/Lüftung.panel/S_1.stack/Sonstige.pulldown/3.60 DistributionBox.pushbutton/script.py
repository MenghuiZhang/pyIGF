# coding: utf8
from IGF_log import getlog,getloglocal
from rpw import revit
from pyrevit import script, forms
from eventhandler import CHANGEFAMILY,Familien,LISTE_IS,ExternalEvent,NUMMERIEREN
from time import strftime, localtime
import os
from System.Windows import Visibility


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
        self.minvalue = 0
        self.maxvalue = 100
        self.value = 1
        self.PB_text = ''
        self.changefamily = CHANGEFAMILY()
        self.visible = Visibility.Visible
        self.hidden = Visibility.Hidden
        self.nummerieren = NUMMERIEREN()
        self.strftime = strftime
        self.localtime = localtime
        self.changefamilyEvent = ExternalEvent.Create(self.changefamily)
        self.nummerierenEvent = ExternalEvent.Create(self.nummerieren)
        
        self.classFamilien = Familien
        self.liste_is = LISTE_IS                
        forms.WPFWindow.__init__(self,'window.xaml',handle_esc=False)
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

    def neunummerieren(self, sender, args):
        self.neu_nummer.Visibility = self.hidden
        self.progress_neu.Visibility = self.visible
        self.title_p_neu.Visibility = self.visible

        self.nummerieren.zeit = self.get_zeit()
        self.nummerierenEvent.Raise()  

    
    def start(self, sender, args):
        

        self.changefamilyEvent.Raise()   
    
       
wind = AktuelleBerechnung()
wind.changefamily.class_GUI = wind
wind.nummerieren.class_GUI = wind
wind.Show()

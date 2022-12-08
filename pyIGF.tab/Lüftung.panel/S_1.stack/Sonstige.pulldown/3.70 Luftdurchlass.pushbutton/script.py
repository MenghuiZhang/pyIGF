# coding: utf8
from IGF_log import getlog,getloglocal
from rpw import revit
from pyrevit import script, forms
from eventhandler import CHANGEPARAM,ExternalEvent 
import os
from System.Windows import Visibility
import time


__title__ = "3.70 Volumenstrom Luftauslässe"

__doc__ = """

Min/Max/Nacht/TiefeNacht in Parameter Volumenstrom übertragen

[2022.05.19]
Version: 1.0
"""
__authors__ = "Menghui Zhang"

logger = script.get_logger()
try:getlog()
except:pass
try:getloglocal()
except:pass

class GUI(forms.WPFWindow):
    def __init__(self):
        self.changefamily = CHANGEPARAM()
        self.changefamilyEvent = ExternalEvent.Create(self.changefamily)
        forms.WPFWindow.__init__(self,'window.xaml',handle_esc=False)
        self.set_icon(os.path.join(os.path.dirname(__file__), 'icon.png'))
        self.visible = Visibility.Visible
        self.hidden = Visibility.Hidden
        self.Liste = [
            'IGF_RLT_AuslassVolumenstromMin',
            'IGF_RLT_AuslassVolumenstromMax',
            'IGF_RLT_AuslassVolumenstromNacht',
            'IGF_RLT_AuslassVolumenstromTiefeNacht',
            'Volumenstrom'
        ]
        self.In.ItemsSource = self.Liste
        self.Aus.ItemsSource = self.Liste
        
    def start(self, sender, args):
        
        self.button.Visibility = self.hidden
        self.progress.Visibility = self.visible
        self.title_p.Visibility = self.visible
        # n = 0
        # for n in range(10000):
        #     self.progress.Value = n+1
        #     print(n)
        #     if n%100 == 0:
                
        #         self.title_p.Text = str(n+1) + ' / 10000'
                
        self.changefamilyEvent.Raise()
        
    def valuechanged(self, sender, args):
        if self.progress.Value == self.progress.Maximum:
            self.progress.Value = self.progress.Minimum
            self.button.Visibility = self.visible
            self.progress.Visibility = self.hidden
            self.title_p.Visibility = self.hidden
       
gui = GUI()
gui.changefamily.class_GUI = gui
gui.Show()

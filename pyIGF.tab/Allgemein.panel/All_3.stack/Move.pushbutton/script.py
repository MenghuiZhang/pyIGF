# coding: utf8
from pyrevit import forms
from eventhandler import LINKS,RECHTS,OBEN,UNTEN,ExternalEvent

__title__ = "Verschieben"
__doc__ = """
Elements verschieben

[2022.04.19]
Version: 1.0
"""
__authors__ = "Menghui Zhang"



class Suche(forms.WPFWindow):
    def __init__(self):
        self.l = LINKS()
        self.r = RECHTS()
        self.o = OBEN()
        self.u = UNTEN()
        self.levent = ExternalEvent.Create(self.l)
        self.revent = ExternalEvent.Create(self.r)
        self.oevent = ExternalEvent.Create(self.o)
        self.uevent = ExternalEvent.Create(self.u)
        forms.WPFWindow.__init__(self, 'window.xaml')


    def link(self, sender, args):
        temp_zahl = self.zahl.Text
        zahl  = float(temp_zahl) * (-1)
        self.l.zahl = zahl
        self.levent.Raise()

    def recht(self, sender, args):
        temp_zahl = self.zahl.Text
        zahl  = float(temp_zahl)
        self.r.zahl = zahl
        self.revent.Raise()
    
    def oben(self, sender, args):
        temp_zahl = self.zahl.Text
        zahl  = float(temp_zahl)
        self.o.zahl = zahl
        self.oevent.Raise()
    def unten(self, sender, args):
        temp_zahl = self.zahl.Text
        zahl  = float(temp_zahl) * (-1)
        self.u.zahl = zahl
        self.uevent.Raise()


Suche = Suche()
Suche.Show()
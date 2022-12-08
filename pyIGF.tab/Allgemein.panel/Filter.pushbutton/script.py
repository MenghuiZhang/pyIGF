# coding: utf8
from IGF_log import getloglocal
from rpw import revit,DB
from pyrevit import script, forms
from externalevent import ExternalEvent,ExternalEventListe,ObservableCollection,Bauteil,Systemtyp
import os


__title__ = "Filter"
__doc__ = """

ergänzt: Systemtyp auswählen
[2022.12.02]
Version: 1.1
"""
__authors__ = "Menghui Zhang"

try:
    getloglocal(__title__)
except:
    pass

logger = script.get_logger()

uidoc = revit.uidoc
doc = revit.doc

class GUI(forms.WPFWindow):
    def __init__(self):
        self.temp_coll = ObservableCollection[Bauteil]()    
        self.temp_coll_Systemtyp = ObservableCollection[Systemtyp]()   
        self.externaleventhandler = ExternalEventListe()
        self.event = ExternalEvent.Create(self.externaleventhandler)
        forms.WPFWindow.__init__(self,'window.xaml',handle_esc=False)
        self.set_icon(os.path.join(os.path.dirname(__file__), 'IGF.png'))
        self.LV_System.ItemsSource = self.temp_coll_Systemtyp
        self.DG_Familie.ItemsSource = self.temp_coll
  
    def catechanged(self,sender,args):
        item = sender.DataContext
        checked = sender.IsChecked
        if self.LV_category.SelectedIndex != -1:
            if item in self.LV_category.SelectedItems:
                for el in self.LV_category.SelectedItems:el.checked = checked
        try:
            self.get_Systemtyp()
            self.get_item()
        except Exception as e:print(e)
    
    def get_Systemtyp(self):
        self.temp_coll_Systemtyp.Clear()
        Dict = {}
        for el in self.LV_category.Items:
            if el.checked:
                for system in el.Systems:
                    if system.systemtyp != '' and system.systemtyp not in Dict.keys():Dict[system.systemtyp] = system
        for systemname in sorted(Dict.keys()):
            self.temp_coll_Systemtyp.Add(Dict[systemname])
        
    def get_item(self):
        self.temp_coll.Clear()
        text = self.suche.Text
        for el in self.LV_category.Items:
            if el.checked:
                for item in el.Items:
                    item.elems = []
                    for system in el.Systems:
                        if system.checked:
                            if item.familyname in el.elem_mit_system[system.systemtyp].keys():
                                if item.typname in el.elem_mit_system[system.systemtyp][item.familyname].keys():
                                    item.elems.extend(el.elem_mit_system[system.systemtyp][item.familyname][item.typname])
                    item.anzahl = item.get_anzahl()
                    if item.anzahl > 0:
                        if not text:self.temp_coll.Add(item)
                        else:
                            if item.category.upper().find(text.upper()) != -1 or item.familyname.upper().find(text.upper()) != -1 or item.typname.upper().find(text.upper()) != -1:
                                self.temp_coll.Add(item)

    
    def systemchanged(self,sender,args):
        if self.LV_System.SelectedIndex == -1:
            return 
        item = sender.DataContext
        checked = sender.IsChecked
        if self.LV_System.SelectedIndex != -1:
            if item in self.LV_System.SelectedItems:
                for el in self.LV_System.SelectedItems:
                    el.checked = checked
                    for cate in self.LV_category.Items:
                        if cate.checked:
                            if el.systemtyp in cate.dict_Systems.keys():
                                cate.dict_Systems[el.systemtyp].checked = checked

        try:self.get_item()
        except Exception as e:print(e)


    def familiechanged(self,sender,e):
        item = sender.DataContext
        checked = sender.IsChecked
        if self.DG_Familie.SelectedIndex != -1:
            if item in self.DG_Familie.SelectedItems:
                for el in self.DG_Familie.SelectedItems:el.checked = checked

    def ausrevit(self,sender,e):
        self.externaleventhandler.Executeapp = self.externaleventhandler.Reset
        self.event.Raise()
  
    def inrevit(self,sender,e):
        self.externaleventhandler.Executeapp = self.externaleventhandler.PostSelect
        self.event.Raise()
    
    def suchechanged(self,sender,e):
        self.get_item()


gui = GUI()
gui.externaleventhandler.GUI = gui
gui.externaleventhandler.Executeapp = gui.externaleventhandler.Reset
gui.event.Raise()
gui.Show()
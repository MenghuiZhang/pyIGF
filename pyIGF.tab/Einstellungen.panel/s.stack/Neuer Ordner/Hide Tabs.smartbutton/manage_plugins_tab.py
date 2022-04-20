# coding: utf8
import sys
sys.path.append(r'R:\pyRevit\xx_Skripte\libs\IGF_libs')
from System.Collections.ObjectModel import ObservableCollection
from pyrevit import forms,revit
from pyrevit import script
from pyrevit.coreutils import ribbon

all_module = ribbon.get_current_ui()
logger = script.get_logger()


class Tab(object):
    def __init__(self,name,checked):
        self.name = name
        self.checked = checked



def get_all_plugins():
    """return all Plugins"""
    dict_Module = {}
    all_module = ribbon.get_current_ui()
    for modul in all_module:
        revit_obj = modul.get_rvtapi_object()
        if revit_obj.ToString() == 'Autodesk.Windows.RibbonTab' and modul.name != 'pyRevit':
            dict_Module[modul.name] = modul
    return dict_Module

def get_itemssource(_liste):
    """return Itemssorce für WPF"""
    Liste_Tab = ObservableCollection[Tab]()
    for el in _liste:
        Liste_Tab.Add(Tab(el,False))
    return Liste_Tab

def change_icon(hidden_tabs,button):
    if hidden_tabs > 0:
        icon = script.get_bundle_file('on.png')
    else:
        icon = script.get_bundle_file('off.png')
    button.set_icon(icon)

class WPF_UI(forms.WPFWindow):
    def __init__(self, xaml_file_name,liste_module,config):
        self.liste_module = liste_module
        self.config = config
        forms.WPFWindow.__init__(self, xaml_file_name)
        self.LVModule.ItemsSource = liste_module
        self.read_config()

    def read_config(self):
        try:
            hidetabs = self.config.hidetabs
            for el in self.liste_module:
                if el.name in hidetabs:
                    el.checked = True
            self.LVModule.Items.Refresh()

        except:
            try:
                self.config.hidetabs = []
            except:
                pass

    def write_config(self):
        try:
            hide_tabs = []
            for el in self.liste_module:
                if el.checked:
                    hide_tabs.append(el.name)
            self.config.hidetabs = hide_tabs
        except:
            pass
        script.save_config()
    
    def checkedchanged(self, sender, args):
        Checked = sender.IsChecked
        if self.LVModule.SelectedItem is not None:
            try:
                if sender.DataContext in self.LVModule.SelectedItems:
                    for item in self.LVModule.SelectedItems:
                        try:
                            item.checked = Checked
                        except:
                            pass
                    self.LVModule.Items.Refresh()
                else:
                    pass
            except:
                pass


    def ok(self,sender,args):
        self.Close()
        self.write_config()
    def checkall(self,sender,args):
        for item in self.LVModule.Items:
            item.checked = True
        self.LVModule.Items.Refresh()

    def uncheckall(self,sender,args):
        for item in self.LVModule.Items:
            item.checked = False
        self.LVModule.Items.Refresh()

    def toggleall(self,sender,args):
        for item in self.LVModule.Items:
            value = item.checked
            item.checked = not value
        self.LVModule.Items.Refresh()

    def abbrechen(self,sender,args):
        self.Close()
        script.exit()

def run_wpf(Liste_Tab,config):
    ui_tabs = WPF_UI("window.xaml",Liste_Tab,config)
    try:
        ui_tabs.ShowDialog()
    except Exception as e:
        logger.error(e)
        ui_tabs.Close()
        script.exit()

def setup_plugins(config,all_plugins):
    hide_plugins = config.hidetabs
    for el in all_plugins:
        if el.name in hide_plugins:
            el.visible = False
        else:
            el.visible = True
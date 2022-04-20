# coding: utf8
import sys
sys.path.append(r'R:\pyRevit\xx_Skripte\libs\IGF_libs')
import clr
clr.AddReference('AdWindows')
from Autodesk.Windows import ComponentManager
from System.Collections.ObjectModel import ObservableCollection
from pyrevit import forms,revit
from pyrevit import script
from pyrevit.coreutils import ribbon

__title__ = "ausgewählte Tab ausblenden"
__doc__ = """

blendet ausgewählte Module aus.

[2021.12.06]
Version: 1.1
"""
__author__ = "Menghui Zhang"

logger = script.get_logger()
config = script.get_config()

class Children(object):
    def __init__(self,name,checked,obj):
        self.name = name
        self.checked = checked
        self.object = obj
        self.children = ObservableCollection[Children]()
        self.parent = None
        self.expand = False

def get_all_plugins():
    """return all Plugins"""
    dict_Module = {}
    all_module = ComponentManager.Ribbon.Tabs
    for modul in all_module:
        revit_obj = modul.GetType()
        if revit_obj.ToString() == 'Autodesk.Windows.RibbonTab':
            dict_Module[modul.Title] = modul
    return dict_Module

def get_itemssource(_dict):
    """return Itemssorce für WPF"""
    Liste_Tab = ObservableCollection[Children]()
    for el in sorted(_dict.keys()):
        tab = _dict[el]
        Tab = Children(el,False,_dict[el])
        dict_objs = {panel.Source.AutomationName:panel for panel in tab.Panels}
        for sub_el in sorted(dict_objs.keys()):
            if el == 'pyRevit':
                if sub_el not in ['pyRevit','Einstellungen']:
                    panel = Children(sub_el,False,dict_objs[sub_el])
                    panel.parent = Tab
                    Tab.children.Add(panel)
            else:
                panel = Children(sub_el,False,dict_objs[sub_el])
                panel.parent = Tab
                Tab.children.Add(panel)
                
        Liste_Tab.Add(Tab)
    return Liste_Tab

def change_icon(config,button):
    hidden_tabs = config.hidetabs
    hidden_panels = config.hidepanels

    if hidden_tabs:
        if len(hidden_tabs) > 0:
            icon = script.get_bundle_file('on.png')
    else:
        if hidden_panels:
            if len(hidden_panels.keys()) > 0:
                icon = script.get_bundle_file('on.png')
            else:
                icon = script.get_bundle_file('off.png')
        else:
            icon = script.get_bundle_file('off.png')

    button.set_icon(icon)


class WPF_UI(forms.WPFWindow):
    def __init__(self, xaml_file_name,liste_module,config):
        self.liste_module = liste_module
        self.config = config
        forms.WPFWindow.__init__(self, xaml_file_name)
        self.treeView1.ItemsSource = liste_module
        self.read_config()
        self.update()

    def read_config(self):
        try:
            hidetabs = self.config.hidetabs
            for el in self.liste_module:
                if el.name in hidetabs:
                    el.checked = True
                    for panel in el.children:
                        panel.checked = True
        except:
            try:
                self.config.hidetabs = []
            except:
                pass
        try:
            hidepanels = self.config.hidepanels
            if len(hidepanels.keys()) > 0:
                for el in self.liste_module:
                    if el.name in hidepanels.keys():                 
                        for child in el.children:
                            if child.name in hidepanels[el.name]:
                                child.checked = True


        except Exception as e:
            print(e)
            try:
                self.config.hidetabs = {}
            except:
                pass

        self.treeView1.Items.Refresh()

    def write_config(self):
        try:
            hide_tabs = []
            for el in self.liste_module:
                if el.checked and el.name != 'pyRevit':
                    hide_tabs.append(el.name)
            self.config.hidetabs = hide_tabs
        except:
            pass

        try:
            hide_panels = {}
            for el in self.liste_module:
                if el.checked:
                    if el.name == 'pyRevit':
                        hide_panels['pyRevit'] = []
                        for sub_el in el.children:
                            if sub_el.checked:
                                hide_panels['pyRevit'].append(sub_el.name)
                else:
                    for sub_el in el.children:
                        if sub_el.checked:
                            if el.name not in hide_panels.keys():
                                hide_panels[el.name] = [sub_el.name]
                            else:
                                hide_panels[el.name].append(sub_el.name)
                
            self.config.hidepanels = hide_panels
        except:
            pass
        script.save_config()
    
    def checkedchanged(self, sender, args):
        Checked = sender.IsChecked
        item = sender.DataContext
        for child in sender.DataContext.children:
            child.checked = Checked
            child.expand = True
        parent = item.parent
        
        if parent:
            parent.expand = True
            results = [child.checked for child in parent.children]
            if all(results):
                parent.checked = True
            elif any(results):
                parent.checked = None
            elif not any(results):
                parent.checked = False
        self.treeView1.Items.Refresh()

    def update(self):
        for el in self.liste_module:
            results = [panel.checked for panel in el.children]
            if all(results):
                el.checked = True
            elif any(results):
                el.checked = None
            elif not any(results):
                el.checked = False
        self.treeView1.Items.Refresh()
    
    def all(self,sender,args):
        for item in self.treeView1.Items:
            item.checked = True
            for panel in item.children:
                panel.checked = True
        self.treeView1.Items.Refresh()

    def kein(self,sender,args):
        for item in self.treeView1.Items:
            item.checked = False
            for panel in item.children:
                panel.checked = False
        self.treeView1.Items.Refresh()

    def zu(self,sender,args):
        for item in self.treeView1.Items:
            item.expand = False
        self.treeView1.Items.Refresh()

    def aus(self,sender,args):
        for item in self.treeView1.Items:
            item.expand = True
        self.treeView1.Items.Refresh()

    def ab(self,sender,args):
        self.Close()
        script.exit()
    
    def ok(self,sender,args):
        self.Close()
        self.write_config()

def run_wpf(Liste_Tab,config):
    ui_tabs = WPF_UI("window.xaml",Liste_Tab,config)
    try:
        ui_tabs.ShowDialog()
    except Exception as e:
        logger.error(e)
        ui_tabs.Close()
        script.exit()

def set_visible(item,visible):
    try:
        item.Visible = visible
    except:
        pass
    try:
        item.IsVisible = visible
    except:
        pass

def setup_plugins(all_plugins):
    for el in all_plugins:
        if el.checked:
            if el.name != 'pyRevit':
                set_visible(el.object,False)
            else:
                for item in el.children:
                    if item.checked:
                        set_visible(item.object,False)
                    else:
                        set_visible(item.object,True)
        else:
            for item in el.children:
                if item.checked:
                    set_visible(item.object,False)
                else:
                    set_visible(item.object,True)


if __name__ == '__main__':
    all_plugins = get_all_plugins()
    itemssource = get_itemssource(all_plugins)
    run_wpf(itemssource,config)
    setup_plugins(itemssource)
    change_icon(config,script.get_button())
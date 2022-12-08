# coding: utf8
import clr
clr.AddReference('AdWindows')
from Autodesk.Windows import ComponentManager
from System.Collections.ObjectModel import ObservableCollection
from pyrevit import forms,revit
from pyrevit import script
import os.path as op

logger = script.get_logger()
FILES_DIR = op.dirname(__file__)
config = script.get_config('Module ausblenden')

class Children(object):
    def __init__(self,name,checked,obj):
        self.name = name
        self.checked = checked
        self.object = obj
        self.children = ObservableCollection[Children]()
        self.parent = None
        self.expand = False

DICT_ALLTABS = {}
ITEMSSOURCE = ObservableCollection[Children]()

def get_all_plugins():
    all_module = ComponentManager.Ribbon.Tabs
    for modul in all_module:
        revit_obj = modul.GetType()
        if revit_obj.ToString() == 'Autodesk.Windows.RibbonTab':
            DICT_ALLTABS[modul.Title] = modul

get_all_plugins()  

def get_itemssource():
    """return Itemssorce für WPF"""
    for el in sorted(DICT_ALLTABS.keys()):
        tab = DICT_ALLTABS[el]
        Tab = Children(el,True,DICT_ALLTABS[el])
        dict_objs = {}
        for panel in tab.Panels:
            if panel.Source.Items.Count > 0:
                dict_objs[panel.Source.AutomationName] = panel

        for sub_el in sorted(dict_objs.keys()):
            if el in ['pyIGF']:
                if sub_el not in ['Einstellungen']:
                    panel = Children(sub_el,True,dict_objs[sub_el])
                    panel.parent = Tab
                    Tab.children.Add(panel)
            else:
                panel = Children(sub_el,True,dict_objs[sub_el])
                panel.parent = Tab
                Tab.children.Add(panel)          
        ITEMSSOURCE.Add(Tab)

get_itemssource()

def change_icon(button = None, script = script):
    try:
        hidden_tabs = config.hidetabs
    except:
        hidden_tabs = None
    try:
        hidden_panels = config.hidepanels
    except:
        hidden_panels = None

    if hidden_tabs:
        if len(hidden_tabs) > 0:
            icon = script.get_bundle_file(op.join(FILES_DIR, 'on.png'))
    else:
        if hidden_panels:
            if len(hidden_panels.keys()) > 0:
                icon = script.get_bundle_file(op.join(FILES_DIR, 'on.png'))
            else:
                icon = script.get_bundle_file(op.join(FILES_DIR, 'off.png'))
        else:
            icon = script.get_bundle_file(op.join(FILES_DIR, 'off.png'))

    button.set_icon(icon)

def setup_tooltip(button = None,script = script):
    tips = "1. Button klicken\n2. Module auswählen\n3. OK klicken"
    button.set_tooltip_ext(tips)
    try:
        tooltip_image = script.get_bundle_file(op.join(FILES_DIR,'image_tip.png'))
        button.set_tooltip_image(tooltip_image)
    except:
        pass

    try:
        tooltip_video = script.get_bundle_file(op.join(FILES_DIR,'video_tip.mp4'))
        button.set_tooltip_video(tooltip_video)
    except:
        pass

class WPF_UI(forms.WPFWindow):
    def __init__(self):
        self.liste_module = ITEMSSOURCE
        self.config = script.get_config('Module ausblenden')
        forms.WPFWindow.__init__(self,op.join(FILES_DIR, 'window.xaml'))
        self.treeView1.ItemsSource = ITEMSSOURCE
        self.read_config()
        self.update()

    def read_config(self):
        try:
            hidetabs = self.config.hidetabs
            for el in self.liste_module:
                if el.name not in hidetabs:
                    el.checked = True
                    for panel in el.children:
                        panel.checked = True
                else:
                    el.checked = False
                    for panel in el.children:
                        panel.checked = False
        except:
            try:
                self.config.hidetabs = []
            except:pass
        try:
            hidepanels = self.config.hidepanels
            if len(hidepanels.keys()) > 0:
                for el in self.liste_module:
                    if el.name in hidepanels.keys():                 
                        for child in el.children:
                            if child.name in hidepanels[el.name]:
                                child.checked = False

        except:
            try:
                self.config.hidetabs = {}
            except Exception as e:
                print(e)

        self.treeView1.Items.Refresh()

    def write_config(self):
        try:
            hide_tabs = []
            for el in self.liste_module:
                if el.checked == False and el.name != 'pyIGF':
                    hide_tabs.append(el.name)
            self.config.hidetabs = hide_tabs
        except:pass

        try:
            hide_panels = {}
            for el in self.liste_module:
                if el.checked is False:
                    if el.name == 'pyIGF':
                        hide_panels[el.name] = []
                        for sub_el in el.children:
                            if sub_el.checked is False:
                                hide_panels[el.name].append(sub_el.name)
                else:
                    for sub_el in el.children:
                        if sub_el.checked is False:
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

def run_wpf():
    ui_tabs = WPF_UI()
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

def setup_plugins():
    for el in ITEMSSOURCE:
        if el.checked is False:
            if el.name != 'pyIGF':
                set_visible(el.object,False)
            else:
                for item in el.children:
                    if not item.checked:
                        set_visible(item.object,False)
                    else:
                        set_visible(item.object,True)
        else:
            set_visible(el.object,True)
            for item in el.children:
                if not item.checked:
                    set_visible(item.object,False)
                else:
                    set_visible(item.object,True)

def setup_plugins_init():
    config = script.get_config('Module ausblenden')
    try:
        hidden_tabs = config.hidetabs
    except:
        hidden_tabs = []
    try:
        hidden_panels = config.hidepanels
    except:
        hidden_panels = {}
    for el in ITEMSSOURCE:
        if len(hidden_tabs) > 0:
            if el.name in hidden_tabs:
                set_visible(el.object,False)
            else:
                if len(hidden_panels.keys()) >0:
                    if el.name in hidden_panels.keys():
                        for item in el.children:
                            if item.name in hidden_panels[el.name]:
                                set_visible(item.object,False)
                            else:
                                set_visible(item.object,True)
                    else:
                        set_visible(el.object,True)
                else:
                    set_visible(el.object,True)
        else:
            set_visible(el.object,True)
            if len(hidden_panels.keys()) >0:
                if el.name in hidden_panels.keys():
                    for item in el.children:
                        if item.name in hidden_panels[el.name]:
                            set_visible(item.object,False)
                        else:
                            set_visible(item.object,True)
                else:
                    set_visible(el.object,True)
            else:
                set_visible(el.object,True)             
# coding: utf8
import clr
from pyrevit import revit, DB, script, UI
from System.Collections.ObjectModel import ObservableCollection
from pyrevit.forms import WPFWindow
from System.Windows import Visibility,Thickness  
from System.Windows import FontWeights
from pyIGF_config import Server_config

__title__ = "Einstellungen"
__doc__ = """Grundeinstellungen von Pläne festlegen"""
__author__ = "Menghui Zhang"

from pyIGF_logInfo import getlog
getlog(__title__)

logger = script.get_logger()
output = script.get_output()
config = script.get_config('Plaene')
Server_Config = Server_config()
server_config = Server_Config.get_config('Plaene')

uidoc = revit.uidoc
doc = revit.doc

hide = Visibility.Hidden
show = Visibility.Visible

coll = DB.FilteredElementCollector(doc).OfClass(clr.GetClrType(DB.Family))
titleblock_family = []
titleblock = []

coll1 = DB.FilteredElementCollector(doc).OfClass(clr.GetClrType(DB.ElementType))
viewport = []

for el in coll:
    if el.FamilyCategoryId.ToString() == '-2000280':
        titleblock_family.append(el)

for el in coll1:
    if el.FamilyName == 'Ansichtsfenster':
        viewport.append(el.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString())
viewport.sort()

for el in titleblock_family:
    symids = el.GetFamilySymbolIds()
    for id in symids:
        name = doc.GetElement(id).get_Parameter(DB.BuiltInParameter.SYMBOL_FAMILY_AND_TYPE_NAMES_PARAM).AsString()
        titleblock.append(name)
titleblock.sort()

class PlaeneUI(WPFWindow):
    def __init__(self, xaml_file_name):
        WPFWindow.__init__(self, xaml_file_name)
        try:
            self.confi = config.getconfig
            if self.confi == True:
                self.serv.FontWeight = FontWeights.Bold
                self.conf = server_config
            else:
                self.loca.FontWeight = FontWeights.Bold
                self.conf = config
        except:
            self.conf = server_config
            self.serv.FontWeight = FontWeights.Bold
        
        try:
            self.read_config(self.conf)
        except:
            pass
        self.backAll()
        self.planerstellen.FontWeight = FontWeights.Bold
        self.planerstellen.BorderThickness = Thickness(1,1,2,0)
        self.erstellung.Visibility = show
        self.erstellung_plankopf.ItemsSource = titleblock
        self.LG_erstellung.ItemsSource = viewport
        self.HA_erstellung.ItemsSource = viewport
        self.LG_Ansicht_anpassen.ItemsSource = viewport
        self.HA_Ansicht_anpassen.ItemsSource = viewport
            

    def backAll(self):
        self.planerstellen.FontWeight = FontWeights.Normal
        self.planerstellen.BorderThickness = Thickness(1,1,2,2)
        self.plananpassen.FontWeight = FontWeights.Normal
        self.plananpassen.BorderThickness = Thickness(0,1,2,2)
        self._3_co.FontWeight = FontWeights.Normal
        self._3_co.BorderThickness = Thickness(0,1,2,2)
        self._4_co.FontWeight = FontWeights.Normal
        self._4_co.BorderThickness = Thickness(0,1,2,2)
        self.erstellung.Visibility = hide
        self.anpassung.Visibility = hide
        # self.test.Visibility = hide
        # self.test.Visibility = hide
    
    def erstellen(self,sender,args):
        if config.getconfig:
            if self.Server_Pruefen():
                if self.server_changed():
                    self.write_config(server_config)
                    Server_Config.save_config()
                else:
                    self.read_config(server_config)
        else:
            self.write_config(config)


        self.backAll()
        self.planerstellen.FontWeight = FontWeights.Bold
        self.planerstellen.BorderThickness = Thickness(1,1,2,0)
        self.erstellung.Visibility = show
        
    def anpassen(self,sender,args):
        if config.getconfig:
            if self.Server_Pruefen():
                if self.server_changed():
                    self.write_config(server_config)
                    Server_Config.save_config()
                else:
                    self.read_config(server_config)
        else:
            self.write_config(config)
        self.backAll()
        self.plananpassen.FontWeight = FontWeights.Bold
        self.plananpassen.BorderThickness = Thickness(0,1,2,0)
        self.anpassung.Visibility = show
    def _3(self,sender,args):
        self.backAll()
        self._3_co.FontWeight = FontWeights.Bold
        self._3_co.BorderThickness = Thickness(0,1,2,0)
        
    def _4(self,sender,args):
        self.backAll()
        self._4_co.FontWeight = FontWeights.Bold
        self._4_co.BorderThickness = Thickness(0,1,2,0)

    def read_config(self,configdatei):
        # Pläne erstellen
        try:
            self.bz_l_erstellung.Text = configdatei.bz_l_erstellung
        except:
            self.bz_l_erstellung.Text = configdatei.bz_l_erstellung = '10'
        try:
            self.bz_r_erstellung.Text = configdatei.bz_r_erstellung
        except:
            self.bz_r_erstellung.Text = configdatei.bz_r_erstellung = '10'
        try:
            self.bz_o_erstellung.Text = configdatei.bz_o_erstellung
        except:
            self.bz_o_erstellung.Text = configdatei.bz_o_erstellung = '10'
        try:
            self.bz_u_erstellung.Text = configdatei.bz_u_erstellung
        except:
            self.bz_u_erstellung.Text = configdatei.bz_u_erstellung = '10'
        
        try:
            if configdatei.HA_erstellung in viewport:
                self.HA_erstellung.Text = configdatei.HA_erstellung
            else:
                self.HA_erstellung.Text = configdatei.HA_erstellung = ''
        except:
            self.HA_erstellung.Text = configdatei.HA_erstellung = ''
        try:
            if configdatei.LG_erstellung in viewport:
                self.LG_erstellung.Text = configdatei.LG_erstellung
            else:
                self.LG_erstellung.Text = configdatei.LG_erstellung = ''
        except:
            self.LG_erstellung.Text = configdatei.LG_erstellung = ''
        
        try:
            if configdatei.erstellung_plankopf in titleblock:
                self.erstellung_plankopf.Text = configdatei.erstellung_plankopf
            else:
                self.erstellung_plankopf.Text = configdatei.erstellung_plankopf = ''
        except:
            self.erstellung_plankopf.Text = configdatei.erstellung_plankopf = ''
        try:
            self.pk_l_erstellung.Text = configdatei.pk_l_erstellung
        except:
            self.pk_l_erstellung.Text = configdatei.pk_l_erstellung = '20'
        try:
            self.pk_r_erstellung.Text = configdatei.pk_r_erstellung
        except:
            self.pk_r_erstellung.Text = configdatei.pk_r_erstellung = '5'
        try:
            self.pk_o_erstellung.Text = configdatei.pk_o_erstellung
        except:
            self.pk_o_erstellung.Text = configdatei.pk_o_erstellung = '5'
        try:
            self.pk_u_erstellung.Text = configdatei.pk_u_erstellung
        except:
            self.pk_u_erstellung.Text = configdatei.pk_u_erstellung = '5'
        
        try: 
            self.gewerke_erstellung.Text = configdatei.gewerke_erstellung
        except:
            self.gewerke_erstellung.Text = configdatei.gewerke_erstellung = 'Koordination, Heizung/Kälte, Lüftung, Sanitär, technische Gase, Gebäudeautomation'
        
        # Raster anpassen
        try:
            self.bz_l_anpassen.Text = configdatei.bz_l_anpassen
        except:
            self.bz_l_anpassen.Text = configdatei.bz_l_anpassen = '10'
        try:
            self.bz_r_anpassen.Text = configdatei.bz_r_anpassen
        except:
            self.bz_r_anpassen.Text = configdatei.bz_r_anpassen = '10'
        try:
            self.bz_o_anpassen.Text = configdatei.bz_o_anpassen
        except:
            self.bz_o_anpassen.Text = configdatei.bz_o_anpassen = '10'
        try:
            self.bz_u_anpassen.Text = configdatei.bz_u_anpassen
        except:
            self.bz_u_anpassen.Text = configdatei.bz_u_anpassen = '10'
        
        try:
            if configdatei.HA_Ansicht_anpassen in viewport:
                self.HA_Ansicht_anpassen.Text = configdatei.HA_Ansicht_anpassen
            else:
                self.HA_Ansicht_anpassen.Text = configdatei.HA_Ansicht_anpassen = ''
        except:
            self.HA_Ansicht_anpassen.Text = configdatei.HA_Ansicht_anpassen = ''
        try:
            if configdatei.LG_Ansicht_anpassen in viewport:
                self.LG_Ansicht_anpassen.Text = configdatei.LG_Ansicht_anpassen
            else:
                self.LG_Ansicht_anpassen.Text = configdatei.LG_Ansicht_anpassen = ''
        except:
            self.LG_Ansicht_anpassen.Text = configdatei.LG_Ansicht_anpassen = ''
        

        try:
            self.pk_l_anpassen.Text = configdatei.pk_l_anpassen
        except:
            self.pk_l_anpassen.Text = configdatei.pk_l_anpassen = '20'
        try:
            self.pk_r_anpassen.Text = configdatei.pk_r_anpassen
        except:
            self.pk_r_anpassen.Text = configdatei.pk_r_anpassen = '5'
        try:
            self.pk_o_anpassen.Text = configdatei.pk_o_anpassen
        except:
            self.pk_o_anpassen.Text = configdatei.pk_o_anpassen = '5'
        try:
            self.pk_u_anpassen.Text = configdatei.pk_u_anpassen
        except:
            self.pk_u_anpassen.Text = configdatei.pk_u_anpassen = '5'
        
        try:
            self.raster_anpassen.IsChecked = configdatei.raster_anpassen
        except:
            self.raster_anpassen.IsChecked = configdatei.raster_anpassen = False
        
        try:
            self.Haupt_anpassen.IsChecked = configdatei.Haupt_anpassen
        except:
            self.Haupt_anpassen.IsChecked = configdatei.Haupt_anpassen = False
        
        try:
            self.legend_anpassen.IsChecked = configdatei.legend_anpassen
        except:
            self.legend_anpassen.IsChecked = configdatei.legend_anpassen = False
        
        
    def rueck(self, sender, args):
        self.read_config(server_config)
        self.write_config(config)
        self.loca.FontWeight = FontWeights.Bold
        self.serv.FontWeight = FontWeights.Normal
    
    def server_changed(self):
        buttons = UI.TaskDialogCommonButtons.No | UI.TaskDialogCommonButtons.Yes
        Task = UI.TaskDialog.Show('Warnung','Möchten Sie die Server-Einstellungen ändern?',buttons)
        if Task.ToString() == 'Yes':
            return True
        elif Task.ToString() == 'No':
            return False
    
    def Server_Pruefen(self):
        # Pläne erstellen
        if self.erstellung.Visibility == show:
            Pr1 = server_config.bz_l_erstellung == self.bz_l_erstellung.Text
            Pr2 = server_config.bz_r_erstellung == self.bz_r_erstellung.Text
            Pr3 = server_config.bz_o_erstellung == self.bz_o_erstellung.Text
            Pr4 = server_config.bz_u_erstellung == self.bz_u_erstellung.Text
            Pr5 = server_config.pk_l_erstellung == self.pk_l_erstellung.Text
            Pr6 = server_config.pk_r_erstellung == self.pk_r_erstellung.Text
            Pr7 = server_config.pk_o_erstellung == self.pk_o_erstellung.Text
            Pr8 = server_config.pk_u_erstellung == self.pk_u_erstellung.Text
            try:
                Pr9 = server_config.erstellung_plankopf == self.erstellung_plankopf.SelectedItem.ToString()
            except:
                Pr9 = server_config.erstellung_plankopf == self.erstellung_plankopf.Text

            Pr10 = server_config.gewerke_erstellung == self.gewerke_erstellung.Text
            try:
                Pr11 = server_config.HA_erstellung == self.HA_erstellung.SelectedItem.ToString()
            except:
                Pr11 = server_config.HA_erstellung == self.HA_erstellung.Text
            try:
                Pr12 = server_config.LG_erstellung == self.LG_erstellung.SelectedItem.ToString()
            except:
                Pr12 = server_config.LG_erstellung == self.LG_erstellung.Text

            if all([Pr1,Pr2,Pr3,Pr4,Pr5,Pr6,Pr7,Pr8,Pr9,Pr10,Pr11,Pr12]):
                return False
            else:
                return True
        if self.anpassung.Visibility == show:
            Pr1 = server_config.bz_l_anpassen == self.bz_l_anpassen.Text
            Pr2 = server_config.bz_r_anpassen == self.bz_r_anpassen.Text
            Pr3 = server_config.bz_o_anpassen == self.bz_o_anpassen.Text
            Pr4 = server_config.bz_u_anpassen == self.bz_u_anpassen.Text
            Pr5 = server_config.pk_l_anpassen == self.pk_l_anpassen.Text
            Pr6 = server_config.pk_r_anpassen == self.pk_r_anpassen.Text
            Pr7 = server_config.pk_o_anpassen == self.pk_o_anpassen.Text
            Pr8 = server_config.pk_u_anpassen == self.pk_u_anpassen.Text
            Pr9 = server_config.Haupt_anpassen == self.Haupt_anpassen.IsChecked
            Pr10 = server_config.raster_anpassen == self.raster_anpassen.IsChecked
            Pr11 = server_config.legend_anpassen == self.legend_anpassen.IsChecked

            try:
                Pr12 = server_config.HA_Ansicht_anpassen == self.HA_Ansicht_anpassen.SelectedItem.ToString()
            except:
                Pr12 = server_config.HA_Ansicht_anpassen == self.HA_Ansicht_anpassen.Text
            try:
                Pr13 = server_config.LG_Ansicht_anpassen == self.LG_Ansicht_anpassen.SelectedItem.ToString()
            except:
                Pr13 = server_config.LG_Ansicht_anpassen == self.LG_Ansicht_anpassen.Text

            if all([Pr1,Pr2,Pr3,Pr4,Pr5,Pr6,Pr7,Pr8,Pr9,Pr10,Pr11,Pr12,Pr13]):
                return False
            else:
                return True


    def write_config(self,configdatei):
        # Pläne erstellen
        if self.erstellung.Visibility == show:
            try:
                configdatei.bz_l_erstellung = self.bz_l_erstellung.Text
            except:
                pass

            try:
                configdatei.bz_r_erstellung = self.bz_r_erstellung.Text
            except:
                pass

            try:
                configdatei.bz_o_erstellung = self.bz_o_erstellung.Text
            except:
                pass

            try:
                configdatei.bz_u_erstellung = self.bz_u_erstellung.Text
            except:
                pass

            try:
                configdatei.pk_u_erstellung = self.pk_u_erstellung.Text
            except:
                pass

            try:
                configdatei.pk_o_erstellung = self.pk_o_erstellung.Text
            except:
                pass

            try:
                configdatei.pk_l_erstellung = self.pk_l_erstellung.Text
            except:
                pass

            try:
                configdatei.pk_r_erstellung = self.pk_r_erstellung.Text
            except:
                pass

            try:
                configdatei.HA_erstellung = self.HA_erstellung.Text
            except:
                pass

            try:
                configdatei.LG_erstellung = self.LG_erstellung.Text
            except:
                pass

            try:
                configdatei.erstellung_plankopf = self.erstellung_plankopf.Text
            except:
                pass

            try: 
                configdatei.gewerke_erstellung = self.gewerke_erstellung.Text
            except:
                pass
        # Raster anpassen
        if self.anpassung.Visibility == show:
            try:
                configdatei.raster_anpassen = self.raster_anpassen.IsChecked
            except:
                pass
            try:
                configdatei.Haupt_anpassen = self.Haupt_anpassen.IsChecked
            except:
                pass
            try:
                configdatei.legend_anpassen = self.legend_anpassen.IsChecked
            except:
                pass
            
            try:
                configdatei.bz_l_anpassen = self.bz_l_anpassen.Text
            except:
                pass

            try:
                configdatei.bz_r_anpassen = self.bz_r_anpassen.Text
            except:
                pass

            try:
                configdatei.bz_o_anpassen = self.bz_o_anpassen.Text
            except:
                pass

            try:
                configdatei.bz_u_anpassen = self.bz_u_anpassen.Text
            except:
                pass

            try:
                configdatei.pk_u_anpassen = self.pk_u_anpassen.Text
            except:
                pass

            try:
                configdatei.pk_o_anpassen = self.pk_o_anpassen.Text
            except:
                pass

            try:
                configdatei.pk_l_anpassen = self.pk_l_anpassen.Text
            except:
                pass

            try:
                configdatei.pk_r_anpassen = self.pk_r_anpassen.Text
            except:
                pass

            try:
                configdatei.HA_Ansicht_anpassen = self.HA_Ansicht_anpassen.Text
            except:
                pass

            try:
                configdatei.LG_Ansicht_anpassen = self.LG_Ansicht_anpassen.Text
            except:
                pass


    def ok(self, sender, args):
        if config.getconfig:
            self.write_config(server_config)
            Server_Config.save_config()
            script.save_config()
        else:
            self.write_config(config)
            script.save_config()
            Server_Config.save_config()
        self.Close()
    def anwenden(self, sender, args):
        if config.getconfig:
            if self.Server_Pruefen():
                if self.server_changed():
                    self.write_config(server_config)
                    Server_Config.save_config()
                else:
                    self.read_config(server_config)
        else:
            self.write_config(config)
            script.save_config()
  
    def abbrechen(self, sender, args):
        self.Close()

    def serve(self, sender, args):
        self.serv.FontWeight = FontWeights.Bold
        self.loca.FontWeight = FontWeights.Normal
        try:
            if not config.getconfig:
                self.write_config(config)
        except:
            pass

        try:
            config.getconfig = True
        except:
            pass
        self.read_config(server_config)
       
    def local(self, sender, args):
        self.loca.FontWeight = FontWeights.Bold
        self.serv.FontWeight = FontWeights.Normal
        try:
            if config.getconfig:
                if self.Server_Pruefen():
                    if self.server_changed():
                        self.write_config(server_config)
                        Server_Config.save_config()
                    else:
                        pass
        except:
            pass
        try:
            config.getconfig = False
        except:
            pass
        self.read_config(config)
        
Planfenster = PlaeneUI("window.xaml")

try:
    Planfenster.ShowDialog()
except Exception as e:
    logger.error(e)
    Planfenster.Close()

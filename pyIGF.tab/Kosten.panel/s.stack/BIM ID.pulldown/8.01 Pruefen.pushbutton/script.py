# coding: utf8
from pyrevit import revit, UI, DB, script, forms
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from System.Collections.ObjectModel import ObservableCollection
from System.Windows import FontWeights, FontStyles
from System.Windows.Controls import DataGridEditingUnit
from System.Windows.Media import Brushes, BrushConverter
from pyrevit.forms import WPFWindow, SelectFromList
import clr

__title__ = "8.01 BIM-ID Prüfen"
__doc__ = """..."""
__author__ = "Menghui Zhang"

from pyIGF_logInfo import getlog
getlog(__title__)


logger = script.get_logger()
output = script.get_output()
Element_config = script.get_config()

uidoc = revit.uidoc
doc = revit.doc

system_luft = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctSystem).WhereElementIsElementType()
system_rohr = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipingSystem).WhereElementIsElementType()
system_elek = FilteredElementCollector(doc).OfClass(clr.GetClrType(Electrical.ElectricalSystem)).WhereElementIsElementType()

class System(object):
    def __init__(self):
        self.Gewerke = ''
        self.ID = ''
    @property
    def SysName(self):
        return self._SysName
    @SysName.setter
    def SysName(self, value):
        self._SysName = value
    @property
    def Gewerke(self):
        return self._Gewerke
    @Gewerke.setter
    def Gewerke(self, value):
        self._Gewerke = value
    @property
    def KG(self):
        return self._KG
    @KG.setter
    def KG(self, value):
        self._KG = value
    @property
    def KN_01(self):
        return self._KN_01
    @KN_01.setter
    def KN_01(self, value):
        self._KN_01 = value
    @property
    def KN_02(self):
        return self._KN_02
    @KN_02.setter
    def KN_02(self, value):
        self._KN_02 = value
    @property
    def ID(self):
        return self._ID
    @ID.setter
    def ID(self, value):
        self._ID = value

SysListe = {}
Liste_Luft = ObservableCollection[System]()
Liste_Rohr = ObservableCollection[System]()
Liste_Elektro = ObservableCollection[System]()
Liste_Alle = ObservableCollection[System]()

def DatenErmitteln(Sys_coll,datatabelle):
    for el in Sys_coll:
        temp_system = System()
        sys_Name = el.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString()
        KG = el.LookupParameter('IGF_X_Kostengruppe').AsValueString()
        Gewerke = el.LookupParameter('IGF_X_Gewerkkürzel').AsString()
        KN_01 = el.LookupParameter('IGF_X_Kennnummer_1').AsValueString()
        KN_02 = el.LookupParameter('IGF_X_Kennnummer_2').AsValueString()
        ID = el.LookupParameter('IGF_X_BIM-ID').AsString()
        if sys_Name:
            temp_system.SysName = sys_Name
        if Gewerke:
            temp_system.Gewerke = Gewerke
        if KG:
            temp_system.KG = KG
        if KN_01:
            temp_system.KN_01 = KN_01
        if KN_02:
            temp_system.KN_02 = KN_02
        if ID:
            temp_system.ID = ID
        datatabelle.Add(temp_system)
        Liste_Alle.Add(temp_system)
        SysListe[sys_Name] = el

class Pruefen(WPFWindow):
    def __init__(self, xaml_file_name,liste_Luft,liste_Rohr,liste_Elektro,liste_All):
        self.liste_Luft = liste_Luft
        self.liste_Rohr = liste_Rohr
        self.liste_Elektro = liste_Elektro
        self.liste_All = liste_All
        WPFWindow.__init__(self, xaml_file_name)
        try:
            self.dataGrid.ItemsSource = self.liste_All
            self.backAll()
            self.click(self.all)
        except Exception as e:
            logger.error(e)

    def CellEditEnding(self, sender, args):
        newValue = args.EditingElement.Text
        Header = args.Column.Header
        item = args.Row.Item
        if Header == "Gewerke":
            ge = newValue
            KN01 = item.KN_01
            KN02 = item.KN_02
            KG = item.KG
            id = ge + '_' + KG + '_' + KN01 + ' ' + KN02
            args.Row.Item.ID = id
        elif Header == "KG":
            ge = item.Gewerke
            KN01 = item.KN_01
            KN02 = item.KN_02
            KG = newValue
            id = ge + '_' + KG + '_' + KN01 + ' ' + KN02
            args.Row.Item.ID = id
        elif Header == "KN-01":
            ge = item.Gewerke
            KN01 = newValue
            KN02 = item.KN_02
            KG = item.KG
            id = ge + '_' + KG + '_' + KN01 + ' ' + KN02
            args.Row.Item.ID = id
        elif Header == "KN-02":
            ge = item.Gewerke
            KN01 = item.KN_01
            KN02 = newValue
            KG = item.KG
            id = ge + '_' + KG + '_' + KN01 + ' ' + KN02
            args.Row.Item.ID = id

        # elif Header == "BIM-ID":
        #     ge =
        #     KN01 = item.KN_01
        #     KN02 = newValue
        #     KG = item.KG
        #     id = ge + '_' + KG + '_' + KN01 + ' ' + KN02
        #     args.Row.Item.ID = id

    def close(self,sender,args):
        self.Close()

    def update(self,sender,args):
        self.dataGrid.Items.Refresh()

    def ok(self,sender,args):
        self.Close()
        with forms.ProgressBar(title="{value}/{max_value} Systeme",
                               cancellable=True, step=10) as pb:

            n_1 = 0
            t = Transaction(doc,'BIM-ID schreiben')
            t.Start()
            for el in Liste_Alle:
                if pb.cancelled:
                    t.RollBack()
                    script.exit()
                n_1 += 1
                pb.update_progress(n_1, len(SysListe.Keys))
                systname = el.SysName
                item = SysListe[systname]
                item.LookupParameter('IGF_X_Gewerkkürzel').Set(str(el.Gewerke))
                item.LookupParameter('IGF_X_Kennnummer_1').Set(str(el.KN_01))
                item.LookupParameter('IGF_X_Kennnummer_2').Set(str(el.KN_02))
                item.LookupParameter('IGF_X_Kostengruppe').Set(str(el.KG))
                item.LookupParameter('IGF_X_BIM-ID').Set(str(el.ID))
            t.Commit()

    def click(self,button):
        button.Background = BrushConverter().ConvertFromString("#FF707070")
        button.FontWeight = FontWeights.Bold
        button.FontStyle = FontStyles.Italic

# Button Farbe zurücksetzen
    def back(self,button):
        button.Background  = Brushes.White
        button.FontWeight = FontWeights.Normal
        button.FontStyle = FontStyles.Normal

    def backAll(self):
        self.back(self.all)
        self.back(self.luft)
        self.back(self.rohr)
        self.back(self.elek)

    def lueftung(self,sender,args):
        self.backAll()
        self.click(self.luft)

        self.dataGrid.ItemsSource = self.liste_Luft
        self.temp = self.liste_Luft
        self.dataGrid.Items.Refresh()

    def rohre(self,sender,args):
        self.backAll()
        self.click(self.rohr)

        self.dataGrid.ItemsSource = self.liste_Rohr
        self.temp = self.liste_Rohr
        self.dataGrid.Items.Refresh()

    def elektro(self,sender,args):
        self.backAll()
        self.click(self.elek)

        self.dataGrid.ItemsSource = self.liste_Elektro
        self.temp = self.liste_Elektro
        self.dataGrid.Items.Refresh()

    def alle(self,sender,args):
        self.backAll()
        self.click(self.all)

        self.dataGrid.ItemsSource = self.liste_All
        self.temp = self.liste_All
        self.dataGrid.Items.Refresh()

DatenErmitteln(system_luft,Liste_Luft)
system_luft.Dispose()
DatenErmitteln(system_rohr,Liste_Rohr)
system_rohr.Dispose()
DatenErmitteln(system_elek,Liste_Elektro)
system_elek.Dispose()

Sys_Pruefen = Pruefen('Pruefen.xaml', Liste_Luft,Liste_Rohr,Liste_Elektro,Liste_Alle)
Sys_Pruefen.ShowDialog()

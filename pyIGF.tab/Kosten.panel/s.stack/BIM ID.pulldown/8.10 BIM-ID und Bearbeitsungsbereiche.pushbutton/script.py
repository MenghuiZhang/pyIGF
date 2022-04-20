# coding: utf8
import sys
sys.path.append(r'R:\pyRevit\xx_Skripte\libs\IGF_libs')
from IGF_log import getlog
from rpw import revit, DB
from pyrevit import script, forms
import clr
from System.Collections.ObjectModel import ObservableCollection
from System.Windows import FontWeights, FontStyles
from System.Windows.Media import Brushes, BrushConverter
from System.Windows.Forms import OpenFileDialog,DialogResult
from System.Windows.Controls import *
from pyrevit.forms import WPFWindow
clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as Excel
from System.Runtime.InteropServices import Marshal
from IGF_forms import abfrage

__title__ = "8.10 BIM-ID und Bearbeitsungsbereiche"
__doc__ = """BIM-ID und Bearbeitunsbereich aus excel in Modell schreiben

[2021.11.23]
Version: 2.0
"""
__authors__ = "Menghui Zhang"

try:
    getlog(__title__)
except:
    pass

uidoc = revit.uidoc
doc = revit.doc

logger = script.get_logger()
output = script.get_output()

name = doc.ProjectInformation.Name
number = doc.ProjectInformation.Number

bimid_config = script.get_config(name+number+'BIM-ID und Bearbeitsungsbereiche')

muster_bb = ['KG4xx_Musterbereich']

# Bearbeitungsbereich
worksets = DB.FilteredWorksetCollector(doc).OfKind(DB.WorksetKind.UserWorkset)
Workset_dict = {}
for el in worksets:
    Workset_dict[el.Name] = el.Id.ToString()

# Exceldaten
class Exceldaten(object):
    def __init__(self):
        self.checked = False
        self.bb = False
        self.Systemname = ''
        self.GK = ''
        self.KG = ''
        self.KN01 = ''
        self.KN02 = ''
        self.BIMID = ''
        self.Workset = ''

    @property
    def checked(self):
        return self._checked
    @checked.setter
    def checked(self, value):
        self._checked = value
    @property
    def Systemname(self):
        return self._Systemname
    @Systemname.setter
    def Systemname(self, value):
        self._Systemname = value
    @property
    def GK(self):
        return self._GK
    @GK.setter
    def GK(self, value):
        self._GK = value
    @property
    def KG(self):
        return self._KG
    @KG.setter
    def KG(self, value):
        self._KG = value
    @property
    def KN01(self):
        return self._KN01
    @KN01.setter
    def KN01(self, value):
        self._KN01 = value
    @property
    def KN02(self):
        return self._KN02
    @KN02.setter
    def KN02(self, value):
        self._KN02 = value
    @property
    def BIMID(self):
        return self._BIMID
    @BIMID.setter
    def BIMID(self, value):
        self._BIMID = value
    @property
    def Workset(self):
        return self._Workset
    @Workset.setter
    def Workset(self, value):
        self._Workset = value

Liste_Luft = ObservableCollection[Exceldaten]()
Liste_Alle = ObservableCollection[Exceldaten]()
Liste_Rohr = ObservableCollection[Exceldaten]()

def datenlesen(filepath,sheetname,Liste):
    ex = Excel.ApplicationClass()
    book = ex.Workbooks.Open(filepath)
    sheet = book.Worksheets[sheetname]
    rows = sheet.UsedRange.Rows.Count
    for row in range(3,rows+1):
        tempclass = Exceldaten()
        sysname = sheet.Cells[row, 1].Value2
        GK = sheet.Cells[row, 2].Value2
        KG = str(int(sheet.Cells[row, 3].Value2))
        KN01 = str(int(sheet.Cells[row, 4].Value2))
        if len(KN01) == 1:
            KN01 = '0' + KN01
        KN02 = str(int(sheet.Cells[row, 5].Value2))
        if len(KN02) == 1:
            KN02 = '0' + KN02
        workset = sheet.Cells[row, 7].Value2
        bimid = sheet.Cells[row, 6].Value2
        tempclass.Systemname = sysname
        tempclass.GK = GK
        tempclass.KG = KG
        tempclass.KN01 = KN01
        tempclass.KN02 = KN02
        tempclass.BIMID = bimid
        tempclass.Workset = workset
        Liste.Add(tempclass)
        Liste_Alle.Add(tempclass)
    book.Save()
    book.Close()
    Marshal.FinalReleaseComObject(sheet)
    Marshal.FinalReleaseComObject(book)
    ex.Quit()
    Marshal.FinalReleaseComObject(ex)


try:
    Adresse = bimid_config.bimid
    datenausexcel = {}
    try:
        datenlesen(Adresse,'Luft',Liste_Luft)
        datenlesen(Adresse,'Rohr',Liste_Rohr)
    except Exception as e:
        logger.error(e)
except:
    pass


# ExcelBimId GUI
class ExcelBimId(WPFWindow):
    def __init__(self, xaml_file_name,liste_Luft,liste_Rohr,liste_All):
        self.liste_Luft = liste_Luft
        self.liste_All = liste_All
        self.liste_Rohr = liste_Rohr
        WPFWindow.__init__(self, xaml_file_name)
        self.tempcoll = ObservableCollection[Exceldaten]()
        self.altdatagrid = None
        self.read_config()

        try:
            self.dataGrid.ItemsSource = self.liste_All
            self.altdatagrid = self.liste_All
            self.backAll()
            self.click(self.alle)
        except Exception as e:
            logger.error(e)

        self.systemsuche.TextChanged += self.search_txt_changed
        self.Adresse.TextChanged += self.excel_changed
    def click(self,button):
        button.Background = BrushConverter().ConvertFromString("#FF707070")
        button.FontWeight = FontWeights.Bold
        button.FontStyle = FontStyles.Italic
    def back(self,button):
        button.Background  = Brushes.White
        button.FontWeight = FontWeights.Normal
        button.FontStyle = FontStyles.Normal
    def backAll(self):
        self.back(self.luft)
        self.back(self.rohr)
        self.back(self.alle)

    def rohr(self,sender,args):
        self.backAll()
        self.click(self.rohr)
        self.dataGrid.ItemsSource = self.liste_Rohr
        self.altdatagrid = self.liste_Rohr
        self.dataGrid.Items.Refresh()
    def luft(self,sender,args):
        self.backAll()
        self.click(self.luft)
        self.dataGrid.ItemsSource = self.liste_Luft
        self.altdatagrid = self.liste_Luft
        self.dataGrid.Items.Refresh()
    def alle(self,sender,args):
        self.backAll()
        self.click(self.alle)
        self.dataGrid.ItemsSource = self.liste_All
        self.altdatagrid = self.liste_All
        self.dataGrid.Items.Refresh()

    def read_config(self):
        try:
            self.Adresse.Text = str(bimid_config.bimid)
        except:
            self.Adresse.Text = bimid_config.bimid = ""

    def write_config(self):
        bimid_config.bimid = self.Adresse.Text.encode('utf-8')
        script.save_config()

    def search_txt_changed(self, sender, args):
        """Handle text change in search box."""
        self.tempcoll.Clear()
        text_typ = self.systemsuche.Text.upper()
        if text_typ in ['',None]:
            self.dataGrid.ItemsSource = self.altdatagrid

        else:
            if text_typ == None:
                text_typ = ''
            for item in self.altdatagrid:
                if item.Systemname.upper().find(text_typ) != -1:
                    self.tempcoll.Add(item)
            self.dataGrid.ItemsSource = self.tempcoll
        self.dataGrid.Items.Refresh()

    def excel_changed(self, sender, args):
        Liste_Luft.Clear()
        Liste_Rohr.Clear()
        Liste_Alle.Clear()
        try:
            datenlesen(self.Adresse.Text,'Luft',Liste_Luft)
            datenlesen(self.Adresse.Text,'Rohr',Liste_Rohr)
        except:
            pass
        self.liste_Luft = Liste_Luft
        self.dataGrid.ItemsSource = Liste_Luft

    def durchsuchen(self,sender,args):
        dialog = OpenFileDialog()
        dialog.Multiselect = False
        dialog.Title = "BIM-ID Datei suchen"
        dialog.Filter = "Excel Dateien|*.xls;*.xlsx"
        if dialog.ShowDialog() == DialogResult.OK:
            self.Adresse.Text = dialog.FileName
        self.write_config()

    def checkall(self,sender,args):
        for item in self.dataGrid.Items:
            item.checked = True
        self.dataGrid.Items.Refresh()

    def uncheckall(self,sender,args):
        for item in self.dataGrid.Items:
            item.checked = False
        self.dataGrid.Items.Refresh()

    def toggleall(self,sender,args):
        for item in self.dataGrid.Items:
            value = item.checked
            item.checked = not value
        self.dataGrid.Items.Refresh()
    def checkallbb(self,sender,args):
        for item in self.dataGrid.Items:
            item.bb = True
        self.dataGrid.Items.Refresh()

    def uncheckallbb(self,sender,args):
        for item in self.dataGrid.Items:
            item.bb = False
        self.dataGrid.Items.Refresh()

    def toggleallbb(self,sender,args):
        for item in self.dataGrid.Items:
            value = item.bb
            item.bb = not value
        self.dataGrid.Items.Refresh()

    def ok(self,sender,args):
        self.Close()

windowExcelBimId = ExcelBimId("Window.xaml",Liste_Luft,Liste_Rohr,Liste_Alle)
try:
    windowExcelBimId.ShowDialog()
except Exception as e:
    logger.error(e)
    windowExcelBimId.Close()
    script.exit()

muster_bb_bearbeiten = windowExcelBimId.muster.IsChecked

worksset_Excel = [e.Workset for e in Liste_Alle]
worksset_Excel = list(worksset_Excel)
fehlendeworkset = []
if len(worksset_Excel) > 0:
    for item in worksset_Excel:
        if not item in Workset_dict.keys():
           fehlendeworkset.append(item)
fehlendeworkset = set(fehlendeworkset)
fehlendeworkset = list(fehlendeworkset)

# Bearbeitungsbereich erstellen
if any(fehlendeworkset):
    if forms.alert("fehlende Bearbeitungsbereiche erstellen?", ok=False, yes=True, no=True):
        t = DB.Transaction(doc)
        t.Start('Bearbeitungsbereich erstellen')
        for el in fehlendeworkset:
            logger.info(30*'-')
            logger.info(el)
            try:
                item = DB.Workset.Create(doc,el)
                Workset_dict[el] = item.Id.ToString()
                logger.info('Bearbeitungsbereich {} erstellt'.format(el))
            except Exception as e:
                logger.error(e)
        doc.Regenerate()
        t.Commit()
        t.Dispose()

# Luft System
luftsys = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_DuctSystem).WhereElementIsNotElementType()
luftsysids = list(luftsys.ToElementIds())
luftsys.Dispose()

# Rohr System
rohrsys = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_PipingSystem).WhereElementIsNotElementType()
rohrsysids = list(rohrsys.ToElementIds())
rohrsys.Dispose()

systemliste = luftsysids[:]
systemliste.extend(rohrsysids)

class MEP_System:
    def __init__(self,System_Excel):
        self.bimid = System_Excel.checked
        self.GK = System_Excel.GK
        self.KG = System_Excel.KG
        self.KN01 = System_Excel.KN01
        self.KN02 = System_Excel.KN02
        self.BIMID = System_Excel.BIMID
        self.Workset = System_Excel.Workset
        self.bb = System_Excel.bb
        self.systemname = System_Excel.Systemname
        self.liste_system = []
        self.liste_bauteile = []
        self.typ = None
        
    def get_elemente(self):
        for el in self.liste_system:
            systemid = el.Id.ToString()
            try:
                elemente = el.DuctNetwork
            except:
                elemente = el.PipingNetwork

            for elem in elemente:
                cate = elem.Category.Id.ToString()
                # Leitung
                if cate in ['-2008000','-2008020','-2008050','-2008044']:
                    if elem.MEPSystem.Id.ToString() == systemid:
                        if elem.Id.ToString() not in self.liste_bauteile:
                            self.liste_bauteile.append(elem.Id.ToString())
                # Auslass, Sprinkler
                elif cate in ['-2008013','-2008099']:
                    if list(elem.MEPModel.ConnectorManager.Connectors)[0].MEPSystem.Id.ToString() == systemid:
                        if elem.Id.ToString() not in self.liste_bauteile:
                            self.liste_bauteile.append(elem.Id.ToString())
                # Bauteile
                elif cate in ['-2008010','-2008016','-2001140','-2008049','-2008055','-2001160']:
                    conns = elem.MEPModel.ConnectorManager.Connectors
                    In = {}
                    Out = {}
                    Unverbunden = {}
                    for conn in conns:
                        if conn.IsConnected:
                            if conn.Direction.ToString() == 'In':
                                In[conn.Id] = conn
                            else:
                                Out[conn.Id] = conn
                        else:
                            Unverbunden[conn.Id] = conn
                    sorted(In)
                    sorted(Out)
                    sorted(Unverbunden)
                    conns = In.values()[:]
                    connouts = Out.values()[:]
                    connunvers = Unverbunden.values()[:]
                    conns.extend(connouts)
                    conns.extend(connunvers)
                    
                    try:
                        for conn in conns:
                            if not conn.MEPSystem:
                                continue
                            else:
                                if conn.MEPSystem.Id.ToString() == systemid:
                                    if elem.Id.ToString() not in self.liste_bauteile:
                                        self.liste_bauteile.append(elem.Id.ToString())
                    except:
                        pass

    def get_systemtyp(self):
        for el in self.liste_system:
            typ = el.get_Parameter(DB.BuiltInParameter.ELEM_TYPE_PARAM).AsElementId()
            self.typ = doc.GetElement(typ)
            break
    
    def wert_schreiben(self, elem, param_name, wert):
        if not wert is None:
            if elem.LookupParameter(param_name):
                elem.LookupParameter(param_name).Set(wert)

    def wert_schreiben_system(self):
        self.wert_schreiben(self.typ,'IGF_X_Gewerkkürzel',str(self.GK))
        self.wert_schreiben(self.typ,'IGF_X_Kostengruppe',int(self.KG))
        self.wert_schreiben(self.typ,'IGF_X_Kennnummer_1',int(self.KN01))
        self.wert_schreiben(self.typ,'IGF_X_Kennnummer_2',int(self.KN02))
        self.wert_schreiben(self.typ,'IGF_X_BIM-ID',str(self.BIMID))

class Bauteil:
    def __init__(self,elemid,system):
        self.elemid = DB.ElementId(int(elemid))
        self.elem = doc.GetElement(self.elemid)
        self.system = system
        self.bb = self.elem.LookupParameter('Bearbeitungsbereich').AsValueString()

    def wert_schreiben(self, elem, param_name, wert):
            if not wert is None:
                if elem.LookupParameter(param_name):
                    elem.LookupParameter(param_name).Set(wert)
    
    def werte_schreiben_bimid(self):
        self.wert_schreiben(self.elem,'IGF_X_Gewerkkürzel_Exemplar',str(self.system.GK))
        self.wert_schreiben(self.elem,'IGF_X_KG_Exemplar',int(self.system.KG))
        self.wert_schreiben(self.elem,'IGF_X_KN01_Exemplar',int(self.system.KN01))
        self.wert_schreiben(self.elem,'IGF_X_KN02_Exemplar',int(self.system.KN02))
        self.wert_schreiben(self.elem,'IGF_X_BIM-ID_Exemplar',str(self.system.BIMID))

    def werte_schreiben_BB(self):
        try:
            self.wert_schreiben(self.elem,'Bearbeitungsbereich',int(Workset_dict[self.system.Workset]))
        except:
            pass


dict_systeme = {}
for system in Liste_Alle:
    if system.checked or system.bb:
        dict_systeme[system.Systemname] = MEP_System(system)

if len(dict_systeme.keys()) == 0:
    script.exit()

for sysid in systemliste:
    elem = doc.GetElement(sysid)
    systyp = elem.get_Parameter(DB.BuiltInParameter.ELEM_TYPE_PARAM).AsValueString()
    if systyp in dict_systeme.keys():
        system = dict_systeme[systyp]
        system.liste_system.append(elem)

liste = dict_systeme.keys()[:]
liste_neu = []
with forms.ProgressBar(title="{value}/{max_value} Systeme --- Datenermittlung",cancellable=True, step=1) as pb:
    for n,typ in enumerate(liste):
        if pb.cancelled:
            frage_schicht = abfrage(title= __title__,
                    info = 'Vorgang abrechen oder ermittlete Daten behalten?' , 
                    ja = True,ja_text= 'abbrechen',nein_text='weiter').ShowDialog()
            if frage_schicht.antwort == 'abbrechen':
                script.exit()
            else:
                break
        mepsystem = dict_systeme[typ]
        mepsystem.get_elemente()
        logger.info('{} Elemente in System {}'.format(len(mepsystem.liste_bauteile),mepsystem.systemname))
        liste_neu.append(mepsystem)
        pb.update_progress(n+1,len(liste))

bearbeitet = []
nichtbearbeitet = liste_neu[:]
t = DB.Transaction(doc,'BIM-ID')
t.Start()
if forms.alert('Daten schreiben?',yes=True,no=True,ok=False):
    with forms.ProgressBar(title='Systeme --- Datene schreiben',cancellable=True, step=10) as pb2:
        for n,mepsystem in enumerate(liste_neu):
            systeme_title = str(mepsystem.systemname) + ' --- ' + str(n+1) + '/' + str(len(liste_neu)) + 'Systeme'
            pb2.title = '{value}/{max_value} Elemente in System ' +  systeme_title
            pb2.step = int((len(mepsystem.liste_bauteile)) /1000) + 10
            mepsystem.get_systemtyp()
            try:
                mepsystem.wert_schreiben_system()
            except:
                pass
            for n1,bauteilid in enumerate(mepsystem.liste_bauteile):
                if pb2.cancelled:
                    if forms.alert('bisherige Änderung behalten?',yes=True,no=True,ok=False):
                        t.Commit()
                        logger.info('Folgenede Systeme sind bereits bearbeitet.')
                        for el in bearbeitet:
                            logger.info(el.systemname)
                        logger.info('---------------------------------------')
                        logger.info('Folgenede Systeme sind nicht bearbeitet.')
                        for el in nichtbearbeitet:
                            logger.info(el.systemname)
                    else:
                        t.RollBack()
                    
                    script.exit()
                bauteil = Bauteil(bauteilid,mepsystem)
                if bauteil.system.bimid:
                    bauteil.werte_schreiben_bimid()
                if bauteil.system.bb:
                    if bauteil.bb in muster_bb and not muster_bb_bearbeiten:
                        pass
                    else:
                        bauteil.werte_schreiben_BB()

                pb2.update_progress(n1+1, len(mepsystem.liste_bauteile))
            bearbeitet.append(mepsystem)
            nichtbearbeitet.remove(mepsystem)
            
t.Commit()
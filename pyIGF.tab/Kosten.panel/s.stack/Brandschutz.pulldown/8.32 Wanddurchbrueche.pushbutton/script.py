# coding: utf8
import sys
sys.path.append(r'R:\pyRevit\xx_Skripte\libs\IGF_libs')
from IGF_log import getlog
from pyrevit.forms import WPFWindow
from pyrevit import script, forms
from rpw import revit,DB
from System import Guid
from System.Collections.ObjectModel import ObservableCollection

__title__ = "8.32 ergänzt Brandschutzklasse für WD"
__doc__ = """

ergänzt Brandschutzklasse für Wanddurchbrüche.

F30 wird in ein WD eingetragen wenn es in ein F30 Wand platziert.

Parameter: IGF_HLS_Brandschutz (276f654e-b5ea-4ead-ac28-fdc326b12e12)

Kategorie: HLS-Bauteile

[2021.11.29]
Version: 2.0
"""
__authors__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()
uidoc = revit.uidoc
doc = revit.doc
projectinfo = doc.ProjectInformation.Name + ' - '+ doc.ProjectInformation.Number

config = script.get_config('WD-' + projectinfo)

global abbrechen
abbrechen = False
try:
    getlog(__title__)
except:
    pass

# Revitlinkmodel
revitLinks_collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_RvtLinks).WhereElementIsNotElementType()
revitLinks = revitLinks_collector.ToElementIds()

revitLinksDict = {}
for el in revitLinks_collector:
    revitLinksDict[el.Name] = el
revitLinks_collector.Dispose()

rvtLink = forms.SelectFromList.show(revitLinksDict.keys(), button_name='Select RevitLink')
rvtdoc = None
if not rvtLink:
    logger.error("Keine Revitverknüpfung gewählt")
    script.exit()

rvtdoc = revitLinksDict[rvtLink].GetLinkDocument()
if not rvtdoc:
    logger.error("Keine Revitverknüpfung in aktueller Projekt gefunden")
    script.exit()

walls = DB.FilteredElementCollector(rvtdoc).OfCategory(DB.BuiltInCategory.OST_Walls).WhereElementIsNotElementType()
wallsName = []
for el in walls:
    if not el.Name in wallsName:
        wallsName.append(el.Name)

BrandWalls = forms.SelectFromList.show(wallsName,multiselect=True, button_name='Select Walls')

BrandWallEles = []
for el in walls:
    if el.Name in BrandWalls:
        BrandWallEles.append(el)

# Wanddurchbrüche
HLS_Bauteile = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MechanicalEquipment).WhereElementIsNotElementType()
HLS_Ids = HLS_Bauteile.ToElementIds()
HLS_Bauteile.Dispose()

HLS_Dict = {}
for el in HLS_Ids:
    elem = doc.GetElement(el)
    Family = elem.Symbol.FamilyName
    if not Family in HLS_Dict.keys():
        HLS_Dict[Family] = [elem]
    else:
        HLS_Dict[Family].append(elem)

class HLSBauteile(object):
    def __init__(self,Name):
        self.checked = False
        self.Name = Name

Liste_WD = ObservableCollection[HLSBauteile]()
HLS_Bauteile_Liste = HLS_Dict.keys()[:]
HLS_Bauteile_Liste.sort()

for name in HLS_Bauteile_Liste:
    hlsbauteil = HLSBauteile(name)
    Liste_WD.Add(hlsbauteil)

# GUI Systemauswahl
class Wanddurchbrueche(WPFWindow):
    def __init__(self, xaml_file_name,Liste_WandD):
        self.Liste_WandD = Liste_WandD
        WPFWindow.__init__(self, xaml_file_name)
        self.tempcoll = ObservableCollection[HLSBauteile]()
        self.altdatagrid = Liste_WandD
        self.dataGrid.ItemsSource = Liste_WandD
        self.read_config()
        
        self.suche.TextChanged += self.search_txt_changed
    
    def read_config(self):
        try:
            wddict = config.Wd_dict
            for item in self.Liste_WandD:
                if item.Name in wddict.keys():
                    item.checked = wddict[item.Name]
            self.dataGrid.Items.Refresh()
        except:
            pass
        try:
            if config.Hide:
                tempcoll = ObservableCollection[HLSBauteile]()
                for item in self.dataGrid.Items:
                    if item.checked:
                        tempcoll.Add(item)
                self.dataGrid.ItemsSource = tempcoll
                self.dataGrid.Items.Refresh()
                   
        except Exception as e:
            pass
        
        
    def write_config(self):
        dict_ = {}
        for item in self.Liste_WandD:
            dict_[item.Name] = item.checked

        config.Wd_dict = dict_
        script.save_config()

    def search_txt_changed(self, sender, args):
        """Handle text change in search box."""
        self.tempcoll.Clear()
        text_typ = self.suche.Text.upper()
        if text_typ in ['',None]:
            self.dataGrid.ItemsSource = self.altdatagrid

        else:
            if text_typ == None:
                text_typ = ''
            for item in self.altdatagrid:
                if item.Name.upper().find(text_typ) != -1:
                    self.tempcoll.Add(item)
            self.dataGrid.ItemsSource = self.tempcoll
        self.dataGrid.Items.Refresh()

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

    def select(self,sender,args):
        global abbrechen
        abbrechen = True
        self.write_config()
        self.Close()

    def hide(self,sender,args):
        tempcoll = ObservableCollection[HLSBauteile]()
        for item in self.dataGrid.Items:
            if item.checked:
                tempcoll.Add(item)
        self.dataGrid.ItemsSource = tempcoll
        config.Hide = True
        script.save_config()

    def show(self,sender,args):
        self.dataGrid.ItemsSource = self.altdatagrid
        config.Hide = False
        script.save_config()
    
Systemwindows = Wanddurchbrueche("window.xaml",Liste_WD)
try:
    Systemwindows.ShowDialog()
except Exception as e:
    logger.error(e)
    Systemwindows.Close()
    script.exit()

Familie = []
for ele in Liste_WD:
    if ele.checked:
        Familie.extend(HLS_Dict[ele.Name])

opt = DB.Options()
opt.ComputeReferences = True
opt.IncludeNonVisibleObjects = True


def getSolids(elem):
    lstSolid = []
    ge = elem.get_Geometry(opt)
    if ge != None:
        lstSolid.extend(GetSolid(ge))
    return lstSolid
def GetSolid(GeoEle):
    lstSolid = []
    for el in GeoEle:
        if el.GetType().ToString() == 'Autodesk.Revit.DB.Solid':
            if el.SurfaceArea > 0 and el.Volume > 0 and el.Faces.Size > 1 and el.Edges.Size > 1:
                lstSolid.append(el)
        elif el.GetType().ToString() == 'Autodesk.Revit.DB.GeometryInstance':
            ge = el.GetInstanceGeometry()
            lstSolid.extend(GetSolid(ge))
    return lstSolid

def TransformSolid(elem):
    m_lstModels = []
    listSolids = getSolids(elem)
    for solid in listSolids:
        tempSolid = solid
        tempSolid = DB.SolidUtils.CreateTransformed(solid,revitLinksDict[rvtLink].GetTransform())
        m_lstModels.append(tempSolid)
    return m_lstModels

def HLSkurve(elem):
    BB = elem.get_BoundingBox(None)

    Cen_X = (BB.Max.X + BB.Min.X) / 2
    Cen_Y = (BB.Max.Y + BB.Min.Y) / 2
    Cen_Z = (BB.Max.Z + BB.Min.Z) / 2

    Line_x = DB.Line.CreateBound(DB.XYZ(BB.Max.X,Cen_Y,Cen_Z), DB.XYZ(BB.Min.X,Cen_Y,Cen_Z))
    Line_y = DB.Line.CreateBound(DB.XYZ(Cen_X,BB.Max.Y,Cen_Z), DB.XYZ(Cen_X,BB.Min.Y,Cen_Z))

    return Line_x,Line_y

def EbenenUmbenennen(ebene):
    out = ''
    if not ebene:
        out = ''
    if ebene.find('EG') != -1:
        out = 'EG'
    elif ebene.find('SAN') != -1:
        out = 'EG'
    elif ebene.find('1.OG') != -1:
        out = '1.OG'
    elif ebene.find('2.OG') != -1:
        out = '2.OG'
    elif ebene.find('3.OG') != -1:
        out = '3.OG'
    elif ebene.find('1.UG') != -1:
        out = '1.UG'
    elif ebene.find('2.UG') != -1:
        out = '2.UG'
    elif ebene.find('3.UG') != -1:
        out = '3.UG'
    else:
        out = '4.OG'
    return out

def NameUmbenennen(Name):
    out = ''
    if Name.find('F30') != -1:
        out = 'F30'
    elif Name.find('F90') != -1:
        out = 'F90'
    elif Name.find('F120') != -1:
        out = 'F120'
    else:
        out = 'F0'
    return out

# RvtLinkElem
RvtLinkElemSolids = {}
# ProElemCurve = []
step = int(len(BrandWallEles)/200)
with forms.ProgressBar(title='{value}/{max_value} Wände in RVT-Link Model',cancellable=True, step=step) as pb:
    for n_1, ele in enumerate(BrandWallEles):
        if pb.cancelled:
            script.exit()
        pb.update_progress(n_1, len(BrandWallEles))
        models = TransformSolid(ele)
        ebenename = rvtdoc.GetElement(ele.LevelId).Name
        WallName = ele.Name
        Klasse = NameUmbenennen(WallName)
        Ebenen = EbenenUmbenennen(ebenename)
        if not Ebenen in RvtLinkElemSolids.keys():
            RvtLinkElemSolids[Ebenen] = {}
        if not Klasse in RvtLinkElemSolids[Ebenen].keys():
            RvtLinkElemSolids[Ebenen][Klasse] = []
        if not models in RvtLinkElemSolids[Ebenen][Klasse]:
            RvtLinkElemSolids[Ebenen][Klasse].append(models)

rvtdoc.Dispose()

class Wanddurchbruch:
    def __init__(self,elem):
        self.elem = elem
        self.Klass = ''
    @property
    def Klass(self):
        return self._Klass
    @Klass.setter
    def Klass(self,value):
        self._Klass = value
    def wert_schreiben(self):
        param = self.elem.LookupParameter('IGF_HLS_Brandschutz')
        if param:
            param.Set(self.Klass)

opt2 = DB.SolidCurveIntersectionOptions()
opt2.ResultType = DB.SolidCurveIntersectionMode.CurveSegmentsInside

wanddurchbruch_liste = []
with forms.ProgressBar(title='{value}/{max_value} Wanddurchbrüche',cancellable=True, step=10) as pb2:
    for n,elem in enumerate(Familie):
        if pb2.cancelled:
            script.exit()
        pb2.update_progress(n + 1, len(Familie))

        ebene = doc.GetElement(elem.LevelId).Name
        neu_ebene = EbenenUmbenennen(ebene)
        curve_x,curve_y = HLSkurve(elem)
        wanddurchbruch = Wanddurchbruch(elem)

        if neu_ebene in RvtLinkElemSolids.keys():
            models_dict = RvtLinkElemSolids[neu_ebene]
            for klasse in ['F120','F90','F30','F0']:
                if klasse in models_dict.keys():
                    models = models_dict[klasse]
                    for items in models:
                        for solid in items:
                            result1 = solid.IntersectWithCurve(curve_x,opt2)
                            result2 = solid.IntersectWithCurve(curve_y,opt2)
                            if result1.SegmentCount > 0 or result2.SegmentCount > 0:
                                wanddurchbruch.Klass = klasse
                                wanddurchbruch_liste.append(wanddurchbruch)
                                break
                        if wanddurchbruch.Klass:
                            break
                if wanddurchbruch.Klass:
                    break
        curve_x.Dispose()
        curve_y.Dispose()

# Daten schreiben
if forms.alert("Daten schreiben?", ok=False, yes=True, no=True):
    t = DB.Transaction(doc,'Brandschutzklasse der Wanddurchbrüche')
    t.Start()
    with forms.ProgressBar(title='{value}/{max_value} Wanddurchbrüche',cancellable=True, step=int(len(wanddurchbruch_liste)/1000)+10) as pb3:
        for n,wanddurchbruch in enumerate(wanddurchbruch_liste):
            if pb3.cancelled:
                t.RollBack()
                script.exit()
            pb3.update_progress(n, len(wanddurchbruch_liste))
            wanddurchbruch.wert_schreiben()
            
    t.Commit()
    t.Dospose()
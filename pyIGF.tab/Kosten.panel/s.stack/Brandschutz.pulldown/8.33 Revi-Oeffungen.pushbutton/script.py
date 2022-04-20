# coding: utf8
import sys
sys.path.append(r'R:\pyRevit\xx_Skripte\libs\IGF_libs')
from IGF_log import getlog
from rpw import revit, DB
from pyrevit import script, forms

__title__ = "8.33 ermittelt Deckentyp für Revi-Öffnungen"
__doc__ = """

Familie: Revi-Öffnungen(Deckenspiegel)

Parameter: IGF_BIG_Basisbauteil Revi-öffnungen

Kategorie: Allgemeines Modell

[2021.11.29]
Version: 1.1
"""
__author__ = "Menghui Zhang"

try:
    getlog(__title__)
except:
    pass

logger = script.get_logger()
output = script.get_output()

uidoc = revit.uidoc
doc = revit.doc
projectinfo = doc.ProjectInformation.Name + ' - '+ doc.ProjectInformation.Number
config = script.get_config('Revi-Öffnungen-' + projectinfo)

# Revitlinkmodel
revitLinks_collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_RvtLinks).WhereElementIsNotElementType()
revitLinks = revitLinks_collector.ToElementIds()

revitLinksDict = {}
for el in revitLinks_collector:
    revitLinksDict[el.Name] = el

rvtLink = forms.SelectFromList.show(revitLinksDict.keys(), button_name='Select RevitLink')
rvtdoc = None
if not rvtLink:
    logger.error("Keine Revitverknüpfung gewählt")
    script.exit()
rvtdoc = revitLinksDict[rvtLink].GetLinkDocument()
if not rvtdoc:
    logger.error("Keine Revitverknüpfung in aktueller Projekt gefunden")
    script.exit()


Cellings = DB.FilteredElementCollector(rvtdoc).OfCategory(DB.BuiltInCategory.OST_Ceilings).WhereElementIsNotElementType()
CellingsName = []

for el in Cellings:
    if not el.Name in CellingsName:
        CellingsName.append(el.Name)
Cellings_select = forms.SelectFromList.show(CellingsName,
multiselect=True, button_name='Select Decken')

cellingsListe = []
for el in Cellings:
    if el.Name in Cellings_select:
        cellingsListe.append(el)

# Allgemeines Modell

Bauteile = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_GenericModel).WhereElementIsNotElementType()
Bauteile_Ids = Bauteile.ToElementIds()
Bauteile.Dispose()

# revi-Öffnungen
Liste_Bauteile = []
for el in Bauteile_Ids:
    elem = doc.GetElement(el)
    Family = elem.Symbol.FamilyName
    if Family == 'Revi-Öffnungen':
        Liste_Bauteile.append(elem)


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

def HLSkurve(elem,erweite):
    BB = elem.get_BoundingBox(None)
    Cen_X = (BB.Max.X + BB.Min.X) / 2
    Cen_Y = (BB.Max.Y + BB.Min.Y) / 2
    hlscurve = DB.Line.CreateBound(DB.XYZ(Cen_X,Cen_Y,BB.Max.Z + erweite), DB.XYZ(Cen_X,Cen_Y,BB.Min.Z - erweite))
    return hlscurve
    

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


# RvtLinkElem
RvtLinkElemSolids = {}
# ProElemCurve = []
step = int(len(cellingsListe)/200)
with forms.ProgressBar(title='{value}/{max_value} Decken in RVT-Link Model',cancellable=True, step=step) as pb:
    for n_1, ele in enumerate(cellingsListe):
        if pb.cancelled:
            script.exit()
        pb.update_progress(n_1, len(cellingsListe))
        models = TransformSolid(ele)
        ebenename = rvtdoc.GetElement(ele.LevelId).Name
        deckenName = ele.Name
        Ebenen = EbenenUmbenennen(ebenename)
        if not Ebenen in RvtLinkElemSolids.keys():
            RvtLinkElemSolids[Ebenen] = {}
        if not deckenName in RvtLinkElemSolids[Ebenen].keys():
            RvtLinkElemSolids[Ebenen][deckenName] = []
        if not models in RvtLinkElemSolids[Ebenen][deckenName]:
            RvtLinkElemSolids[Ebenen][deckenName].append(models)

rvtdoc.Dispose()

class Oeffnung:
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
        param = self.elem.LookupParameter('IGF_BIG_Basisbauteil Revi-öffnungen')
        if param:
            param.Set(self.Klass)

opt2 = DB.SolidCurveIntersectionOptions()
opt2.ResultType = DB.SolidCurveIntersectionMode.CurveSegmentsInside

oeffnung_liste = []

with forms.ProgressBar(title='{value}/{max_value} Bauteile',cancellable=True, step=10) as pb2:
    for n_1,elem in enumerate(Liste_Bauteile):
        if pb2.cancelled:
            script.exit()
        pb2.update_progress(n_1+1, len(Liste_Bauteile))
        ebene = elem.LookupParameter('Bauteillistenebene').AsValueString()
        neu_ebene = EbenenUmbenennen(ebene)
        oeffnung = Oeffnung(elem)
        if neu_ebene in RvtLinkElemSolids.keys():
            models_dict = RvtLinkElemSolids[neu_ebene]
            for lin in [0.2,0.5,1,1.5]:
                curve = HLSkurve(elem,lin)
                for klasse in models_dict.keys():
                    models = models_dict[klasse]
                    for item in models:
                        for solid in item:
                            result = solid.IntersectWithCurve(curve,opt2)
                            if result.SegmentCount > 0:
                                oeffnung.Klass = klasse
                                oeffnung_liste.append(oeffnung)
                                break
                            result.Dispose()
                        if oeffnung.Klass:
                            break
                    if oeffnung.Klass:
                        break
                if oeffnung.Klass:
                    break
                curve.Dispsoe()

# Daten schreiben
if forms.alert("Daten schreiben?", ok=False, yes=True, no=True):
    t = DB.Transaction(doc,'Revi-Öffnungen')
    t.Start()
    with forms.ProgressBar(title='{value}/{max_value} Revi-Öffnungen',cancellable=True, step=int(len(oeffnung_liste)/1000)+10) as pb3:
        for n,oeffnung in enumerate(oeffnung_liste):
            if pb3.cancelled:
                t.RollBack()
                script.exit()
            pb3.update_progress(n, len(oeffnung_liste))
            oeffnung.wert_schreiben()
    t.Commit()
    t.Dispose()
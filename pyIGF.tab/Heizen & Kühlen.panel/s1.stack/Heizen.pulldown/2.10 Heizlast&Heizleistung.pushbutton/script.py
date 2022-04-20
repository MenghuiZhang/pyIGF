# coding: utf8
import sys

sys.path.append(r'R:\pyRevit\xx_Skripte\libs\IGF_libs')
from rpw import revit, DB
from pyrevit import script, forms
from IGF_log import getlog
from IGF_libKopie import get_value

__title__ = "2.10 Heizlast & Heizleistung MEP Räume"
__doc__ = """berechnet Heizlast und Heizleistung von MEP Räume

input Parameter:
--------------------
LIN_BA_CALCULATED_HEATING_LOAD: Heizlast Gebäude

IGF_H_DeS_Leistung: Heizleistung DeS

IGF_H_HK_Leistung: Heizleistung Heizkörper

IGF_H_ULH_Leistung: Heizleistung Umlufthitzer

IGF_H_HZ_Leistung: sonstige Heizleistung
--------------------

output Parameter:
--------------------
IGF_H_HeizleistungRaum: Summe von HK & DeS- & ULH- & HZ-Heizleistung

IGF_H_HeizlastGesamt: Heizlast Gebäude

IGF_H_HeizleistungBilanz: gesamte Heizleistung - gesamte Heizlast

IGF_H_HeizBilanzProzent: gesamte Heizleistung / gesamte Heizlast
--------------------

[Version: 1.1]
[2021.11.18]
"""

__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()

try:
    getlog(__title__)
except:
    pass

uidoc = revit.uidoc
doc = revit.doc
name = doc.ProjectInformation.Name
number = doc.ProjectInformation.Number
active_view = uidoc.ActiveView

# MEP Räume aus aktueller Ansicht
spaces_collector = DB.FilteredElementCollector(doc).OfCategory(
    DB.BuiltInCategory.OST_MEPSpaces).WhereElementIsNotElementType()
spaces = spaces_collector.ToElementIds()
spaces_collector.Dispose()

if not spaces:
    logger.error("Keine MEP Räume in aktueller Projekt gefunden")
    script.exit()


class MEPRaum:
    def __init__(self, element_id):
        self.element_id = element_id
        self.element = doc.GetElement(self.element_id)
        self.name = self.get_parameter('Name')
        self.nummer = self.get_parameter('Nummer')
        self.Heizlast_Gebaeude = self.get_parameter('LIN_BA_CALCULATED_HEATING_LOAD')
        self.HL_DeS = self.get_parameter('IGF_H_DeS_Leistung')
        self.HL_HK = self.get_parameter('IGF_H_HK_Leistung')
        self.HL_HZ = self.get_parameter('IGF_H_HZ_Leistung')
        self.HL_ULH = self.get_parameter('IGF_H_ULH_Leistung')

        self.Heizlast_gesamt = self.HeizlastGesamt_Berechnen()
        self.HL_gesamt = self.HL_gesamt_berechnen()
        self.HL_Bilanz = self.HL_Bilanz_Berechnen()
        self.Prozent = self.Prozent_Berechnen()

    def Werte_Pruefen(self, wert):
        if not wert:
            wert = 0
        return wert

    def HL_Bilanz_Berechnen(self):

        HL_Bilanz = 0
        HL_Bilanz = float(self.HL_gesamt) - self.Heizlast_gesamt

        return HL_Bilanz

    def get_parameter(self, param_name):
        param = self.element.LookupParameter(param_name)
        if not param:
            logger.error(
                "Parameter {} konnte nicht gefunden werden".format(param_name))
            return ''
        return get_value(param)

    def HL_gesamt_berechnen(self):
        Heizleistung = 0
        self.HL_DeS = self.Werte_Pruefen(self.HL_DeS)
        self.HL_ULH = self.Werte_Pruefen(self.HL_ULH)
        self.HL_HZ = self.Werte_Pruefen(self.HL_HZ)
        self.HL_HK = self.Werte_Pruefen(self.HL_HK)

        Heizleistung = self.HL_DeS + self.HL_ULH + self.HL_HZ + self.HL_HK

        return round(Heizleistung, 2)

    def HeizlastGesamt_Berechnen(self):
        HeizlastGesamt = self.Heizlast_Gebaeude

        if not HeizlastGesamt:
            HeizlastGesamt = 0.0
        return HeizlastGesamt

    def KuelleistungULK_Berechnen(self):

        KuelleistungULK = 0
        KuelleistungULK = self.Heizlast_gesamt - self.P_zu

        return KuelleistungULK

    def Prozent_Berechnen(self):
        prozent = 0
        if not self.Heizlast_gesamt:
            return 0

        prozent = float(self.HL_gesamt) / self.Heizlast_gesamt
        return round(prozent, 3)

    def werte_schreiben(self):
        """Schreibt die berechneten Werte zurück in das Modell."""

        def wert_schreiben(param_name, wert):
            if not wert is None:
                self.element.LookupParameter(
                    param_name).SetValueString(str(wert))

        wert_schreiben("IGF_H_HeizlastGesamt", self.Heizlast_gesamt)
        wert_schreiben("IGF_H_HeizleistungBilanz", self.HL_Bilanz)
        wert_schreiben("IGF_H_HeizleistungRaum", self.HL_gesamt)
        self.element.LookupParameter('IGF_H_HeizBilanzProzent').Set(self.Prozent)


mepraum_liste = []
with forms.ProgressBar(title='{value}/{max_value} MEP-Räume', cancellable=True, step=10) as pb:
    for n, space_id in enumerate(spaces):
        if pb.cancelled:
            script.exit()

        pb.update_progress(n + 1, len(spaces))
        mepraum = MEPRaum(space_id)

        mepraum_liste.append(mepraum)

# Werte zuückschreiben + Abfrage
if forms.alert('Berechnete Werte in Modell schreiben?', ok=False, yes=True, no=True):
    with forms.ProgressBar(title='{value}/{max_value} Werte schreiben', cancellable=True, step=10) as pb2:
        t = DB.Transaction(doc)
        t.Start('Werte schreiben')

        for n, mepraum in enumerate(mepraum_liste):
            if pb2.cancelled:
                t.RollBack()
                script.exit()
            pb2.update_progress(n + 1, len(mepraum_liste))
            mepraum.werte_schreiben()
        t.Commit()

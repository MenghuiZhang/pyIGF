# coding: utf8

from pyrevit import revit, UI, DB
from pyrevit import script, forms
import rpw
import time
from Autodesk.Revit.DB import Transaction

start = time.time()


__title__ = "9.02 BWZ_Planschlüssel"
__doc__ = """BLB_Planschlüssel,Plannummer und Planname eines Planes werden generiert anhand die golgenden Parameter.
'IGF_Ersteller','IGF_Leistungsphase','BLB_Bauteil_Trakt','BLB_Planart',
'BLB_Geschoss_Anlage','BLB_Planinhalt','IGF_Freitext','Aktuelle Änderung',
'IGF_Planstatus','BLB_Planbezeichnung 1','BLB_Planbezeichnung 2'"""
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()

uidoc = rpw.revit.uidoc
doc = rpw.revit.doc


from pyIGF_logInfo import getlog
getlog(__title__)


selection = revit.get_selection()
sel_sheets = forms.select_sheets(title='Select Sheets')

if sel_sheets:
    selection.set_to(sel_sheets)


if not sel_sheets:
    logger.error('Keine Pläne ausgewählt!')
    script.exit()

WE_Nr = doc.ProjectInformation.LookupParameter('BLB_WE-Nr.').AsString()


Parameters = ['IGF_Ersteller','IGF_Leistungsphase','BLB_Bauteil_Trakt','BLB_Planart',
              'BLB_Geschoss_Anlage','BLB_Planinhalt','IGF_Freitext','Aktuelle Änderung',
              'IGF_Planstatus','BLB_Planbezeichnung 1','BLB_Planbezeichnung 2']

def Neunummer(indexnummer):
    index = '--'
    if indexnummer:
        if len(indexnummer) == 1:
            index = '0' + str(indexnummer)
        elif len(indexnummer) == 2:
            index = str(indexnummer)
    return index

def get_value(param):
    value = revit.query.get_param_value(param)
    try:
        unit = param.DisplayUnitType
        value = DB.UnitUtils.ConvertFromInternalUnits(
            value,
            unit)
    except Exception as e:
        pass
    return value
def Pruefung(inWerte,Ziffer):
    out = Ziffer * '-'
    if Ziffer == 9:
        out = Neunummer(inWerte)
        return out
    else:
        if Ziffer == 8:
            out = 'xxxxx'
        if inWerte:
            return inWerte
        else:
            return out

with forms.ProgressBar(title='{value}/{max_value} Planschlüssel',
                       cancellable=True, step=1) as pb:
    n = 0
    t = Transaction(doc, "Planschlüssel")
    t.Start()
    for Item in sel_sheets:
        if pb.cancelled:
            script.exit()
        pb.update_progress(n + 1, len(sel_sheets))
        ParemeterListe = [WE_Nr]
        ParemeterListeNeu = []
        ZifferListe = [0,3,1,4,2,2,3,8,9,1,3,3]
        for param in Parameters:
            wertParam = get_value(Item.LookupParameter(param))
            ParemeterListe.append(wertParam)
        for i in range(len(ParemeterListe)):
            ParemeterListeNeu.append(Pruefung(ParemeterListe[i],ZifferListe[i]))
        Nummer = ParemeterListeNeu[0] + '_' + ParemeterListeNeu[1] + '_' +\
                 ParemeterListeNeu[2] + '_' + ParemeterListeNeu[3] + '_' +\
                 ParemeterListeNeu[4] + '_' + ParemeterListeNeu[5] + '_' +\
                 ParemeterListeNeu[6] + '_' + ParemeterListeNeu[7] + '_' +\
                 ParemeterListeNeu[8] + '_' + ParemeterListeNeu[9]
        Name = ParemeterListeNeu[10] + '_' + ParemeterListeNeu[11] + '_' + ParemeterListeNeu[5]
        Item.LookupParameter("BLB_Planschlüssel").Set(Nummer)
        Item.LookupParameter("Plannummer").Set(Nummer)
        Item.LookupParameter("Planname").Set(Name)
    t.Commit()
total = time.time() - start
logger.info("total time: {} {}".format(total, 100 * "_"))

# coding: utf8

from pyrevit import revit, UI, DB
from pyrevit import script, forms
import rpw
import time
from Autodesk.Revit.DB import Transaction

start = time.time()


__title__ = "9.11 MFC_Planindex"
__doc__ = """MFC_Planindex"""
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

# Planköpfe aus aktueller Projekt
TitleBlocks_collector = DB.FilteredElementCollector(doc) \
    .OfCategory(DB.BuiltInCategory.OST_TitleBlocks) \
    .WhereElementIsNotElementType()
TitleBlocks = TitleBlocks_collector.ToElementIds()


def get_value(param):
    """Konvertiert Einheiten von internen Revit Einheiten in Projekteinheiten"""

    value = revit.query.get_param_value(param)

    try:
        unit = param.DisplayUnitType

        value = DB.UnitUtils.ConvertFromInternalUnits(
            value,
            unit)

    except Exception as e:
        pass

    return value


Sheets_collector = []

for sheet in sel_sheets:
    Sheetnumber = sheet.get_Parameter(DB.BuiltInParameter.SHEET_NUMBER).AsString()
    for titleblock in TitleBlocks_collector:
        titlenumber = titleblock.get_Parameter(DB.BuiltInParameter.SHEET_NUMBER).AsString()
        if titlenumber == Sheetnumber:
            Sheets_collector.append([sheet,titleblock])


with forms.ProgressBar(title='{value}/{max_value} Planindex',
                       cancellable=True, step=1) as pb:
    n = 0

    t = Transaction(doc, "Planindex")
    t.Start()
    def Neunummer(indexnummer):
        index = '00'
        if indexnummer:
            if len(indexnummer) == 1:
                index = '0' + str(indexnummer)
            elif len(indexnummer) == 2:
                index = str(indexnummer)
        return index

    def Vorabzug(input_para):
        out = 'P'
        if input_para == True:
            out = 'V'
        return out

    list = []
    for Item in Sheets_collector:
        if pb.cancelled:
            script.exit()
        n += 1
        pb.update_progress(n, len(Sheets_collector))
        Nummer = Item[0].get_Parameter(DB.BuiltInParameter.SHEET_NUMBER).AsString()
        tempel = get_value(Item[1].Symbol.LookupParameter('Vorabzugsstempel'))
        if not tempel:
            tempel = get_value(Item[1].LookupParameter('Vorabzugsstempel'))
        V_or_P = Vorabzug(tempel)

        Index = Item[0].get_Parameter(DB.BuiltInParameter.SHEET_CURRENT_REVISION).AsString()
        Neu_Index = Neunummer(Index)

        if len(Nummer) > 10:
            KN01 = Nummer[:len(Nummer) - 3]
            Neu_Nummer = KN01 + V_or_P + Neu_Index
            if Neu_Nummer in list:
                logger.error('{} ist bereits verwendet. Aktuelle Nummer: {}'.format(Neu_Nummer,Nummer))
            else:
                if Item[0].LookupParameter('Plannummer'):
                    Item[0].LookupParameter('Plannummer').Set(Neu_Nummer)
                    list.append(Neu_Nummer)
    t.Commit()

total = time.time() - start
logger.info("total time: {} {}".format(total, 100 * "_"))

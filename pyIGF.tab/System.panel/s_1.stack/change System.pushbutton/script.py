# coding: utf8
import sys
sys.path.append(r'R:\pyRevit\xx_Skripte\libs\IGF_libs')
from IGF_log import getlog
from rpw import revit, DB
from pyrevit import script, forms

__context__ = 'Selection'
__title__ = "Systeme löschen"
__doc__ = """

löscht die Systeme, die mit Bauteile verknüpft sind.


"""
__authors__ = "Menghui Zhang"

try:
    getlog(__title__)
except:
    pass

uidoc = revit.uidoc
doc = revit.doc

cl = [doc.GetElement(x) for x in uidoc.Selection.GetElementIds()]

if forms.alert('System löschen?',ok=False,yes=True,no=True):
    t = DB.Transaction(doc,'Systemtyp ändern')
    t.Start()
    with forms.ProgressBar(title = "{value}/{max_value} Elemente",cancellable=True, step=int(len(cl)/100)+1) as pb:
        for n,el in enumerate(cl):
            if pb.cancelled:
                t.RollBack()
                script.exit()
            pb.update_progress(n+1,len(cl))
            try:
                el.LookupParameter('Systemtyp').Set(DB.ElementId(-1))
            except:
                pass
    doc.Regenerate()
    t.Commit()
    t.Dispose()

    t = DB.Transaction(doc,'Systemname ändern')
    t.Start()
    
    for el in cl:
        try:
            doc.Delete(el.MEPSystem.Id)
        except:
            pass
    t.Commit()
    t.Dispose()
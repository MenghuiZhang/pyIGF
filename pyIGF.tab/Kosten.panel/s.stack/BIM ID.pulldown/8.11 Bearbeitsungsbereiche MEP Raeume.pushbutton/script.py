# coding: utf8
from pyrevit import revit, UI, DB
from pyrevit import script, forms
from Autodesk.Revit.DB import *
from pyrevit.forms import WPFWindow

__title__ = "8.11 Bearbeitungsbereich MEP Räume"
__doc__ = """Bearbeitungsbereich der MEP Räume ändern

[2021.11.23]
Version: 1.1
"""
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()
bb_config = script.get_config()
try:
    from IGF_log import getlog
    getlog(__title__)
except:pass
uidoc = revit.uidoc
doc = revit.doc

# Bearbeitungsbereich
worksets = FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset)
Workset_dict = {}
for el in worksets:
    Workset_dict[el.Name] = el.Id.ToString()

MEP = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_MEPSpaces).WhereElementIsNotElementType()
MEPids = MEP.ToElementIds()
if len(MEPids) < 1:
    logger.error("Kein MEP Raum in aktueller Projekt")
    script.exit()

Raumtrenn = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_RoomSeparationLines).WhereElementIsNotElementType()
Trenn = Raumtrenn.ToElementIds()
if len(Trenn) < 1:
    logger.error("Keine Raumtrennlinien in aktueller Projekt")

liste_wpf = Workset_dict.Keys[:]
liste_wpf.sort()


# Bearbeitungsbereich Pläne
class Bearbeitungsbereich(WPFWindow):
    def __init__(self, xaml_file_name):
        WPFWindow.__init__(self, xaml_file_name)
        self.bearbeitungsb.ItemsSource = liste_wpf
        self.read_config()


    def read_config(self):
        try:
            if not str(bb_config.bb) in liste_wpf:
                self.bearbeitungsb.Text = bb_config.bb = ""

            try:
                self.bearbeitungsb.Text = str(bb_config.bb)
            except:
                self.bearbeitungsb.Text = bb_config.bb = ""
        except:
            pass

    def write_config(self):
        bb_config.bb = self.bearbeitungsb.SelectedItem.ToString().encode('utf-8')
        script.save_config()

    def ok(self,sender,args):
        self.write_config()
        self.Close()

Bearbeitungsbereichwpf = Bearbeitungsbereich("Window.xaml")
Bearbeitungsbereichwpf.ShowDialog()

Workset = None
try:
    Workset = Workset_dict[str(bb_config.bb)]
except Exception as e:
    logger.error(e)
    script.exit()


t = Transaction(doc,'Bearbeitungsbereich MEP Räume')
t.Start()
with forms.ProgressBar(title='{value}/{max_value} MEP Räume',cancellable=True, step=10) as pb:
    for n,mepid in enumerate(MEPids):
        if pb.cancelled:
            t.RollBack()
            script.exit()
        pb.update_progress(n+1, len(MEPids))
        elem = doc.GetElement(mepid)
        try:
            elem.LookupParameter('Bearbeitungsbereich').Set(int(Workset))
        except:
            pass
t.Commit()

t1 = Transaction(doc,'Bearbeitungsbereich Raumtrennlinie')
t1.Start()
with forms.ProgressBar(title='{value}/{max_value} Raumtrennlinie',cancellable=True, step=1) as pb1:
    for n,trennid in enumerate(Trenn):
        if pb1.cancelled:
            t1.RollBack()
            script.exit()
        pb1.update_progress(n+1, len(Trenn))
        elem = doc.GetElement(trennid)
        try:
            elem.LookupParameter('Bearbeitungsbereich').Set(int(Workset))
        except:
            pass
t1.Commit()
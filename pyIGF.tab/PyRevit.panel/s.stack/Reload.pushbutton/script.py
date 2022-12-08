"""Reload pyRevit into new session."""
# -*- coding=utf-8 -*-
#pylint: disable=import-error,invalid-name,broad-except
from pyrevit import EXEC_PARAMS
from pyrevit import script
from pyrevit import forms
from pyrevit.loader import sessionmgr
from pyrevit.loader import sessioninfo
import clr
clr.AddReference('AdWindows')
from Autodesk.Windows import ComponentManager


res = True
if EXEC_PARAMS.executed_from_ui:
    res = forms.alert('Reloading increases the memory footprint and is '
                      'automatically called by pyRevit when necessary.\n\n'
                      'pyRevit developers can manually reload when:\n'
                      '    - New buttons are added.\n'
                      '    - Buttons have been removed.\n'
                      '    - Button icons have changed.\n'
                      '    - Base C# code has changed.\n'
                      '    - Value of pyRevit parameters\n'
                      '      (e.g. __title__, __doc__, ...) have changed.\n'
                      '    - Cached engines need to be cleared.\n\n'
                      'Are you sure you want to reload?',
                      ok=False, yes=True, no=True)

if res:
    logger = script.get_logger()
    results = script.get_results()

    # re-load pyrevit session.
    logger.info('Reloading....')
    sessionmgr.reload_pyrevit()

    results.newsession = sessioninfo.get_session_uuid()

# et Visibility
def set_visible(item,visible):
    try:
        item.Visible = visible
    except:
        pass
    try:
        item.IsVisible = visible
    except:
        pass

DICT_ALLTABS = {}
all_module = ComponentManager.Ribbon.Tabs
for modul in all_module:
    revit_obj = modul.GetType()
    if revit_obj.ToString() == 'Autodesk.Windows.RibbonTab':
        DICT_ALLTABS[modul.Title] = modul

config = script.get_config('Module ausblenden')
try:
    try:
        hidden_tabs = config.hidetabs
    except:
        hidden_tabs = []
    try:
        hidden_panels = config.hidepanels
    except:
        hidden_panels = {}
    for el_Tab in DICT_ALLTABS.keys():
        if el_Tab in hidden_tabs:
            set_visible(DICT_ALLTABS[el_Tab],False)
            continue
        else:
            set_visible(DICT_ALLTABS[el_Tab],True)
        if el_Tab in hidden_panels.keys():
            hidden_items = hidden_panels[el_Tab]
            for el_panel in DICT_ALLTABS[el_Tab].Panels:
                panelname = el_panel.Source.AutomationName
                if el_panel.Source.Items.Count == 0:
                    set_visible(el_panel,False)
                else:
                    if panelname in hidden_items:
                        set_visible(el_panel,False)
                    else:
                        set_visible(el_panel,True)
except Exception as e:print(e)
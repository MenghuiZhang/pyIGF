"""Reload pyRevit into new session."""
# -*- coding=utf-8 -*-
#pylint: disable=import-error,invalid-name,broad-except
from pyrevit import EXEC_PARAMS
from pyrevit import script
from pyrevit import forms
from pyrevit.loader import sessionmgr
from pyrevit.loader import sessioninfo


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

try:
    import sys
    sys.path.append(r'R:\pyRevit\xx_Skripte\Final\pyrevit.extension\pyIGF.tab\Einstellungen.panel\s.stack\Hide Module.smartbutton')
    from manage_plugins import get_all_plugins,get_itemssource,setup_plugins_init
    all_plugins = get_all_plugins()
    itemssource = get_itemssource(all_plugins)
    setup_plugins_init(itemssource)
except Exception as e:
    script.get_logger().error(e)
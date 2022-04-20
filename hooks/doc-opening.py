# pylint: skip-file
from pyrevit import HOST_APP, EXEC_PARAMS
from pyrevit import revit, script

args = EXEC_PARAMS.event_args

try:
    import sys
    sys.path.append(r'R:\pyRevit\xx_Skripte\Final\pyrevit.extension\pyIGF.tab\Einstellungen.panel\s.stack\Hide Module.smartbutton')
    from manage_plugins import get_all_plugins,get_itemssource,setup_plugins_init
    all_plugins = get_all_plugins()
    itemssource = get_itemssource(all_plugins)
    setup_plugins_init(itemssource)
except Exception as e:
    script.get_logger().error(e)


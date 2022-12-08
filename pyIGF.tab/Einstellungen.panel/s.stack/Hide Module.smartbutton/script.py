# coding: utf8
from manage_plugins import DICT_ALLTABS,ITEMSSOURCE,run_wpf,setup_plugins,change_icon,script,setup_tooltip,setup_plugins_init

__title__ = "PlugIn Menü \neinblenden"

logger = script.get_logger()

def schreiben_Log():
    """write all Log information"""
    try:
        from IGF_log import getlog
        getlog(__title__)
    except:
        pass

if __name__ == '__main__':
    schreiben_Log()
    run_wpf()
    setup_plugins()
    change_icon(button = script.get_button())

def __selfinit__(script_cmp, ui_button_cmp, __rvt__):
    setup_tooltip(button = ui_button_cmp,script=script_cmp)
    try:
        setup_plugins_init()
        change_icon(button = ui_button_cmp,script=script_cmp)
    except:
        pass
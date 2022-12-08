# coding: utf8
from manage_plugins import *

__title__ = "PlugIn Menü ausblenden"

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
    all_plugins = get_all_plugins()
    itemssource = get_itemssource(all_plugins)
    run_wpf(itemssource)
    setup_plugins(itemssource)
    change_icon(button = script.get_button())


def __selfinit__(script_cmp, ui_button_cmp, __rvt__):
    setup_tooltip(button = ui_button_cmp,script=script_cmp)
    try:
        all_plugins = get_all_plugins()
        itemssource = get_itemssource(all_plugins)
        setup_plugins_init(itemssource)
        
        change_icon(button = ui_button_cmp,script=script_cmp)
    except:
        pass
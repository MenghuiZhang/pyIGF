# coding: utf8
from manage_plugins_tab import *

__title__ = "PlugIns ausblenden"

logger = script.get_logger()
config = script.get_config("PlugIns ausblenden")

doc = None

def get_doc():
    """return document"""
    doc = revit.doc
    return doc

def schreiben_Log():
    """write all Log information"""
    try:
        from IGF_log import getlog
        getlog(__title__)
    except:
        pass


if __name__ == '__main__':
    doc = get_doc()
    button = script.get_button()
    dict_plugins = get_all_plugins()
    is_plugins = get_itemssource(sorted(dict_plugins.keys()))
    run_wpf(is_plugins,config)
    setup_plugins(config,dict_plugins.values())
    change_icon(len(config.hidetabs),button)


def __selfinit__(script_cmp, ui_button_cmp, __rvt__):
    tips = "1. Button klicken\n2. Plug-Ins auswählen"
    ui_button_cmp.set_tooltip_ext(tips)
    tooltip_image = script_cmp.get_bundle_file('tooltip.png')
    ui_button_cmp.set_tooltip_image(tooltip_image)

    tooltip_video = script_cmp.get_bundle_file('temp.mp4')
    ui_button_cmp.set_tooltip_video(tooltip_video)

    config_hide = script.get_config("PlugIns ausblenden")

    dict_plugins = get_all_plugins()

    try:
        if len(config_hide.hidetabs) > 0:
            icon = script_cmp.get_bundle_file('on.png')
        else:
            icon = script_cmp.get_bundle_file('off.png')
        ui_button_cmp.set_icon(icon)
    except Exception as e:
        logger.error(e)   
    try:
        setup_plugins(config_hide,dict_plugins.values())
    except Exception as e:
        logger.error(e)
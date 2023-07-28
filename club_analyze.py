#!/usr/bin/python3
#
# club_analyze.py - https://github.com/dmanusrex/club_analyze.py
# Copyright (C) 2021 - Darren Richer
#
# Analyze SWON data and generate a club compliance report
#

import customtkinter as ctk
import club_analyzer_ui as ui
from config import AnalyzerConfig
import swon_version
import os
import sys
import logging
from requests.exceptions import RequestException

from version import ANALYZER_VERSION

def check_for_update() -> None:
    """Notifies if there's a newer released version"""
    current_version = ANALYZER_VERSION
    try:
        latest_version = swon_version.latest()
        if latest_version is not None and not swon_version.is_latest_version(
            latest_version, current_version
        ):
            logging.info(
                f"New version available {latest_version.tag}"
            )
            logging.info(f"Download URL: {latest_version.url}")
#           Make it clickable???  webbrowser.open(latest_version.url))
    except RequestException as ex:
        logging.warning("Error checking for update: %s", ex)


def main():
    '''Runs the Offiicals Analyzer'''

    bundle_dir = getattr(sys, '_MEIPASS', os.path.abspath(os.path.dirname(__file__)))

    root = ctk.CTk()
    config = AnalyzerConfig()
    ctk.set_appearance_mode(config.get_str("Theme"))  # Modes: "System" (standard), "Dark", "Light"
    ctk.set_default_color_theme(config.get_str("Colour"))  # Themes: "blue" (standard), "green", "dark-blue"
    new_scaling_float = int(config.get_str("Scaling").replace("%", "")) / 100
    ctk.set_widget_scaling(new_scaling_float)
    root.title("Swim Ontario - Officials Analyzer")
    icon_file = os.path.abspath(os.path.join(bundle_dir, 'media','swon-analyzer.ico'))
    root.iconbitmap(icon_file)
    root.geometry(f"{1100}x{800}")
    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)
    root.resizable(True, True)
    content = ui.SwonApp(root, config)
    content.grid(column=0, row=0, sticky="news")
    check_for_update()

    try:
        root.update()
        # pylint: disable=import-error,import-outside-toplevel
        import pyi_splash  # type: ignore

        if pyi_splash.is_alive():
            pyi_splash.close()
    except ModuleNotFoundError:
        pass
    except RuntimeError:
        pass

    root.mainloop()

    config.save()

if __name__ == "__main__":
    main()


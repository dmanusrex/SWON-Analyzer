# Club Analyzer - https://github.com/dmanusrex/SWON-Analyzer
# Copyright (C) 2021 - Darren Richer
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
# EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
# MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
# IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
# DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
# OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE
# OR OTHER DEALINGS IN THE SOFTWARE.

'''Analyze SWON data and generate a club compliance report'''


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


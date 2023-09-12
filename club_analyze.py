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

"""Analyze SWON data and generate a club compliance report"""


import customtkinter as ctk   # type: ignore
import club_analyzer_ui as ui
from config import AnalyzerConfig
import os
import sys

from rtr import RTR


def main():
    """Runs the Offiicals Utilities"""

    bundle_dir = getattr(sys, "_MEIPASS", os.path.abspath(os.path.dirname(__file__)))

    root = ctk.CTk()
    config = AnalyzerConfig()
    rtr_data = RTR(config)

    ctk.set_appearance_mode(config.get_str("Theme"))  # Modes: "System" (standard), "Dark", "Light"
    ctk.set_default_color_theme(config.get_str("Colour"))  # Themes: "blue" (standard), "green", "dark-blue"

    root.title("Swim Ontario - Officials Utilities")
    icon_file = os.path.abspath(os.path.join(bundle_dir, "media", "swon-analyzer.ico"))

    root.iconbitmap(icon_file)
    root.columnconfigure(0, weight=1, minsize=400)
    root.rowconfigure(0, weight=1, minsize=600)
    root.resizable(True, True)

    content = ui.SwonApp(root, config, rtr_data)
    content.grid(column=0, row=0, sticky="news")

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

    # Scaling seems to work better after the root.update() call
    new_scaling_float = int(config.get_str("Scaling").replace("%", "")) / 100
#    ctk.set_widget_scaling(new_scaling_float)
#    ctk.set_window_scaling(new_scaling_float)

    root.mainloop()

    config.save()


if __name__ == "__main__":
    main()

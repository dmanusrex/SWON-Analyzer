# Club Analyzer - https://github.com/dmanusrex/SWON-Analyzer
# Copyright (C) 2023 - Darren Richer
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


""" SWON Officials Utilities Main Screen """

import logging
from logging.handlers import RotatingFileHandler
import os
import tkinter as tk
from tkinter import StringVar
import webbrowser
from typing import Any
from platformdirs import user_log_dir
import pathlib

import customtkinter as ctk  # type: ignore
import pandas as pd
from requests.exceptions import RequestException

import swon_version
from config import AnalyzerConfig
from odp import Generate_Documents_Frame
from pathway import Pathway_Documents_Frame, Pathway_ROR_Frame
from rtr import RTR, RTR_Frame
from rtrbrowse import RTR_Browse_Frame
from sanction import Sanction_COA_CoHost, Sanction_ROR
from version import ANALYZER_VERSION, UNLOCK_CODE
from ui_common import Application_Preferences

tkContainer = Any


class TextHandler(logging.Handler):
    # This class allows you to log to a Tkinter Text or ScrolledText widget
    # Adapted from Moshe Kaplan: https://gist.github.com/moshekaplan/c425f861de7bbf28ef06

    def __init__(self, text):
        # run the regular Handler __init__
        logging.Handler.__init__(self)
        # Store a reference to the Text it will log to
        self.text = text

    def emit(self, record):
        msg = self.format(record)

        def append():
            self.text.configure(state="normal")
            self.text.insert(tk.END, msg + "\n")
            self.text.configure(state="disabled")
            # Autoscroll to the bottom
            self.text.yview(tk.END)

        # This is necessary because we can't modify the Text from other threads
        self.text.after(0, append)


class _Logging(ctk.CTkFrame):
    """Logging Window"""

    def __init__(self, container: ctk.CTk, config: AnalyzerConfig):
        super().__init__(container)
        self._config = config
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        ctk.CTkLabel(self, text="Messages").grid(column=0, row=0, sticky="ws", pady=10)

        self.logwin = ctk.CTkTextbox(self, state="disabled")
        self.logwin.grid(column=0, row=1, sticky="nsew")
        # Logging configuration
        logdir = user_log_dir("swon-analyzer", "Swim Ontario")
        pathlib.Path(logdir).mkdir(parents=True, exist_ok=True)
        logfile = os.path.abspath(os.path.join(logdir, "swon-analyzer.log"))

        simple_formatter = logging.Formatter("%(asctime)s - %(message)s")
        detailed_formatter = logging.Formatter("%(asctime)s %(name)s[%(process)d]: %(levelname)s - %(message)s")

        # get a top-level logger,
        # set its log level to DEBUG,
        # BUT PREVENT IT from propagating messages to the root logger
        # This allows different levels to be logged to the file and the text window

        log = logging.getLogger()
        log.setLevel(logging.INFO)
        log.propagate = False

        # create a file handler
        file_handler = RotatingFileHandler(logfile, maxBytes=1000000, backupCount=5)
        file_handler.setLevel(logging.DEBUG)
        file_handler.setFormatter(detailed_formatter)

        # Create textLogger
        text_handler = TextHandler(self.logwin)
        text_handler.setLevel(logging.INFO)
        text_handler.setFormatter(simple_formatter)

        # Add the handler to logger
        log.addHandler(file_handler)
        log.addHandler(text_handler)


class SwonApp(ctk.CTkFrame):
    """Main Appliction"""

    def __init__(self, container: ctk.CTk, config: AnalyzerConfig, rtr_data: RTR):
        super().__init__(container)
        self._config = config
        self._rtr_data = rtr_data
        self.df = pd.DataFrame()
        self.affiliates = pd.DataFrame()
        self.unlocked: bool = False
        self.menu_mode: StringVar = StringVar(value=self._config.get_str("DefaultMenu"))

        # Configure the main window 1x2
        self.grid_rowconfigure(0, weight=1, minsize=400)
        self.grid_columnconfigure(1, weight=1, minsize=400)

        # create navigation frame
        self.navigation_frame = ctk.CTkFrame(self, corner_radius=0)
        self.navigation_frame.grid(row=0, column=0, sticky="nsew")
        self.navigation_frame.grid_rowconfigure(8, weight=1)

        self.navigation_frame_label = ctk.CTkLabel(
            self.navigation_frame,
            text=f"Officials Utilities\n{ANALYZER_VERSION}",
            font=ctk.CTkFont(size=15, weight="bold"),
        )
        self.navigation_frame_label.grid(row=0, column=0, padx=20, pady=20)

        # The menu organizes buttons by Job Function

        self.mode_menu = ctk.CTkOptionMenu(
            self.navigation_frame,
            values=["COA/Co-Host", "ROR/POA"],
            variable=self.menu_mode,
            corner_radius=0,
            command=self.change_buttons,
        )
        self.mode_menu.grid(row=1, column=0, padx=5, pady=(0, 5), sticky="ew")

        # Create buttons common to all jobs

        self.app_preferences_button = ctk.CTkButton(
            self.navigation_frame,
            corner_radius=0,
            height=40,
            border_spacing=10,
            text="App Preferences",
            fg_color="transparent",
            text_color=("gray10", "gray90"),
            hover_color=("gray70", "gray30"),
            anchor="w",
            command=self.app_preferences_button_event,
        )
        self.app_preferences_button.grid(row=2, column=0, sticky="ew")

        self.rtr_button = ctk.CTkButton(
            self.navigation_frame,
            corner_radius=0,
            height=40,
            border_spacing=10,
            text="RTR Data",
            fg_color="transparent",
            text_color=("gray10", "gray90"),
            hover_color=("gray70", "gray30"),
            anchor="w",
            command=self.rtr_button_event,
        )
        self.rtr_button.grid(row=3, column=0, sticky="ew")

        self.rtr_browser_button = ctk.CTkButton(
            self.navigation_frame,
            corner_radius=0,
            height=40,
            border_spacing=10,
            text="RTR Browser",
            fg_color="transparent",
            text_color=("gray10", "gray90"),
            hover_color=("gray70", "gray30"),
            anchor="w",
            command=self.rtr_browser_button_event,
        )

        self.rtr_browser_button.grid(row=4, column=0, sticky="ew")

        # COA/Co-Host Options

        self.sanction_coa_button = ctk.CTkButton(
            self.navigation_frame,
            corner_radius=0,
            height=40,
            border_spacing=10,
            text="Sanctioning (COA/Co-Host)",
            fg_color="transparent",
            text_color=("gray10", "gray90"),
            hover_color=("gray70", "gray30"),
            anchor="w",
            command=self.sanction_coa_button_event,
        )

        self.odp_doc_button = ctk.CTkButton(
            self.navigation_frame,
            corner_radius=0,
            height=40,
            border_spacing=10,
            text="Recommendations (COA)",
            fg_color="transparent",
            text_color=("gray10", "gray90"),
            hover_color=("gray70", "gray30"),
            anchor="w",
            command=self.odp_doc_button_event,
        )

        self.pathway_doc_button = ctk.CTkButton(
            self.navigation_frame,
            corner_radius=0,
            height=40,
            border_spacing=10,
            text="New Pathway (COA)",
            fg_color="transparent",
            text_color=("gray10", "gray90"),
            hover_color=("gray70", "gray30"),
            anchor="w",
            command=self.pathway_doc_button_event,
        )

        # ROR/POA Options

        self.sanction_ror_button = ctk.CTkButton(
            self.navigation_frame,
            corner_radius=0,
            height=40,
            border_spacing=10,
            text="Sanctioning (ROR/POA)",
            fg_color="transparent",
            text_color=("gray10", "gray90"),
            hover_color=("gray70", "gray30"),
            anchor="w",
            command=self.sanction_ror_button_event,
        )

        self.pathway_ror_button = ctk.CTkButton(
            self.navigation_frame,
            corner_radius=0,
            height=40,
            border_spacing=10,
            text="New Pathway (ROR/POA)",
            fg_color="transparent",
            text_color=("gray10", "gray90"),
            hover_color=("gray70", "gray30"),
            anchor="w",
            command=self.pathway_ror_button_event,
        )

        """Turn on the Unlock Code button to enable new features"""

        self.unlock_code_button = ctk.CTkButton(
            self.navigation_frame,
            corner_radius=0,
            height=40,
            border_spacing=10,
            text="Unlock Code",
            fg_color="transparent",
            text_color=("gray10", "gray90"),
            hover_color=("gray70", "gray30"),
            anchor="w",
            command=self.unlock_code_button_event,
        )
        if UNLOCK_CODE is not None:
            self.unlock_code_button.grid(row=9, column=0, sticky="sew")

        """Turn on the New Version button if there's a newer released version"""

        try:
            latest_version = swon_version.latest()
            if latest_version is not None and not swon_version.is_latest_version(latest_version, ANALYZER_VERSION):
                self.new_ver_button = ctk.CTkButton(
                    self.navigation_frame,
                    corner_radius=0,
                    height=40,
                    border_spacing=10,
                    text=f"New Version {latest_version.tag}",
                    fg_color="transparent",
                    text_color=("gray10", "gray90"),
                    hover_color=("gray70", "gray30"),
                    anchor="w",
                    command=self.new_ver_button_event,
                )
                self.new_ver_url = latest_version.url
                self.new_ver_button.grid(row=10, column=0, sticky="sew")
        except RequestException as ex:
            logging.warning("Error checking for update: %s", ex)

        self.log_button = ctk.CTkButton(
            self.navigation_frame,
            corner_radius=0,
            height=40,
            border_spacing=10,
            text="Log Messages",
            fg_color="transparent",
            text_color=("gray10", "gray90"),
            hover_color=("gray70", "gray30"),
            anchor="w",
            command=self.log_button_event,
        )
        self.log_button.grid(row=11, column=0, sticky="sew")

        self.help_button = ctk.CTkButton(
            self.navigation_frame,
            corner_radius=0,
            height=40,
            border_spacing=10,
            text="Help",
            fg_color="transparent",
            text_color=("gray10", "gray90"),
            hover_color=("gray70", "gray30"),
            anchor="w",
            command=self.help_button_event,
        )
        self.help_button.grid(row=12, column=0, sticky="sew")

        # The RTR frame is common to all applications

        self.app_preferences_frame = Application_Preferences(self, self._config)
        self.app_preferences_frame.configure(corner_radius=0, fg_color="transparent")
        self.app_preferences_frame.grid_columnconfigure(0, weight=1)
        self.app_preferences_frame.grid(row=0, column=1, sticky="new")

        self.rtr_frame = RTR_Frame(self, self._config, self._rtr_data)
        self.rtr_frame.configure(corner_radius=0, fg_color="transparent")
        self.rtr_frame.grid_columnconfigure(0, weight=1)
        self.rtr_frame.grid(row=0, column=1, sticky="new")

        self.help_frame = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")

        # create the subframes - Sanctioning Application

        self.sanction_ror_frame = Sanction_ROR(self, self._config, self._rtr_data)
        self.sanction_ror_frame.configure(corner_radius=0, fg_color="transparent")
        self.sanction_ror_frame.grid_columnconfigure(0, weight=1)

        self.sanction_coa_frame = Sanction_COA_CoHost(self, self._config, self._rtr_data)
        self.sanction_coa_frame.configure(corner_radius=0, fg_color="transparent")
        self.sanction_coa_frame.grid_columnconfigure(0, weight=1)

        # create the subframes - ODP Application

        self.odp_doc_frame = Generate_Documents_Frame(self, self._config, self._rtr_data)
        self.odp_doc_frame.configure(corner_radius=0, fg_color="transparent")
        self.odp_doc_frame.grid_columnconfigure(0, weight=1)

        self.rtr_browser_frame = RTR_Browse_Frame(self, self._config, self._rtr_data)
        self.rtr_browser_frame.configure(corner_radius=0, fg_color="transparent")
        self.rtr_browser_frame.grid_columnconfigure(0, weight=1)

        # create the subframes - New Pathway Application

        self.pathway_ror_frame = Pathway_ROR_Frame(self, self._config, self._rtr_data)
        self.pathway_ror_frame.configure(corner_radius=0, fg_color="transparent")
        self.pathway_ror_frame.grid_columnconfigure(0, weight=1)

        self.pathway_doc_frame = Pathway_Documents_Frame(self, self._config, self._rtr_data)
        self.pathway_doc_frame.configure(corner_radius=0, fg_color="transparent")
        self.pathway_doc_frame.grid_columnconfigure(0, weight=1)

        # Logging Window
        self.log_frame = _Logging(self, self._config)
        self.log_frame.configure(corner_radius=0, fg_color="transparent")
        self.log_frame.grid_columnconfigure(0, weight=1)

        # Default to the RTR button pressed
        self.change_buttons(self._config.get_str("DefaultMenu"))
        self.rtr_button_event()

    def change_buttons(self, value):
        if value == "COA/Co-Host":
            self.sanction_ror_button.grid_forget()
            self.pathway_ror_button.grid_forget()
            self.sanction_coa_button.grid(row=5, column=0, sticky="ew")
            self.odp_doc_button.grid(row=6, column=0, sticky="ew")
            if self.unlocked:
                self.pathway_doc_button.grid(row=7, column=0, sticky="ew")
            else:
                self.pathway_doc_button.grid_forget()
            self.rtr_button_event()
        elif value == "ROR/POA":
            self.sanction_coa_button.grid_forget()
            self.odp_doc_button.grid_forget()
            self.pathway_doc_button.grid_forget()
            self.sanction_ror_button.grid(row=5, column=0, sticky="ew")
            if self.unlocked:
                self.pathway_ror_button.grid(row=6, column=0, sticky="ew")
            else:
                self.pathway_ror_button.grid_forget()
            self.rtr_button_event()
        return

    def select_frame_by_name(self, name):
        # set button color for selected button
        self.app_preferences_button.configure(
            fg_color=("gray75", "gray25") if name == "app-preferences" else "transparent"
        )
        self.rtr_button.configure(fg_color=("gray75", "gray25") if name == "rtr" else "transparent")
        self.sanction_ror_button.configure(fg_color=("gray75", "gray25") if name == "sanction-ror" else "transparent")
        self.sanction_coa_button.configure(fg_color=("gray75", "gray25") if name == "sanction-coa" else "transparent")
        self.odp_doc_button.configure(fg_color=("gray75", "gray25") if name == "odp-doc" else "transparent")
        self.rtr_browser_button.configure(fg_color=("gray75", "gray25") if name == "rtr-browser" else "transparent")
        self.log_button.configure(fg_color=("gray75", "gray25") if name == "log" else "transparent")
        self.pathway_doc_button.configure(fg_color=("gray75", "gray25") if name == "pathway-doc" else "transparent")

        # show selected frame
        if name == "app-preferences":
            self.app_preferences_frame.grid(row=0, column=1, sticky="new")
        else:
            self.app_preferences_frame.grid_forget()
        if name == "rtr":
            self.rtr_frame.grid(row=0, column=1, sticky="new")
        else:
            self.rtr_frame.grid_forget()
        if name == "sanction-ror":
            self.sanction_ror_frame.grid(row=0, column=1, sticky="new")
        else:
            self.sanction_ror_frame.grid_forget()
        if name == "sanction-coa":
            self.sanction_coa_frame.grid(row=0, column=1, sticky="new")
        else:
            self.sanction_coa_frame.grid_forget()
        if name == "odp-doc":
            self.odp_doc_frame.grid(row=0, column=1, sticky="new")
        else:
            self.odp_doc_frame.grid_forget()
        if name == "rtr-browser":
            self.rtr_browser_frame.grid(row=0, column=1, sticky="new")
        else:
            self.rtr_browser_frame.grid_forget()
        if name == "log":
            self.log_frame.grid(row=0, column=1, sticky="nsew")
        else:
            self.log_frame.grid_forget()
        if name == "pathway-ror":
            self.pathway_ror_frame.grid(row=0, column=1, sticky="new")
        else:
            self.pathway_ror_frame.grid_forget()
        if name == "pathway-doc":
            self.pathway_doc_frame.grid(row=0, column=1, sticky="new")
        else:
            self.pathway_doc_frame.grid_forget()

    def app_preferences_button_event(self) -> None:
        self.select_frame_by_name("app-preferences")

    def rtr_button_event(self) -> None:
        self.select_frame_by_name("rtr")

    def sanction_ror_button_event(self) -> None:
        self.select_frame_by_name("sanction-ror")

    def sanction_coa_button_event(self) -> None:
        self.select_frame_by_name("sanction-coa")

    def odp_doc_button_event(self) -> None:
        self.select_frame_by_name("odp-doc")

    def rtr_browser_button_event(self) -> None:
        self.select_frame_by_name("rtr-browser")

    def pathway_ror_button_event(self) -> None:
        self.select_frame_by_name("pathway-ror")

    def pathway_doc_button_event(self) -> None:
        self.select_frame_by_name("pathway-doc")

    def new_ver_button_event(self) -> None:
        webbrowser.open(self.new_ver_url)

    def unlock_code_button_event(self) -> None:
        unlockcode = ctk.CTkInputDialog(title="Unlock Experimental Features", text="Enter the Unlock Code")
        # Workaround not being able to pass the show="*" method to the dialog constructor
        unlockcode.after(20, lambda: unlockcode._entry.configure(show="*"))

        if unlockcode.get_input() == UNLOCK_CODE:
            self.rtr_frame.enable_features()
            self.unlocked = True
            self.unlock_code_button.grid_forget()
            self.rtr_button_event()
            logging.info("Experimental Features Unlocked")
        else:
            logging.info("Invalid Unlock Code")

    def log_button_event(self) -> None:
        self.select_frame_by_name("log")

    def help_button_event(self) -> None:
        webbrowser.open("https://swon-analyzer.readthedocs.io")


def main():
    """testing"""
    root = ctk.CTk()
    root.columnconfigure(0, weight=1, minsize=400)
    root.rowconfigure(0, weight=1, minsize=600)
    root.resizable(True, True)
    options = AnalyzerConfig()
    rtr_data = RTR(options)
    settings = SwonApp(root, options, rtr_data)
    settings.grid(column=0, row=0, sticky="news")
    logging.info("Hello World")
    root.mainloop()


if __name__ == "__main__":
    main()

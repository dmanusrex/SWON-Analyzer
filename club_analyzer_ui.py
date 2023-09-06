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

import os
import pandas as pd
import logging
import webbrowser
import tkinter as tk
import customtkinter as ctk
from typing import Any
from requests.exceptions import RequestException

# Appliction Specific Imports
from config import AnalyzerConfig
from version import ANALYZER_VERSION, UNLOCK_CODE
from rtr import RTR, RTR_Frame
from odp import Generate_Documents_Frame, Email_Documents_Frame
from sanction import Sanction_Preferences, Sanction_ROR, Sanction_COA_CoHost
from pathway import Pathway_Documents_Frame, Pathway_ROR_Frame

import swon_version

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


class _Logging(ctk.CTkFrame):  # pylint: disable=too-many-ancestors,too-many-instance-attributes
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
        logfile = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "swon-analyzer.log"))

        logging.basicConfig(filename=logfile, level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
        # Create textLogger
        text_handler = TextHandler(self.logwin)
        text_handler.setFormatter(logging.Formatter("%(levelname)s - %(message)s"))
        # Add the handler to logger
        logger = logging.getLogger()
        logger.addHandler(text_handler)


class SwonApp(ctk.CTkFrame):  # pylint: disable=too-many-ancestors
    """Main Appliction"""

    # pylint: disable=too-many-arguments,too-many-locals
    def __init__(self, container: ctk.CTk, config: AnalyzerConfig, rtr_data: RTR):
        super().__init__(container)
        self._config = config
        self._rtr_data = rtr_data
        self.df = pd.DataFrame()
        self.affiliates = pd.DataFrame()

        # Configure the main window 1x2
        self.grid_rowconfigure(0, weight=1, minsize=400)
        self.grid_columnconfigure(1, weight=1, minsize=400)

        # create navigation frame
        self.navigation_frame = ctk.CTkFrame(self, corner_radius=0)
        self.navigation_frame.grid(row=0, column=0, sticky="nsew")
        self.navigation_frame.grid_rowconfigure(6, weight=1)

        self.navigation_frame_label = ctk.CTkLabel(
            self.navigation_frame,
            text=f"Officials Utilities\n{ANALYZER_VERSION}",
            font=ctk.CTkFont(size=15, weight="bold"),
        )
        self.navigation_frame_label.grid(row=0, column=0, padx=20, pady=20)

        # The menu will allow the user to switch between sanctioning and personal ODP

        self.mode_menu = ctk.CTkOptionMenu(
            self.navigation_frame,
            values=["Sanctioning", "Officals Recomendations"],
            corner_radius=0,
            command=self.change_buttons,
        )
        self.mode_menu.grid(row=1, column=0, padx=5, pady=(0, 5), sticky="ew")

        self.sanction_preferences_button = ctk.CTkButton(
            self.navigation_frame,
            corner_radius=0,
            height=40,
            border_spacing=10,
            text="Preferences",
            fg_color="transparent",
            text_color=("gray10", "gray90"),
            hover_color=("gray70", "gray30"),
            anchor="w",
            command=self.sanction_preferences_button_event,
        )
        self.sanction_preferences_button.grid(row=2, column=0, sticky="ew")

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

        self.sanction_ror_button = ctk.CTkButton(
            self.navigation_frame,
            corner_radius=0,
            height=40,
            border_spacing=10,
            text="ROR/POA Reports",
            fg_color="transparent",
            text_color=("gray10", "gray90"),
            hover_color=("gray70", "gray30"),
            anchor="w",
            command=self.sanction_ror_button_event,
        )
        self.sanction_ror_button.grid(row=4, column=0, sticky="ew")

        self.sanction_coa_button = ctk.CTkButton(
            self.navigation_frame,
            corner_radius=0,
            height=40,
            border_spacing=10,
            text="COA/Co-Host",
            fg_color="transparent",
            text_color=("gray10", "gray90"),
            hover_color=("gray70", "gray30"),
            anchor="w",
            command=self.sanction_coa_button_event,
        )
        self.sanction_coa_button.grid(row=5, column=0, sticky="ew")

        # Create the ODP buttons

        self.odp_preferences_button = ctk.CTkButton(
            self.navigation_frame,
            corner_radius=0,
            height=40,
            border_spacing=10,
            text="Preferences",
            fg_color="transparent",
            text_color=("gray10", "gray90"),
            hover_color=("gray70", "gray30"),
            anchor="w",
            command=self.odp_preferences_button_event,
        )

        self.odp_doc_button = ctk.CTkButton(
            self.navigation_frame,
            corner_radius=0,
            height=40,
            border_spacing=10,
            text="Documents",
            fg_color="transparent",
            text_color=("gray10", "gray90"),
            hover_color=("gray70", "gray30"),
            anchor="w",
            command=self.odp_doc_button_event,
        )

        self.odp_email_button = ctk.CTkButton(
            self.navigation_frame,
            corner_radius=0,
            height=40,
            border_spacing=10,
            text="Emails",
            fg_color="transparent",
            text_color=("gray10", "gray90"),
            hover_color=("gray70", "gray30"),
            anchor="w",
            command=self.odp_email_button_event,
        )

        # Create the Pathway Buttons

        self.pathway_ror_button = ctk.CTkButton(
            self.navigation_frame,
            corner_radius=0,
            height=40,
            border_spacing=10,
            text="ROR/POA Reports",
            fg_color="transparent",
            text_color=("gray10", "gray90"),
            hover_color=("gray70", "gray30"),
            anchor="w",
            command=self.pathway_ror_button_event,
        )

        self.pathway_doc_button = ctk.CTkButton(
            self.navigation_frame,
            corner_radius=0,
            height=40,
            border_spacing=10,
            text="Pathway Docs",
            fg_color="transparent",
            text_color=("gray10", "gray90"),
            hover_color=("gray70", "gray30"),
            anchor="w",
            command=self.pathway_doc_button_event,
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
            self.unlock_code_button.grid(row=7, column=0, sticky="sew")

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
                self.new_ver_button.grid(row=8, column=0, sticky="sew")
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
        self.log_button.grid(row=9, column=0, sticky="sew")

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
        self.help_button.grid(row=10, column=0, sticky="sew")

        # The RTR frame is common to all applications

        self.rtr_frame = RTR_Frame(self, self._config, self._rtr_data)
        self.rtr_frame.configure(corner_radius=0, fg_color="transparent")
        self.rtr_frame.grid_columnconfigure(0, weight=1)
        self.rtr_frame.grid(row=0, column=1, sticky="new")

        self.help_frame = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")

        # create the subframes - Sanctioning Application

        self.sanction_preferences_frame = Sanction_Preferences(self, self._config)
        self.sanction_preferences_frame.configure(corner_radius=0, fg_color="transparent")
        self.sanction_preferences_frame.grid_columnconfigure(0, weight=1)

        self.sanction_ror_frame = Sanction_ROR(self, self._config, self._rtr_data)
        self.sanction_ror_frame.configure(corner_radius=0, fg_color="transparent")
        self.sanction_ror_frame.grid_columnconfigure(0, weight=1)

        self.sanction_coa_frame = Sanction_COA_CoHost(self, self._config, self._rtr_data)
        self.sanction_coa_frame.configure(corner_radius=0, fg_color="transparent")
        self.sanction_coa_frame.grid_columnconfigure(0, weight=1)

        # create the subframes - ODP Application

        #        self.odp_preferences_frame = Sanction_Preferences(self, self._config)
        #        self.odp_preferences_frame.configure(corner_radius=0, fg_color="transparent")
        #        self.odp_preferences_frame.grid_columnconfigure(0, weight=1)

        self.odp_doc_frame = Generate_Documents_Frame(self, self._config, self._rtr_data)
        self.odp_doc_frame.configure(corner_radius=0, fg_color="transparent")
        self.odp_doc_frame.grid_columnconfigure(0, weight=1)

        self.odp_email_frame = Email_Documents_Frame(self, self._config)
        self.odp_email_frame.configure(corner_radius=0, fg_color="transparent")
        self.odp_email_frame.grid_columnconfigure(0, weight=1)

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
        self.rtr_button_event()

    def change_buttons(self, value):
        if value == "Sanctioning":
            self.sanction_preferences_button.grid(row=2, column=0, sticky="ew")
            self.sanction_ror_button.grid(row=4, column=0, sticky="ew")
            self.sanction_coa_button.grid(row=5, column=0, sticky="ew")
            #            self.odp_preferences_button.grid_forget()
            self.odp_doc_button.grid_forget()
            self.odp_email_button.grid_forget()
            self.pathway_ror_button.grid_forget()
            self.pathway_doc_button.grid_forget()
            self.rtr_button_event()
        elif value == "Officals Recomendations":
            self.sanction_preferences_button.grid_forget()
            self.sanction_ror_button.grid_forget()
            self.sanction_coa_button.grid_forget()
            #            self.odp_preferences_button.grid(row=2, column=0, sticky="ew")
            self.odp_doc_button.grid(row=4, column=0, sticky="ew")
            self.odp_email_button.grid(row=5, column=0, sticky="ew")
            self.pathway_ror_button.grid_forget()
            self.pathway_doc_button.grid_forget()
            self.rtr_button_event()
        else:
            self.sanction_preferences_button.grid_forget()
            self.sanction_ror_button.grid_forget()
            self.sanction_coa_button.grid_forget()
            self.odp_doc_button.grid_forget()
            self.odp_email_button.grid_forget()
            self.pathway_ror_button.grid(row=4, column=0, sticky="ew")
            self.pathway_doc_button.grid(row=5, column=0, sticky="ew")
            self.rtr_button_event()
        return

    def select_frame_by_name(self, name):
        # set button color for selected button
        self.sanction_preferences_button.configure(
            fg_color=("gray75", "gray25") if name == "preferences" else "transparent"
        )
        self.rtr_button.configure(fg_color=("gray75", "gray25") if name == "rtr" else "transparent")
        self.sanction_ror_button.configure(fg_color=("gray75", "gray25") if name == "sanction-ror" else "transparent")
        self.sanction_coa_button.configure(fg_color=("gray75", "gray25") if name == "sanction-coa" else "transparent")
        self.odp_preferences_button.configure(
            fg_color=("gray75", "gray25") if name == "odp-preferences" else "transparent"
        )
        self.odp_doc_button.configure(fg_color=("gray75", "gray25") if name == "odp-doc" else "transparent")
        self.odp_email_button.configure(fg_color=("gray75", "gray25") if name == "odp-email" else "transparent")
        self.log_button.configure(fg_color=("gray75", "gray25") if name == "log" else "transparent")
        self.pathway_doc_button.configure(fg_color=("gray75", "gray25") if name == "pathway-doc" else "transparent")

        # show selected frame
        if name == "sanction-preferences":
            self.sanction_preferences_frame.grid(row=0, column=1, sticky="new")
        else:
            self.sanction_preferences_frame.grid_forget()
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
        #        if name == "odp-preferences":
        #            self.odp_preferences_frame.grid(row=0, column=1, sticky="new")
        #        else:
        #            self.odp_preferences_frame.grid_forget()
        if name == "odp-doc":
            self.odp_doc_frame.grid(row=0, column=1, sticky="new")
        else:
            self.odp_doc_frame.grid_forget()
        if name == "odp-email":
            self.odp_email_frame.grid(row=0, column=1, sticky="new")
        else:
            self.odp_email_frame.grid_forget()
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

    def sanction_preferences_button_event(self) -> None:
        self.select_frame_by_name("sanction-preferences")

    def rtr_button_event(self) -> None:
        self.select_frame_by_name("rtr")

    def sanction_ror_button_event(self) -> None:
        self.select_frame_by_name("sanction-ror")

    def sanction_coa_button_event(self) -> None:
        self.select_frame_by_name("sanction-coa")

    def odp_preferences_button_event(self) -> None:
        self.select_frame_by_name("odp-preferences")

    def odp_doc_button_event(self) -> None:
        self.select_frame_by_name("odp-doc")

    def odp_email_button_event(self) -> None:
        self.select_frame_by_name("odp-email")

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
            self.mode_menu.configure(values=["Sanctioning", "Officals Recomendations", "Pathway Check"])
            self.unlock_code_button.grid_forget()
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

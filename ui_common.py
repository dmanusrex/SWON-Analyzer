# Club Analyzer - https://github.com/dmanusrex/SWON-Analyzer
#
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


""" Common UI Elements """

import logging
from tkinter import BooleanVar, StringVar
from typing import Any

import customtkinter as ctk  # type: ignore

# Appliction Specific Imports
from config import AnalyzerConfig


class Officials_Status_Frame(ctk.CTkFrame):
    """Status Codes for Officials"""

    def __init__(self, container: ctk.CTk, config: AnalyzerConfig):
        super().__init__(container)
        self._config = config

        self._incl_inv_pending = BooleanVar(value=self._config.get_bool("incl_inv_pending"))
        self._incl_pso_pending = BooleanVar(value=self._config.get_bool("incl_pso_pending"))
        self._incl_account_pending = BooleanVar(value=self._config.get_bool("incl_account_pending"))

        # self is a vertical container that will contain 1 frame
        self.columnconfigure(0, weight=1)

        # Options Frame - Left and Right Panels

        optionsframe = self
#        optionsframe.grid(column=0, row=0, sticky="news", padx=10, pady=10)
#        optionsframe.rowconfigure(0, weight=1)

        ctk.CTkLabel(optionsframe, text="RTR Officials Status").grid(column=0, row=0, sticky="w", padx=10)

        ctk.CTkSwitch(
            optionsframe,
            text="PSO Pending",
            variable=self._incl_pso_pending,
            onvalue=True,
            offvalue=False,
            command=self._handle_incl_pso_pending,
        ).grid(column=0, row=1, sticky="w", padx=20, pady=10)

        ctk.CTkSwitch(
            optionsframe,
            text="Account Pending",
            variable=self._incl_account_pending,
            onvalue=True,
            offvalue=False,
            command=self._handle_incl_account_pending,
        ).grid(column=0, row=2, sticky="w", padx=20, pady=10)

        ctk.CTkSwitch(
            optionsframe,
            text="Invoice Pending",
            variable=self._incl_inv_pending,
            onvalue=True,
            offvalue=False,
            command=self._handle_incl_inv_pending,
        ).grid(column=0, row=3, sticky="w", padx=20, pady=10)

    def _handle_incl_pso_pending(self, *_arg) -> None:
        self._config.set_bool("incl_pso_pending", self._incl_pso_pending.get())

    def _handle_incl_account_pending(self, *_arg) -> None:
        self._config.set_bool("incl_account_pending", self._incl_account_pending.get())

    def _handle_incl_inv_pending(self, *_arg) -> None:
        self._config.set_bool("incl_inv_pending", self._incl_inv_pending.get())


class Application_Preferences(ctk.CTkFrame):
    """Application Wide Preferences"""

    def __init__(self, container: ctk.CTk, config: AnalyzerConfig):
        super().__init__(container)
        self._config = config
        self._ctk_theme = StringVar(value=self._config.get_str("Theme"))
        self._ctk_size = StringVar(value=self._config.get_str("Scaling"))
        self._ctk_colour = StringVar(value=self._config.get_str("Colour"))
        self._default_menu = StringVar(value=self._config.get_str("DefaultMenu"))

        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=0)

        ctk.CTkLabel(self, text="Application Preferences", font=ctk.CTkFont(weight="bold")).grid(
            column=0, row=0, padx=10, pady=10
        )

        ui_appearence = ctk.CTkFrame(self)
        ui_appearence.grid(column=0, row=1, sticky="news", padx=10, pady=10)
        ui_appearence.rowconfigure(0, weight=0)
        ui_appearence.rowconfigure(1, weight=0)
        ui_appearence.rowconfigure(2, weight=0)
        ui_appearence.columnconfigure(0, weight=0)
        ui_appearence.columnconfigure(1, weight=0)

        self.ui_appearence_label = ctk.CTkLabel(ui_appearence, text="UI Appearance", font=ctk.CTkFont(weight="bold"))
        self.ui_appearence_label.grid(row=0, column=0, columnspan=2, sticky="w")

        self.appearance_mode_label = ctk.CTkLabel(ui_appearence, text="Appearance Mode")
        self.appearance_mode_label.grid(row=1, column=1, sticky="w", padx=0)
        self.appearance_mode_optionemenu = ctk.CTkOptionMenu(
            ui_appearence,
            values=["Light", "Dark", "System"],
            command=self.change_appearance_mode_event,
            variable=self._ctk_theme,
        )
        self.appearance_mode_optionemenu.grid(row=1, column=0, padx=(20, 0), pady=10, sticky="w")

        self.scaling_label = ctk.CTkLabel(ui_appearence, text="UI Scaling")
        self.scaling_label.grid(row=2, column=1, sticky="w")
        self.scaling_optionemenu = ctk.CTkOptionMenu(
            ui_appearence,
            values=["80%", "90%", "100%", "110%", "120%"],
            command=self.change_scaling_event,
            variable=self._ctk_size,
        )
        self.scaling_optionemenu.grid(row=2, column=0, padx=20, pady=10, sticky="w")

        self.colour_label = ctk.CTkLabel(ui_appearence, text="Colour (Restart Required)")
        self.colour_label.grid(row=3, column=1, sticky="w")
        self.colour_optionemenu = ctk.CTkOptionMenu(
            ui_appearence,
            values=["blue", "green", "dark-blue"],
            command=self.change_colour_event,
            variable=self._ctk_colour,
        )
        self.colour_optionemenu.grid(row=3, column=0, padx=20, pady=10, sticky="w")

        def_menu_fr = ctk.CTkFrame(self)
        def_menu_fr.grid(column=0, row=2, sticky="news", padx=10, pady=10)
        def_menu_fr.rowconfigure(0, weight=1)
        def_menu_fr.rowconfigure(1, weight=1)
        def_menu_fr.rowconfigure(2, weight=1)
        def_menu_fr.rowconfigure(3, weight=1)
        def_menu_fr.columnconfigure(0, weight=1)
        def_menu_fr.columnconfigure(1, weight=0)

        self.def_menu_fr_label = ctk.CTkLabel(def_menu_fr, text="Default Menu", font=ctk.CTkFont(weight="bold"))
        self.def_menu_fr_label.grid(row=0, column=0, columnspan=2, sticky="w")

        self.mode_menu = ctk.CTkOptionMenu(
            def_menu_fr,
            values=["COA/Co-Host", "ROR/POA"],
            variable=self._default_menu,
            corner_radius=0,
            command=self.set_default_menu,
        )
        self.mode_menu.grid(row=1, column=0, padx=20, pady=10, sticky="w")

    def change_appearance_mode_event(self, new_appearance_mode: str):
        ctk.set_appearance_mode(new_appearance_mode)
        self._config.set_str("Theme", new_appearance_mode)

    def change_scaling_event(self, new_scaling: str):
        new_scaling_float = int(new_scaling.replace("%", "")) / 100
        ctk.set_widget_scaling(new_scaling_float)
        ctk.set_window_scaling(new_scaling_float)
        self._config.set_str("Scaling", new_scaling)

    def change_colour_event(self, new_colour: str):
        logging.info("Changing colour to : " + new_colour)
        ctk.set_default_color_theme(new_colour)
        self._config.set_str("Colour", new_colour)

    def set_default_menu(self, new_default_menu: str):
        self._config.set_str("DefaultMenu", new_default_menu)


def main():
    """testing"""
    root = ctk.CTk()
    root.resizable(True, True)
    options = AnalyzerConfig()
    settings = Application_Preferences(root, options)
    #    settings = Officials_Status_Frame(root, options)
    settings.grid(column=0, row=0, sticky="news")
    root.mainloop()


if __name__ == "__main__":
    main()

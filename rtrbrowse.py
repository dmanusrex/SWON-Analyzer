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


""" RTR Browser Module """

import logging
import os
import tkinter as tk
from datetime import datetime
from tkinter import BooleanVar, StringVar, filedialog
from typing import Any

import customtkinter as ctk  # type: ignore
import pandas as pd

# Appliction Specific Imports
from config import AnalyzerConfig
from CTkMessagebox import CTkMessagebox  # type: ignore
from rtr import RTR
from rtr_fields import RTR_CLINICS
from tooltip import ToolTip

tkContainer = Any


class RTR_Browse_Frame(ctk.CTkFrame):
    """Allow COAs to browse/export RTR Data"""

    def __init__(self, container: tkContainer, config: AnalyzerConfig, rtr: RTR):
        super().__init__(container)
        self._config = config
        self._rtr = rtr

        self._incl_inv_pending = BooleanVar(value=self._config.get_bool("incl_inv_pending"))
        self._incl_pso_pending = BooleanVar(value=self._config.get_bool("incl_pso_pending"))
        self._incl_account_pending = BooleanVar(value=self._config.get_bool("incl_account_pending"))

        # Add support for the club selection
        self._club_list = ["None"]
        self._club_selected = ctk.StringVar(value="None")

        # List of positions in long form and the associated short form
        self._positions = [
            ("Introduction to Swimming Officiating", "Intro"),
            ("Safety Marshal", "Safety"),
            ("Judge of Stroke/Inspector of Turns (Combo)", "ST"),
            ("Inspector of Turns", "IT"),
            ("Judge of Stroke", "JoS"),
            ("Administration Desk (formerly Clerk of Course) Clinic", "AdminDesk"),
            ("Chief Timekeeper", "CT"),
            ("Meet Manager", "MM"),
            ("Starter", "Starter"),
            ("Chief Finish Judge/Chief Judge Electronics", "CFJ"),
            ("Chief Recorder", "ChiefRec"),
        ]

        self._positions_long = [position[0] for position in self._positions]
        self._position_selected = ctk.StringVar(value=self._positions_long[0])

        self._cert_list = ["Both", "Qualified", "Certified"]
        self._cert_selected = ctk.StringVar(value=self._cert_list[0])

        # self is a vertical container that will contain 3 frames
        self.columnconfigure(0, weight=1)

        buttonsframe = ctk.CTkFrame(self)
        buttonsframe.grid(column=0, row=0, sticky="news")
        buttonsframe.rowconfigure(0, weight=0)

        optionsframe = ctk.CTkFrame(self)
        optionsframe.grid(column=0, row=2, sticky="news")

        # Add Command Buttons

        ctk.CTkLabel(buttonsframe, text="Club Selection").grid(column=0, row=0, sticky="w", padx=10)

        self.club_dropdown = ctk.CTkOptionMenu(
            buttonsframe,
            dynamic_resizing=True,
            values=self._club_list,
            variable=self._club_selected,
            command=self._handle_club_change,
        )
        self.club_dropdown.grid(row=2, column=0, padx=20, pady=(20, 10), sticky="w")

        self.reports_btn = ctk.CTkButton(buttonsframe, text="Full Export (CSV)", command=self._handle_reports_btn)
        self.reports_btn.grid(column=1, row=2, sticky="w", padx=20, pady=(20, 10))

        self.bar = ctk.CTkProgressBar(master=buttonsframe, orientation="horizontal", mode="indeterminate")

        # Options Frame - Left and Right Panels

        left_optionsframe = ctk.CTkFrame(optionsframe)
        left_optionsframe.grid(column=0, row=0, sticky="news", padx=10, pady=10)
        left_optionsframe.rowconfigure(0, weight=1)
        right_optionsframe = ctk.CTkFrame(optionsframe)
        right_optionsframe.grid(column=1, row=0, sticky="news", padx=10, pady=10)
        right_optionsframe.rowconfigure(0, weight=1)
        lower_optionsframe = ctk.CTkFrame(optionsframe)
        lower_optionsframe.grid(column=0, row=1, columnspan=2, sticky="news", padx=10, pady=10)
        lower_optionsframe.rowconfigure(0, weight=1)
        lower_optionsframe.columnconfigure(0, weight=1)

        # Program Options on the left frame

        ctk.CTkLabel(left_optionsframe, text="Position Selection").grid(column=0, row=0, sticky="w", padx=10)

        self.position_dropdown = ctk.CTkOptionMenu(
            left_optionsframe,
            dynamic_resizing=True,
            values=self._positions_long,
            variable=self._position_selected,
            command=self._handle_position_change,
        )
        self.position_dropdown.grid(row=2, column=0, padx=20, pady=(10, 10), sticky="w")

        self.qual_dropdown = ctk.CTkOptionMenu(
            left_optionsframe,
            dynamic_resizing=True,
            values=self._cert_list,
            variable=self._cert_selected,
            command=self._handle_qual_change,
        )
        self.qual_dropdown.grid(row=4, column=0, padx=20, pady=(10, 10), sticky="w")

        # Right options frame for status options

        ctk.CTkLabel(right_optionsframe, text="RTR Officials Status").grid(column=0, row=0, sticky="w", padx=10)

        ctk.CTkSwitch(
            right_optionsframe,
            text="PSO Pending",
            variable=self._incl_pso_pending,
            onvalue=True,
            offvalue=False,
            command=self._handle_incl_pso_pending,
        ).grid(column=0, row=1, sticky="w", padx=20, pady=10)

        ctk.CTkSwitch(
            right_optionsframe,
            text="Account Pending",
            variable=self._incl_account_pending,
            onvalue=True,
            offvalue=False,
            command=self._handle_incl_account_pending,
        ).grid(column=0, row=2, sticky="w", padx=20, pady=10)

        ctk.CTkSwitch(
            right_optionsframe,
            text="Invoice Pending",
            variable=self._incl_inv_pending,
            onvalue=True,
            offvalue=False,
            command=self._handle_incl_inv_pending,
        ).grid(column=0, row=3, sticky="w", padx=20, pady=10)

        # Lower options frame for club selection

        ctk.CTkLabel(lower_optionsframe, text="Matching Officials   ").grid(column=0, row=0, sticky="ws", padx=10, pady=(10,0))

        self.officials_list = ctk.CTkTextbox(lower_optionsframe, state="disabled")
        self.officials_list.grid(column=0, row=1, sticky="ew", padx=10, pady=10)

        # Register Callback
        self._rtr.register_update_callback(self.refresh_club_list)

    def refresh_club_list(self):
        self._club_list = ["None"]
        self._club_list = self._club_list + [club[1] for club in self._rtr.club_list_names]

        self.club_dropdown.configure(values=self._club_list)
        if len(self._club_list) == 2:  # There is only one club loaded, change default
            self.club_dropdown.set(self._club_list[1])
        else:
            self.club_dropdown.set(self._club_list[0])
        self._update_officials_list()
        logging.info("RTR Browser - Club List Refreshed")

    def _handle_incl_pso_pending(self, *_arg) -> None:
        self._config.set_bool("incl_pso_pending", self._incl_pso_pending.get())

    def _handle_incl_account_pending(self, *_arg) -> None:
        self._config.set_bool("incl_account_pending", self._incl_account_pending.get())

    def _handle_incl_inv_pending(self, *_arg) -> None:
        self._config.set_bool("incl_inv_pending", self._incl_inv_pending.get())

    def _handle_club_change(self, *_arg) -> None:
        self._update_officials_list()

    def _handle_position_change(self, *_arg) -> None:
        self._update_officials_list()

    def _handle_qual_change(self, *_arg) -> None:
        self._update_officials_list()

    def _update_officials_list(self) -> None:
        """Update the officials list"""
        club = self._club_selected.get()
        position_long = self._position_selected.get()
        position = [position[1] for position in self._positions if position[0] == position_long][0]

        cert = self._cert_selected.get()[0]  # First letter of the selected value
        pos_status = RTR_CLINICS[position]["status"]
        pos_signoffs = RTR_CLINICS[position]["signoffs"]
        status_values = ["Active"]
        if self._incl_inv_pending.get():
            status_values.append("Invoice Pending")
        if self._incl_account_pending.get():
            status_values.append("Account Pending")
        if self._incl_pso_pending.get():
            status_values.append("PSO Pending")


        if club == "None":
            self.officials_list.configure(state="disabled")
            self.officials_list.delete("1.0", "end")
            self.officials_list.insert("1.0", "Select a club first")
            return

        # Filter on status values, club and position status.  If both are selected the filter is != "N"
        if cert == "B":
            self._rtr_filtered = self._rtr.rtr_data.loc[
                (self._rtr.rtr_data["Status"].isin(status_values))
                & (self._rtr.rtr_data["Club"] == club)
                & (self._rtr.rtr_data[pos_status] != "N")
            ]
        else:
            self._rtr_filtered = self._rtr.rtr_data.loc[
                (self._rtr.rtr_data["Status"].isin(status_values))
                & (self._rtr.rtr_data["Club"] == club)
                & (self._rtr.rtr_data[pos_status] == cert)
            ]

        # Sort by last name
#        self._rtr_filtered = self._rtr_filtered.sort_values(by=["Last Name"])

        # Build the list of officials and number of signoffs
        officials = ""
        for index, row in self._rtr_filtered.iterrows():
            officials = officials + row["Full Name"] + " (" + str(row[pos_signoffs]) + " signoffs)\n"

        self.officials_list.configure(state="normal")
        self.officials_list.delete("1.0", "end")
        self.officials_list.insert("1.0", officials)
        self.officials_list.configure(state="disabled")

    def buttons(self, newstate) -> None:
        """Enable/disable all buttons"""
        self.reports_btn.configure(state=newstate)

    def _handle_reports_btn(self) -> None:
        if self._rtr.rtr_data.empty:
            logging.info("Load data first...")
            CTkMessagebox(master=self, title="Error", message="Load RTR Data First", icon="cancel", corner_radius=0)
            return
        club = self._club_selected.get()
        if club == "None":
            logging.info("Select a club first...")
            CTkMessagebox(master=self, title="Error", message="Select a club first", icon="cancel", corner_radius=0)
            return
        self.buttons("disabled")
        self.bar.grid(row=2, column=0, pady=10, padx=20, sticky="s")
        self.bar.set(0)
        self.bar.start()

        key_columns = [
            "Registration Id",
            "First Name",
            "Last Name",
            "Club",
            "Region",
            "Province",
            "Status",
            "Current_CertificationLevel",
            "Intro_Status",
            "Intro_Count",
            "ST_Status",
            "ST_Count",
            "IT_Status",
            "IT_Count",
            "JoS_Status",
            "JoS_Count",
            "CT_Status",
            "CT_Count",
            "Admin_Status",
            "Admin_Count",
            "MM_Status",
            "MM_Count",
            "Starter_Status",
            "Starter_Count",
            "CFJ_Status",
            "CFJ_Count",
            "ChiefRec_Status",
            "ChiefRec_Count",
            "Referee_Status",
            "Para Swimming eModule",
        ]

        status_values = ["Active"]
        if self._config.get_bool("incl_inv_pending"):
            status_values.append("Invoice Pending")
        if self._config.get_bool("incl_account_pending"):
            status_values.append("Account Pending")
        if self._config.get_bool("incl_pso_pending"):
            status_values.append("PSO Pending")

        # Filter on status values and club
        self._rtr_filtered = self._rtr.rtr_data.loc[
            (self._rtr.rtr_data["Status"].isin(status_values)) & (self._rtr.rtr_data["Club"] == club)
        ]

        try:
            report_file = filedialog.asksaveasfilename(
                filetypes=[("CSV Files", "*.csv")],
                defaultextension=".csv",
                title="Export File",
                initialdir=self._config.get_str("odp_report_directory"),
            )
            if len(report_file) == 0:
                self.bar.stop()
                self.bar.grid_forget()
                self.buttons("enable")
                return

            self._rtr_filtered.to_csv(report_file, columns=key_columns, index=False)
            logging.info("CSV Report Saved: {}".format(report_file))

        except Exception as e:
            logging.info("Unable to save CSV file: {}".format(type(e).__name__))
            logging.info("Exception message: {}".format(e))
            CTkMessagebox(title="Error", message="Unable to save CSV file", icon="cancel", corner_radius=0)

        self.bar.stop()
        self.bar.grid_forget()
        self.buttons("enable")

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


""" RTR Datafile Handling """
import pandas as pd
import numpy as np
import logging
import tkinter as tk
import customtkinter as ctk   # type: ignore
import chardet
from CTkMessagebox import CTkMessagebox
from tkinter import filedialog, ttk, BooleanVar, StringVar
from typing import Any, Callable
from threading import Thread
from datetime import datetime
from copy import deepcopy, copy
from tooltip import ToolTip
from time import sleep

# Appliction Specific Imports
from config import AnalyzerConfig
from rtr_fields import REQUIRED_RTR_FIELDS, RTR_POSITION_FIELDS

NoneFn = Callable[[], None]

tkContainer = Any


class _Data_Loader(Thread):
    """Load RTR Data files"""

    _RTR_Fields = RTR_POSITION_FIELDS

    def __init__(self, config: AnalyzerConfig):
        super().__init__()
        self._config = config
        self.rtr_data: pd.DataFrame
        self.failure_reason = ""

    def run(self):
        html_file = self._config.get_str("officials_list")
        self.club_list_names_df = pd.DataFrame
        self.club_list_names = []
        logging.info("Loading RTR Data")

        # Check if the RTR file is a CSV or HTML file by reading the first line and looking for "Registration Id"

        try:
            with open(html_file, "r") as f:
                first_line = f.readline()
        except:
            logging.info("Unable to open data file")
            self.failure_reason = "Unable to open data file"
            self.rtr_data = pd.DataFrame
            return

        if "Registration Id" in first_line:
            logging.info("CSV Formatted File Detected")
            # Re-read the first megabyte of data to try to determine the character encoding
            with open(html_file, "rb") as f:
                result = chardet.detect(f.read(1000000))
            logging.info("Detected encoding: {}".format(result["encoding"]))
            try:
                self._rtr_data = pd.read_csv(
                    html_file, usecols=REQUIRED_RTR_FIELDS, na_values=["0001-01-01"], encoding=result["encoding"]
                )
            except Exception as e:
                logging.info("Unable to load CSV file: {}".format(type(e).__name__))
                logging.info("Exception message: {}".format(e))
                self.failure_reason = "Unable to load CSV data file - See log for details"
                self.rtr_data = pd.DataFrame
                return
        else:
            logging.info("HTML Formatted File Detected")
            try:
                self._rtr_data = pd.read_html(html_file, na_values=["0001-01-01"])[0]
            except:
                logging.info("Unable to load data file")
                self.failure_reason = "Unable to load data file"
                self.rtr_data = pd.DataFrame
                return
            self._rtr_data.columns = self._rtr_data.iloc[0]  # The first row is the column names
            self._rtr_data = self._rtr_data[1:]

        # Check if required rtr fields are present

        if not all(item in self._rtr_data.columns for item in REQUIRED_RTR_FIELDS):
            self.rtr_data = pd.DataFrame
            logging.info("Missing Fields - Please use a RTR export after September 1, 2023")
            self.failure_reason = "Missing Fields - Please use a RTR export after September 1, 2023"
            return

        # Club Level exports include blank rows, purge those out
        self._rtr_data.drop(index=self._rtr_data[self._rtr_data["Registration Id"].isnull()].index, inplace=True)

        # The RTR has 2 types of "empty" dates.  One is blank the other is 0001-01-01.  Fix that.
        self._rtr_data.replace("0001-01-01", np.nan, inplace=True)

        # The RTR export is inconsistent on column values for certifications. Fix that.
        self._rtr_data.replace("Yes", "yes", inplace=True)  # We don't use the no value so no need to fix it

        # Filter to the required RTR Fields

        self._rtr_data = self._rtr_data[REQUIRED_RTR_FIELDS]

        # Extend the RTR dataset with the new pathway checks

        # Normalize the current RTR status fields

        self._rtr_data = self._rtr_data.join(self._rtr_data.apply(self._add_new_columns, axis=1))

        # Update the pathway columns
        self._rtr_data["NP_Official"] = self._rtr_data.apply(lambda row: self._np_official(row), axis=1)
        self._rtr_data["NP_Ref1"] = self._rtr_data.apply(lambda row: self._np_ref_1(row), axis=1)
        self._rtr_data["NP_Ref2"] = self._rtr_data.apply(lambda row: self._np_ref_2(row), axis=1)
        self._rtr_data["NP_Starter1"] = self._rtr_data.apply(lambda row: self._np_starter_1(row), axis=1)
        self._rtr_data["NP_Starter2"] = self._rtr_data.apply(lambda row: self._np_starter_2(row), axis=1)
        self._rtr_data["NP_MM1"] = self._rtr_data.apply(lambda row: self._np_mm_1(row), axis=1)
        self._rtr_data["NP_MM2"] = self._rtr_data.apply(lambda row: self._np_mm_2(row), axis=1)

        # final verion - Filter for all valid statuses and make a copy

        self.rtr_data = self._rtr_data.loc[
            self._rtr_data["Status"].isin(["Active", "PSO Pending", "Invoice Pending", "Account Pending"])
        ].copy()

        logging.info("Loaded %d officials" % self.rtr_data.shape[0])

        logging.info("Loading Complete")

    def _add_new_columns(self, row: Any) -> pd.Series:
        return pd.Series(
            [
                self._cert_status(row, "Intro"),
                self._cert_status(row, "ST"),
                self._cert_status(row, "IT"),
                self._cert_status(row, "JoS"),
                self._cert_status(row, "CT"),
                self._cert_status(row, "Clerk"),
                self._cert_status(row, "MM"),
                self._cert_status(row, "Starter"),
                self._cert_status(row, "CFJ"),
                self._cert_status(row, "ChiefRec"),
                self._cert_status(row, "Referee"),
                "N",
                "N",
                "N",
                "N",
                "N",
                "N",
                "N",
            ],
            index=[
                "Intro_Status",
                "ST_Status",
                "IT_Status",
                "JoS_Status",
                "CT_Status",
                "Clerk_Status",
                "MM_Status",
                "Starter_Status",
                "CFJ_Status",
                "ChiefRec_Status",
                "Referee_Status",
                "NP_Official",
                "NP_Ref1",
                "NP_Ref2",
                "NP_Starter1",
                "NP_Starter2",
                "NP_MM1",
                "NP_MM2",
            ],
        )

    def _is_valid_date(self, date_string) -> bool:
        if pd.isnull(date_string):
            return False
        if date_string == "0001-01-01":
            return False
        try:
            datetime.strptime(date_string, "%Y-%m-%d")
            return True
        except ValueError:
            return False

    def _cert_status(self, row: Any, skill: str) -> str:
        """Returns N for not qualfied, Q for Qualfied and C for Certified"""

        rtr_fields = self._RTR_Fields[skill]

        if row["Current_CertificationLevel"] in ["LEVEL IV - GREEN PIN", "LEVEL V - BLUE PIN"]:
            return "C"  # Qualified - All Level IV/Vs are certified, detail records may not exist

        if row[rtr_fields[0]].lower() == "no":  # No Clinic Taken
            return "N"

        if len(rtr_fields) == 1:  # No certifications required
            return "C"

        if len(rtr_fields) == 2:
            if self._is_valid_date(row[rtr_fields[1]]):
                return "C"
        elif self._is_valid_date(row[rtr_fields[1]]) and self._is_valid_date(row[rtr_fields[2]]):
            return "C"

        return "Q"  # Qualified

    def _np_official(self, row: Any) -> str:
        """Check if a certified official in the new pathway"""

        if (
            row["Intro_Status"] == "C"
            and (row["ST_Status"] == "Q" or row["ST_Status"] == "C")
            and row["CT_Status"] == "C"
        ):
            return "Yes"
        return "No"

    def _np_ref_1(self, row: Any) -> str:
        """Referee 1 in the new pathway"""

        if row["NP_Official"] == "No":  # You have to be a certified official first
            return "No"

        # New Pathway - ST&T, Starter and Admin Desk and Referee Certified, Chief Recorder, CFJ and MM Qualfiied or certified

        C_or_Q = ["C", "Q"]

        if (
            row["ST_Status"] == "C"
            and row["Starter_Status"] == "C"
            and row["Clerk_Status"] == "C"
            and row["Referee_Status"] in C_or_Q
            and row["ChiefRec_Status"] in C_or_Q
            and row["CFJ_Status"] in C_or_Q
            and row["MM_Status"] in C_or_Q
            and (row["Para Swimming eModule"].lower() == "yes" or row["Para Domestic"] == "Trained Official")
        ):
            return "Yes"

        return "No"

    def _np_ref_2(self, row: Any) -> str:
        """Referee 1 in the new pathway"""

        if row["NP_Ref1"] == "No":  # You have to be a certified Refree 1 first
            return "No"

        # New Pathway - Ref 1 + Certified in Chief Recorder, CFJ and MM, must be signed off as ref (which is currently IV/V)

        C_or_Q = ["C", "Q"]

        if (
            row["ChiefRec_Status"] == "C"
            and row["CFJ_Status"] == "C"
            and row["MM_Status"] == "C"
            and row["Current_CertificationLevel"] in ["LEVEL IV - GREEN PIN", "LEVEL V - BLUE PIN"]
        ):
            return "Yes"

        return "No"

    def _np_starter_1(self, row: Any) -> str:
        """Starter 1 in the new pathway"""

        if row["NP_Official"] == "No":  # You have to be a certified official first
            return "No"

        # New Pathway - S&T, Starter Certified plus on-line para clinic

        if (
            row["ST_Status"] == "C"
            and row["Starter_Status"] == "C"
            and (row["Para Swimming eModule"].lower() == "yes" or row["Para Domestic"] == "Trained Official")
        ):
            return "Yes"

        return "No"

    def _np_starter_2(self, row: Any) -> str:
        """Starter 2 in the new pathway"""

        if row["NP_Starter1"] == "No":  # You have to be a certified Starter 1 first
            return "No"

        # New Pathway - Starter 1 + Certified as CFJ

        if row["CFJ_Status"] == "C":
            return "Yes"

        return "No"

    def _np_mm_1(self, row: Any) -> str:
        """MM 1 in the new pathway"""

        if row["NP_Official"] == "No":  # You have to be a certified official first
            return "No"

        # New Pathway - Certifed as Meet Manager

        if row["MM_Status"] == "C":
            return "Yes"

        return "No"

    def _np_mm_2(self, row: Any) -> str:
        """MM 2 in the new pathway"""

        if row["NP_MM1"] == "No":  # You have to be a certified MM 1 first
            return "No"

        # New Pathway - Certifed in Chief Recorder, CFJ, Clerk of Course

        if row["ChiefRec_Status"] == "C" and row["CFJ_Status"] == "C" and row["Clerk_Status"] == "C":
            return "Yes"

        return "No"


class RTR:
    """RTR Application Data"""

    def __init__(self, config: AnalyzerConfig, **kwargs):
        self._config = config
        self.rtr_data = pd.DataFrame()
        self.affiliates = pd.DataFrame()
        self.club_list_names_df = pd.DataFrame()
        self.club_list_names: list = []
        self._update_fn: list = []

        # Pre-calculate some statistics on the loaded data

        self.total_officials = StringVar(value="0")
        self.total_clubs = StringVar(value="0")
        self.total_affilated_officials = StringVar(value="0")
        self.total_active = StringVar(value="0")
        self.total_pso_pending = StringVar(value="0")
        self.total_inv_pending = StringVar(value="0")
        self.total_account_pending = StringVar(value="0")
        self.total_NoLevel = StringVar(value="0")
        self.total_Level_I = StringVar(value="0")
        self.total_Level_II = StringVar(value="0")
        self.total_Level_III = StringVar(value="0")
        self.total_Level_IV = StringVar(value="0")
        self.total_Level_V = StringVar(value="0")
        self.total_np_official = StringVar(value="0")
        self.total_np_ref1 = StringVar(value="0")
        self.total_np_ref2 = StringVar(value="0")
        self.total_np_starter1 = StringVar(value="0")
        self.total_np_starter2 = StringVar(value="0")
        self.total_np_mm1 = StringVar(value="0")
        self.total_np_mm2 = StringVar(value="0")

    def load_rtr_data(self, new_data: pd.DataFrame) -> None:
        if self.rtr_data.empty:
            self.rtr_data = new_data.copy()
            logging.info("%d officials records loaded" % self.rtr_data.shape[0])
        else:
            self.rtr_data = pd.concat([self.rtr_data, new_data], axis=0).drop_duplicates()
            logging.info("%d officials records merged" % self.rtr_data.shape[0])

        # We exclude affiliated offiicals from determining the list of clubs. This is important for club level exports.
        self.club_list_names_df = self.rtr_data.loc[
            self.rtr_data["AffiliatedClubs"].isnull(), ["ClubCode", "Club"]
        ].drop_duplicates()
        self.club_list_names = self.club_list_names_df.values.tolist()
        self.club_list_names.sort(key=lambda x: x[0])

        logging.info("Extracting Affiliation Data")

        # Find officials with affiliated clubs. For sanctioning affiliated offiicals must have a certification level
        self.affiliates = self.rtr_data[
            ~self.rtr_data["AffiliatedClubs"].isnull() & ~self.rtr_data["Current_CertificationLevel"].isnull()
        ].copy()
        if self.affiliates.empty:
            logging.info("No affiliation records found")
        else:
            self.affiliates["AffiliatedClubs"] = self.affiliates["AffiliatedClubs"].str.split(",")
            self.affiliates = self.affiliates.explode("AffiliatedClubs")
            logging.info("Extracted %d affiliation records" % self.affiliates.shape[0])

        self.calculate_stats()

    #            self.cohost.refresh_club_list(self.club_list_names)

    def calculate_stats(self) -> None:
        """Calculate statistics on the loaded data"""
        self.total_officials.set(str(self.rtr_data.shape[0]))
        self.total_clubs.set(str(len(self.club_list_names)))
        self.total_affilated_officials.set(str(self.affiliates.shape[0]))
        if self.rtr_data.empty:
            self.total_active.set("0")
            self.total_pso_pending.set("0")
            self.total_inv_pending.set("0")
            self.total_account_pending.set("0")
            self.total_NoLevel.set("0")
            self.total_Level_I.set("0")
            self.total_Level_II.set("0")
            self.total_Level_III.set("0")
            self.total_Level_IV.set("0")
            self.total_Level_V.set("0")
            self.total_np_official.set("0")
            self.total_np_ref1.set("0")
            self.total_np_ref2.set("0")
            self.total_np_starter1.set("0")
            self.total_np_starter2.set("0")
            self.total_np_mm1.set("0")
            self.total_np_mm2.set("0")
        else:
            self.total_active.set(str(self.rtr_data.loc[self.rtr_data["Status"] == "Active"].shape[0]))
            self.total_pso_pending.set(str(self.rtr_data.loc[self.rtr_data["Status"] == "PSO Pending"].shape[0]))
            self.total_inv_pending.set(str(self.rtr_data.loc[self.rtr_data["Status"] == "Invoice Pending"].shape[0]))
            self.total_account_pending.set(
                str(self.rtr_data.loc[self.rtr_data["Status"] == "Account Pending"].shape[0])
            )
            self.total_NoLevel.set(
                str(self.rtr_data.loc[self.rtr_data["Current_CertificationLevel"].isnull()].shape[0])
            )
            self.total_Level_I.set(
                str(self.rtr_data.loc[self.rtr_data["Current_CertificationLevel"] == "LEVEL I - RED PIN"].shape[0])
            )
            self.total_Level_II.set(
                str(self.rtr_data.loc[self.rtr_data["Current_CertificationLevel"] == "LEVEL II - WHITE PIN"].shape[0])
            )
            self.total_Level_III.set(
                str(self.rtr_data.loc[self.rtr_data["Current_CertificationLevel"] == "LEVEL III - ORANGE PIN"].shape[0])
            )
            self.total_Level_IV.set(
                str(self.rtr_data.loc[self.rtr_data["Current_CertificationLevel"] == "LEVEL IV - GREEN PIN"].shape[0])
            )
            self.total_Level_V.set(
                str(self.rtr_data.loc[self.rtr_data["Current_CertificationLevel"] == "LEVEL V - BLUE PIN"].shape[0])
            )
            self.total_np_official.set(str(self.rtr_data.loc[self.rtr_data["NP_Official"] == "Yes"].shape[0]))
            self.total_np_ref1.set(str(self.rtr_data.loc[self.rtr_data["NP_Ref1"] == "Yes"].shape[0]))
            self.total_np_ref2.set(str(self.rtr_data.loc[self.rtr_data["NP_Ref2"] == "Yes"].shape[0]))
            self.total_np_starter1.set(str(self.rtr_data.loc[self.rtr_data["NP_Starter1"] == "Yes"].shape[0]))
            self.total_np_starter2.set(str(self.rtr_data.loc[self.rtr_data["NP_Starter2"] == "Yes"].shape[0]))
            self.total_np_mm1.set(str(self.rtr_data.loc[self.rtr_data["NP_MM1"] == "Yes"].shape[0]))
            self.total_np_mm2.set(str(self.rtr_data.loc[self.rtr_data["NP_MM2"] == "Yes"].shape[0]))

        self.run_update_callbacks()  # Update other UI elements

    def register_update_callback(self, updatefn: NoneFn) -> None:
        """Register a callback function to be called when the data is updated"""
        self._update_fn.append(updatefn)

    def run_update_callbacks(self) -> None:
        """Run all registered callbacks"""
        for fn in self._update_fn:
            fn()

    def reset_data(self) -> None:
        self.rtr_data = pd.DataFrame()
        self.affiliates = pd.DataFrame()
        self.club_list_names_df = pd.DataFrame()
        self.club_list_names = []
        self.calculate_stats()
        logging.info("Reset Complete")


class RTR_Frame(ctk.CTkFrame):  # pylint: disable=too-many-ancestors
    """Load and manage RTR data files"""

    def __init__(self, container: tkContainer, config: AnalyzerConfig, rtr_data: RTR):
        super().__init__(container)
        self._config = config
        self._rtr_data = rtr_data
        self._officials_list = StringVar(value=self._config.get_str("officials_list"))

        # self is a vertical container
        self.columnconfigure(0, weight=1)
        self.columnconfigure(1, weight=0)
        self.rowconfigure(0, weight=1)

        filesframe = ctk.CTkFrame(self)
        filesframe.grid(column=0, row=1, columnspan=2, sticky="news", padx=10, pady=10)
        filesframe.columnconfigure(0, weight=1)
        filesframe.columnconfigure(1, weight=1)

        statsframe = ctk.CTkFrame(self)
        statsframe.grid(column=0, row=12, columnspan=2, sticky="news")
        statsframe.columnconfigure(0, weight=1)
        statsframe.columnconfigure(1, weight=1)

        self.stats1left = ctk.CTkFrame(statsframe)
        self.stats1left.grid(column=0, row=0, sticky="news", padx=10, pady=10)
        self.stats1left.columnconfigure(0, weight=1)
        self.stats1left.columnconfigure(1, weight=1)

        self.stats1right = ctk.CTkFrame(statsframe)
        self.stats1right.grid(column=1, row=0, sticky="news", padx=10, pady=10)
        self.stats1right.columnconfigure(0, weight=1)
        self.stats1right.columnconfigure(1, weight=1)

        self.stats2left = ctk.CTkFrame(statsframe)
        self.stats2left.grid(column=0, row=1, sticky="news", padx=10, pady=10)
        self.stats2left.columnconfigure(0, weight=1)
        self.stats2left.columnconfigure(1, weight=1)

        self.stats2right = ctk.CTkFrame(statsframe)
        self.stats2right.columnconfigure(0, weight=1)
        self.stats2right.columnconfigure(1, weight=1)

        frlabel = ctk.CTkLabel(self, text="RTR Data Loading", font=ctk.CTkFont(weight="bold"))
        frlabel.grid(column=0, row=0, columnspan=2)

        self.rtrbtn = ctk.CTkButton(filesframe, text="RTR List", command=self._handle_officials_browse)
        self.rtrbtn.grid(column=0, row=2, padx=20, pady=10)
        ToolTip(self.rtrbtn, text="Select the RTR officials export file")  # pylint: disable=C0330
        self.rtrfileentry = ctk.CTkLabel(filesframe, textvariable=self._officials_list)
        self.rtrfileentry.grid(column=1, row=2, sticky="w")

        self.load_btn = ctk.CTkButton(filesframe, text="Load Datafile", command=self._handle_load_btn)
        self.load_btn.grid(column=0, row=4, padx=20, pady=10)
        self.load_txt = ctk.CTkLabel(filesframe, text="Load the RTR Officials Datafile")
        self.load_txt.grid(column=1, row=4, sticky="w")

        self.reset_btn = ctk.CTkButton(filesframe, text="Reset", command=self._handle_reset_btn)
        self.reset_btn.grid(column=0, row=6, padx=20, pady=10)
        ctk.CTkLabel(filesframe, text="Restart data loading").grid(column=1, row=6, sticky="w")

        self.bar = ctk.CTkProgressBar(master=filesframe, orientation="horizontal", mode="indeterminate")

        ctk.CTkLabel(self.stats1left, text="Overall Summary", font=ctk.CTkFont(weight="bold")).grid(
            column=0, row=0, columnspan=2, sticky="news"
        )

        ctk.CTkLabel(self.stats1left, text="Officials:  ").grid(column=0, row=14, sticky="e")
        ctk.CTkLabel(self.stats1left, textvariable=self._rtr_data.total_officials).grid(column=1, row=14, sticky="w")

        ctk.CTkLabel(self.stats1left, text="Clubs:  ").grid(column=0, row=15, sticky="e")
        ctk.CTkLabel(self.stats1left, textvariable=self._rtr_data.total_clubs).grid(column=1, row=15, sticky="w")

        ctk.CTkLabel(self.stats1left, text="Affiliated Officials:  ").grid(column=0, row=16, sticky="e")
        ctk.CTkLabel(self.stats1left, textvariable=self._rtr_data.total_affilated_officials).grid(
            column=1, row=16, sticky="w"
        )

        ctk.CTkLabel(self.stats1right, text="Status Summary", font=ctk.CTkFont(weight="bold")).grid(
            column=0, row=0, columnspan=2, sticky="news"
        )

        ctk.CTkLabel(self.stats1right, text="Active:  ").grid(column=0, row=17, sticky="e")
        ctk.CTkLabel(self.stats1right, textvariable=self._rtr_data.total_active).grid(column=1, row=17, sticky="w")

        ctk.CTkLabel(self.stats1right, text="PSO Pending:  ").grid(column=0, row=18, sticky="e")
        ctk.CTkLabel(self.stats1right, textvariable=self._rtr_data.total_pso_pending).grid(column=1, row=18, sticky="w")

        ctk.CTkLabel(self.stats1right, text="Invoice Pending:  ").grid(column=0, row=19, sticky="e")
        ctk.CTkLabel(self.stats1right, textvariable=self._rtr_data.total_inv_pending).grid(column=1, row=19, sticky="w")

        ctk.CTkLabel(self.stats1right, text="Account Pending:  ").grid(column=0, row=20, sticky="e")
        ctk.CTkLabel(self.stats1right, textvariable=self._rtr_data.total_account_pending).grid(
            column=1, row=20, sticky="w"
        )

        ctk.CTkLabel(self.stats2left, text="Level Summary (Old Pathway)", font=ctk.CTkFont(weight="bold")).grid(
            column=0, row=0, columnspan=2, sticky="news"
        )

        ctk.CTkLabel(self.stats2left, text="No Level:  ").grid(column=0, row=21, sticky="e")
        ctk.CTkLabel(self.stats2left, textvariable=self._rtr_data.total_NoLevel).grid(column=1, row=21, sticky="w")

        ctk.CTkLabel(self.stats2left, text="Level I:  ").grid(column=0, row=22, sticky="e")
        ctk.CTkLabel(self.stats2left, textvariable=self._rtr_data.total_Level_I).grid(column=1, row=22, sticky="w")

        ctk.CTkLabel(self.stats2left, text="Level II:  ").grid(column=0, row=23, sticky="e")
        ctk.CTkLabel(self.stats2left, textvariable=self._rtr_data.total_Level_II).grid(column=1, row=23, sticky="w")

        ctk.CTkLabel(self.stats2left, text="Level III:  ").grid(column=0, row=24, sticky="e")
        ctk.CTkLabel(self.stats2left, textvariable=self._rtr_data.total_Level_III).grid(column=1, row=24, sticky="w")

        ctk.CTkLabel(self.stats2left, text="Level IV:  ").grid(column=0, row=25, sticky="e")
        ctk.CTkLabel(self.stats2left, textvariable=self._rtr_data.total_Level_IV).grid(column=1, row=25, sticky="w")

        ctk.CTkLabel(self.stats2left, text="Level V:  ").grid(column=0, row=26, sticky="e")
        ctk.CTkLabel(self.stats2left, textvariable=self._rtr_data.total_Level_V).grid(column=1, row=26, sticky="w")

        ctk.CTkLabel(self.stats2right, text="New Pathway Summary", font=ctk.CTkFont(weight="bold")).grid(
            column=0, row=0, columnspan=2, sticky="news"
        )

        ctk.CTkLabel(self.stats2right, text="Certified Officials:  ").grid(column=0, row=27, sticky="e")
        ctk.CTkLabel(self.stats2right, textvariable=self._rtr_data.total_np_official).grid(column=1, row=27, sticky="w")

        ctk.CTkLabel(self.stats2right, text="Referee 1:  ").grid(column=0, row=28, sticky="e")
        ctk.CTkLabel(self.stats2right, textvariable=self._rtr_data.total_np_ref1).grid(column=1, row=28, sticky="w")

        ctk.CTkLabel(self.stats2right, text="Referee 2:  ").grid(column=0, row=29, sticky="e")
        ctk.CTkLabel(self.stats2right, textvariable=self._rtr_data.total_np_ref2).grid(column=1, row=29, sticky="w")

        ctk.CTkLabel(self.stats2right, text="Starter 1:  ").grid(column=0, row=30, sticky="e")
        ctk.CTkLabel(self.stats2right, textvariable=self._rtr_data.total_np_starter1).grid(column=1, row=30, sticky="w")

        ctk.CTkLabel(self.stats2right, text="Starter 2:  ").grid(column=0, row=31, sticky="e")
        ctk.CTkLabel(self.stats2right, textvariable=self._rtr_data.total_np_starter2).grid(column=1, row=31, sticky="w")

        ctk.CTkLabel(self.stats2right, text="Meet Manager 1:  ").grid(column=0, row=32, sticky="e")
        ctk.CTkLabel(self.stats2right, textvariable=self._rtr_data.total_np_mm1).grid(column=1, row=32, sticky="w")

        ctk.CTkLabel(self.stats2right, text="Meet Manager 2:  ").grid(column=0, row=33, sticky="e")
        ctk.CTkLabel(self.stats2right, textvariable=self._rtr_data.total_np_mm2).grid(column=1, row=33, sticky="w")

    def enable_features(self) -> None:
        self.stats2right.grid(column=1, row=1, sticky="news", padx=10, pady=10)

    def _handle_officials_browse(self) -> None:
        directory = filedialog.askopenfilename()
        if len(directory) == 0:
            return
        self._config.set_str("officials_list", directory)
        self._officials_list.set(directory)

    def buttons(self, newstate) -> None:
        """Enable/disable all buttons on the UI"""
        self.load_btn.configure(state=newstate)
        self.reset_btn.configure(state=newstate)

    def _handle_load_btn(self) -> None:
        self.buttons("disabled")
        self.load_txt.grid_forget()
        self.bar.grid(column=1, row=4, sticky="w", pady=10, padx=10)
        self.bar.set(0)
        self.bar.start()

        load_thread = _Data_Loader(self._config)
        load_thread.start()
        self.monitor_load_thread(load_thread)
        self.buttons("enabled")

    def _handle_reset_btn(self) -> None:
        self.buttons("disabled")
        self._rtr_data.reset_data()
        self.buttons("enabled")

    def monitor_load_thread(self, thread):
        """Monitor the loading thread"""
        if thread.is_alive():
            self.update_idletasks()
            # check the thread every 100ms
            self.after(100, lambda: self.monitor_load_thread(thread))
        else:
            # Retrieve data from the loading process and merge it with already loaded data
            if not thread.rtr_data.empty:
                self._rtr_data.load_rtr_data(thread.rtr_data)
            else:
                CTkMessagebox(self, title="Error", message=thread.failure_reason, icon="cancel", corner_radius=0)
            self.bar.stop()
            self.bar.grid_forget()
            self.load_txt.grid(column=1, row=4, sticky="w")

            thread.join()

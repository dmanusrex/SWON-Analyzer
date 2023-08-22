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


''' RTR Datafile Handling '''
import os
import pandas as pd
import numpy as np
import logging
import webbrowser
import tkinter as tk
import customtkinter as ctk
from tkinter import filedialog, ttk, BooleanVar, StringVar
from typing import Any, Callable
from threading import Thread
from datetime import datetime
from copy import deepcopy, copy
from tooltip import ToolTip
from time import sleep

# Appliction Specific Imports
from config import AnalyzerConfig

NoneFn = Callable[[], None]

tkContainer = Any

class _Data_Loader(Thread):
    '''Load RTR Data files'''
    def __init__(self, config: AnalyzerConfig):
        super().__init__()
        self._config = config
        self.rtr_data : pd.DataFrame 

    def run(self):
        html_file = self._config.get_str("officials_list")
        self.club_list_names_df = pd.DataFrame
        self.club_list_names = []
        logging.info("Loading RTR Data")
        try:
            self.rtr_data = pd.read_html(html_file)[0]
        except:
            logging.info("Unable to load data file")
            self.rtr_data = pd.DataFrame
            return
        self.rtr_data.columns = self.rtr_data.iloc[0]   # The first row is the column names
        self.rtr_data = self.rtr_data[1:]

        # Club Level exports include blank rows, purge those out
        self.rtr_data.drop(index=self.rtr_data[self.rtr_data['Registration Id'].isnull()].index, inplace=True)

        # The RTR has 2 types of "empty" dates.  One is blank the other is 0001-01-01.  Fix that.
        self.rtr_data.replace('0001-01-01', np.nan, inplace=True)

        # The RTR export is inconsistent on column values for certifications. Fix that.
        self.rtr_data.replace('Yes','yes', inplace=True)    # We don't use the no value so no need to fix it 

        logging.info("Loaded %d officials" % self.rtr_data.shape[0])

        logging.info("Loading Complete")

class RTR:
    '''RTR Application Data'''
    _RTR = {
        'Intro': ["Introduction to Swimming Officiating", "Introduction to Swimming Officiating-Deck Evaluation #1 Date", "Introduction to Swimming Officiating-Deck Evaluation #2 Date"],
        'ST': ["Judge of Stroke/Inspector of Turns", "Judge of Stroke/Inspector of Turns-Deck Evaluation #1 Date", "Judge of Stroke/Inspector of Turns-Deck Evaluation #2 Date"],
        'IT': ["Judge of Stroke/Inspector of Turns", "Judge of Stroke/Inspector of Turns-Deck Evaluation #1 Date", "Judge of Stroke/Inspector of Turns-Deck Evaluation #2 Date"],
        'JoS': ["Judge of Stroke/Inspector of Turns", "Judge of Stroke/Inspector of Turns-Deck Evaluation #2 Date"],
        'CT': ["Chief Timekeeper", "Chief Timekeeper-Deck Evaluation #1 Date", "Chief Timekeeper-Deck Evaluation #2 Date"],
        'Clerk': ["Clerk of Course", "Clerk of Course-Deck Evaluation #1 Date", "Clerk of Course-Deck Evaluation #2 Date"],
        'MM': ["Meet Manager", "Meet Manager-Deck Evaluation #1 Date", "Meet Manager-Deck Evaluation #2 Date"],
        'Starter': ["Starter", "Starter-Deck Evaluation #1 Date", "Starter-Deck Evaluation #2 Date"],
        'ChiefRec': ["Recorder-Scorer"],
        'CFJ': ["Chief Finish Judge/Chief Judge", "Chief Finish Judge/Chief Judge-Deck Evaluation #1 Date", "Chief Finish Judge/Chief Judge-Deck Evaluation #2 Date"],
        'Referee': ["Referee", "Referee-Deck Evaluation #1 Date", "Referee-Deck Evaluation #2 Date"]
    }
    def __init__(self, config: AnalyzerConfig, **kwargs):
        self._config = config
        self.rtr_data = pd.DataFrame()
        self.affiliates = pd.DataFrame()
        self.club_list_names_df = pd.DataFrame
        self.club_list_names = []
        self._update_fn : NoneFn = None
        
        # Pre-calculate some statistics on the loaded data

        self.total_officials = StringVar(value="0")
        self.total_clubs = StringVar(value="0")
        self.total_affilated_officials = StringVar(value="0")
        self.total_active = StringVar(value="0")
        self.total_pso_pending = StringVar(value="0")
        self.total_inv_pending = StringVar(value="0")
        self.total_account_pending = StringVar(value="0")


    def load_rtr_data(self, new_data: pd.DataFrame) -> None:
        if self.rtr_data.empty:
            self.rtr_data = new_data.copy()
            logging.info("%d officials records loaded" % self.rtr_data.shape[0])
        else:
            self.rtr_data = pd.concat([self.rtr_data,new_data], axis=0).drop_duplicates()
            logging.info("%d officials records merged" % self.rtr_data.shape[0])

        # We exclude affiliated offiicals from determining the list of clubs. This is important for club level exports.
        self.club_list_names_df = self.rtr_data.loc[self.rtr_data['AffiliatedClubs'].isnull(),['ClubCode','Club']].drop_duplicates()
        self.club_list_names = self.club_list_names_df.values.tolist()
        self.club_list_names.sort(key=lambda x:x[0])

        logging.info("Extracting Affiliation Data")

        # Find officials with affiliated clubs. For sanctioning affiliated offiicals must have a certification level
        self.affiliates = self.rtr_data[~self.rtr_data['AffiliatedClubs'].isnull() & ~self.rtr_data['Current_CertificationLevel'].isnull()].copy()
        if self.affiliates.empty:
            logging.info("No affiliation records found")
        else:
            self.affiliates['AffiliatedClubs'] = self.affiliates['AffiliatedClubs'].str.split(',')
            self.affiliates = self.affiliates.explode('AffiliatedClubs')
            logging.info("Extracted %d affiliation records" % self.affiliates.shape[0])
        
        self.calculate_stats()

    #            self.cohost.refresh_club_list(self.club_list_names)

    def calculate_stats(self) -> None:
        '''Calculate statistics on the loaded data'''
        self.total_officials.set(str(self.rtr_data.shape[0]))
        self.total_clubs.set(str(len(self.club_list_names)))
        self.total_affilated_officials.set(str(self.affiliates.shape[0]))
        if self.rtr_data.empty:
            self.total_active.set("0")
            self.total_pso_pending.set("0")
            self.total_inv_pending.set("0")
            self.total_account_pending.set("0")
        else:
            self.total_active.set(str(self.rtr_data.loc[self.rtr_data['Status'] == 'Active'].shape[0]))
            self.total_pso_pending.set(str(self.rtr_data.loc[self.rtr_data['Status'] == 'PSO Pending'].shape[0]))
            self.total_inv_pending.set(str(self.rtr_data.loc[self.rtr_data['Status'] == 'Invoice Pending'].shape[0]))
            self.total_account_pending.set(str(self.rtr_data.loc[self.rtr_data['Status'] == 'Account Pending'].shape[0]))
        self._update_fn()   # Update other UI elements
        
    def register_update_callback(self, updatefn: NoneFn) -> None:
        '''Register a callback function to be called when the data is updated'''
        self._update_fn = updatefn
   
    def reset_data(self) -> None:
        self.rtr_data = pd.DataFrame()
        self.affiliates = pd.DataFrame()
        self.club_list_names_df = pd.DataFrame()
        self.club_list_names = []
#        self.cohost.refresh_club_list(self.club_list_names)
        self.calculate_stats()
        logging.info("Reset Complete")



class RTR_Frame(ctk.CTkFrame):   # pylint: disable=too-many-ancestors
    '''Load and manage RTR data files'''
    def __init__(self, container: tkContainer, config: AnalyzerConfig, rtr_data: RTR):
        super().__init__(container)
        self._config = config
        self._rtr_data = rtr_data
        self._officials_list = StringVar(value=self._config.get_str("officials_list"))

         # self is a vertical container
        self.columnconfigure(0, weight=1)
        self.columnconfigure(1, weight=0)
        self.rowconfigure(0, weight=1)
        
        frlabel = ctk.CTkLabel(self, text="RTR Data Loading", font=ctk.CTkFont(weight="bold"))
        frlabel.grid(column=0, row=0, columnspan=2)

        self.rtrbtn = ctk.CTkButton(self, text="RTR List", command=self._handle_officials_browse)
        self.rtrbtn.grid(column=0, row=2, padx=20, pady=10)
        ToolTip(self.rtrbtn, text="Select the RTR officials export file")   # pylint: disable=C0330
        ctk.CTkLabel(self, textvariable=self._officials_list).grid(column=1, row=2, sticky="ew")

        self.load_btn = ctk.CTkButton(self, text="Load Datafile", command=self._handle_load_btn)
        self.load_btn.grid(column=0, row=4, sticky="news", padx=20, pady=10)
        ctk.CTkLabel(self, text="Load the RTR Officials Datafile").grid(column=1, row=4, sticky="w")

        self.reset_btn = ctk.CTkButton(self, text="Reset", command=self._handle_reset_btn)
        self.reset_btn.grid(column=0, row=6, sticky="news", padx=20, pady=10)
        ctk.CTkLabel(self, text="Restart data loading").grid(column=1, row=6, sticky="w")

        ctk.CTkLabel(self, text="Statistics", font=ctk.CTkFont(weight="bold")).grid(column=0, row=12, columnspan=2, sticky="news")
        
        ctk.CTkLabel(self, text="Officials:  ").grid(column=0, row=14, sticky="e")
        ctk.CTkLabel(self, textvariable=self._rtr_data.total_officials).grid(column=1, row=14, sticky="w")

        ctk.CTkLabel(self, text="Clubs:  ").grid(column=0, row=15, sticky="e")
        ctk.CTkLabel(self, textvariable=self._rtr_data.total_clubs).grid(column=1, row=15, sticky="w")

        ctk.CTkLabel(self, text="Affiliated Officials:  ").grid(column=0, row=16, sticky="e")
        ctk.CTkLabel(self, textvariable=self._rtr_data.total_affilated_officials).grid(column=1, row=16, sticky="w")

        ctk.CTkLabel(self, text="Active:  ").grid(column=0, row=17, sticky="e")
        ctk.CTkLabel(self, textvariable=self._rtr_data.total_active).grid(column=1, row=17, sticky="w")

        ctk.CTkLabel(self, text="PSO Pending:  ").grid(column=0, row=18, sticky="e")
        ctk.CTkLabel(self, textvariable=self._rtr_data.total_pso_pending).grid(column=1, row=18, sticky="w")

        ctk.CTkLabel(self, text="Invoice Pending:  ").grid(column=0, row=19, sticky="e")
        ctk.CTkLabel(self, textvariable=self._rtr_data.total_inv_pending).grid(column=1, row=19, sticky="w")

        ctk.CTkLabel(self, text="Account Pending:  ").grid(column=0, row=20, sticky="e")
        ctk.CTkLabel(self, textvariable=self._rtr_data.total_account_pending).grid(column=1, row=20, sticky="w")


    def _handle_officials_browse(self) -> None:
        directory = filedialog.askopenfilename()
        if len(directory) == 0:
            return
        self._config.set_str("officials_list", directory)
        self._officials_list.set(directory)

    def buttons(self, newstate) -> None:
        '''Enable/disable all buttons on the UI'''
        self.load_btn.configure(state = newstate)
        self.reset_btn.configure(state = newstate)
#        self.reports_btn.configure(state = newstate)
#        self.cohost_btn.configure(state = newstate)

    def _handle_load_btn(self) -> None:
        self.buttons("disabled")
        load_thread = _Data_Loader(self._config)
        load_thread.start()
        self.monitor_load_thread(load_thread)
        self.buttons("enabled")

    def _handle_reset_btn(self) -> None:
        self.buttons("disabled")
        self._rtr_data.reset_data()
        self.buttons("enabled")

    def monitor_load_thread(self, thread):
        '''Monitor the loading thread'''
        if thread.is_alive():
            # check the thread every 100ms 
            self.after(100, lambda: self.monitor_load_thread(thread))
        else:
            # Retrieve data from the loading process and merge it with already loaded data
            if not thread.rtr_data.empty:
                self._rtr_data.load_rtr_data(thread.rtr_data)
            thread.join()

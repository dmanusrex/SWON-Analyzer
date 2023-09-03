# DocGen - https://github.com/dmanusrex/SWON-Analyzer
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


'''  Experimental Module - Generate new pathway migration documents '''

import os
import pandas as pd
import numpy as np
import logging
import customtkinter as ctk
import CTkMessagebox
from CTkMessagebox import CTkMessagebox
import keyring
import webbrowser
import tkinter as tk
from tkinter import filedialog, ttk, BooleanVar, StringVar,  HORIZONTAL
from typing import Any
from tooltip import ToolTip
import keyring
from slugify import slugify
from docx import Document
import docx
from docxcompose.composer import Composer
from threading import Thread

import smtplib, ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime
from typing import List

# Appliction Specific Imports
from config import AnalyzerConfig
from rtr import RTR

tkContainer = Any

class Pathway_Documents_Frame(ctk.CTkFrame):   # pylint: disable=too-many-ancestors
    '''Generate Word Documents from a supplied RTR file'''
    def __init__(self, container: tkContainer, config: AnalyzerConfig, rtr: RTR):
        super().__init__(container)
        self._config = config
        self._rtr = rtr

        # Get the needed options
        self._officials_list = StringVar(value=self._config.get_str("officials_list"))
        self._officials_list_filename = StringVar(value=os.path.basename(self._officials_list.get()))
        self._np_report_directory = StringVar(value=self._config.get_str("np_report_directory"))
        self._np_report_file = StringVar(value=self._config.get_str("np_report_file_docx"))
        self._ctk_theme = StringVar(value=self._config.get_str("Theme"))
        self._ctk_size = StringVar(value=self._config.get_str("Scaling"))
        self._ctk_colour = StringVar(value=self._config.get_str("Colour"))
        self._incl_inv_pending = BooleanVar(value=self._config.get_bool("incl_inv_pending"))
        self._incl_pso_pending = BooleanVar(value=self._config.get_bool("incl_pso_pending"))
        self._incl_account_pending = BooleanVar(value=self._config.get_bool("incl_account_pending"))

        # Add support for the club selection
        self._club_list = ['None']
        self._club_selected = ctk.StringVar(value='None')

        # self is a vertical container that will contain 3 frames
        self.columnconfigure(0, weight=1)
        filesframe = ctk.CTkFrame(self)
        filesframe.grid(column=0, row=0, sticky="news")
        filesframe.rowconfigure(0, weight=1)
        filesframe.rowconfigure(1, weight=1)
        filesframe.rowconfigure(2, weight=1)

        optionsframe = ctk.CTkFrame(self)
        optionsframe.grid(column=0, row=2, sticky="news")

        buttonsframe = ctk.CTkFrame(self)
        buttonsframe.grid(column=0, row=4, sticky="news")
        buttonsframe.rowconfigure(0, weight=0)

        # Files Section
        ctk.CTkLabel(filesframe,
            text="Files and Directories").grid(column=0, row=0, sticky="w", padx=10)   # pylint: disable=C0330

        btn2 = ctk.CTkButton(filesframe, text="Pathway Docs Folder", command=self._handle_report_dir_browse)
        btn2.grid(column=0, row=1, padx=20, pady=10)
        ToolTip(btn2, text="Select where output files will be sent")   # pylint: disable=C0330
        ctk.CTkLabel(filesframe, textvariable=self._np_report_directory).grid(column=1, row=1, sticky="w")

        btn3 = ctk.CTkButton(filesframe, text="Consolidated Report File", command=self._handle_report_file_browse)
        btn3.grid(column=0, row=2, padx=20, pady=10)
        ToolTip(btn3, text="Set report file name")   # pylint: disable=C0330
        ctk.CTkLabel(filesframe, textvariable=self._np_report_file).grid(column=1, row=2, sticky="w")

        # Options Frame - Left and Right Panels

        left_optionsframe = ctk.CTkFrame(optionsframe)
        left_optionsframe.grid(column=0, row=0, sticky="news", padx=10, pady=10)
        left_optionsframe.rowconfigure(0, weight=1)
        right_optionsframe = ctk.CTkFrame(optionsframe)
        right_optionsframe.grid(column=1, row=0, sticky="news", padx=10, pady=10)
        right_optionsframe.rowconfigure(0, weight=1)
        lower_optionsframe = ctk.CTkFrame(optionsframe)
        lower_optionsframe.grid(column=0, row=1, columnspan=2, sticky="news", padx=10, pady=10)

        # Program Options on the left frame

        ctk.CTkLabel(left_optionsframe,
            text="UI Appearance").grid(column=0, row=0, sticky="w", padx=10)   # pylint: disable=C0330

        ctk.CTkLabel(left_optionsframe, text="Appearance Mode", anchor="w").grid(row=1, column=1, sticky="w")
        ctk.CTkOptionMenu(left_optionsframe, values=["Light", "Dark", "System"],
           command=self.change_appearance_mode_event, variable=self._ctk_theme).grid(row=1, column=0, padx=20, pady=10) # pylint: disable=C0330

        ctk.CTkLabel(left_optionsframe, text="UI Scaling", anchor="w").grid(row=2, column=1, sticky="w")
        ctk.CTkOptionMenu(left_optionsframe, values=["80%", "90%", "100%", "110%", "120%"],
           command=self.change_scaling_event, variable=self._ctk_size).grid(row=2, column=0, padx=20, pady=10) # pylint: disable=C0330

        ctk.CTkLabel(left_optionsframe, text="Colour (Restart Required)", anchor="w").grid(row=3, column=1, sticky="w")
        ctk.CTkOptionMenu(left_optionsframe, values=["blue", "green", "dark-blue"],
           command=self.change_colour_event, variable=self._ctk_colour).grid(row=3, column=0, padx=20, pady=10) # pylint: disable=C0330


        # Right options frame for status options
        ctk.CTkLabel(right_optionsframe,
            text="RTR Officials Status").grid(column=0, row=0, sticky="w", padx=10)   # pylint: disable=C0330

        ctk.CTkSwitch(right_optionsframe, text = "PSO Pending", variable=self._incl_pso_pending, onvalue = True, offvalue=False,
            command=self._handle_incl_pso_pending).grid(column=0, row=1, sticky="w", padx=20, pady=10) # pylint: disable=C0330

        ctk.CTkSwitch(right_optionsframe, text = "Account Pending", variable=self._incl_account_pending, onvalue = True, offvalue=False,
            command=self._handle_incl_account_pending).grid(column=0, row=2, sticky="w", padx=20, pady=10) # pylint: disable=C0330

        ctk.CTkSwitch(right_optionsframe, text = "Invoice Pending", variable=self._incl_inv_pending, onvalue = True, offvalue=False,
               command=self._handle_incl_inv_pending).grid(column=0, row=3, sticky="w", padx=20, pady=10) # pylint: disable=C0330

        # Lower options frame for club selection

        self.club_dropdown = ctk.CTkOptionMenu(lower_optionsframe, dynamic_resizing=True,
                                                        values=self._club_list, variable=self._club_selected)
        self.club_dropdown.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="w")
        ctk.CTkLabel(lower_optionsframe, text="Club", anchor="w").grid(row=0, column=1, sticky="w", padx=20, pady=(20, 10))

        # Add Command Buttons

        ctk.CTkLabel(buttonsframe,
            text="Actions").grid(column=0, row=0, sticky="w", padx=10)   # pylint: disable=C0330

        self.reports_btn = ctk.CTkButton(buttonsframe, text="Generate Reports", command=self._handle_reports_btn)
        self.reports_btn.grid(column=0, row=1, sticky="news", padx=20, pady=10)

        self.bar = ctk.CTkProgressBar(master=buttonsframe, orientation='horizontal', mode='indeterminate')

        # Register Callback
        self._rtr.register_update_callback(self.refresh_club_list)

    def refresh_club_list(self):
        self._club_list = ['None']
        self._club_list = self._club_list + [club[1] for club in self._rtr.club_list_names]

        self.club_dropdown.configure(values=self._club_list)
        if len(self._club_list) == 2:    # There is only one club loaded, change default
            self.club_dropdown.set(self._club_list[1])
        else:
            self.club_dropdown.set(self._club_list[0])
        logging.info("Club List Refreshed")

    def _handle_report_dir_browse(self) -> None:
        directory = filedialog.askdirectory()
        if len(directory) == 0:
            return
        directory = os.path.normpath(directory)
        self._config.set_str("np_report_directory", directory)
        self._np_report_directory.set(directory)

    def _handle_report_file_browse(self) -> None:
        report_file = filedialog.asksaveasfilename( filetypes = [('Word Documents','*.docx')], defaultextension=".docx", title="Report File", 
                                                initialfile=os.path.basename(self._np_report_file.get()), # pylint: disable=C0330
                                                initialdir=self._config.get_str("np_report_directory")) # pylint: disable=C0330
        if len(report_file) == 0:
            return
        self._config.set_str("np_report_file_docx", report_file)
        self._np_report_file.set(report_file)

    def _handle_incl_pso_pending(self, *_arg) -> None:
        self._config.set_bool("incl_pso_pending", self._incl_pso_pending.get())

    def _handle_incl_account_pending(self, *_arg) -> None:
        self._config.set_bool("incl_account_pending", self._incl_account_pending.get())

    def _handle_incl_inv_pending(self, *_arg) -> None:
        self._config.set_bool("incl_inv_pending", self._incl_inv_pending.get())

    def change_appearance_mode_event(self, new_appearance_mode: str):
        ctk.set_appearance_mode(new_appearance_mode)
        self._config.set_str("Theme", new_appearance_mode)

    def change_scaling_event(self, new_scaling: str) -> None:
        new_scaling_float = int(new_scaling.replace("%", "")) / 100
        ctk.set_widget_scaling(new_scaling_float)
        self._config.set_str("Scaling", new_scaling)

    def change_colour_event(self, new_colour: str) -> None:
        logging.info("Changing colour to : " + new_colour)
        ctk.set_default_color_theme(new_colour)
        self._config.set_str("Colour", new_colour)

    def buttons(self, newstate) -> None:
        '''Enable/disable all buttons'''
        self.reports_btn.configure(state = newstate)

    def _handle_reports_btn(self) -> None:
        if self._rtr.rtr_data.empty:
            logging.info ("Load data first...")
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
        reports_thread = Generate_Reports(self._rtr, self._config, club)
        reports_thread.start()
        self.monitor_reports_thread(reports_thread)

    def monitor_reports_thread(self, thread):
        if thread.is_alive():
            # check the thread every 100ms 
            self.after(100, lambda: self.monitor_reports_thread(thread))
        else:
            self.buttons("enabled")
            self.bar.stop()
            self.bar.grid_forget()
            thread.join()

class Pathway_ROR_Frame(ctk.CTkFrame):  # pylint: disable=too-many-ancestors,too-many-instance-attributes
    '''Reports Setting and Generation for RORs and POAs'''
    def __init__(self, container: ctk.CTk, config: AnalyzerConfig, rtr: RTR):
        super().__init__(container) # , text="General Settings", padding=5
        self._config = config
        self._rtr = rtr
        self._report_file = StringVar(value=self._config.get_str("np_ror_file_docx"))
        self._report_csv = StringVar(value=self._config.get_str("np_report_file_csv"))

        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)
        self.rowconfigure(2, weight=1)

        # Setup the sub-frames
        ctk.CTkLabel(self, text="ROR/POA Pathway Reports (Multi-Club)").grid(column=0, row=0, sticky="we", pady=10)

        optionsframe = ctk.CTkFrame(self)
        optionsframe.grid(column=0, row=2, sticky="news", padx=10, pady=10)

        filesframe = ctk.CTkFrame(self)
        filesframe.grid(column=0, row=4, sticky="news", padx=10, pady=10)
        filesframe.rowconfigure(0, weight=1)
        filesframe.rowconfigure(1, weight=1)
        filesframe.rowconfigure(2, weight=1)

        buttonsframe = ctk.CTkFrame(self)
        buttonsframe.grid(column=0, row=6, sticky="news", padx=10, pady=10)
        buttonsframe.rowconfigure(0, weight=1)
        buttonsframe.columnconfigure(0, weight=1)

        # Report Options

        self._gen_np_csv_var = BooleanVar(optionsframe, value=self._config.get_bool("gen_np_csv"))
        ctk.CTkSwitch(optionsframe, text="Create CSV File", variable=self._gen_np_csv_var, onvalue = True, offvalue=False,
            command=self._handle_gen_csv_file).grid(column=1, row=0, sticky="news", padx=20, pady=10) # pylint: disable=C0330

        self._gen_np_warnings_var = BooleanVar(optionsframe, value=self._config.get_bool("gen_np_warnings"))
        ctk.CTkSwitch(optionsframe, text="Generate Warnings Report", variable=self._gen_np_warnings_var, onvalue = True, offvalue=False,
            command=self._handle_gen_warnings).grid(column=1, row=2, sticky="news", padx=20, pady=10) # pylint: disable=C0330
        

        # Report File

        rptbtn = ctk.CTkButton(filesframe, text="Warnings Filename", command=self._handle_report_file_browse)
        rptbtn.grid(column=0, row=0, padx=20, pady=10)
        ToolTip(rptbtn, text="The name of the warnings/exceptions file")   # pylint: disable=C0330
        ctk.CTkLabel(filesframe, textvariable=self._report_file).grid(column=1, row=0, sticky="w")

        csvbtn = ctk.CTkButton(filesframe, text="CSV Filename", command=self._handle_csv_browse)
        csvbtn.grid(column=0, row=2, padx=20, pady=10)
        ToolTip(csvbtn, text="The name of the consolidated CSV file")   # pylint: disable=C0330
        ctk.CTkLabel(filesframe, textvariable=self._report_csv).grid(column=1, row=2, sticky="w")

        # Action Button
        
        self.reports_btn = ctk.CTkButton(buttonsframe, text="Generate Reports", command=self._handle_reports_btn)
        self.reports_btn.grid(column=0, row=0, sticky="ew", padx=20, pady=10)

        self.bar = ctk.CTkProgressBar(master=buttonsframe, orientation='horizontal', mode='indeterminate')

    def buttons(self, newstate) -> None:
        '''Enable/disable all buttons on the UI'''
        self.reports_btn.configure(state = newstate)

    def _handle_gen_csv_file(self, *_arg) -> None:
        self._config.set_bool("gen_np_csv", self._gen_np_csv_var.get())

    def _handle_gen_warnings(self, *_arg) -> None:
        self._config.set_bool("gen_np_warnings", self._gen_np_warnings_var.get())

    def _handle_report_file_browse(self) -> None:
        report_file = filedialog.asksaveasfilename( filetypes = [('Word Documents','*.docx')], defaultextension=".docx", title="Warnings/Exceptions File", 
                                                initialfile=os.path.basename(self._report_file.get()))
        if len(report_file) == 0:
            return
        self._config.set_str("np_ror_file_docx", report_file)
        self._report_file.set(report_file)

    def _handle_csv_browse(self) -> None:
        report_file = filedialog.asksaveasfilename( filetypes = [('CSV Files','*.csv')], defaultextension=".csv", title="CSV Filename", 
                                                initialfile=os.path.basename(self._report_file.get()))
        if len(report_file) == 0:
            return
        self._config.set_str("np_report_file_csv", report_file)
        self._report_file.set(report_file)

    def _handle_reports_btn(self) -> None:
        if self._rtr.rtr_data.empty:
            logging.info ("Load data first...")
            CTkMessagebox(master=self, title="Error", message="Load RTR Data First", icon="cancel", corner_radius=0)
            return
        self.buttons("disabled")
        self.bar.grid(row=2, column=0, pady=10, padx=20, sticky="s")
        self.bar.set(0)
        self.bar.start()
        reports_thread = _Generate_NP_ROR_Reports(self._rtr, self._config)
        reports_thread.start()
        self.monitor_reports_thread(reports_thread)

    def monitor_reports_thread(self, thread) -> None:
        if thread.is_alive():
            # check the thread every 100ms 
            self.after(100, lambda: self.monitor_reports_thread(thread))
        else:
            self.bar.stop()
            self.bar.grid_forget()
            self.buttons("enabled")
            thread.join()

class _Generate_NP_ROR_Reports(Thread):
    def __init__(self, rtr: RTR, config: AnalyzerConfig):
        super().__init__()
        self._rtr : pd.DataFrame = rtr.rtr_data
        self._config : AnalyzerConfig = config

    def run(self):
        logging.info("Reporting in Progress...")

        self._report_file = self._config.get_str("np_ror_file_docx")
        self._report_csv = self._config.get_str("np_report_file_csv")

        _gen_np_warnings = self._config.get_bool("gen_np_warnings")
        _gen_np_csv = self._config.get_bool("gen_np_csv")

        status_values = ["Active"]
        if self._config.get_bool("incl_inv_pending"):
            status_values.append("Invoice Pending")
        if self._config.get_bool("incl_account_pending"):
            status_values.append("Account Pending")
        if self._config.get_bool("incl_pso_pending"):
            status_values.append("PSO Pending")

        self._rtr_filtered = self._rtr[self._rtr["Status"].isin(status_values)]

        report_time = datetime.now().strftime("%B %d %Y %I:%M%p")

        if _gen_np_csv:
            logging.info("Generating CSV Report")
            key_columns = ['Registration Id', 'First Name','Last Name', 'Club','Region','Province','Status','Current_CertificationLevel',
                    'Intro_Status','ST_Status','IT_Status','JoS_Status','CT_Status','Clerk_Status','MM_Status','Starter_Status','CFJ_Status','ChiefRec_Status','Referee_Status',
                    'Para Swimming eModule','NP_Official','NP_Ref1','NP_Ref2','NP_Starter1','NP_Starter2','NP_MM1','NP_MM2']
            try:
                self._rtr_filtered.to_csv(self._report_csv, columns = key_columns, index=False)
            except Exception as e:
                logging.info("Unable to save CSV file: {}".format(type(e).__name__))
                logging.info("Exception message: {}".format(e))
                CTkMessagebox(title="Error", message="Unable to save CSV file", icon="cancel", corner_radius=0)

            logging.info("CSV Report Complete")


        if _gen_np_warnings:
            doc = Document()
            logging.info("Generating Warnings Report")

#           Place holder for warnings report
        
        try:
            doc.save(self._report_file)
        except Exception as e:
            logging.info("Unable to save full report: {}".format(type(e).__name__))
            logging.info("Exception message: {}".format(e))
            CTkMessagebox(title="Error", message="Unable to save full report file", icon="cancel", corner_radius=0)

        CTkMessagebox(title="Reports", message="Reports complete", icon="check", option_1="OK", corner_radius=0)
        
        logging.info("Reports Complete")

class Email_Pathway_Docs_Frame(ctk.CTkFrame):   # pylint: disable=too-many-ancestors
    '''E-Mail Completed list of Word Documents'''
    def __init__(self, container: tkContainer, config: AnalyzerConfig):
        super().__init__(container)
        self._config = config

        self._np_report_directory = StringVar(value=self._config.get_str("np_report_directory"))
        self._email_list_csv = StringVar(value=self._config.get_str("email_list_csv"))
        self._email_smtp_server = StringVar(value=self._config.get_str("email_smtp_server"))
        self._email_smtp_port = StringVar(value=self._config.get_str("email_smtp_port"))
        self._email_smtp_user = StringVar(value=self._config.get_str("email_smtp_user"))
        self._email_from = StringVar(value=self._config.get_str("email_from"))
        self._email_subject = StringVar(value=self._config.get_str("email_subject"))
        self._email_body = self._config.get_str("email_body")

         # self is a vertical container that will contain 3 frames
        self.columnconfigure(0, weight=1)

        filesframe = ctk.CTkFrame(self)
        filesframe.grid(column=0, row=0, sticky="news")
        filesframe.rowconfigure(0, weight=1)
        filesframe.rowconfigure(1, weight=1)

        optionsframe = ctk.CTkFrame(self)
        optionsframe.grid(column=0, row=2, sticky="news")

        buttonsframe = ctk.CTkFrame(self)
        buttonsframe.grid(column=0, row=4, sticky="news")
        buttonsframe.rowconfigure(0, weight=0)

        # Files Section
        ctk.CTkLabel(filesframe,
            text="E-mail Configuration").grid(column=0, row=0, sticky="w", padx=10)   # pylint: disable=C0330

        # options Section


        entry_width = 500

#        reg_email_smtp_server = self.register(self._handle_email_smtp_server)
#       A registered validation function seems to disable the interactive logging window. Need to investigate

        ctk.CTkLabel(optionsframe, text="SMTP Server", anchor="w").grid(row=1, column=0, sticky="w")

        smtp_server_entry = ctk.CTkEntry(optionsframe, textvariable=self._email_smtp_server, width=entry_width)
        smtp_server_entry.grid(column=1, row=1, sticky="w", padx=10, pady=10)
        smtp_server_entry.bind('<FocusOut>', self._handle_email_smtp_server)

        ctk.CTkLabel(optionsframe, text="SMTP Port", anchor="w").grid(row=2, column=0, sticky="w")
        smtp_port_entry = ctk.CTkEntry(optionsframe, textvariable=self._email_smtp_port, width=entry_width)
        smtp_port_entry.grid(column=1, row=2, sticky="w", padx=10, pady=10)
        smtp_port_entry.bind('<FocusOut>', self._handle_email_smtp_port)
        
        ctk.CTkLabel(optionsframe, text="SMTP Username", anchor="w").grid(row=3, column=0, sticky="w")
        smtp_user_entry = ctk.CTkEntry(optionsframe, textvariable=self._email_smtp_user, width=entry_width)
        smtp_user_entry.grid(column=1, row=3, sticky="w", padx=10, pady=10)
        smtp_user_entry.bind('<FocusOut>', self._handle_email_smtp_user)

        ctk.CTkLabel(optionsframe, text="SMTP Password", anchor="w").grid(row=4, column=0, sticky="w")
        self.password_entry = ctk.CTkEntry(optionsframe, placeholder_text="Password", show="*", width=entry_width)
        self.password_entry.grid(column=1, row=4, sticky="w", padx=10, pady=10)
        self.password_entry.bind('<FocusOut>', self._handle_email_smtp_password)

        ctk.CTkLabel(optionsframe, text="E-mail From", anchor="w").grid(row=5, column=0, sticky="w")
        email_from_entry = ctk.CTkEntry(optionsframe, textvariable=self._email_from, width=entry_width)
        email_from_entry.grid(column=1, row=5, sticky="w", padx=10, pady=10)
        email_from_entry.bind('<FocusOut>', self._handle_email_from)

        ctk.CTkLabel(optionsframe, text="E-mail Subject", anchor="w").grid(row=6, column=0, sticky="w")
        email_subject_entry = ctk.CTkEntry(optionsframe, textvariable=self._email_subject, width=entry_width)
        email_subject_entry.grid(column=1, row=6, sticky="w", padx=10, pady=10)
        email_subject_entry.bind('<FocusOut>', self._handle_email_subject)

        # Body Text
        ctk.CTkLabel(optionsframe, text="E-mail Body", anchor="w").grid(row=7, column=0, sticky="w")

        self.txtbodybox = ctk.CTkTextbox(master=optionsframe, state='normal', width=entry_width)
        self.txtbodybox.grid(column=1, row=7, sticky="w", padx=10, pady=10)
        self.txtbodybox.insert(tk.END, self._email_body)
        self.txtbodybox.bind('<FocusOut>', self._handle_email_body)

        # Add Command Buttons

        ctk.CTkLabel(buttonsframe,
            text="Actions").grid(column=0, row=0, sticky="w", padx=10)   # pylint: disable=C0330

        self.emailtest_btn = ctk.CTkButton(buttonsframe, text="Send Test EMails", command=self._handle_email_test_btn)
        self.emailtest_btn.grid(column=0, row=1, sticky="news", padx=20, pady=10)
        self.emailall_btn = ctk.CTkButton(buttonsframe, text="Send All Emails", command=self._handle_email_all_btn)
        self.emailall_btn.grid(column=1, row=1, sticky="news", padx=20, pady=10)

        self.bar = ctk.CTkProgressBar(master=self, orientation='horizontal', mode='indeterminate')


    def _handle_report_dir_browse(self) -> None:
        directory = filedialog.askdirectory()
        if len(directory) == 0:
            return
        directory = os.path.normpath(directory)
        self._config.set_str("np_report_directory", directory)
        self._np_report_directory.set(directory)

    def _handle_email_smtp_server(self, event) -> bool:
        self._config.set_str("email_smtp_server", event.widget.get())
        return True

    def _handle_email_smtp_port(self, event) -> bool:
        self._config.set_str("email_smtp_port", event.widget.get())
        return True
    
    def _handle_email_smtp_user(self, event) -> bool:
        self._config.set_str("email_smtp_user", event.widget.get())
        self.password_entry.delete(0, tk.END)
        return True

    def _handle_email_smtp_password(self, event) -> bool:
        if event.widget.get() != "Password":
            keyring.set_password("SWON-DOCGEN", self._email_smtp_user.get(), event.widget.get())
            logging.info("Password Changed for %s" % self._email_smtp_user.get())
        return True
    
    def _handle_email_from(self, event) -> bool:
        self._config.set_str("email_from", event.widget.get())
        return True
    
    def _handle_email_subject(self, event) -> bool:
        self._config.set_str("email_subject", event.widget.get())
        return True
    
    def _handle_email_body(self, event) -> bool:
        self._config.set_str("email_body", event.widget.get("0.0", "end"))
        return True

    def buttons(self, newstate) -> None:
        '''Enable/disable all buttons'''
        self.emailtest_btn.configure(state = newstate)
        self.emailall_btn.configure(state = newstate)

    def _handle_email_test_btn(self) -> None:
        self.buttons("disabled")
        email_thread = Email_Reports(True, self._config)
        email_thread.start()
        self.monitor_email_thread(email_thread)

    def _handle_email_all_btn(self) -> None:
        self.buttons("disabled")
        email_thread = Email_Reports(False, self._config)
        email_thread.start()
        self.monitor_email_thread(email_thread)

    def monitor_email_thread(self, thread):
        if thread.is_alive():
            # check the thread every 100ms 
            self.after(100, lambda: self.monitor_email_thread(thread))
        else:
            self.buttons("enabled")
            thread.join()




class NewPathway:

    def __init__(self, club: str, club_data_set : pd.DataFrame, config: AnalyzerConfig, **kwargs):
        self._club_data = club_data_set.copy()
        self.club_code = club
        self._config = config

    def _is_valid_date(self, date_string) -> bool:
        if pd.isnull(date_string): return False
        if date_string == "0001-01-01": return False 
        try:
            datetime.strptime(date_string, '%Y-%m-%d')
            return True
        except ValueError:
            return False


    def _get_date(self, date_string) -> str: 
        if pd.isnull(date_string): return ""
        if date_string == "0001-01-01": return "" 
        return date_string
  
    def _count_signoffs(self, clinic_date_1, clinic_date_2) -> int:
        count = 0
        if self._is_valid_date(clinic_date_1): count += 1
        if self._is_valid_date(clinic_date_2): count += 1
        return count
    
    def add_clinic(self, table, clinic_name, clinic_date, signoff_1, signoff_2) -> None:
        row = table.add_row().cells
        row[0].text = clinic_name
        row[1].text = self._get_date(clinic_date)
        row[2].text = self._get_date(signoff_1)
        row[3].text = self._get_date(signoff_2)
        row[0].paragraphs[0].alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.LEFT
        row[1].paragraphs[0].alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
        row[2].paragraphs[0].alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
        row[3].paragraphs[0].alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER

    def add_pathway(self, table, pathway_progression, certified) -> None:
        row = table.add_row().cells
        row[0].text = pathway_progression
        row[1].text = certified 
        row[0].paragraphs[0].alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.LEFT
        row[1].paragraphs[0].alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER

    def export_csv(self, filename: str) -> None:
        '''Export the club data to a CSV file'''

        key_columns = ['Registration Id', 'First Name','Last Name', 'Club','Region','Province','Status','Current_CertificationLevel',
                   'Intro_Status','ST_Status','IT_Status','JoS_Status','CT_Status','Clerk_Status','MM_Status','Starter_Status','CFJ_Status','ChiefRec_Status','Referee_Status',
                   'Para Swimming eModule','NP_Official','NP_Ref1','NP_Ref2','NP_Starter1','NP_Starter2','NP_MM1','NP_MM2']
        try:
            self._club_data.to_csv(filename, columns = key_columns, index=False)
        except Exception as e:
            logging.info("Unable to save CSV file: {}".format(type(e).__name__))
            logging.info("Exception message: {}".format(e))
            CTkMessagebox(title="Error", message="Unable to save CSV file", icon="cancel", corner_radius=0)

    def dump_data_docx(self, club_fullname: str, reportdate: str) -> List:
        '''Produce the Word Document for the club and return a list of files'''
 
        _report_directory = self._config.get_str("np_report_directory")
        _email_list_csv = self._config.get_str("email_list_csv")
        csv_list = []    # CSV entries for email list (Lastname, Firstname, E-Mail address and Filename)
 
        for index, entry in self._club_data.iterrows():

            # create a filename from the last and firstnames using slugify and the report directory

            filename = os.path.abspath(os.path.join(_report_directory, slugify(entry["Last Name"] + "_" + entry["First Name"]) + ".docx"))
            csv_list.append([entry["Last Name"], entry["First Name"], entry["Email"], filename])

            doc = Document()

            doc.add_heading("2023/24 New Pathway Mapping", 0)
 
            p = doc.add_paragraph()
            p.add_run("Report Date: "+reportdate)
            p.add_run("\n\nName: "+ entry["Last Name"] + ", " + entry["First Name"] + " (SNC ID # " + entry["Registration Id"] + ")")
            p.add_run("\n\nClub: "+ club_fullname + " (" + self.club_code + ")")
            p.add_run("\n\nCurrent Certification Level: ")
            p.add_run("NONE" if pd.isnull(entry["Current_CertificationLevel"]) else entry["Current_CertificationLevel"])



            table = doc.add_table(rows=1, cols=4)
            row = table.rows[0].cells
            row[0].text = "Clinic"
            row[1].text = "Clinic Date"
            row[2].text = "Sign Off #1"
            row[3].text = "Sign Off #2" 
            row[0].paragraphs[0].alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.LEFT
            row[1].paragraphs[0].alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
            row[2].paragraphs[0].alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
            row[3].paragraphs[0].alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER


            self.add_clinic(table, "Intro to Swimming", entry["Introduction to Swimming Officiating-ClinicDate"], entry["Introduction to Swimming Officiating-Deck Evaluation #1 Date"], entry["Introduction to Swimming Officiating-Deck Evaluation #2 Date"])
            self.add_clinic(table, "Safety Marshal", entry["Safety Marshal-ClinicDate"], "N/A", "N/A")
            self.add_clinic(table, "Stroke & Turn", entry["Judge of Stroke/Inspector of Turns-ClinicDate"], entry["Judge of Stroke/Inspector of Turns-Deck Evaluation #1 Date"], entry["Judge of Stroke/Inspector of Turns-Deck Evaluation #2 Date"])
            self.add_clinic(table, "Inspector of Turns", entry("Inspector of Turns-ClinicDate"), entry["Inspector of Turns-Deck Evaluation #1 Date"], entry["Inspector of Turns-Deck Evaluation #2 Date"])
            self.add_clinic(table, "Judge of Stroke", entry["Judge of Stroke-ClinicDate"], entry["Judge of Stroke-Deck Evaluation #1 Date"], entry["Judge of Stroke-Deck Evaluation #2 Date"])
            self.add_clinic(table, "Chief Timekeeper", entry["Chief Timekeeper-ClinicDate"], entry["Chief Timekeeper-Deck Evaluation #1 Date"], entry["Chief Timekeeper-Deck Evaluation #2 Date"])
            self.add_clinic(table, "Admin Desk (Clerk)", entry["Clerk of Course-ClinicDate"], entry["Clerk of Course-Deck Evaluation #1 Date"], entry["Clerk of Course-Deck Evaluation #2 Date"])
            self.add_clinic(table, "Meet Manager", entry["Meet Manager-ClinicDate"], entry["Meet Manager-Deck Evaluation #1 Date"], entry["Meet Manager-Deck Evaluation #2 Date"])
            self.add_clinic(table, "Starter", entry["Starter-ClinicDate"], entry["Starter-Deck Evaluation #1 Date"], entry["Starter-Deck Evaluation #2 Date"])
            self.add_clinic(table, "CFJ/CJE", entry["Chief Finish Judge/Chief Judge-ClinicDate"], entry["Chief Finish Judge/Chief Judge-Deck Evaluation #1 Date"], entry["Chief Finish Judge/Chief Judge-Deck Evaluation #2 Date"])
            self.add_clinic(table, "Chief Recorder/Recorder", entry["Recorder-Scorer-ClinicDate"], "N/A", "N/A")
            self.add_clinic(table, "Referee", entry["Referee-ClinicDate"], "N/A", "N/A")
            self.add_clinic(table, "Para eModule", entry["Para Swimming eModule-ClinicDate"], "N/A", "N/A")
            
            table.style = "Light Grid Accent 5"
            table.autofit = True

            # Add logic to define pathway progression

            doc.add_heading("New Pathway Progression", 2)
            nptable = doc.add_table(rows=1, cols=4)
            row = nptable.rows[0].cells
            row[0].text = "New Pathway Level"
            row[1].text = "Certifed?"
            row[0].paragraphs[0].alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.LEFT
            row[1].paragraphs[0].alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER

            self.add_pathway(nptable, "Certified Official", entry["NP_Official"])
            self.add_pathway(nptable, "Referee 1", entry["NP_Ref1"])
            self.add_pathway(nptable, "Referee 2", entry["NP_Ref2"])
            self.add_pathway(nptable, "Starter 1", entry["NP_Starter1"])
            self.add_pathway(nptable, "Starter 2", entry["NP_Starter2"])
            self.add_pathway(nptable, "Meet Manager 1", entry["NP_MM1"])
            self.add_pathway(nptable, "Meet Manager 2", entry["NP_MM2"])

            nptable.style = "Light Grid Accent 5"
            nptable.autofit = True


            # Recommendations to be added here

            try:
                doc.save(filename)

            except Exception as e:
                logging.info(f'Error processing offiical {entry["Last Name"]}, {entry["First Name"]}: {type(e).__name__} - {e}')

        return csv_list


class Generate_Reports(Thread):
    def __init__(self, rtr: RTR, config: AnalyzerConfig, selected_club: str):
        super().__init__()
        self._rtr = rtr
        self._df : pd.DataFrame = self._rtr.rtr_data
        self._config : AnalyzerConfig = config
        self._selected_club = selected_club

    def run(self):
        # This still has code elements to run multiple clubs. For single clubs a simple filter has been applied.
        # Once the final specifications are set, this can be greatly simplified.

        logging.info("Reporting in Progress...")

        _report_directory = self._config.get_str("np_report_directory")
        _report_file_docx = self._config.get_str("np_report_file_docx")
        _full_report_file = os.path.abspath(os.path.join(_report_directory, _report_file_docx))
        _email_list_csv = self._config.get_str("email_list_csv")
        _full_csv_file = os.path.abspath(os.path.join(_report_directory, _email_list_csv))

        club_list_names = self._rtr.club_list_names

        club_summaries = []

        status_values = ["Active"]
        if self._config.get_bool("incl_inv_pending"):
            status_values.append("Invoice Pending")
        if self._config.get_bool("incl_account_pending"):
            status_values.append("Account Pending")
        if self._config.get_bool("incl_pso_pending"):
            status_values.append("PSO Pending")

        report_time = datetime.now().strftime("%B %d %Y %I:%M%p")

        all_csv_entries = []

        for club, club_full in filter(lambda x:x[1]==self._selected_club, club_list_names):
            logging.info("Processing %s" % club_full)
            club_data = self._df[(self._df["ClubCode"] == club)]
            club_data = club_data[club_data["Status"].isin(status_values)]
            club_stat = NewPathway(club, club_data, self._config)
            club_csv = club_stat.dump_data_docx(club_full, report_time)
            all_csv_entries.extend(club_csv)
            club_summaries.append ([club, club_full, club_stat])

        # Create the email list CSV file    
        # 
        # The email list is a CSV file with the following columns:
        #  Last Name, First Name, E-Mail address, Filename

        logging.info("Creating email list CSV file")
        email_list_df = pd.DataFrame(all_csv_entries, columns=["Last Name", "First Name", "EMail", "Filename"])

        try:
            email_list_df.to_csv(_full_csv_file, index=False)
        except Exception as e:
            logging.info("Unable to save email list: {}".format(type(e).__name__))
            logging.info("Exception message: {}".format(e))
        
        # Create the master document

        logging.info("Creating master document")

        number_of_sections=len(all_csv_entries)
        master = Document()
        composer = Composer(master)
        for i in range(0, number_of_sections):
            doc_temp = Document(all_csv_entries[i][3])
            doc_temp.add_page_break()
            composer.append(doc_temp)

        try:
            composer.save(_full_report_file)
        except Exception as e:
            logging.info("Unable to save full report: {}".format(type(e).__name__))
            logging.info("Exception message: {}".format(e))

        logging.info("Report Complete")


class Email_Reports(Thread):
    def __init__(self, testmode:bool, config: AnalyzerConfig):
        super().__init__()
        self._testmode : bool = testmode
        self._config : AnalyzerConfig = config
        self._email_password : str = "EMPTY"

        self._email_smtp_server = self._config.get_str("email_smtp_server")
        self._email_smtp_port = self._config.get_str("email_smtp_port")
        self._email_smtp_user = self._config.get_str("email_smtp_user")
        self._email_from = self._config.get_str("email_from")
        self._email_subject = self._config.get_str("email_subject")
        self._email_body = self._config.get_str("email_body")

    def run(self):
        logging.info("Sending E-Mails...")

        _report_directory = self._config.get_str("np_report_directory")
        _email_list_csv = self._config.get_str("email_list_csv")
        _full_csv_file = os.path.abspath(os.path.join(_report_directory, _email_list_csv))

        try:
            self._email_password = keyring.get_password("SWON-ANALZER", self._email_smtp_user)
        except Exception as e:
            logging.info("Unable to retrieve email password: {}".format(type(e).__name__))
            logging.info("Exception message: {}".format(e))
            return
        
        if self._testmode:
            logging.info("Test Mode - Sending max (3) mails to {}".format(self._email_from))

        try:
            email_list_df = pd.read_csv(_full_csv_file)
        except Exception as e:    
            logging.info("Unable to load email list: {}".format(type(e).__name__))
            logging.info("Exception message: {}".format(e))
            return    

        context = ssl.create_default_context()

        if self._email_smtp_port == "465":
            try:
                server = smtplib.SMTP_SSL(self._email_smtp_server, self._email_smtp_port, context=context)
                server.login(self._email_smtp_user, self._email_password)
            except Exception as e:
                logging.info("Unable to connect to email server: {}".format(type(e).__name__))
                logging.info("Exception message: {}".format(e))
                return
        else:  # Port 587
            try:
                server = smtplib.SMTP(self._email_smtp_server, self._email_smtp_port)
                server.starttls(context=context)
                server.login(self._email_smtp_user, self._email_password)
            except Exception as e:
                logging.info("Unable to connect to email server: {}".format(type(e).__name__))
                logging.info("Exception message: {}".format(e))
                return
    
        # For each entry in the list encode the Document and send it.  In test mode, use sender address for to and limit to 5 files

        for index, entry in email_list_df.iterrows():
            if self._testmode and index > 2:
                break
            logging.info(f'Sending email to {entry["Last Name"]}, {entry["First Name"]}  E-Mail: {entry["EMail"]}')

            if self._testmode:
                self._send_email(self._email_from, entry["Filename"], server)
            else:
                self._send_email(entry["EMail"], entry["Filename"], server)
                

        logging.info("Email Complete")
    
    def _send_email(self, email_address: str, filename: str, server) -> None:
        # Create a multipart message and set headers
        message = MIMEMultipart()
        message["From"] = self._email_from
        message["To"] = email_address
        message["Subject"] = self._email_subject

        # Add body to email
        message.attach(MIMEText(self._email_body, "plain"))

        # Open document file in binary mode
        with open(filename, "rb") as attachment:
            # Add file as application/octet-stream
            # Email client can usually download this automatically as attachment
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())

        # Encode file in ASCII characters to send by email
        encoders.encode_base64(part)

        # Add header as key/value pair to attachment part
        basename = os.path.basename(filename)
        part.add_header("Content-Disposition",f"attachment; filename= {basename}",)

        # Add attachment to message and convert message to string
        message.attach(part)
        text = message.as_string()

        # Log in to server using secure context and send email
        try:
            server.sendmail(self._email_from, email_address, text)
        except Exception as e:
            logging.info("Unable to send email: {}".format(type(e).__name__))
            logging.info("Exception message: {}".format(e))

  
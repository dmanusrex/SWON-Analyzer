# Club Analyzer - https://github.com/dmanusrex/SWON-Analyzer
# Copyright (C) 2023 - Darren Richer
#
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

''' Club Summary

Main object that does the heavy lifting.  This will produce the required dataset for a given dataframe.
The dataframe does not have to be an individual club. It can be a combo to assess compliances as well.

General Notes:

The RTR data exported for officials is not a complete data set.  Further, the underlying data has a number
of inconsistencies for historical reasons and bad data validation.   The officials export is a normalization of
a number of SQL server tables. 

The principle goal is to allow for automated analysis to support a number of activities:

1)  Basic statistical information for clubs
2)  Identify data errors for offiicials
3)  Identify "easy wins" for officials to advances for club
4)  Determine what level of meet sanction a club can apply for.

This class performs all of the necessary functions to create the data for those objectives.

TO DO: Re-factor code to be cleaner. The current version is the result of a number of spec changes over a few
  days and is messy

'''

import pandas as pd
import itertools as it
from datetime import datetime
from typing import List
from config import AnalyzerConfig
from copy import deepcopy, copy
from docx import Document
import docx
import logging

class club_summary:
    _Qual_Fields = {
        'Intro': ["Introduction to Swimming Officiating", "Introduction to Swimming Officiating-Deck Evaluation #1 Date", "Introduction to Swimming Officiating-Deck Evaluation #2 Date"],
        'ST': ["Judge of Stroke/Inspector of Turns", "Judge of Stroke/Inspector of Turns-Deck Evaluation #1 Date", "Judge of Stroke/Inspector of Turns-Deck Evaluation #2 Date"],
        'CT': ["Chief Timekeeper", "Chief Timekeeper-Deck Evaluation #1 Date", "Chief Timekeeper-Deck Evaluation #2 Date"],
        'Clerk': ["Clerk of Course", "Clerk of Course-Deck Evaluation #1 Date", "Clerk of Course-Deck Evaluation #2 Date"],
        'MM': ["Meet Manager", "Meet Manager-Deck Evaluation #1 Date", "Meet Manager-Deck Evaluation #2 Date"],
        'Starter': ["Starter", "Starter-Deck Evaluation #1 Date", "Starter-Deck Evaluation #2 Date"],
        'RS': ["Recorder-Scorer", "Recorder-Scorer-Deck Evaluation #1 Date", "Recorder-Scorer-Deck Evaluation #2 Date"],
        'CFJ': ["Chief Finish Judge/Chief Judge", "Chief Finish Judge/Chief Judge-Deck Evaluation #1 Date", "Chief Finish Judge/Chief Judge-Deck Evaluation #2 Date"],
        'Referee': ["Referee", "Referee-Deck Evaluation #1 Date", "Referee-Deck Evaluation #2 Date"]
    }


    def __init__(self, club: str, club_data_set : pd.DataFrame, config: AnalyzerConfig, **kwargs):
        self._club_data_full = club_data_set
        self._club_data = self._club_data_full.query("Current_CertificationLevel not in ['LEVEL IV - GREEN PIN','LEVEL V - BLUE PIN']")
        self.club_code = club
        self.Level_None : int = 0
        self.Level_1s : int = 0
        self.Level_2s : int = 0
        self.Level_3s : int = 0
        self.Level_4s : int = 0
        self.Level_5s : int = 0
        self.Qual_Refs : int = 0
        
        # Each list contains a summary count of the number of offiicals with 0, 1, or 2 certification dates
        self.Intro : List = []
        self.SandT : List = []
        self.ChiefT : List = []
        self.Clerk : List = []
        self.MeetM : List = []
        self.Starter : List = []
        self.CFJ : List = []
        self.RecSec : List = []
        self.Referee: List = []
        self.Qualified_Refs: List = []
        self.Level_4_5s : List = []
        self.Level_3_list : List = []
        self.NoLevel_Missing_Cert : List = []
        self.NoLevel_Missing_SM : List = []
        self.NoLevel_Has_II : List = []
        self.Sanction_Level : List = []

        self.Failed_Sanctions : List = []

        self._count_levels()
        self._count_certifications()
        self._find_qualfied_refs()
        self._find_all_level4_5s()
        self._check_no_levels()
        self._check_sanctions()
        self._config = config


    def _is_valid_date(self, date_string):
        if pd.isnull(date_string): return False
        if date_string == "0001-01-01": return False 
        try:
            datetime.strptime(date_string, '%Y-%m-%d')
            return True
        except ValueError:
            return False
        
    def _count_levels(self):
        self.Level_None = self._club_data.query("Current_CertificationLevel.isnull()").shape[0]
        self.Level_1s = self._club_data.query("Current_CertificationLevel == 'LEVEL I - RED PIN'").shape[0]
        self.Level_2s = self._club_data.query("Current_CertificationLevel == 'LEVEL II - WHITE PIN'").shape[0]
        self.Level_3s = self._club_data.query("Current_CertificationLevel == 'LEVEL III - ORANGE PIN'").shape[0]
        self.Level_4s = self._club_data_full.query("Current_CertificationLevel == 'LEVEL IV - GREEN PIN'").shape[0]
        self.Level_5s = self._club_data_full.query("Current_CertificationLevel == 'LEVEL V - BLUE PIN'").shape[0]


    def _find_qualfied_refs(self):
        # To be a Level III referee you need CT, Clerk, Starter and one of CFJ or MM
        # Also check domestic clinic status
        level3_list = self._club_data.query("Current_CertificationLevel == 'LEVEL III - ORANGE PIN'")
        self.Qualified_Refs = []
        self.Level_3_list= []

        for index, row in level3_list.iterrows():
            ref_name = row["Last Name"] + ", " + row["First Name"]
            self.Level_3_list.append(ref_name)
            if ((row["Referee"].lower() == "yes") and
                self._is_valid_date(row["Chief Timekeeper-Deck Evaluation #2 Date"]) and
                self._is_valid_date(row["Starter-Deck Evaluation #2 Date"]) and
                self._is_valid_date(row["Clerk of Course-Deck Evaluation #2 Date"]) and
                ( self._is_valid_date(row["Chief Finish Judge/Chief Judge-Deck Evaluation #2 Date"]) or
                self._is_valid_date(row["Meet Manager-Deck Evaluation #2 Date"]) ) ):
                    para_dom = row["Para Domestic"]
                    para_emodule = row["Para Swimming eModule"]
                    self.Qualified_Refs.append([ref_name, para_dom, para_emodule])
        self.Qual_Refs = len(self.Qualified_Refs)

    def _find_all_level4_5s(self):
        '''In the RTR Level 4/5s may not have the underlying detail but by definition they must be certified in all positions'''
        level45_list = self._club_data_full.query("Current_CertificationLevel == 'LEVEL IV - GREEN PIN' or Current_CertificationLevel == 'LEVEL V - BLUE PIN'")
        self.Level_4_5s = []

        # Get all their names

        for index, row in level45_list.iterrows():
            L45_name = row["Last Name"] + ", " + row["First Name"]
            self.Level_4_5s.append(L45_name)

        # Add them to all the qualification lists
        if self.Level_4_5s:
            self.ChiefT[3].extend(self.Level_4_5s)
            self.ChiefT[4].extend(self.Level_4_5s)
            self.MeetM[3].extend(self.Level_4_5s)
            self.MeetM[4].extend(self.Level_4_5s)
            self.Clerk[3].extend(self.Level_4_5s)
            self.Clerk[4].extend(self.Level_4_5s)
            self.Starter[3].extend(self.Level_4_5s)
            self.Starter[4].extend(self.Level_4_5s)
            self.CFJ[3].extend(self.Level_4_5s)
            self.CFJ[4].extend(self.Level_4_5s)
            self.SandT[3].extend(self.Level_4_5s)
            self.SandT[4].extend(self.Level_4_5s)

    def _check_no_levels(self):
        no_level_list = self._club_data.query("Current_CertificationLevel.isnull()")
        has_both = []
        has_intro_only = []
        has_level_ii = []

        for index, row in no_level_list.iterrows():
            official_name = row["Last Name"] + ", " + row["First Name"]
            if row["Introduction to Swimming Officiating"].lower() == "yes":
                if row["Safety Marshal"].lower() == "yes":
                    has_both.append(official_name)
                else:
                    has_intro_only.append(official_name)
            if (row["Judge of Stroke/Inspector of Turns"].lower() == "yes" or
               row["Chief Timekeeper"].lower() == "yes" or
               row["Clerk of Course"].lower() == "yes" or
               row["Starter"].lower() == "yes" or
               row["Chief Finish Judge/Chief Judge"].lower() == "yes" or
               row["Meet Manager"].lower() == "yes"):
                has_level_ii.append(official_name)

        self.NoLevel_Missing_Cert = has_both
        self.NoLevel_Missing_SM = has_intro_only
        self.NoLevel_Has_II = has_level_ii
        
    def _count_certifications_detail(self, cert_name, cert_date_1, cert_date_2):
        cert_total = 0  # Count of clinics taken
        cert_1so = 0    # Has 1 Sign-Off
        cert_2so = 0    # Has 2 Sign-Offs (fully qualified)
        qual_list = []  # List of officials fully qualified
        cert_list = []  # List of offiicals certified (excludes full qualified officials)

        for index, row in self._club_data.iterrows():
            if row[cert_name].lower() == "yes":
                cert_total += 1
                qual_list.append(row["Last Name"]+", "+row["First Name"])
                if self._is_valid_date(row[cert_date_1]) and self._is_valid_date(row[cert_date_2]):
                    cert_2so += 1
                    cert_list.append(row["Last Name"]+", "+row["First Name"])
                elif self._is_valid_date(row[cert_date_1]) or self._is_valid_date(row[cert_date_2]):
                    cert_1so += 1

        return [cert_total, cert_1so, cert_2so, qual_list, cert_list]
    
    def _count_certifications(self):
        self.Intro = self._count_certifications_detail("Introduction to Swimming Officiating", "Introduction to Swimming Officiating-Deck Evaluation #1 Date", "Introduction to Swimming Officiating-Deck Evaluation #2 Date")
        self.SandT = self._count_certifications_detail("Judge of Stroke/Inspector of Turns", "Judge of Stroke/Inspector of Turns-Deck Evaluation #1 Date", "Judge of Stroke/Inspector of Turns-Deck Evaluation #2 Date")
        self.ChiefT = self._count_certifications_detail("Chief Timekeeper", "Chief Timekeeper-Deck Evaluation #1 Date", "Chief Timekeeper-Deck Evaluation #2 Date")
        self.Clerk = self._count_certifications_detail("Clerk of Course", "Clerk of Course-Deck Evaluation #1 Date", "Clerk of Course-Deck Evaluation #2 Date")
        self.MeetM = self._count_certifications_detail("Meet Manager", "Meet Manager-Deck Evaluation #1 Date", "Meet Manager-Deck Evaluation #2 Date")
        self.Starter = self._count_certifications_detail("Starter", "Starter-Deck Evaluation #1 Date", "Starter-Deck Evaluation #2 Date")
        self.RecSec = self._count_certifications_detail("Recorder-Scorer", "Recorder-Scorer-Deck Evaluation #1 Date", "Recorder-Scorer-Deck Evaluation #2 Date")
        self.CFJ = self._count_certifications_detail("Chief Finish Judge/Chief Judge", "Chief Finish Judge/Chief Judge-Deck Evaluation #1 Date", "Chief Finish Judge/Chief Judge-Deck Evaluation #2 Date")
        self.Referee = self._count_certifications_detail("Referee", "Referee-Deck Evaluation #1 Date", "Referee-Deck Evaluation #2 Date")

    
    def _check_sanctions(self):
        approved_sanctions = []

        '''Sanctioning Parameters to pass:

        # of Level 4/5s needed
        # of Level 3 Referees Need
        # of Level 3s needed
        # of CTs needed (qualified, certified)
        # of MMs needed (qualified, certified)
        # of Clerks needed (qualified, certified)
        # of Starters needed (qualified, certified)
        # of CFJs needed (qualified, certified)
        # of Stoke & Turn needed (qualified, certified)

        ''' 

        T1Opt1 = self._check_sanctions_detail(1,0,0,1,0,1,0,0,0,1,0,1,0,4,0, "TIER I - A")
        T1Opt2 = self._check_sanctions_detail(0,1,0,1,0,0,1,0,0,1,0,1,0,3,1, "TIER I - B")

        if T1Opt1 or T1Opt2: 
            opts = "TIER I - Class II Time Trial + In-House Competition (Option(s):"+ (" A" if T1Opt1 else "") + (" B" if T1Opt2 else "") + ")"
            approved_sanctions.append(opts)

        T2Opt1 = self._check_sanctions_detail(1,1,0,2,0,1,0,1,0,1,0,1,0,4,2, "TIER II - A")
        T2Opt2 = self._check_sanctions_detail(1,0,1,1,1,1,0,1,0,1,0,1,0,4,2, "TIER II - B")
        T2Opt3 = self._check_sanctions_detail(0,0,1,1,1,0,1,1,0,1,0,1,0,4,2, "TIER II - C")

        if T2Opt1 or T2Opt2 or T2Opt3:
            opts = "TIER II - Closed Invitational (limited to 4 sessions) (Option(s): "+ (" A" if T2Opt1 else "") + (" B" if T2Opt2 else "") + (" C" if T2Opt3 else "") + ")"
            approved_sanctions.append(opts)

        T3Opt1 = self._check_sanctions_detail(1,1,1,2,0,1,0,1,0,1,0,1,0,6,2, "TIER III - A")
        T3Opt2 = self._check_sanctions_detail(1,0,1,1,1,0,1,1,0,1,1,1,0,6,2, "TIER III - B")

        if T3Opt1 or T3Opt2: 
            opts = "TIER III - Open Invitational (limited to 6 sessions, no standards) (Option(s):"+ (" A" if T3Opt1 else "") + (" B" if T3Opt2 else "") + ")"
            approved_sanctions.append(opts)

        T4Opt1 = self._check_sanctions_detail(2,0,1,1,1,0,1,1,1,1,1,1,1,8,4, "TIER IV - A")
        T4Opt2 = self._check_sanctions_detail(1,1,2,0,2,0,1,1,1,1,1,1,1,8,4, "TIER IV - B")

        if T4Opt1 or T4Opt2: 
            opts = "TIER IV - Open or Closed Invitational + Regionals & Provincials (no session limits, any double-ended meet) (Option(s):"+ (" A" if T4Opt1 else "") + (" B" if T4Opt2 else "") + ")"
            approved_sanctions.append(opts)

        self.Sanction_Level = approved_sanctions

    def _check_sanctions_detail(self, Level4_5 : int, Qual_Ref: int, Level3: int, 
                                 Qual_CT: int, Cert_CT: int,
                                 Qual_MM: int, Cert_MM: int,
                                 Qual_Clerk: int, Cert_Clerk: int,
                                 Qual_Starter: int, Cert_Starter: int,
                                 Qual_CFJ: int, Cert_CFJ: int,
                                 Qual_SandT: int, Cert_SandT: int,
                                 dbg_scenario_name) -> dict: 
        '''build and test sanction application - returns a dictionary of a valid staffing result if found'''
        my_scenario = self._build_staffing_scenario(Level4_5, Qual_Ref, Level3, Qual_CT, Cert_CT, Qual_MM, Cert_MM, Qual_Clerk, Cert_Clerk,
                                                    Qual_Starter, Cert_Starter, Qual_CFJ, Cert_CFJ, Qual_SandT, Cert_SandT)
        

        if my_scenario:
            staff_list = self._find_staffing_scenario(my_scenario,[],len(my_scenario))
            if staff_list:   # Passed Sr. Checks - Check S&T then continue
                qual_SandT_left = list(filter(lambda i: i not in staff_list, self.SandT[3]))
                cert_SandT_left = list(filter(lambda i: i not in staff_list, self.SandT[4]))

                if len(cert_SandT_left) < Cert_SandT or len(qual_SandT_left) < Qual_SandT + Cert_SandT:
                    msg = dbg_scenario_name +" : Unable to staff stroke & turn"
                    logging.debug(self.club_code + ": "+msg)
                    self.Failed_Sanctions.append(msg)
                    return {}
                staff_list.reverse()
                staff_jobs = list(my_scenario.keys())
                return dict(zip(staff_jobs, staff_list))
            else:
                msg = dbg_scenario_name + " : Unable to staff senior grid"
                logging.debug(self.club_code + ": "+msg)
                self.Failed_Sanctions.append(msg)
                return {}
        msg = dbg_scenario_name +" : Minimum available skills not met"
        logging.debug(self.club_code + ": "+msg)
        self.Failed_Sanctions.append(msg)
        return {}
    
    def _build_staffing_scenario(self, Level4_5 : int, Qual_Ref: int, Level3: int, 
                                 Qual_CT: int, Cert_CT: int,
                                 Qual_MM: int, Cert_MM: int,
                                 Qual_Clerk: int, Cert_Clerk: int,
                                 Qual_Starter: int, Cert_Starter: int,
                                 Qual_CFJ: int, Cert_CFJ: int,
                                 Qual_SandT: int, Cert_SandT: int) -> dict:
        ''' Build the dictionary for the staffing scenario'''

        # Higher Level Positions can backfill lower level positions to achieve the desired outcome

        scenario = {}
        
        # No need to performance optimize this, we also do some pre-checks to avoid analysis on scenarios 
        # we know would fail

        # Check Quick Failure Conditions
        if (len(self.ChiefT[3]) < Qual_CT + Cert_CT or
                len(self.ChiefT[4]) < Cert_CT or
                len(self.MeetM[3]) < Qual_MM + Cert_MM or
                len(self.MeetM[4]) < Cert_MM or
                len(self.Clerk[3]) < Qual_Clerk + Cert_Clerk or
                len(self.Clerk[4]) < Cert_Clerk or 
                len(self.Starter[3]) < Qual_Starter + Cert_Starter or
                len(self.Starter[4]) < Cert_Starter or
                len(self.CFJ[3]) < Qual_CFJ + Cert_CFJ or
                len(self.CFJ[4]) < Cert_CFJ or
                len(self.SandT[3]) < Qual_SandT or
                len(self.SandT[4]) < Cert_SandT or
                Level4_5 > (self.Level_4s + self.Level_5s) or
                (Qual_Ref + Level4_5) >  (self.Level_4s + self.Level_5s + self.Qual_Refs) or
                (Level3 + Level4_5 + Qual_Ref) > (self.Level_5s + self.Level_4s + self.Level_3s)) :
            return {}

        for x in range(Qual_CT): 
            if self.ChiefT[3]: scenario['CT_Q'+str(x)] = self.ChiefT[3]
        for x in range(Cert_CT): 
            if self.ChiefT[4]: scenario["CT_C"+str(x)] = self.ChiefT[4]
        for x in range(Qual_MM): 
            if self.MeetM[3]: scenario["MM_Q"+str(x)] = self.MeetM[3]
        for x in range(Cert_MM): 
            if self.MeetM[4]: scenario["MM_C"+str(x)] = self.MeetM[4]
        for x in range(Qual_Clerk): 
            if self.Clerk[3]: scenario["Clerk_Q"+str(x)] = self.Clerk[3]
        for x in range(Cert_Clerk): 
            if self.Clerk[4]: scenario["Clerk_C"+str(x)] = self.Clerk[4]
        for x in range(Qual_Starter):
            if self.Starter[3]: scenario["Starter_Q"+str(x)] = self.Starter[3]
        for x in range(Cert_Starter):
            if self.Starter[4]: scenario["Starter_C"+str(x)] = self.Starter[4]
        for x in range(Qual_CFJ):
            if self.CFJ[3]: scenario["CFJ_Q"+str(x)] = self.CFJ[3]
        for x in range(Cert_CFJ):
            if self.CFJ[4]: scenario["CFJ_C"+str(x)] = self.CFJ[4]
        for x in range(Level4_5):
            if self.Level_4_5s: scenario["L45_"+str(x)] = self.Level_4_5s
        for x in range(Level3):
            scenario["L3_"+str(x)] = []            
            if self.Level_3_list: scenario["L3_"+str(x)].extend(self.Level_3_list)
            if self.Level_4_5s: scenario["L3_"+str(x)].extend(self.Level_4_5s)
        for x in range(Qual_Ref):
            scenario["L3Ref_"+str(x)] = []
            if self.Qualified_Refs: scenario["L3Ref_"+str(x)].extend([item[0] for item in self.Qualified_Refs])
            if self.Level_4_5s: scenario["L3Ref_"+str(x)].extend(self.Level_4_5s)
        return scenario

    def _find_staffing_scenario(self, scenario: dict, current_plan: List, required_staff: int) -> List:
        """Try to find a workable staffing scenario"""
        working_plan : List
        current_skill : dict
        scenario_copy : dict

        if not scenario: 
            return []
        
        scenario_copy = deepcopy(scenario)
        current_skill = scenario_copy.popitem()
        working_plan = deepcopy(current_plan)

        for name in current_skill[1]:
            working_plan = deepcopy(current_plan)
            if name not in current_plan and scenario_copy:
                working_plan = copy(current_plan)
                working_plan.append(name)
                sub = self._find_staffing_scenario (scenario_copy, working_plan, required_staff)
                if sub and len(sub) == required_staff: 
                    return sub
            elif name not in current_plan:
                working_plan.append(name)
                return working_plan
        return []    



    def dump_data_docx(self, doc : Document, club_fullname: str, reportdate: str, affiliates: list):
        '''Produce the Word Document for the club'''
        doc.add_heading(club_fullname + " (" + self.club_code + ")", 0)
        doc.add_heading("Provisionally Approved Sanction Types (as of "+reportdate+")", level = 2)
        sanction_p = doc.add_paragraph()
        for sanction in self.Sanction_Level:
            sanction_p.add_run("\n%s" % sanction)
        if len(self.Sanction_Level) == 0:
            sanction_p.add_run("NO APPROVED SANCTION TYPES")
        doc.add_heading("Officials Summary", level = 2)
        ps = doc.add_paragraph("No Level: %d\n" % self.Level_None)
        ps.add_run("Level 1 : %d\n" % self.Level_1s)
        ps.add_run("Level 2 : %d\n" % self.Level_2s)
        ps.add_run("Level 3 : %d\n" % self.Level_3s)
        ps.add_run("Level 4 : %d\n" % self.Level_4s)
        ps.add_run("Level 5 : %d" % self.Level_5s)

        table = doc.add_table(rows=1, cols=4)
        row = table.rows[0].cells
        row[0].text = "Qualficiation"
        row[1].text = "Total Clinics"
        row[2].text = "1 Sign-Off"
        row[3].text = "2 Sign-Offs" 
        row[0].paragraphs[0].alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.RIGHT
        row[1].paragraphs[0].alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
        row[2].paragraphs[0].alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
        row[3].paragraphs[0].alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER

        table_data = [
            ("Intro to Swimming", str(self.Intro[0]), str(self.Intro[1]), str(self.Intro[2])),
            ("Stroke & Turn", str(self.SandT[0]), str(self.SandT[1]), str(self.SandT[2])),
            ("Chief Timekeeper", str(self.ChiefT[0]), str(self.ChiefT[1]), str(self.ChiefT[2])),
            ("Admin Desk (Clerk)", str(self.Clerk[0]), str(self.Clerk[1]), str(self.Clerk[2])) ,
            ("Meet Manager", str(self.MeetM[0]), str(self.MeetM[1]), str(self.MeetM[2])) ,
            ("Starter", str(self.Starter[0]), str(self.Starter[1]), str(self.Starter[2])),
            ("CFJ/CJE",str(self.CFJ[0]), str(self.CFJ[1]), str(self.CFJ[2])),
            ("Recorder/Scorer", str(self.RecSec[0]), "", ""),
            ("Referee Clinics", str(self.Referee[0]), "", "") ]
        

        for entry in table_data:
            row = table.add_row().cells
            for k in range(len(entry)):
                row[k].text = entry[k]
                if k > 0:
                    row[k].paragraphs[0].alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
                else:
                    row[k].paragraphs[0].alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.RIGHT

        table.style = "Light Grid Accent 5"
        table.autofit = True


        if len(self.Qualified_Refs) > 0:
            doc.add_heading("Qualified Level III Referees", level = 3)
            refp = doc.add_paragraph()
            for ref in self.Qualified_Refs:
                refname = ref[0]
                if ref[2] == "yes":
                    refname += " (Para eModule)"
                if not pd.isnull(ref[1]):
                    refname += " (Para National: "+ref[1]+")"
                refp.add_run("\n"+refname)

        if self._config.get_bool("incl_affiliates") and affiliates:
            affiliated_officials = self._club_data_full[self._club_data_full["Registration Id"].isin(affiliates)]
            if not affiliated_officials.empty:       # We could have affiliated officials but no supporting data (PSO vs ROR vs Club views)
                doc.add_heading("Affiliated Officials", level = 3)

                atable = doc.add_table(rows=1, cols=3)
                row = atable.rows[0].cells
                row[0].text = "Registration Id"
                row[1].text = "Name"
                row[2].text = "Certification Level" 
                row[0].paragraphs[0].alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.LEFT
                row[1].paragraphs[0].alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.LEFT
                row[2].paragraphs[0].alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.LEFT

                for index, affiliated_official in affiliated_officials.iterrows():
                    row = atable.add_row().cells
                    row[0].text = affiliated_official["Registration Id"]
                    row[1].text = affiliated_official["Last Name"]+", "+affiliated_official["First Name"]+" ("+affiliated_official["ClubCode"] +")"
                    row[2].text = affiliated_official["Current_CertificationLevel"]
                    row[0].paragraphs[0].alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.LEFT
                    row[1].paragraphs[0].alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.LEFT
                    row[2].paragraphs[0].alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.LEFT

                atable.style = "Light Grid Accent 5"
#                atable.autofit = True
                atable.allow_autofit = True

        if self._config.get_bool("incl_sanction_errors") and self.Failed_Sanctions:
            doc.add_heading("Sanctioning Issues", level = 3)
            error_p = doc.add_paragraph()
            newline = False
            for errmsg in self.Failed_Sanctions:
                if newline:
                    error_p.add_run("\n"+errmsg)
                else:
                    error_p.add_run(errmsg)
                    newline = True

        if self._config.get_bool("incl_errors"):
            if self.NoLevel_Missing_Cert:
                doc.add_heading("RTR Error Detected - Official(s) missing Level I Certification Record", level = 2)
                error_p = doc.add_paragraph()
                error_p.add_run('\n'.join(self.NoLevel_Missing_Cert))

            if self.NoLevel_Missing_SM:
                doc.add_heading("RTR Warning - Level I Partially Complete - Need Safety Marshal", level = 2)
                warn_p = doc.add_paragraph()
                warn_p.add_run('\n'.join(self.NoLevel_Missing_SM))

            if self.NoLevel_Has_II:
                doc.add_heading("RTR Warning - Official has Level II clinics - Missing Level I Certification", level = 2)
                warn_p2 = doc.add_paragraph()
                warn_p2.add_run('\n'.join(self.NoLevel_Has_II))

if __name__ == '__main__':
    ''' Testing -- Incomplete'''
    my_dict = {
        "CT": ["John", "Alice", "Peter", "Mary", "Nicole", "Dean","Trevor"],
        "MM": ["Bob", "Charlie", "John", "Peter", "Mary", "Nicole", "Dean","Trevor"],
        "AD": ["Eve", "Curly", "John"],
        "CJE/CJE": ["Alice", "Peter", "Mary", "Nicole", "Dean","Trevor"],
        "Starter": ["John", "Peter", "Mary", "Nicole", "Dean","Trevor"]
    }

        
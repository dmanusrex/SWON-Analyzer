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
# The above copyright notice and this permission notice shall be included in
# all copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
# EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
# MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
# IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
# DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
# OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE
# OR OTHER DEALINGS IN THE SOFTWARE.

"""RTR Fields"""

# There are over 200 fields in the RTR export. These are the needed ones.

REQUIRED_RTR_FIELDS = [
    "Registration Id",
    "First Name",
    "Last Name",
    "Email",
    "Club",
    "Region",
    "Province",
    "Status",
    "Level 1 Date of Certification",
    "Level 2 Date of Certification",
    "Level 3 Date of Certification",
    "Level 4 Date of Certification",
    "Level 5 Date of Certification",
    "Introduction to Swimming Officiating",
    "Introduction to Swimming Officiating-ClinicDate",
    "Introduction to Swimming Officiating-Deck Evaluation #1 Date",
    "Introduction to Swimming Officiating-Deck Evaluation #2 Date",
    "Safety Marshal",
    "Safety Marshal-ClinicDate",
    "Safety Marshal-Deck Evaluation #1 Date",
    "Safety Marshal-Deck Evaluation #2 Date",
    "Judge of Stroke/Inspector of Turns",
    "Judge of Stroke/Inspector of Turns-ClinicDate",
    "Judge of Stroke/Inspector of Turns-Deck Evaluation #1 Date",
    "Judge of Stroke/Inspector of Turns-Deck Evaluation #2 Date",
    "Administration Desk (formerly Clerk of Course) Clinic",
    "Administration Desk (formerly Clerk of Course) Clinic-ClinicDate",
    "Administration Desk (formerly Clerk of Course) Clinic-Deck Evaluation #1 Date",
    "Administration Desk (formerly Clerk of Course) Clinic-Deck Evaluation #2 Date",
    "Chief Timekeeper",
    "Chief Timekeeper-ClinicDate",
    "Chief Timekeeper-Deck Evaluation #1 Date",
    "Chief Timekeeper-Deck Evaluation #2 Date",
    "Meet Manager",
    "Meet Manager-ClinicDate",
    "Meet Manager-Deck Evaluation #1 Date",
    "Meet Manager-Deck Evaluation #2 Date",
    "Chief Finish Judge/Chief Judge",
    "Chief Finish Judge/Chief Judge-ClinicDate",
    "Chief Finish Judge/Chief Judge-Deck Evaluation #1 Date",
    "Chief Finish Judge/Chief Judge-Deck Evaluation #2 Date",
    "Chief Judge Electronics",
    "Chief Judge Electronics-ClinicDate",
    "Chief Judge Electronics-Deck Evaluation #1 Date",
    "Chief Judge Electronics-Deck Evaluation #2 Date",
    "Chief Recorder and Recorder (formerly Recorder/Scorer) Clinic",
    "Chief Recorder and Recorder (formerly Recorder/Scorer) Clinic-ClinicDate",
    "Chief Recorder and Recorder (formerly Recorder/Scorer) Clinic-Deck Evaluation #1 Date",
    "Chief Recorder and Recorder (formerly Recorder/Scorer) Clinic-Deck Evaluation #2 Date",
    "Starter",
    "Starter-ClinicDate",
    "Starter-Deck Evaluation #1 Date",
    "Starter-Deck Evaluation #2 Date",
    "Referee",
    "Referee-ClinicDate",
    "Referee-Deck Evaluation #1 Date",
    "Referee-Deck Evaluation #2 Date",
    "Para Swimming eModule",
    "Para Swimming eModule-ClinicDate",
    "Judge of Stroke",
    "Judge of Stroke-ClinicDate",
    "Judge of Stroke-Deck Evaluation #1 Date",
    "Judge of Stroke-Deck Evaluation #2 Date",
    "Inspector of Turns",
    "Inspector of Turns-ClinicDate",
    "Inspector of Turns-Deck Evaluation #1 Date",
    "Inspector of Turns-Deck Evaluation #2 Date",
    "Para Domestic",
    "Para Domestic Course Date",
    "ClubCode",
    "Current_CertificationLevel",
    "AffiliatedClubs",
]

RTR_POSITION_FIELDS = {
    "Intro": [
        "Introduction to Swimming Officiating",
        "Introduction to Swimming Officiating-Deck Evaluation #1 Date",
        "Introduction to Swimming Officiating-Deck Evaluation #2 Date",
    ],
    "ST": [
        "Judge of Stroke/Inspector of Turns",
        "Judge of Stroke/Inspector of Turns-Deck Evaluation #1 Date",
        "Judge of Stroke/Inspector of Turns-Deck Evaluation #2 Date",
    ],
    "IT": [
        "Inspector of Turns",
        "Inspector of Turns-Deck Evaluation #1 Date",
        "Inspector of Turns-Deck Evaluation #2 Date",
    ],
    "JoS": ["Judge of Stroke", "Judge of Stroke-Deck Evaluation #1 Date"],
    "CT": ["Chief Timekeeper", "Chief Timekeeper-Deck Evaluation #1 Date", "Chief Timekeeper-Deck Evaluation #2 Date"],
    "Clerk": [
        "Administration Desk (formerly Clerk of Course) Clinic",
        "Administration Desk (formerly Clerk of Course) Clinic-Deck Evaluation #1 Date",
        "Administration Desk (formerly Clerk of Course) Clinic-Deck Evaluation #2 Date",
    ],
    "MM": ["Meet Manager", "Meet Manager-Deck Evaluation #1 Date", "Meet Manager-Deck Evaluation #2 Date"],
    "Starter": ["Starter", "Starter-Deck Evaluation #1 Date", "Starter-Deck Evaluation #2 Date"],
    "ChiefRec": ["Chief Recorder and Recorder (formerly Recorder/Scorer) Clinic"],
    "CFJ": [
        "Chief Finish Judge/Chief Judge",
        "Chief Finish Judge/Chief Judge-Deck Evaluation #1 Date",
        "Chief Finish Judge/Chief Judge-Deck Evaluation #2 Date",
    ],
    "Referee": ["Referee"],
}


# Abstract the RTR fields so that they can be changed easily if the RTR export changes

RTR_CLINICS = {
    "Intro": {
        "hasClinic": "Introduction to Swimming Officiating",
        "clinicDate": "Introduction to Swimming Officiating-ClinicDate",
        "deckEvals": [
            "Introduction to Swimming Officiating-Deck Evaluation #1 Date",
            "Introduction to Swimming Officiating-Deck Evaluation #2 Date",
        ],
        "status": "Intro_Status",
        "signoffs": "Intro_Count",
    },
    "ST": {
        "hasClinic": "Judge of Stroke/Inspector of Turns",
        "clinicDate": "Judge of Stroke/Inspector of Turns-ClinicDate",
        "deckEvals": [
            "Judge of Stroke/Inspector of Turns-Deck Evaluation #1 Date",
            "Judge of Stroke/Inspector of Turns-Deck Evaluation #2 Date",
        ],
        "status": "ST_Status",
        "signoffs": "ST_Count",
    },
    "IT": {
        "hasClinic": "Inspector of Turns",
        "clinicDate": "Inspector of Turns-ClinicDate",
        "deckEvals": ["Inspector of Turns-Deck Evaluation #1 Date", "Inspector of Turns-Deck Evaluation #2 Date"],
        "status": "IT_Status",
        "signoffs": "IT_Count",
    },
    "JoS": {
        "hasClinic": "Judge of Stroke",
        "clinicDate": "Judge of Stroke-ClinicDate",
        "deckEvals": ["Judge of Stroke-Deck Evaluation #1 Date", "Judge of Stroke-Deck Evaluation #2 Date"],
        "status": "JoS_Status",
        "signoffs": "JoS_Count",
    },
    "CT": {
        "hasClinic": "Chief Timekeeper",
        "clinicDate": "Chief Timekeeper-ClinicDate",
        "deckEvals": ["Chief Timekeeper-Deck Evaluation #1 Date", "Chief Timekeeper-Deck Evaluation #2 Date"],
        "status": "CT_Status",
        "signoffs": "CT_Count",
    },
    "AdminDesk": {
        "hasClinic": "Administration Desk (formerly Clerk of Course) Clinic",
        "clinicDate": "Administration Desk (formerly Clerk of Course) Clinic-ClinicDate",
        "deckEvals": [
            "Administration Desk (formerly Clerk of Course) Clinic-Deck Evaluation #1 Date",
            "Administration Desk (formerly Clerk of Course) Clinic-Deck Evaluation #2 Date",
        ],
        "status": "Admin_Status",
        "signoffs": "Admin_Count",
    },
    "MM": {
        "hasClinic": "Meet Manager",
        "clinicDate": "Meet Manager-ClinicDate",
        "deckEvals": ["Meet Manager-Deck Evaluation #1 Date", "Meet Manager-Deck Evaluation #2 Date"],
        "status": "MM_Status",
        "signoffs": "MM_Count",
    },
    "Starter": {
        "hasClinic": "Starter",
        "clinicDate": "Starter-ClinicDate",
        "deckEvals": ["Starter-Deck Evaluation #1 Date", "Starter-Deck Evaluation #2 Date"],
        "status": "Starter_Status",
        "signoffs": "Starter_Count",
    },
    "ChiefRec": {
        "hasClinic": "Chief Recorder and Recorder (formerly Recorder/Scorer) Clinic",
        "clinicDate": "Chief Recorder and Recorder (formerly Recorder/Scorer) Clinic-ClinicDate",
        "deckEvals": [],
        "status": "ChiefRec_Status",
        "signoffs": "ChiefRec_Count",
    },
    "CFJ": {
        "hasClinic": "Chief Finish Judge/Chief Judge",
        "clinicDate": "Chief Finish Judge/Chief Judge-ClinicDate",
        "deckEvals": [
            "Chief Finish Judge/Chief Judge-Deck Evaluation #1 Date",
            "Chief Finish Judge/Chief Judge-Deck Evaluation #2 Date",
        ],
        "status": "CFJ_Status",
        "signoffs": "CFJ_Count",
    },
    "Referee": {"hasClinic": "Referee", "clinicDate": "Referee-ClinicDate", "deckEvals": [], "status": "Referee_Status", "signoffs": "Referee_Count"},
    "Para": {"hasClinic": "Para Swimming eModule", "clinicDate": "Para Swimming eModule-ClinicDate", "deckEvals": []},
    "ParaDom": {"hasClinic": "Para Domestic", "clinicDate": "Para Domestic Course Date", "deckEvals": []},
}

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
    "Clerk": ["Administration Desk (formerly Clerk of Course) Clinic", "Administration Desk (formerly Clerk of Course) Clinic-Deck Evaluation #1 Date", "Administration Desk (formerly Clerk of Course) Clinic-Deck Evaluation #2 Date"],
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

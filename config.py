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
#

'''Config parsing and options'''

import configparser

class AnalyzerConfig:
    '''Get/Set program options'''

    # Name of the configuration file
    _CONFIG_FILE = "swon-analyzer.ini"
    # Name of the section we use in the ini file
    _INI_HEADING = "swon-analyzer"
    # Configuration defaults if not present in the config file
    _CONFIG_DEFAULTS = {_INI_HEADING: {
        "officials_list": "./officials_list.xls",   # Location of RTR export file
        "report_directory": ".",                    # Report output directory
        "report_file": "club_analysis.txt",         # test output file 
        "report_file_docx": "club_analysis.docx",   # Word File name
        "report_file_cohost": "cohosting.docx",     # Co-hosting filename
        "incl_errors": "True",                      # Include Errors
        "incl_inv_pending": "True",                 # Include Invoice Pending Status
        "incl_pso_pending": "True",                 # Include PSO Pending Status
        "incl_account_pending": "True",             # Include Account Pending Status
        "incl_affiliates": "True",                  # Include Affiliated Officials
        "incl_sanction_errors": "True",             # Include Sanctioning Errors in Reports
        "gen_1_per_club": "False",                  # Generate 1 Word Doc / Club
        "gen_word": "True",                         # Generate Master Word Doc
        "Theme": "System",                          # Theme- System, Dark or Light
        "Scaling": "100%",                          # Display Zoom Level
        "Colour" : "blue",                          # Colour Theme
    }}

    def __init__(self):
        self._config = configparser.ConfigParser(interpolation=None)
        self._config.read_dict(self._CONFIG_DEFAULTS)
        self._config.read(self._CONFIG_FILE)

    def save(self) -> None:
        '''Save the (updated) configuration to the ini file'''
        with open(self._CONFIG_FILE, 'w') as configfile:
            self._config.write(configfile)

    def get_str(self, name: str) -> str:
        '''Get a string option'''
        return self._config.get(self._INI_HEADING, name)

    def set_str(self, name: str, value: str) -> str:
        '''Set a string option'''
        self._config.set(self._INI_HEADING, name, value)
        return self.get_str(name)

    def get_float(self, name: str) -> float:
        '''Get a float option'''
        return self._config.getfloat(self._INI_HEADING, name)

    def set_float(self, name: str, value: float) -> float:
        '''Set a float option'''
        self._config.set(self._INI_HEADING, name, str(value))
        return self.get_float(name)

    def get_int(self, name: str) -> int:
        '''Get an integer option'''
        return self._config.getint(self._INI_HEADING, name)

    def set_int(self, name: str, value: int) -> int:
        '''Set an integer option'''
        self._config.set(self._INI_HEADING, name, str(value))
        return self.get_int(name)

    def get_bool(self, name: str) -> bool:
        '''Get a boolean option'''
        return self._config.getboolean(self._INI_HEADING, name)

    def set_bool(self, name: str, value: bool) -> bool:
        '''Set a boolean option'''
        self._config.set(self._INI_HEADING, name, str(value))
        return self.get_bool(name)
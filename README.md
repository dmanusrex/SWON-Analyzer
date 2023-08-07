# SWON-Analyzer

   Swim Ontario is changing sanctioning requirements to be focus on roles and the host club(s) ability to staff the meets.  This utility is designed to analyze an offiicals export from Swimming Canada's Registration Tracking and Results (RTR) system to determine what sanctioning level a club/clubs may qualify for.  In addition, it will identify any issues with the RTR data export.  

## Requirements

This will run on on Windows 10 or Windows 11 PC.   The user must have COA, ROR or POA access to the RTR to be able to generate the needed export(s).  RTR export files must be the original files and have been done after July 27, 2023.

## Features

- Allows users to import one (or more) RTR export files. The system will merge the files.  This is needed for meet co-hosting.
- Official status selection
- Ability to produces various reports
- Ability to include/exclude RTR errors/warning from the report
- Ability to incldue/exclude reasons sanction types could not be approved

## Installation

- see ![INSTALL](INSTALL)

## How it works

1. Select an RTR export file and click the "Load Datafile" button.
2. If you are co-hosting repeat step 1 as necessary to load the other club files.
3. On the Analyzer/Report Settings tab choose your options
4. The reports tab will generate standard reports. One is a master all-in-one file and the other will generate 1 doc/club.
5. To check meet co-hosting, use the co-hosting tab. Select the co-host clubs and generate the report.

## License
This software is licensed under the GNU Affero General Public License version
3. See the [LICENSE](LICENSE) file for full details.



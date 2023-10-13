## Changelog

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/)

### [Unreleased]
- :bug: Fix UI scaling


### [0.5.5] - 2023-07-28
- :sparkles: Baseline release for testing
- :bug: Fixed a bunch of bugs

### [0.6.0] - 2023-08-07
- :sparkles: Add checks for missing Level II / III Certifications

### [0.7.0] - 2023-08-17
- :sparkles: Updated for the latest sanctioning matrix
- :sparkles: Changes to stroke & turn staffing to better support upcoming Fall 2023 SNC Changes
- :bug: Fixed a bug where the RTR filename would not update on UI when re-selected

### [0.9.0] - 2023-08-12
- :sparkles: First version of the integrated app

### [1.0.0] - 2023-08-13
- :sparkles: Added the ability to select a club for recommendations generation to allow ROR/POAs to use their master files

### [1.0.1] - 2023-08-24
- :bug: Fixed issue with pyinstaller not picking up icons for CTKMessageBox

### [1.0.2] - 2023-08-27
- :bug: Fixed issue with SMTP not working for port 587 (STARTTLS)

### [1.1.0] - 2023-09-05
- :sparkles: Basic [docs](http://SWON-Analyzer.readthedocs.io/)
- :sparkles: Experimental Support for the New Pathway
- :sparkles: Support for RTR file format changes as of September 1, 2023
- :sparkles: Support for CSV format in addition to the standard RTR HTML format

### [1.1.1] - 2023-09-05
- :bug: Fix race condition on Stroke & Turn staffing scenarios

### [1.2.0] - 2023-09-07
- :sparkles: Generate a setup exe for windows (uses NSIS installer)


### [1.2.1]
- :bug: Stop using ONEFILE option with pyinstaller
- :bug: Fixed some scaling issues
- :sparkles: Add option to create desktop icon during installation

### [1.2.2]
- :sparkles: Support RTR Export format changes

### [1.3.0]
- :bug: Show new clinics properly in Sanctioning Module
- :sparkles: First phase of internal re-work to full abstract RTR export format

### [1.3.1]
- :bug: Fix bug introduced in 1.3.0 that caused summary stats to be incorrect
- :sparkles: RTR data abstraction is complete
- :sparkles: Recommendations are now in line with new split clinic requirements

### [1.4.0]
- :sparkles: New RTR Browser and simplied data export

### [1.5.0]

- :sparkles: Migrate to internal fields for certifications to abstract RTR changes
- :sparkles: Move configs and logs to user directories
- :sparkles: Clarify Level 3/4/5 failure reasons
- :sparkles: Add Sentry SDK support
- :bug: Correct Intro recommendation
- :sparkles: Add verbose failure reasons (# of missing skills, which position couldn't be staffed)
- :sparkles: Allow for All Users/This User Only installs on Windows
- :bug: Fix pyinstaller spec not picking up supporting docxcompose files
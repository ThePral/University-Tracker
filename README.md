UNIVERSITY TRACKER
=================

A desktop application planner for university and graduate program applications.
The interface uses a blue theme and stores data locally in an Excel workbook.

OVERVIEW
- Dashboard cards summarize programs, submitted or active applications, offers,
  and actionable deadlines in the next 14 days.
- Application pipeline counts are clickable to filter by stage.
- Deadline center lists overdue items and deadlines in the next 60 days. Double-click
  a row to open that program for editing.

APPLICATION WORKSPACE
- Keep university, program, admissions, funding, research, and application details
  together in one record.
- Search across saved details and filter by status, priority, country, degree,
  or deadline range. Sort by recent check, deadline, university, or priority.
- Add, edit, duplicate, and delete one or multiple programs.
- Select up to four programs to compare key admissions, cost, funding, and fit
  details side by side.
- Right-click selected programs to change application status without opening the form.

DATA TOOLS
- Import CSV files that include Country, University, and Program Name columns.
  Matching university/program records can be updated or skipped during import.
- Export the currently filtered programs to UTF-8 CSV.
- Create a dated backup copy of the Excel workbook.
- Source and portable copies keep Excel_Sheets\universities.xlsx beside the app.
  Installed Qt copies use %LOCALAPPDATA%\University Tracker\Excel_Sheets\universities.xlsx.
  The workbook is written safely and grouped into sheets by country.

KEYBOARD SHORTCUTS
- Ctrl+N: add a program
- Ctrl+S: save the current program
- Ctrl+F: focus shortlist search
- Ctrl+E: export filtered programs to CSV
- Ctrl+I: import CSV
- Ctrl+D: duplicate the selected program
- Delete: delete selected program(s), after confirmation

RUN FROM SOURCE
1. Install Python 3.11 or newer.
2. Run: python -m pip install -r requirements.txt
3. Run: python university_tracker.py

BUILD THE WINDOWS EXECUTABLE
Run build_exe.bat. The executable is created at dist\UniversityTracker.exe.
Keep the Excel_Sheets folder beside the executable to retain existing data.
Close universities.xlsx in Excel before saving from the app.

FRESH SIDEBAR EDITION
Run the separate redesign with: python university_tracker_v2.py
It uses the same Excel_Sheets\universities.xlsx file and preserves the original
program fields. The existing build_exe.bat still packages university_tracker.py.

QT EDITION AND WINDOWS INSTALLER
Run the Qt application with: python university_tracker_qt.py
Install its build dependencies with: python -m pip install -r requirements_qt.txt
The portable Qt executable is built by build_university_tracker_qt.bat.

To make a shareable installer, run build_university_tracker_installer.bat.
The builder needs 64-bit Python 3.11 or newer on Windows. On the first build it
downloads a pinned official Inno Setup compiler, checks its SHA-256, and prepares
it in portable mode under .installer-tools. This does not install the compiler
system-wide. It also installs the Python build dependencies if needed.
An installed Inno Setup compiler is used when available. For an offline build,
preinstall requirements_qt.txt and use:
  build_university_tracker_installer.bat --skip-dependencies --compiler "C:\path\to\ISCC.exe"

The finished file is dist_installer\UniversityTracker-Setup-1.0.1.exe.
Share this file with others. They do not need Python, Excel, or Inno Setup.
It supports Windows 10 (1809 or newer) and Windows 11 with x64 compatibility.
Setup installs for the current Windows user without requiring administrator
rights, adds a Start Menu shortcut, offers a desktop shortcut, and registers
an uninstaller in Windows Settings > Apps. After setup, the user can launch
University Tracker from the finish page or Start Menu.

The installer bundles the application and its dependencies, with an empty
tracker for each user. It never copies your existing shortlist into the setup
file. Personal workbooks are stored outside the installation folder, and are
kept during upgrades and uninstalling. installer\QUICK_START.txt explains how
to move an existing workbook from a portable copy.

For a new release, update APP_VERSION in university_tracker_qt.py and rebuild.
The builder reads that version for the app resource and installer filename.
Keep AppId in installer\university_tracker.iss unchanged so upgrades reuse the
existing installation. Build files are under build_installer; setup files and
their SHA-256 checksums are under dist_installer.

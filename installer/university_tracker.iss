; Compile through build_university_tracker_installer.bat.
; Keep AppId unchanged so later versions upgrade the same installation.
#ifndef AppVersion
  #define AppVersion "1.0.1"
#endif
#ifndef AppSource
  #define AppSource AddBackslash(SourcePath) + "..\build_installer\app\UniversityTrackerQt"
#endif

[Setup]
AppId={{145A015C-2BD9-4BBD-BDF1-B91532872833}
AppName=University Tracker
AppVersion={#AppVersion}
AppPublisher=University Tracker
VersionInfoDescription=University Tracker Installer
VersionInfoProductName=University Tracker
VersionInfoVersion={#AppVersion}
DefaultDirName={localappdata}\Programs\University Tracker
DefaultGroupName=University Tracker
PrivilegesRequired=lowest
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
MinVersion=10.0.17763
DisableProgramGroupPage=yes
DisableWelcomePage=no
AllowNoIcons=yes
UsePreviousAppDir=yes
UninstallDisplayIcon={app}\UniversityTrackerQt.exe
CloseApplications=yes
CloseApplicationsFilter=UniversityTrackerQt.exe
RestartApplications=no
WizardStyle=modern
Compression=lzma2
SolidCompression=yes
OutputDir=..\dist_installer
OutputBaseFilename=UniversityTracker-Setup-{#AppVersion}
InfoBeforeFile=QUICK_START.txt

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; Flags: unchecked

[Files]
; Deliberately package only the application payload, never Excel_Sheets.
Source: "{#AppSource}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "installed-mode.txt"; DestDir: "{app}"; Flags: ignoreversion
Source: "QUICK_START.txt"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\University Tracker"; Filename: "{app}\UniversityTrackerQt.exe"; WorkingDir: "{app}"
Name: "{autodesktop}\University Tracker"; Filename: "{app}\UniversityTrackerQt.exe"; WorkingDir: "{app}"; Tasks: desktopicon

[Run]
Filename: "{app}\UniversityTrackerQt.exe"; Description: "Launch University Tracker"; Flags: nowait postinstall skipifsilent

; User workbooks live outside {app} and are never listed in [UninstallDelete].

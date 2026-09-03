#ifndef MyAppVersion
  #define MyAppVersion "0.2.0"
#endif

#define MyAppName "Excel Scenario Analysis Tool"
#define MyAppPublisher "Leo Goldberg"
#define MyAppExeName "ExcelScenarioAnalysisTool.exe"

[Setup]
AppId={{6D443E25-FFB1-4D54-913B-980A552B8748}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL=https://github.com/yogoldy/excel-diff-tracker
AppSupportURL=https://github.com/yogoldy/excel-diff-tracker/issues
AppUpdatesURL=https://github.com/yogoldy/excel-diff-tracker/releases
DefaultDirName={localappdata}\Programs\Excel Scenario Analysis Tool
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
MinVersion=10.0.22000
ArchitecturesAllowed=arm64
ArchitecturesInstallIn64BitMode=arm64
OutputDir=..\artifacts\installer
OutputBaseFilename=ExcelScenarioAnalysisTool-Setup-arm64
SetupIconFile=..\artifacts\branding\app-icon.ico
UninstallDisplayIcon={app}\{#MyAppExeName}
LicenseFile=..\LICENSE
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
CloseApplications=yes
CloseApplicationsFilter={#MyAppExeName},ExcelDiffTracker.exe
RestartApplications=no
ChangesAssociations=no
ChangesEnvironment=no
VersionInfoVersion={#MyAppVersion}.0
VersionInfoCompany={#MyAppPublisher}
VersionInfoDescription={#MyAppName} installer
VersionInfoProductName={#MyAppName}
VersionInfoProductVersion={#MyAppVersion}

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "startup"; Description: "Start Excel Scenario Analysis Tool with Windows"; GroupDescription: "Startup:"; Flags: checkedonce
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Shortcuts:"; Flags: unchecked

[Files]
Source: "..\artifacts\publish\win-arm64\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[InstallDelete]
; Retire only the explicitly listed runtime files owned by the prior release.
; User workbooks, history, reports, settings, and uninstall metadata are excluded.
#include "LegacyRuntimeFiles-0.1.1.iss"
Type: files; Name: "{app}\ExcelDiffTracker.exe"
Type: files; Name: "{userprograms}\Excel Diff Tracker\Excel Diff Tracker.lnk"
Type: dirifempty; Name: "{userprograms}\Excel Diff Tracker"
Type: files; Name: "{userdesktop}\Excel Diff Tracker.lnk"

[Icons]
Name: "{group}\Excel Scenario Analysis Tool"; Filename: "{app}\{#MyAppExeName}"
Name: "{userdesktop}\Excel Scenario Analysis Tool"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Registry]
Root: HKCU; Subkey: "Software\Microsoft\Windows\CurrentVersion\Run"; ValueType: none; ValueName: "ExcelDiffTracker"; Flags: deletevalue
Root: HKCU; Subkey: "Software\Microsoft\Windows\CurrentVersion\Run"; ValueType: string; ValueName: "ExcelScenarioAnalysisTool"; ValueData: """{app}\{#MyAppExeName}"" --background"; Tasks: startup; Flags: uninsdeletevalue

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Launch Excel Scenario Analysis Tool"; Flags: nowait postinstall skipifsilent

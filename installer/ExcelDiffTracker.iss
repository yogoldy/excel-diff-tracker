#ifndef MyAppVersion
  #define MyAppVersion "0.1.1"
#endif

#define MyAppName "Excel Diff Tracker"
#define MyAppPublisher "Leo Goldberg"
#define MyAppExeName "ExcelDiffTracker.exe"

[Setup]
AppId={{6D443E25-FFB1-4D54-913B-980A552B8748}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL=https://github.com/yogoldy/excel-diff-tracker
AppSupportURL=https://github.com/yogoldy/excel-diff-tracker/issues
AppUpdatesURL=https://github.com/yogoldy/excel-diff-tracker/releases
DefaultDirName={localappdata}\Programs\Excel Diff Tracker
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
MinVersion=10.0.22000
ArchitecturesAllowed=arm64
ArchitecturesInstallIn64BitMode=arm64
OutputDir=..\artifacts\installer
OutputBaseFilename=ExcelDiffTracker-Setup-arm64
SetupIconFile=..\artifacts\branding\app-icon.ico
UninstallDisplayIcon={app}\{#MyAppExeName}
LicenseFile=..\LICENSE
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
CloseApplications=yes
CloseApplicationsFilter={#MyAppExeName}
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
Name: "startup"; Description: "Start Excel Diff Tracker with Windows"; GroupDescription: "Startup:"; Flags: checkedonce
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Shortcuts:"; Flags: unchecked

[Files]
Source: "..\artifacts\publish\win-arm64\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\Excel Diff Tracker"; Filename: "{app}\{#MyAppExeName}"
Name: "{userdesktop}\Excel Diff Tracker"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Registry]
Root: HKCU; Subkey: "Software\Microsoft\Windows\CurrentVersion\Run"; ValueType: string; ValueName: "ExcelDiffTracker"; ValueData: """{app}\{#MyAppExeName}"" --background"; Tasks: startup; Flags: uninsdeletevalue

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Launch Excel Diff Tracker"; Flags: nowait postinstall skipifsilent

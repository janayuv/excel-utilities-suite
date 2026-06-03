[Setup]
AppName=MyApp
AppVersion=1.0
DefaultDirName={pf}\MyApp
DefaultGroupName=MyApp
OutputDir=Output
OutputBaseFilename=setup

[Files]
Source: "MyApp.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "MyApp.dll"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\MyApp"; Filename: "{app}\MyApp.exe"

[Run]
Filename: "{app}\MyApp.exe"; Description: "Launch MyApp"; Flags: nowait postinstall
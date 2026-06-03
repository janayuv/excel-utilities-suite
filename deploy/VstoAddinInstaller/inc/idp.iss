[Setup]
AppName=IDP Application
AppVersion=2.1
DefaultDirName={pf}\IDP Application
DefaultGroupName=IDP Application
OutputDir=Output
OutputBaseFilename=idp_installer
Compression=lzma
SolidCompression=yes

[Files]
Source: "bin\IDPApp.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "bin\config.json"; DestDir: "{app}"; Flags: ignoreversion
Source: "bin\*.dll"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs

[Icons]
Name: "{group}\IDP Application"; Filename: "{app}\IDPApp.exe"
Name: "{group}\Uninstall IDP Application"; Filename: "{uninstallexe}"

[Registry]
Root: HKCU; Subkey: "Software\IDPApplication"; ValueType: string; ValueName: "InstallPath"; ValueData: "{app}"

[Run]
Filename: "{app}\IDPApp.exe"; Description: "Launch IDP Application"; Flags: nowait postinstall
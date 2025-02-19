[Setup]
AppName=App Muter
AppVersion=1.0.0
WizardStyle=modern
DefaultDirName={autopf}\App Muter
DefaultGroupName=App Muter
UninstallDisplayIcon={app}\App Muter.exe
Compression=lzma2
SolidCompression=yes
OutputDir=installer
OutputBaseFilename=AppMuterSetup
PrivilegesRequired=admin

[Files]
Source: "dist\App Muter\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs

[Icons]
Name: "{group}\App Muter"; Filename: "{app}\App Muter.exe"
Name: "{commondesktop}\App Muter"; Filename: "{app}\App Muter.exe"

[Run]
Filename: "{app}\App Muter.exe"; Description: "Launch App Muter"; Flags: nowait postinstall skipifsilent 
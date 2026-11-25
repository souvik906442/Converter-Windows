; Inno Setup script - builds Converter-Setup.exe from dist\Converter.exe
[Setup]
AppId={{B7E2C1A4-5D3F-4E8A-9C61-0F2A7D4B8E13}
AppName=Offline Converter
AppVersion=1.0
AppPublisher=souvik906442
DefaultDirName={autopf}\Offline Converter
DefaultGroupName=Offline Converter
DisableProgramGroupPage=yes
OutputDir=dist
OutputBaseFilename=Converter-Setup
SetupIconFile=icon.ico
UninstallDisplayIcon={app}\Converter.exe
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; Flags: unchecked

[Files]
Source: "dist\Converter.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{autoprograms}\Offline Converter"; Filename: "{app}\Converter.exe"
Name: "{autodesktop}\Offline Converter"; Filename: "{app}\Converter.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\Converter.exe"; Description: "Launch Offline Converter"; Flags: nowait postinstall skipifsilent

[Setup]
AppName=Doc-smart
AppVersion=1.0.0
AppPublisher=Your Name
AppPublisherURL=https://github.com/yourusername/doc-smart
DefaultDirName={autopf}\Doc-smart
DefaultGroupName=Doc-smart
AllowNoIcons=yes
OutputDir=output
OutputBaseFilename=Doc-smart-Setup
Compression=lzma
SolidCompression=yes
WizardStyle=modern

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "dist\Doc-smart.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\Doc-smart"; Filename: "{app}\Doc-smart.exe"
Name: "{group}\{cm:UninstallProgram,Doc-smart}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\Doc-smart"; Filename: "{app}\Doc-smart.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\Doc-smart.exe"; Description: "{cm:LaunchProgram,Doc-smart}"; Flags: nowait postinstall skipifsilent
#ifndef AppName
#define AppName "MarkOneNoteDown"
#endif
#ifndef AppVersion
#define AppVersion "1.0.0.0"
#endif
#ifndef AppExeName
#define AppExeName "MarkOneNoteDown.App.exe"
#endif
#ifndef SourceDir
#define SourceDir "artifacts\\exe"
#endif
#ifndef OutputDir
#define OutputDir "artifacts\\installer"
#endif

[Setup]
AppName={#AppName}
AppVersion={#AppVersion}
DefaultDirName={pf}\{#AppName}
OutputDir={#OutputDir}
OutputBaseFilename={#AppName}-Setup
PrivilegesRequired=admin
ArchitecturesInstallIn64BitMode=x64
DisableProgramGroupPage=yes
WizardStyle=modern

[Files]
Source: "{#SourceDir}\\*"; DestDir: "{app}"; Flags: recursesubdirs createallsubdirs ignoreversion

[Icons]
Name: "{autoprograms}\\{#AppName}"; Filename: "{app}\\{#AppExeName}"
Name: "{autodesktop}\\{#AppName}"; Filename: "{app}\\{#AppExeName}"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Create a desktop icon"; GroupDescription: "Additional icons:"

[Run]
Filename: "{app}\\{#AppExeName}"; Description: "Launch {#AppName}"; Flags: nowait postinstall skipifsilent

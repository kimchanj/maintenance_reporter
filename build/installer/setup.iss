; Inno Setup Script - Maintenance Reporter
#define MyAppName "Maintenance Reporter"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "SoftNet"
#define MyAppExeName "MaintenanceReporter.exe"

[Setup]
AppId={{9A3C9B4B-2C31-4A4B-9C31-5E5B0C6B5B2A}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppPublisher}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputDir=..\installer_output
OutputBaseFilename=MaintenanceReporter_Setup_{#MyAppVersion}
Compression=lzma
SolidCompression=yes
SetupIconFile=..\..\assets\app.ico
WizardStyle=modern

[Languages]
Name: "korean"; MessagesFile: "compiler:Languages\Korean.isl"

[Tasks]
Name: "desktopicon"; Description: "바탕화면 아이콘 만들기"; GroupDescription: "추가 아이콘:"; Flags: unchecked

[Files]
Source: "..\pyinstaller\dist\MaintenanceReporter\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "프로그램 실행"; Flags: nowait postinstall skipifsilent

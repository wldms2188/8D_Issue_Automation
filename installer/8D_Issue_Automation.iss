; 8D Issue Automation - stable internal installer
; Application behavior baseline: 4ea1958
; EXE entry point: app\main_enterprise_final.py (via 8D_Issue_Automation.spec)
;
; Manual build:
;   1) Run build_release.bat
;   2) Or compile this file with Inno Setup 6 after dist\8D_Issue_Automation.exe exists.

#define MyAppName "8D Issue Automation"
#define MyAppVersion "1.0.0-stable-4ea1958"
#define MyAppPublisher "Pack 개발품질"
#define MyAppExeName "8D_Issue_Automation.exe"

[Setup]
AppId={{8D5A8E9D-6E52-4F91-A77B-8D2026091501}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={localappdata}\Programs\8D Issue Automation
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputDir=output
OutputBaseFilename=8D_Issue_Automation_Stable_4ea1958_Setup
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=lowest
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
UninstallDisplayName={#MyAppName}
UninstallDisplayIcon={app}\{#MyAppExeName}
SetupLogging=yes
CloseApplications=yes
RestartApplications=no
VersionInfoVersion=1.0.0.0
VersionInfoTextVersion={#MyAppVersion}
VersionInfoCompany={#MyAppPublisher}
VersionInfoDescription={#MyAppName} stable installer

[Languages]
Name: "korean"; MessagesFile: "compiler:Languages\Korean.isl"

[Files]
Source: "..\dist\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion

[Tasks]
Name: "desktopicon"; Description: "바탕화면 바로가기 만들기"; GroupDescription: "추가 바로가기:"; Flags: unchecked

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{#MyAppName} 실행"; WorkingDir: "{app}"; Flags: nowait postinstall skipifsilent

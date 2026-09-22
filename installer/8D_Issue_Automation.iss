; 8D Issue Automation Advanced - isolated distribution installer
; Verified application baseline: 3dc878a
; EXE entry point: app\main_enterprise_final.py (via 8D_Issue_Automation.spec)
;
; This installer intentionally uses its own AppId and install directory so it
; can coexist with older/internal builds without upgrading or overwriting them.

#define MyAppName "8D Issue Automation Advanced"
#define MyAppVersion "2.0.0-advanced-3dc878a"
#define MyAppPublisher "Pack 개발품질"
#define MyAppExeName "8D_Issue_Automation.exe"

[Setup]
AppId={{A83DC878-2026-4A8D-9C30-8D2026092201}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={localappdata}\Programs\8D Issue Automation Advanced
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputDir=output
OutputBaseFilename=8D_Issue_Automation_Advanced_3dc878a_Setup
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
VersionInfoVersion=2.0.0.0
VersionInfoTextVersion={#MyAppVersion}
VersionInfoCompany={#MyAppPublisher}
VersionInfoDescription={#MyAppName}

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

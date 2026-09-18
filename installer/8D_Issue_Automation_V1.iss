[Setup]
AppId={{8D-ISSUE-AUTOMATION-V1}}
AppName=8D Issue Automation
AppVersion=1.0
AppPublisher=LG Energy Solution
DefaultDirName={localappdata}\Programs\8D Issue Automation
DefaultGroupName=8D Issue Automation
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
OutputDir=output
OutputBaseFilename=8D_Issue_Automation_V1_Setup
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
UninstallDisplayName=8D Issue Automation
SetupLogging=yes

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "바탕화면 바로가기 만들기"; GroupDescription: "추가 바로가기:"; Flags: unchecked

[Files]
Source: "..\dist\8D_Issue_Automation_V1.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\8D Issue Automation"; Filename: "{app}\8D_Issue_Automation_V1.exe"
Name: "{autodesktop}\8D Issue Automation"; Filename: "{app}\8D_Issue_Automation_V1.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\8D_Issue_Automation_V1.exe"; Description: "8D Issue Automation 실행"; Flags: nowait postinstall skipifsilent

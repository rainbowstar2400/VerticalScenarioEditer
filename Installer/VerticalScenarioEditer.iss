[Setup]
AppName=VerticalScenarioEditer
AppVersion=1.0.0
DefaultDirName={pf}\VerticalScenarioEditer
DefaultGroupName=VerticalScenarioEditer
SetupIconFile=..\icons\VSE_icon_2.ico
UninstallDisplayIcon={app}\VerticalScenarioEditer.exe
OutputDir=output
OutputBaseFilename=VerticalScenarioEditer-Setup
Compression=lzma
SolidCompression=yes
WizardStyle=modern
[Languages]
Name: "ja"; MessagesFile: "compiler:Languages\Japanese.isl"

[Files]
Source: "..\publish\*"; DestDir: "{app}"; Flags: recursesubdirs createallsubdirs

[Icons]
Name: "{group}\VerticalScenarioEditer"; Filename: "{app}\VerticalScenarioEditer.exe"
Name: "{commondesktop}\VerticalScenarioEditer"; Filename: "{app}\VerticalScenarioEditer.exe"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "デスクトップにアイコンを作成する"; GroupDescription: "追加タスク"

#define MyAppName "PyWeChat GUI 工作台"
#ifndef MyAppVersion
  #define MyAppVersion "0.1.0"
#endif
#ifndef MyAppPublisher
  #define MyAppPublisher "Hello-Mr-Crab"
#endif
#ifndef MyAppURL
  #define MyAppURL "https://github.com/Hello-Mr-Crab/pywechat"
#endif
#ifndef MyAppExeName
  #define MyAppExeName "pywechat-gui.exe"
#endif
#ifndef MyOutputBaseFilename
  #define MyOutputBaseFilename "pywechat-gui-setup"
#endif
#ifndef MySourceDir
  #define MySourceDir "..\dist\pywechat-gui"
#endif

[Setup]
AppId={{5AA73530-FFB9-4E03-A1A8-1A4577BE9C44}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={autopf}\PyWeChat GUI
DefaultGroupName=PyWeChat GUI 工作台
UninstallDisplayIcon={app}\{#MyAppExeName}
SetupIconFile=assets\pywechat-gui.ico
LicenseFile=..\LICENSE
OutputDir=..\dist
OutputBaseFilename={#MyOutputBaseFilename}
Compression=lzma
SolidCompression=yes
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64compatible
PrivilegesRequired=admin

[Languages]
Name: "chinesesimp"; MessagesFile: "ChineseSimplified.isl"

[Tasks]
Name: "desktopicon"; Description: "创建桌面快捷方式"; GroupDescription: "附加任务："; Flags: unchecked

[Files]
Source: "{#MySourceDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\PyWeChat GUI 工作台"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\PyWeChat GUI 工作台"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "启动 PyWeChat GUI"; Flags: nowait postinstall skipifsilent

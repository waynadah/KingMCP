[Setup]
AppId=DesktopAI-MCP
AppName=Desktop AI Agent
AppVersion=5.3
DefaultDirName={localappdata}\Programs\DesktopAI
DefaultGroupName=Desktop AI Agent
PrivilegesRequired=lowest
OutputDir=installer
OutputBaseFilename=DesktopAI-Setup-5.3
SetupIconFile=desktopai.ico
UninstallDisplayIcon={app}\DesktopAI.exe
Compression=lzma2
SolidCompression=yes
[Files]
Source: "dist\DesktopAI\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; Flags: unchecked
[Icons]
Name: "{group}\Desktop AI Agent"; Filename: "{app}\DesktopAI.exe"; WorkingDir: "{app}"
Name: "{autodesktop}\Desktop AI Agent"; Filename: "{app}\DesktopAI.exe"; WorkingDir: "{app}"; Tasks: desktopicon
[Run]
Filename: "{app}\DesktopAI.exe"; Description: "Launch Desktop AI Agent"; Flags: nowait postinstall skipifsilent

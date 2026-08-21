; Outlook PST 搜尋工具 — Inno Setup 安裝腳本（Inno Setup 6）
; 編譯：ISCC.exe installer.iss（於 build 目錄執行；成品輸出到 ..\dist）
#define MyAppName "Outlook PST 搜尋工具"
#define MyAppVersion "1.0.5"
#define MyAppPublisher "PstSearchTool"
#define MyAppExeName "PstSearchTool.exe"

[Setup]
AppId={{F4A9C3E2-1B2D-4E6A-9F0C-8D7B6A5C4E3D}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={localappdata}\Programs\PstSearchTool
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
OutputDir=..\dist
OutputBaseFilename=PstSearchTool-Setup-{#MyAppVersion}
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
UninstallDisplayIcon={app}\{#MyAppExeName}

[Files]
Source: "..\dist\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Tasks]
Name: "desktopicon"; Description: "建立桌面捷徑"; GroupDescription: "其他："; Flags: unchecked

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "立即執行 {#MyAppName}"; Flags: nowait postinstall skipifsilent

[Code]
function IsDotNet48Installed(): Boolean;
var
  ver: Cardinal;
begin
  Result := RegQueryDWordValue(HKLM, 'SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full', 'Release', ver) and (ver >= 528040);
end;

function InitializeSetup(): Boolean;
begin
  if not IsDotNet48Installed() then
  begin
    MsgBox('此程式需要 .NET Framework 4.8（Windows 7 SP1 / 10 / 11 適用）。' + #13#10 +
           '請先安裝 .NET Framework 4.8：https://dotnet.microsoft.com/download/dotnet-framework/net48' + #13#10 +
           '安裝完成後再重新執行本安裝程式。',
           mbError, MB_OK);
    Result := False;
  end
  else
    Result := True;
end;

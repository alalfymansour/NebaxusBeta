; NebaxusBeta — Inno Setup installer script
; Requires: Inno Setup 6+ (https://jrsoftware.org/isinfo.php)
; Usage: ISCC.exe installer.iss

#define MyAppName "NebaxusBeta"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "Nebaxus"
#define MyAppURL "https://nebaxus.com"
#define MyAppExeName "NebaxusBeta.exe"

[Setup]
AppId={{8A2E5B3C-9D4F-4E6A-8B1C-2D3E4F5A6B7C}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={localappdata}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputDir=dist
OutputBaseFilename=NebaxusBeta_Setup
SetupIconFile=nebaxus.ico
UninstallDisplayIcon={app}\nebaxus.ico
Compression=lzma2/ultra64
SolidCompression=yes
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=commandline
ArchitecturesInstallIn64BitMode=x64compatible
MinVersion=10.0.17763

[Languages]
Name: "arabic"; MessagesFile: "compiler:Languages\Arabic.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "إنشاء اختصار على سطح المكتب"; GroupDescription: "اختصارات:"

[Files]
Source: "dist\desktop.dist\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion
Source: "nebaxus.ico"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\nebaxus.ico"
Name: "{group}\إلغاء التثبيت"; Filename: "{uninstallexe}"
Name: "{commondesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\nebaxus.ico"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "تشغيل {#MyAppName}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
Type: filesandordirs; Name: "{localappdata}\{#MyAppName}\logs"
Type: filesandordirs; Name: "{localappdata}\{#MyAppName}\backups"
Type: filesandordirs; Name: "{localappdata}\{#MyAppName}\exports"
; Keep database and instance folder on uninstall (user data)
; Type: filesandordirs; Name: "{localappdata}\{#MyAppName}\instance"

[Code]
function InitializeUninstall: Boolean;
begin
  if MsgBox('سيتم إلغاء تثبيت البرنامج. البيانات المحفوظة (قاعدة البيانات والنسخ الاحتياطية) ستبقى في مجلد التطبيق.', mbConfirmation, MB_YESNO) = IDYES then
    Result := True
  else
    Result := False;
end;

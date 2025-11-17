; =========================
; FacturaScan - Inno Setup
; =========================
#define MyAppName "FacturaScan"
#define MyAppVersion "1.9.3"
#define MyAppPublisher "Departamento TI"
#define MyAppExeName "FacturaScan.exe"

[Setup]
; ⚠️ Corregido
AppId={{75FCFDE0-F147-4AAD-86B0-9EA71650ED3A}}

AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}

; --- Si quieres seguir en C:\ (puede requerir UAC en algunos PCs)
DefaultDirName=C:\{#MyAppName}
UsePreviousAppDir=yes
; --- Si quieres evitar UAC SIEMPRE, usa estas dos y comenta las dos de arriba:
;PrivilegesRequired=lowest
;DefaultDirName={localappdata}\{#MyAppName}

UninstallDisplayIcon={app}\{#MyAppExeName}

; Arquitectura
ArchitecturesAllowed=x64
ArchitecturesInstallIn64BitMode=x64

; UI
DisableProgramGroupPage=yes
WizardStyle=modern

; Salida del instalador
OutputDir=C:\Users\TEBA\Desktop
OutputBaseFilename=FacturaScan-{#MyAppVersion}-Setup

; Icono
SetupIconFile=C:\Users\TEBA\Desktop\Nueva carpeta\Control documental 26-08-2025\FacturaScan.ico

; Compresión y logging
SolidCompression=yes
SetupLogging=yes

; Cierre automático y silencioso de FacturaScan
CloseApplications=force
CloseApplicationsFilter=FacturaScan.exe
RestartApplications=yes
AlwaysRestart=no

[Languages]
Name: "spanish"; MessagesFile: "compiler:Languages\Spanish.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
; Ejecutable principal explícito
Source: "C:\ProyectoFacturaScan\FacturaScan\dist\FacturaScan.dist\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion
; Resto de dependencias generadas por Nuitka
Source: "C:\ProyectoFacturaScan\FacturaScan\dist\FacturaScan.dist\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
; ▶ Relanza SIEMPRE (también en /VERYSILENT): sin postinstall
Filename: "{app}\{#MyAppExeName}"; Flags: nowait

; Si prefieres solo en modo normal: usa esto en su lugar
; Filename: "{app}\{#MyAppExeName}"; Flags: nowait postinstall skipifsilent

[Code]
procedure KillIfRunning(const ExeName: string);
var
  ResultCode: Integer;
begin
  if FileExists(ExpandConstant('{sys}\taskkill.exe')) then
    Exec(ExpandConstant('{sys}\taskkill.exe'),
         '/IM "' + ExeName + '" /F',
         '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
end;

procedure InitializeWizard();
begin
  KillIfRunning('FacturaScan.exe');
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssInstall then
    KillIfRunning('FacturaScan.exe');
end;

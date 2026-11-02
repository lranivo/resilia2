; ============================================================
; Installateur Resilia2  (Inno Setup 6)  ->  UN SEUL fichier .exe
;
; Place ce fichier à la racine du projet, à côté de :
;     dist\resilia2\resilia2.exe  +  dist\resilia2\_internal\
;     resilia.ico
; Compilation : double-clic sur build_installer.bat
; Résultat    : installer\Resilia2_Setup_<version>.exe
; ============================================================

#define AppName      "Resilia2"
#define AppVersion   "1.0.0"
#define AppPublisher "OBS - MSC IOO"
#define AppExe       "resilia2.exe"

[Setup]
AppId={{B7C1E5A2-4D3F-4E8B-9A61-2F5D8C0E7A19}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
DefaultDirName={autopf}\{#AppName}
DefaultGroupName={#AppName}
DisableProgramGroupPage=yes
UninstallDisplayIcon={app}\{#AppExe}
SetupIconFile=resilia.ico
OutputDir=installer
OutputBaseFilename=Resilia2_Setup_{#AppVersion}
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
; Pas besoin d'être administrateur : installation dans le profil de
; l'utilisateur par défaut (l'option « pour tous les utilisateurs »
; reste proposée si les droits le permettent).
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
; Le lanceur crée ce mutex : l'installateur demande de le fermer.
AppMutex=Resilia2LauncherMutex
CloseApplications=yes
RestartApplications=no

[Languages]
Name: "french"; MessagesFile: "compiler:Languages\French.isl"

[Tasks]
Name: "desktopicon"; Description: "Créer un raccourci sur le Bureau"; \
  GroupDescription: "Raccourcis :"; Flags: checkedonce

[Files]
Source: "dist\resilia2\*"; DestDir: "{app}"; \
  Flags: recursesubdirs createallsubdirs ignoreversion
Source: "resilia.ico"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#AppName}"; Filename: "{app}\{#AppExe}"; \
  IconFilename: "{app}\resilia.ico"
Name: "{group}\Désinstaller {#AppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExe}"; \
  IconFilename: "{app}\resilia.ico"; Tasks: desktopicon

[Run]
Filename: "{app}\{#AppExe}"; Description: "Lancer {#AppName}"; \
  Flags: nowait postinstall skipifsilent

[Code]
{ ---------- Prérequis externes (détection, sans blocage) ---------- }

function ChromeInstalle: Boolean;
begin
  Result :=
    RegKeyExists(HKLM64, 'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe') or
    RegKeyExists(HKLM32, 'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe') or
    RegKeyExists(HKCU,   'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe') or
    FileExists(ExpandConstant('{commonpf64}\Google\Chrome\Application\chrome.exe')) or
    FileExists(ExpandConstant('{commonpf32}\Google\Chrome\Application\chrome.exe')) or
    FileExists(ExpandConstant('{localappdata}\Google\Chrome\Application\chrome.exe'));
end;

function ExcelInstalle: Boolean;
begin
  Result := RegKeyExists(HKCR, 'Excel.Application\CLSID');
end;

function InitializeSetup: Boolean;
var
  Manques: String;
begin
  Result := True;
  Manques := '';

  if not ChromeInstalle then
    Manques := Manques + #13#10 + ' - Google Chrome (utilisé pour piloter WASAC / Salto)';

  if not ExcelInstalle then
    Manques := Manques + #13#10 + ' - Microsoft Excel (lecture et écriture de resil.xlsx)';

  if (Manques <> '') and (not WizardSilent) then
    Result := MsgBox(
      'Prérequis non détectés sur ce poste :' + Manques + #13#10#13#10 +
      'Resilia2 ne fonctionnera pas sans eux. Continuer l''installation ?',
      mbConfirmation, MB_YESNO) = IDYES;
end;

{ ---------- Arrêt des processus Resilia2 restants (mise à jour) ---------- }

function PrepareToInstall(var NeedsRestart: Boolean): String;
var
  Code: Integer;
begin
  Exec(ExpandConstant('{sys}\taskkill.exe'), '/F /T /IM {#AppExe}',
       '', SW_HIDE, ewWaitUntilTerminated, Code);
  Result := '';
end;

{ ---------- Désinstallation ---------- }

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  Code: Integer;
  Donnees: String;
begin
  if CurUninstallStep = usUninstall then
    Exec(ExpandConstant('{sys}\taskkill.exe'), '/F /T /IM {#AppExe}',
         '', SW_HIDE, ewWaitUntilTerminated, Code);

  if CurUninstallStep = usPostUninstall then
  begin
    Donnees := ExpandConstant('{localappdata}\Resilia2');

    if DirExists(Donnees) and (not UninstallSilent) then
      if MsgBox('Supprimer aussi les données utilisateur (journaux, ' +
                'diagnostics, configuration) ?' + #13#10 + Donnees,
                mbConfirmation, MB_YESNO or MB_DEFBUTTON2) = IDYES then
        DelTree(Donnees, True, True, True);
  end;
end;

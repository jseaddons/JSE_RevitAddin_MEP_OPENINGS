; JSE Revit MEP Openings Add-in Installer Script for Inno Setup
; Multi-version installer with automatic Revit version detection
; Uses subfolder structure to avoid cluttering the Addins folder
; Supports: Revit 2022, 2023, 2024, 2025, 2026
; Requires Inno Setup 6+
;
; Deploy sources (build each project separately; no merge required):
;   MEP Openings:  deploy\<year>\  (populated by JSE_MEPOPENING_23 PostBuild)
;   Parameter Svc: ..\JSE_Parameter_Service\deploy\<year>\  (PS PostBuild or BUILD_ALL_FOR_INSTALLER)

#define MepDeploy "deploy"
#define PsDeploy "..\JSE_Parameter_Service\deploy"

#define MyAppName "JSE MEP OPENING"
#define MyAppVersion "2.0"
#define MyAppPublisher "JSE"
#define MyAppURL "https://www.jseaddons.com"

[Setup]
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}/support
AppUpdatesURL={#MyAppURL}/updates
DefaultDirName={userappdata}\JSE\MEP_Openings
DisableProgramGroupPage=yes
OutputBaseFilename=MEP_OPENING_Installer_v{#MyAppVersion}
Compression=lzma
SolidCompression=yes
PrivilegesRequired=lowest
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
UninstallDisplayIcon={app}\Resources\RibbonIcon32.png

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Types]
Name: "full"; Description: "Full installation"
Name: "custom"; Description: "Custom installation"; Flags: iscustom

[Components]
Name: "main"; Description: "Main Application"; Types: full custom; Flags: fixed
Name: "r22"; Description: "Revit 2022 Support"; Types: full custom
Name: "r23"; Description: "Revit 2023 Support"; Types: full custom
Name: "r24"; Description: "Revit 2024 Support"; Types: full custom
Name: "r25"; Description: "Revit 2025 Support"; Types: full custom
Name: "r26"; Description: "Revit 2026 Support"; Types: full custom

[Files]
; === REVIT 2022 - MEP Openings (deploy\) ===
Source: "{#MepDeploy}\2022\JSE_RevitAddin_MEP_OPENINGS.addin"; DestDir: "{userappdata}\Autodesk\Revit\Addins\2022"; Components: r22; Flags: ignoreversion
Source: "{#MepDeploy}\2022\JSE_MEP_OPENINGS\*.dll"; DestDir: "{userappdata}\Autodesk\Revit\Addins\2022\JSE_MEP_OPENINGS"; Components: r22; Flags: ignoreversion; Excludes: "JSE_Parameter_Service.dll"
Source: "{#MepDeploy}\2022\JSE_MEP_OPENINGS\*.config"; DestDir: "{userappdata}\Autodesk\Revit\Addins\2022\JSE_MEP_OPENINGS"; Components: r22; Flags: ignoreversion skipifsourcedoesntexist
Source: "{#MepDeploy}\2022\JSE_MEP_OPENINGS\x64\SQLite.Interop.dll"; DestDir: "{userappdata}\Autodesk\Revit\Addins\2022\JSE_MEP_OPENINGS\x64"; Components: r22; Flags: ignoreversion skipifsourcedoesntexist
; === REVIT 2022 - Parameter Service (JSE_Parameter_Service\deploy\) ===
Source: "{#PsDeploy}\2022\JSE_Parameter_Service.dll"; DestDir: "{userappdata}\Autodesk\Revit\Addins\2022\JSE_MEP_OPENINGS"; Components: r22; Flags: ignoreversion skipifsourcedoesntexist

; === REVIT 2023 - MEP Openings ===
Source: "{#MepDeploy}\2023\JSE_RevitAddin_MEP_OPENINGS.addin"; DestDir: "{userappdata}\Autodesk\Revit\Addins\2023"; Components: r23; Flags: ignoreversion
Source: "{#MepDeploy}\2023\JSE_MEP_OPENINGS\*.dll"; DestDir: "{userappdata}\Autodesk\Revit\Addins\2023\JSE_MEP_OPENINGS"; Components: r23; Flags: ignoreversion; Excludes: "JSE_Parameter_Service.dll"
Source: "{#MepDeploy}\2023\JSE_MEP_OPENINGS\*.config"; DestDir: "{userappdata}\Autodesk\Revit\Addins\2023\JSE_MEP_OPENINGS"; Components: r23; Flags: ignoreversion skipifsourcedoesntexist
Source: "{#MepDeploy}\2023\JSE_MEP_OPENINGS\x64\SQLite.Interop.dll"; DestDir: "{userappdata}\Autodesk\Revit\Addins\2023\JSE_MEP_OPENINGS\x64"; Components: r23; Flags: ignoreversion skipifsourcedoesntexist
; === REVIT 2023 - Parameter Service ===
Source: "{#PsDeploy}\2023\JSE_Parameter_Service.dll"; DestDir: "{userappdata}\Autodesk\Revit\Addins\2023\JSE_MEP_OPENINGS"; Components: r23; Flags: ignoreversion skipifsourcedoesntexist

; === REVIT 2024 - MEP Openings ===
Source: "{#MepDeploy}\2024\JSE_RevitAddin_MEP_OPENINGS.addin"; DestDir: "{userappdata}\Autodesk\Revit\Addins\2024"; Components: r24; Flags: ignoreversion
Source: "{#MepDeploy}\2024\JSE_MEP_OPENINGS\*.dll"; DestDir: "{userappdata}\Autodesk\Revit\Addins\2024\JSE_MEP_OPENINGS"; Components: r24; Flags: ignoreversion; Excludes: "JSE_Parameter_Service.dll"
Source: "{#MepDeploy}\2024\JSE_MEP_OPENINGS\*.config"; DestDir: "{userappdata}\Autodesk\Revit\Addins\2024\JSE_MEP_OPENINGS"; Components: r24; Flags: ignoreversion skipifsourcedoesntexist
Source: "{#MepDeploy}\2024\JSE_MEP_OPENINGS\x64\SQLite.Interop.dll"; DestDir: "{userappdata}\Autodesk\Revit\Addins\2024\JSE_MEP_OPENINGS\x64"; Components: r24; Flags: ignoreversion skipifsourcedoesntexist
; === REVIT 2024 - Parameter Service ===
Source: "{#PsDeploy}\2024\JSE_Parameter_Service.dll"; DestDir: "{userappdata}\Autodesk\Revit\Addins\2024\JSE_MEP_OPENINGS"; Components: r24; Flags: ignoreversion skipifsourcedoesntexist

; === REVIT 2025 - MEP Openings ===
Source: "{#MepDeploy}\2025\JSE_RevitAddin_MEP_OPENINGS.addin"; DestDir: "{userappdata}\Autodesk\Revit\Addins\2025"; Components: r25; Flags: ignoreversion
Source: "{#MepDeploy}\2025\JSE_MEP_OPENINGS\*.dll"; DestDir: "{userappdata}\Autodesk\Revit\Addins\2025\JSE_MEP_OPENINGS"; Components: r25; Flags: ignoreversion; Excludes: "JSE_Parameter_Service.dll"
Source: "{#MepDeploy}\2025\JSE_MEP_OPENINGS\*.json"; DestDir: "{userappdata}\Autodesk\Revit\Addins\2025\JSE_MEP_OPENINGS"; Components: r25; Flags: ignoreversion skipifsourcedoesntexist
Source: "{#MepDeploy}\2025\JSE_MEP_OPENINGS\*.config"; DestDir: "{userappdata}\Autodesk\Revit\Addins\2025\JSE_MEP_OPENINGS"; Components: r25; Flags: ignoreversion skipifsourcedoesntexist
; === REVIT 2025 - Parameter Service ===
Source: "{#PsDeploy}\2025\JSE_Parameter_Service.dll"; DestDir: "{userappdata}\Autodesk\Revit\Addins\2025\JSE_MEP_OPENINGS"; Components: r25; Flags: ignoreversion skipifsourcedoesntexist

; === REVIT 2026 - MEP Openings ===
Source: "{#MepDeploy}\2026\JSE_RevitAddin_MEP_OPENINGS.addin"; DestDir: "{userappdata}\Autodesk\Revit\Addins\2026"; Components: r26; Flags: ignoreversion
Source: "{#MepDeploy}\2026\JSE_MEP_OPENINGS\*.dll"; DestDir: "{userappdata}\Autodesk\Revit\Addins\2026\JSE_MEP_OPENINGS"; Components: r26; Flags: ignoreversion; Excludes: "JSE_Parameter_Service.dll"
Source: "{#MepDeploy}\2026\JSE_MEP_OPENINGS\*.json"; DestDir: "{userappdata}\Autodesk\Revit\Addins\2026\JSE_MEP_OPENINGS"; Components: r26; Flags: ignoreversion skipifsourcedoesntexist
Source: "{#MepDeploy}\2026\JSE_MEP_OPENINGS\*.config"; DestDir: "{userappdata}\Autodesk\Revit\Addins\2026\JSE_MEP_OPENINGS"; Components: r26; Flags: ignoreversion skipifsourcedoesntexist
; === REVIT 2026 - Parameter Service ===
Source: "{#PsDeploy}\2026\JSE_Parameter_Service.dll"; DestDir: "{userappdata}\Autodesk\Revit\Addins\2026\JSE_MEP_OPENINGS"; Components: r26; Flags: ignoreversion skipifsourcedoesntexist

; === SHARED RESOURCES ===
Source: "Resources\*.rfa"; DestDir: "{userappdata}\Autodesk\Revit\Addins\2022\JSE_MEP_OPENINGS\Resources"; Components: r22; Flags: ignoreversion recursesubdirs
Source: "Resources\*.rfa"; DestDir: "{userappdata}\Autodesk\Revit\Addins\2023\JSE_MEP_OPENINGS\Resources"; Components: r23; Flags: ignoreversion recursesubdirs
Source: "Resources\*.rfa"; DestDir: "{userappdata}\Autodesk\Revit\Addins\2024\JSE_MEP_OPENINGS\Resources"; Components: r24; Flags: ignoreversion recursesubdirs
Source: "Resources\*.rfa"; DestDir: "{userappdata}\Autodesk\Revit\Addins\2025\JSE_MEP_OPENINGS\Resources"; Components: r25; Flags: ignoreversion recursesubdirs
Source: "Resources\*.rfa"; DestDir: "{userappdata}\Autodesk\Revit\Addins\2026\JSE_MEP_OPENINGS\Resources"; Components: r26; Flags: ignoreversion recursesubdirs

; === CONFIG FILES (MEP deploy only) ===
Source: "{#MepDeploy}\2022\JSE_MEP_OPENINGS\*.config"; DestDir: "{userappdata}\Autodesk\Revit\Addins\2022\JSE_MEP_OPENINGS"; Components: r22; Flags: skipifsourcedoesntexist
Source: "{#MepDeploy}\2023\JSE_MEP_OPENINGS\*.config"; DestDir: "{userappdata}\Autodesk\Revit\Addins\2023\JSE_MEP_OPENINGS"; Components: r23; Flags: skipifsourcedoesntexist
Source: "{#MepDeploy}\2024\JSE_MEP_OPENINGS\*.config"; DestDir: "{userappdata}\Autodesk\Revit\Addins\2024\JSE_MEP_OPENINGS"; Components: r24; Flags: skipifsourcedoesntexist
Source: "{#MepDeploy}\2025\JSE_MEP_OPENINGS\*.config"; DestDir: "{userappdata}\Autodesk\Revit\Addins\2025\JSE_MEP_OPENINGS"; Components: r25; Flags: skipifsourcedoesntexist
Source: "{#MepDeploy}\2026\JSE_MEP_OPENINGS\*.config"; DestDir: "{userappdata}\Autodesk\Revit\Addins\2026\JSE_MEP_OPENINGS"; Components: r26; Flags: skipifsourcedoesntexist

[Icons]
Name: "{group}\MEP OPENING Uninstall"; Filename: "{uninstallexe}"
Name: "{group}\Visit JSE Website"; Filename: "{#MyAppURL}"

[Code]
var
  Revit2022Installed: Boolean;
  Revit2023Installed: Boolean;
  Revit2024Installed: Boolean;
  Revit2025Installed: Boolean;
  Revit2026Installed: Boolean;

function IsRevitInstalled(Version: String): Boolean;
var
  RevitPath: String;
begin
  RevitPath := ExpandConstant('{commonpf}\Autodesk\Revit ' + Version + '\Revit.exe');
  Result := FileExists(RevitPath);
  
  if not Result then
  begin
    RevitPath := ExpandConstant('{commonpf64}\Autodesk\Revit ' + Version + '\Revit.exe');
    Result := FileExists(RevitPath);
  end;
end;

function InitializeSetup(): Boolean;
var
  DomainName: String;
  AnyRevitInstalled: Boolean;
begin
  Result := False;
  
  // Domain check
  DomainName := UpperCase(GetEnv('USERDOMAIN'));
  if Pos('JSE24', DomainName) = 0 then
  begin
    MsgBox('Installation Error: Cannot be installed without a valid license. Please contact admin@jseeng.in for assistance.', mbError, MB_OK);
    Exit;
  end;
  
  // Check which Revit versions are installed
  Revit2022Installed := IsRevitInstalled('2022');
  Revit2023Installed := IsRevitInstalled('2023');
  Revit2024Installed := IsRevitInstalled('2024');
  Revit2025Installed := IsRevitInstalled('2025');
  Revit2026Installed := IsRevitInstalled('2026');

  AnyRevitInstalled := Revit2022Installed or Revit2023Installed or Revit2024Installed or Revit2025Installed or Revit2026Installed;

  if not AnyRevitInstalled then
  begin
    MsgBox('Warning: No supported Revit versions (2022-2026) were detected on this computer.'#13#10#13#10 +
           'The add-in will be installed, but you need Revit to use it.', mbInformation, MB_OK);
  end;
  
  Result := True;
end;

procedure InitializeWizard;
begin
  // Pre-select components based on installed Revit versions
  if Revit2022Installed then
    WizardForm.ComponentsList.Checked[1] := True  // r22
  else
    WizardForm.ComponentsList.ItemEnabled[1] := False;

  if Revit2023Installed then
    WizardForm.ComponentsList.Checked[2] := True  // r23
  else
    WizardForm.ComponentsList.ItemEnabled[2] := False;

  if Revit2024Installed then
    WizardForm.ComponentsList.Checked[3] := True  // r24
  else
    WizardForm.ComponentsList.ItemEnabled[3] := False;

  if Revit2025Installed then
    WizardForm.ComponentsList.Checked[4] := True  // r25
  else
    WizardForm.ComponentsList.ItemEnabled[4] := False;

  if Revit2026Installed then
    WizardForm.ComponentsList.Checked[5] := True  // r26
  else
    WizardForm.ComponentsList.ItemEnabled[5] := False;
end;

// Unblock files using PowerShell (removes Zone.Identifier)
procedure UnblockFiles(RevitVersion: String);
var
  ResultCode: Integer;
  PowerShellCmd: String;
  AddinFolder: String;
begin
  AddinFolder := ExpandConstant('{userappdata}\Autodesk\Revit\Addins\' + RevitVersion + '\JSE_MEP_OPENINGS');
  
  PowerShellCmd := 'Get-ChildItem -Path "' + AddinFolder + '" -File -Recurse | ' +
                   'Unblock-File -ErrorAction SilentlyContinue';
  
  if Exec('powershell.exe', '-ExecutionPolicy Bypass -Command "' + PowerShellCmd + '"', 
          '', SW_HIDE, ewWaitUntilTerminated, ResultCode) then
  begin
    Log('Unblocked files in: ' + AddinFolder);
  end
  else
  begin
    Log('Failed to unblock files in: ' + AddinFolder);
  end;
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
  InstalledVersions: String;
begin
  if CurStep = ssPostInstall then
  begin
    InstalledVersions := '';

    // Unblock files for each installed version
    if WizardForm.ComponentsList.Checked[1] then
    begin
      UnblockFiles('2022');
      InstalledVersions := InstalledVersions + '2022, ';
    end;

    if WizardForm.ComponentsList.Checked[2] then
    begin
      UnblockFiles('2023');
      InstalledVersions := InstalledVersions + '2023, ';
    end;

    if WizardForm.ComponentsList.Checked[3] then
    begin
      UnblockFiles('2024');
      InstalledVersions := InstalledVersions + '2024, ';
    end;

    if WizardForm.ComponentsList.Checked[4] then
    begin
      UnblockFiles('2025');
      InstalledVersions := InstalledVersions + '2025, ';
    end;

    if WizardForm.ComponentsList.Checked[5] then
    begin
      UnblockFiles('2026');
      InstalledVersions := InstalledVersions + '2026, ';
    end;

    // Remove trailing comma
    if Length(InstalledVersions) > 2 then
      InstalledVersions := Copy(InstalledVersions, 1, Length(InstalledVersions) - 2);

    MsgBox('MEP OPENING Add-in (v' + ExpandConstant('{#MyAppVersion}') + ') installed successfully!'#13#10#13#10 +
           'Installed for Revit: ' + InstalledVersions + #13#10#13#10 +
           'Files are organized in JSE_MEP_OPENINGS subfolder to keep your Addins folder clean.'#13#10 +
           'Please restart Revit to load the add-in.', mbInformation, MB_OK);
  end;
end;

[UninstallDelete]
Type: filesandordirs; Name: "{userappdata}\Autodesk\Revit\Addins\2022\JSE_MEP_OPENINGS"
Type: filesandordirs; Name: "{userappdata}\Autodesk\Revit\Addins\2023\JSE_MEP_OPENINGS"
Type: filesandordirs; Name: "{userappdata}\Autodesk\Revit\Addins\2024\JSE_MEP_OPENINGS"
Type: filesandordirs; Name: "{userappdata}\Autodesk\Revit\Addins\2025\JSE_MEP_OPENINGS"
Type: filesandordirs; Name: "{userappdata}\Autodesk\Revit\Addins\2026\JSE_MEP_OPENINGS"
Type: files; Name: "{userappdata}\Autodesk\Revit\Addins\2022\JSE_RevitAddin_MEP_OPENINGS.addin"
Type: files; Name: "{userappdata}\Autodesk\Revit\Addins\2023\JSE_RevitAddin_MEP_OPENINGS.addin"
Type: files; Name: "{userappdata}\Autodesk\Revit\Addins\2024\JSE_RevitAddin_MEP_OPENINGS.addin"
Type: files; Name: "{userappdata}\Autodesk\Revit\Addins\2025\JSE_RevitAddin_MEP_OPENINGS.addin"
Type: files; Name: "{userappdata}\Autodesk\Revit\Addins\2026\JSE_RevitAddin_MEP_OPENINGS.addin"

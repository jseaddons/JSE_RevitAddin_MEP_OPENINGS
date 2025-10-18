; JSE Revit MEP Openings Add-in Installer Script for Inno Setup
; This script installs the add-in DLL and .addin manifest to both user and machine-wide Revit Addins folders
; Requires Inno Setup 6+


[Setup]
AppName=JSE MEP OPENING (2023-2024)
AppVersion=1.0
AppPublisher=JSE
AppPublisherURL=https://www.jseaddons.com
AppSupportURL=https://www.jseaddons.com/support
AppUpdatesURL=https://www.jseaddons.com/updates
DefaultDirName={userappdata}\Autodesk\Revit\Addins\2024
DisableProgramGroupPage=yes
OutputBaseFilename=MEP_OPENING_Installer
Compression=lzma
SolidCompression=yes
PrivilegesRequired=lowest
; SetupIconFile="JSE.ico" ; Optional: place a JSE.ico in the script folder for branding

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
; Main add-in DLL and manifest
Source: "bin\Release R24\Any CPU\Release R24\JSE_RevitAddin_MEP_OPENINGS.dll"; DestDir: "{userappdata}\Autodesk\Revit\Addins\2024"; Flags: ignoreversion
Source: "bin\Release R24\Any CPU\Release R24\JSE_RevitAddin_MEP_OPENINGS.addin"; DestDir: "{userappdata}\Autodesk\Revit\Addins\2024"; Flags: ignoreversion

; Revit 2023 support
Source: "bin\Release R24\Any CPU\Release R24\JSE_RevitAddin_MEP_OPENINGS.dll"; DestDir: "{userappdata}\Autodesk\Revit\Addins\2023"; Flags: ignoreversion
Source: "bin\Release R24\Any CPU\Release R24\JSE_RevitAddin_MEP_OPENINGS.addin"; DestDir: "{userappdata}\Autodesk\Revit\Addins\2023"; Flags: ignoreversion
; Also include the 2023-targeted release build output
Source: "bin\Release R23\Any CPU\Release R23\JSE_RevitAddin_MEP_OPENINGS.dll"; DestDir: "{userappdata}\Autodesk\Revit\Addins\2023"; Flags: ignoreversion
Source: "bin\Release R23\Any CPU\Release R23\JSE_RevitAddin_MEP_OPENINGS.addin"; DestDir: "{userappdata}\Autodesk\Revit\Addins\2023"; Flags: ignoreversion


[Icons]
Name: "{autoprograms}\MEP OPENING Uninstall"; Filename: "{uninstallexe}"





[Code]

function InitializeSetup(): Boolean;
var
  DomainName: String;
begin
  Result := False;
  // Get the domain name from environment variable
  DomainName := UpperCase(GetEnv('USERDOMAIN'));
  // Check if domain contains "JSE24" (case insensitive)
  if Pos('JSE24', DomainName) > 0 then
  begin
    Result := True;
  end
  else
  begin
    MsgBox('Installation Error: Cannot be installed without a valid license. Please contact admin@jseeng.in for assistance.', mbError, MB_OK);
    Result := False;
  end;
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
  begin
    MsgBox('MEP OPENING Add-in (v1.0) by JSE was installed to your user Addins folders for Revit 2023 and 2024.'#13#10#13#10 +
           'Please restart Revit to load the add-in.'#13#10#13#10 + 'For support, visit https://www.jseaddons.com', mbInformation, MB_OK);
  end;
end;

; JSE_MEP_OPENINGS_23 Inno Setup Script
; This installer copies files to %USERPROFILE%\AppData\Roaming\Autodesk\Revit\Addins\2023 (no admin required)
; Added domain check for "jse24" (case insensitive)

[Setup]
AppName=JSE_MEP_OPENINGS_23
AppVersion=1.0
DefaultDirName={userappdata}\Autodesk\Revit\Addins\2023
DisableDirPage=yes
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
OutputDir=.
OutputBaseFilename=JSE_MEP_OPENINGS_23_Installer
Compression=lzma
SolidCompression=yes

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

[Files]
; Copy the DLL from the latest build path
Source: "bin\Debug R23\Debug R23\JSE_RevitAddin_MEP_OPENINGS.dll"; DestDir: "{app}"; Flags: ignoreversion
; Copy the .addin file to the addins folder
Source: "JSE_RevitAddin_MEP_OPENINGS.addin"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
; No Start Menu or Desktop icons

[Run]
; No post-install actions

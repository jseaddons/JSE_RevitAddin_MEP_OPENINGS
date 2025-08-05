; JSE Revit MEP Openings Add-in Installer Script for Inno Setup
; This script installs the add-in DLL and .addin manifest to both user and machine-wide Revit Addins folders
; Requires Inno Setup 6+


[Setup]
AppName=MEP OPENING
AppVersion=1.0
AppPublisher=JSE
AppPublisherURL=https://www.jseaddons.com
AppSupportURL=https://www.jseaddons.com/support
AppUpdatesURL=https://www.jseaddons.com/updates
DefaultDirName={autopf}\JSE\MEP_OPENING
DisableProgramGroupPage=yes
OutputBaseFilename=MEP_OPENING_Installer
Compression=lzma
SolidCompression=yes
PrivilegesRequired=admin
; SetupIconFile="JSE.ico" ; Optional: place a JSE.ico in the script folder for branding

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
; Main add-in DLL and manifest
Source: "bin\Debug R24\JSE_RevitAddin_MEP_OPENINGS.dll"; DestDir: "{userappdata}\Autodesk\Revit\Addins\2024"; Flags: ignoreversion
Source: "bin\Debug R24\JSE_RevitAddin_MEP_OPENINGS.addin"; DestDir: "{userappdata}\Autodesk\Revit\Addins\2024"; Flags: ignoreversion
Source: "bin\Debug R24\JSE_RevitAddin_MEP_OPENINGS.dll"; DestDir: "{commonappdata}\Autodesk\Revit\Addins\2024"; Flags: ignoreversion
Source: "bin\Debug R24\JSE_RevitAddin_MEP_OPENINGS.addin"; DestDir: "{commonappdata}\Autodesk\Revit\Addins\2024"; Flags: ignoreversion

; Revit 2023 support
Source: "bin\Debug R24\JSE_RevitAddin_MEP_OPENINGS.dll"; DestDir: "{userappdata}\Autodesk\Revit\Addins\2023"; Flags: ignoreversion
Source: "bin\Debug R24\JSE_RevitAddin_MEP_OPENINGS.addin"; DestDir: "{userappdata}\Autodesk\Revit\Addins\2023"; Flags: ignoreversion
Source: "bin\Debug R24\JSE_RevitAddin_MEP_OPENINGS.dll"; DestDir: "{commonappdata}\Autodesk\Revit\Addins\2023"; Flags: ignoreversion
Source: "bin\Debug R24\JSE_RevitAddin_MEP_OPENINGS.addin"; DestDir: "{commonappdata}\Autodesk\Revit\Addins\2023"; Flags: ignoreversion


[Icons]
Name: "{autoprograms}\MEP OPENING Uninstall"; Filename: "{uninstallexe}"




[Code]
procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
    MsgBox('MEP OPENING Add-in (v1.0) by JSE was installed to both user and machine-wide Revit Addins folders.\n\nPlease restart Revit to load the add-in.\n\nFor support, visit https://www.jseaddons.com', mbInformation, MB_OK);
end;

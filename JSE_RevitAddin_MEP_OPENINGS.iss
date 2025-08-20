; JSE Revit MEP Openings Add-in Installer Script for Inno Setup
; This script installs the add-in DLL and .addin manifest to machine-wide Revit Addins folders
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
; PrivilegesRequired=admin
; SetupIconFile="JSE.ico" ; Optional: place a JSE.ico in the script folder for branding

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
; Build-time switches
; By default the installer will include both Revit 2024 and 2023 targets.
; You can override these at compile time with ISCC by passing /D flags. Examples:
;  - Build only for 2024:  ISCC.exe /DINSTALL_FOR_2023=0 /DINSTALL_FOR_2024=1 JSE_RevitAddin_MEP_OPENINGS.iss
;  - Build only for 2023:  ISCC.exe /DINSTALL_FOR_2024=0 /DINSTALL_FOR_2023=1 JSE_RevitAddin_MEP_OPENINGS.iss
;  - Build for both (default): ISCC.exe JSE_RevitAddin_MEP_OPENINGS.iss

; Default switches (can be overridden via ISCC /D...)
#ifndef INSTALL_FOR_2024
#define INSTALL_FOR_2024 1
#endif
#ifndef INSTALL_FOR_2023
#define INSTALL_FOR_2023 1
#endif

; Main add-in DLL and manifest (using build output under bin\Debug R24\Debug R24)
; Note: Using only machine-wide installation to avoid admin/user area conflicts
#if INSTALL_FOR_2024
Source: "bin\Debug R24\Debug R24\JSE_RevitAddin_MEP_OPENINGS.dll"; DestDir: "{commonappdata}\Autodesk\Revit\Addins\2024"; Flags: ignoreversion
Source: "bin\Debug R24\Debug R24\JSE_RevitAddin_MEP_OPENINGS.addin"; DestDir: "{commonappdata}\Autodesk\Revit\Addins\2024"; Flags: ignoreversion
#endif

#if INSTALL_FOR_2023
Source: "bin\Debug R24\Debug R24\JSE_RevitAddin_MEP_OPENINGS.dll"; DestDir: "{commonappdata}\Autodesk\Revit\Addins\2023"; Flags: ignoreversion
Source: "bin\Debug R24\Debug R24\JSE_RevitAddin_MEP_OPENINGS.addin"; DestDir: "{commonappdata}\Autodesk\Revit\Addins\2023"; Flags: ignoreversion
#endif


[Icons]
Name: "{autoprograms}\MEP OPENING Uninstall"; Filename: "{uninstallexe}"




[Code]
procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
    MsgBox('MEP OPENING Add-in (v1.0) by JSE was installed to machine-wide Revit Addins folders.\n\nPlease restart Revit to load the add-in.\n\nFor support, visit https://www.jseaddons.com', mbInformation, MB_OK);
end;

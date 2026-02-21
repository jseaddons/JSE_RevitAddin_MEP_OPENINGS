# MEP Openings Multi-Version Installer

## Overview

This installer supports **Revit 2023, 2024, 2025, and 2026** with automatic detection of installed versions.

## Features

1. **Automatic Revit Version Detection** - Only shows versions that are installed
2. **File Unblocking** - Automatically removes Windows "Zone.Identifier" to prevent security warnings
3. **Selective Installation** - Users can choose which Revit versions to install for
4. **SQLite Support** - Includes SQLite DLLs only for R23 and R24 (newer versions use different database)
5. **Parameter Service Integration** - Includes JSE_Parameter_Service.dll for all versions
6. **Organized Folder Structure** - All DLLs go into `JSE_MEP_OPENINGS` subfolder to avoid cluttering Addins folder

## Quick Start - One Command Build!

```powershell
.\BUILD_ALL_FOR_INSTALLER.ps1
```

This single command:
1. Builds JSE_Parameter_Service (all 4 versions)
2. Copies Parameter Service DLLs to main project
3. Builds JSE_MEPOPENING_23 (all 4 versions)
4. Verifies all files are ready

Then just open `JSE_RevitAddin_MEP_OPENINGS.iss` in Inno Setup and compile!

## Manual Build (if needed)

### Step 1: Build Parameter Service

```powershell
cd C:\Jse_Developments\JSE_Parameter_Service
.\Build-AllVersions.ps1
```

### Step 2: Copy to Main Project

```powershell
cd C:\Jse_Developments\JSE_MEPOPENING_23
.\Build-ParameterService-and-Copy.ps1
```

Or use batch file:
```batch
copy-parameter-service.bat
```

### Step 3: Build Main Project

```powershell
.\Build-AllVersions.ps1
```

### Step 4: Build Installer

Open `JSE_RevitAddin_MEP_OPENINGS.iss` in Inno Setup and compile.

## Folder Structure After Build

```
JSE_MEPOPENING_23/
├── bin/
│   ├── Debug R23/              ← .NET Framework 4.8
│   │   └── Debug R23/          ← Actual output folder
│   │       ├── JSE_RevitAddin_MEP_OPENINGS.dll
│   │       ├── JSE_Parameter_Service.dll      ← Copied from Parameter Service
│   │       ├── CommunityToolkit.Mvvm.dll      ← NuGet (auto-copied)
│   │       ├── Serilog.dll                    ← NuGet (auto-copied)
│   │       ├── System.Text.Json.dll           ← NuGet (auto-copied)
│   │       ├── System.Data.SQLite.dll         ← SQLite
│   │       └── x64/
│   │           └── SQLite.Interop.dll
│   ├── Debug R24/              ← .NET 6
│   │   └── Debug R24/          ← Same structure as R23
│   ├── Debug R25/              ← .NET 8
│   │   └── Debug R25/          
│   │       ├── JSE_RevitAddin_MEP_OPENINGS.dll
│   │       ├── JSE_Parameter_Service.dll      ← Copied from Parameter Service
│   │       └── publish/        ← Dependencies here for .NET 8
│   │           ├── CommunityToolkit.Mvvm.dll
│   │           ├── Serilog.dll
│   │           ├── Microsoft.Data.Sqlite.dll
│   │           └── ...
│   └── Debug R26/              ← .NET 8
│       └── Debug R26/          
│           └── ... (same as R25)
├── Resources/
│   └── *.rfa files
├── JSE_RevitAddin_MEP_OPENINGS.iss
├── BUILD_ALL_FOR_INSTALLER.ps1    ← Master build script
├── Build-AllVersions.ps1          ← Main project build
├── Build-ParameterService-and-Copy.ps1
└── copy-parameter-service.bat
```

## Installation Structure

The ISS installer creates this organized structure:

```
%APPDATA%\Autodesk\Revit\Addins\2023\
├── JSE_RevitAddin_MEP_OPENINGS.addin     ← Points to subfolder
└── JSE_MEP_OPENINGS｜                      ← All DLLs here
    ├── JSE_RevitAddin_MEP_OPENINGS.dll
    ├── JSE_Parameter_Service.dll
    ├── CommunityToolkit.Mvvm.dll
    ├── Serilog.dll
    └── ...
```

## How It Works

### During Installation

1. **Domain Check** - Verifies `USERDOMAIN` contains "JSE24"
2. **Revit Detection** - Checks which Revit versions are installed
3. **Component Selection** - Pre-selects installed versions, disables others
4. **File Copy** - Copies files to `JSE_MEP_OPENINGS` subfolder for each version
5. **File Unblocking** - Runs PowerShell to remove Zone.Identifier from all files

### File Unblocking

The installer automatically runs PowerShell to unblock files:

```powershell
Get-ChildItem -Path "$env:APPDATA\Autodesk\Revit\Addins\202X\JSE_MEP_OPENINGS" -File -Recurse | Unblock-File
```

This prevents the "attempt to load assembly from network location" error.

## Troubleshooting

### "JSE_Parameter_Service.dll not found"

Run the master build script:
```powershell
.\BUILD_ALL_FOR_INSTALLER.ps1
```

Or manually build and copy:
```powershell
cd C:\Jse_Developments\JSE_Parameter_Service
.\Build-AllVersions.ps1
cd C:\Jse_Developments\JSE_MEPOPENING_23
.\Build-ParameterService-and-Copy.ps1
```

### "No Revit versions detected"

The installer will still allow installation, but warn the user. They can select versions manually if Revit is installed in a non-standard location.

### "Files still blocked after installation"

Manually unblock files: Right-click DLL → Properties → Unblock

## Security Notes

- The installer does NOT require Administrator privileges (`PrivilegesRequired=lowest`)
- Files are installed to user-specific folder (`{userappdata}`) not ProgramData
- Domain validation ensures only authorized users can install
- File unblocking prevents runtime security exceptions

## Project Dependencies

**JSE_MEPOPENING_23** requires **JSE_Parameter_Service** DLL to be present in each output folder before building the installer.

The `BUILD_ALL_FOR_INSTALLER.ps1` script handles this automatically by:
1. Building Parameter Service project (all 4 versions)
2. Copying the DLLs to the correct folders in main project
3. Building the main project (all 4 versions)
4. Verifying everything is ready for ISS installer

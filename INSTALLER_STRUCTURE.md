# Installer Folder Structure

## New Organized Structure

To avoid cluttering the Revit Addins folder, all DLLs are now placed in a subfolder:

### Installation Layout

```
%APPDATA%\Autodesk\Revit\Addins\
├── 2023\
│   ├── JSE_RevitAddin_MEP_OPENINGS.addin     ← Addin manifest (points to subfolder)
│   └── JSE_MEP_OPENINGS
│       ├── JSE_RevitAddin_MEP_OPENINGS.dll   ← Main add-in DLL
│       ├── JSE_Parameter_Service.dll         ← Parameter Service
│       ├── CommunityToolkit.Mvvm.dll         ← NuGet dependencies
│       ├── Serilog.dll
│       ├── Serilog.Sinks.Debug.dll
│       ├── System.Text.Json.dll
│       ├── System.Data.SQLite.dll            ← SQLite (R23/R24 only)
│       ├── x64\
│       │   └── SQLite.Interop.dll
│       ├── x86\
│       │   └── SQLite.Interop.dll
│       └── Resources\
│           └── *.rfa family files
├── 2024\
│   ├── JSE_RevitAddin_MEP_OPENINGS.addin
│   └── JSE_MEP_OPENINGS\
│       └── ... (same structure as 2023)
├── 2025\
│   ├── JSE_RevitAddin_MEP_OPENINGS.addin
│   └── JSE_MEP_OPENINGS\
│       ├── JSE_RevitAddin_MEP_OPENINGS.dll
│       ├── JSE_Parameter_Service.dll
│       ├── CommunityToolkit.Mvvm.dll
│       ├── Serilog.dll
│       ├── Microsoft.Data.Sqlite.dll         ← SQLite for .NET 8
│       ├── SQLitePCLRaw.*.dll
│       ├── e_sqlite3.dll
│       └── Resources\
└── 2026\
    ├── JSE_RevitAddin_MEP_OPENINGS.addin
    └── JSE_MEP_OPENINGS\
        └── ... (same structure as 2025)
```

## Benefits

1. **Clean Addins Folder**: Only the `.addin` file is in the root Addins folder
2. **Easy Uninstall**: The entire `JSE_MEP_OPENINGS` folder can be deleted
3. **No Conflicts**: All dependencies are isolated in their own folder
4. **Organized**: Resources and DLLs are neatly contained

## How It Works

The `.addin` file uses a relative path to point to the subfolder:

```xml
<Assembly>JSE_MEP_OPENINGS\JSE_RevitAddin_MEP_OPENINGS.dll</Assembly>
```

Revit loads the add-in from this relative path, so all DLLs are loaded from the subfolder.

## Building the Installer

The ISS script automatically:
1. Creates the `JSE_MEP_OPENINGS` subfolder for each version
2. Copies the `.addin` file to the Addins root
3. Copies all DLLs to the subfolder
4. Runs PowerShell to unblock files in the subfolder
5. Cleans up on uninstall

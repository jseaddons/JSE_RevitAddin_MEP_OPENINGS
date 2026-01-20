# JSE MEP Openings - Deployment Guide

## SQLite Native Library Deployment

This add-in uses SQLite for data persistence, which requires native DLLs to be deployed alongside the main add-in DLL.

## Quick Deployment

### For Developers (Automatic)

The project's PostBuild target automatically copies all required files to:
```
C:\ProgramData\Autodesk\Revit\Addins\{RevitVersion}\
```

Just build the project and restart Revit.

### For Team Deployment (PowerShell Script)

1. Build the project in `Release` configuration
2. Run the deployment script:
   ```powershell
   .\deploy-addin.ps1 -RevitVersion "2023"
   ```
3. Restart Revit

The script will:
- Copy the main add-in DLL
- Copy all SQLite dependencies (`System.Data.SQLite.dll`, `x64\SQLite.Interop.dll`)
- Copy other required dependencies
- Verify all files are present

## Manual Deployment

If you need to deploy manually:

### Required Files

Copy these files to `%APPDATA%\Autodesk\Revit\Addins\{RevitVersion}\JSE_MEP_Openings\`:

**Core Files:**
- `JSE_RevitAddin_MEP_OPENINGS.dll` (main add-in)
- `JSE_RevitAddin_MEP_OPENINGS.addin` (manifest)

**SQLite Dependencies (CRITICAL):**
- `System.Data.SQLite.dll`
- `x64\SQLite.Interop.dll` ← **Native runtime (required!)**

**Other Dependencies:**
- `CommunityToolkit.Mvvm.dll`
- `Serilog.dll`
- `Serilog.Sinks.Debug.dll`

### Verification

After deployment, verify all files are present:

```powershell
$targetPath = "$env:APPDATA\Autodesk\Revit\Addins\2023\JSE_MEP_Openings"
Get-ChildItem $targetPath | Select-Object Name, Length
```

You should see at least:
- `JSE_RevitAddin_MEP_OPENINGS.dll`
- `System.Data.SQLite.dll`
- `x64\SQLite.Interop.dll`

## Troubleshooting

### SQLite Initialization Errors

If you see errors like:
```
Unable to load DLL 'SQLite.Interop.dll'
```
or
```
System.IO.FileNotFoundException: Could not load file or assembly 'System.Data.SQLite'
```

**Solution:**
1. Verify both `System.Data.SQLite.dll` and `x64\SQLite.Interop.dll` exist in the add-in folder
2. Check that the native DLL under `x64` is the 64-bit version (Revit is 64-bit only)
3. Ensure DLLs are not blocked by Windows (right-click → Properties → Unblock)
4. Check Revit journal files for detailed error messages

### Missing Native DLL

If `x64\SQLite.Interop.dll` is missing:

1. **Check NuGet Package:**
   - Verify `System.Data.SQLite.Core` version 1.0.118.0 is restored
   - Restore NuGet packages: `dotnet restore`

2. **Check Build Output:**
   - Look in `bin\Release R{Version}\x64\` folder
   - `SQLite.Interop.dll` should be present after build

3. **Manual Copy:**
   - Find DLL in NuGet cache: `%USERPROFILE%\.nuget\packages\sqlitepclraw.bundle_e_sqlite3\2.1.7\runtimes\win-x64\native\e_sqlite3.dll`
   - Copy to add-in folder manually

### Architecture Mismatch

Ensure you're building for **x64** (Revit is 64-bit only):

```xml
<PropertyGroup>
  <PlatformTarget>x64</PlatformTarget>
</PropertyGroup>
```

## Build Configuration

The project supports multiple Revit versions:
- R21 = Revit 2021 (.NET Framework 4.8)
- R22 = Revit 2022 (.NET Framework 4.8)
- R23 = Revit 2023 (.NET Framework 4.8)
- R24 = Revit 2024 (.NET Framework 4.8)
- R25 = Revit 2025 (.NET 8.0)
- R26 = Revit 2026 (.NET 8.0)

Build with the appropriate configuration:
```
Release R23  (for Revit 2023)
Release R24  (for Revit 2024)
```

## Deployment Checklist

- [ ] Build project in Release configuration
- [ ] Verify `e_sqlite3.dll` is in build output
- [ ] Copy all files to Revit Addins folder
- [ ] Verify `.addin` manifest path is correct
- [ ] Test on clean machine (if possible)
- [ ] Document deployment location for team

## Team Rollout

For team deployment:

1. **Create deployment package:**
   - Build in Release configuration
   - Package all files from `bin\Release R{Version}\`
   - Include deployment script

2. **Distribute to team:**
   - Share ZIP file with all DLLs
   - Include deployment instructions
   - Provide PowerShell script for easy installation

3. **Verify installation:**
   - Check add-in loads in Revit
   - Test SQLite functionality (refresh service)
   - Check logs for any errors

## Support

If SQLite errors persist:
1. Check `refresh.log` for detailed error messages
2. Verify all DLLs are present and not corrupted
3. Try rebuilding from clean solution
4. Contact development team with error logs


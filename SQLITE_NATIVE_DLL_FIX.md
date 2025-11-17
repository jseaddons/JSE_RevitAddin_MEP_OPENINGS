# SQLite Native DLL Deployment Fix

> **Update – November 2025**  
> The add-in now uses the `System.Data.SQLite.Core` package. Ensure the following files are copied alongside `JSE_RevitAddin_MEP_OPENINGS.dll` (and into the Revit execution temp folder):  
> - `System.Data.SQLite.dll`  
> - `x64\SQLite.Interop.dll`
>
> The guidance below (referencing `Microsoft.Data.Sqlite` and `e_sqlite3.dll`) is retained for historical context while we finish migrating documentation. Prefer the updated instructions above when diagnosing current builds.

## Problem

SQLite initialization fails with:
```
The type initializer for 'Microsoft.Data.Sqlite.SqliteConnection' threw an exception
```

Even though `Batteries.Init()` succeeds, the native DLL (`e_sqlite3.dll`) is not found when `SqliteConnection` tries to load it.

## Root Cause

**ILRepack merges .NET assemblies, but native DLLs cannot be merged.** The native DLL (`e_sqlite3.dll`) must be in the **same directory** as the merged add-in DLL when Revit loads it.

## Two Deployment Scenarios

### Scenario 1: Add-in Manager (Testing)
When using **Revit Add-in Manager** to load the DLL directly:
- DLL is loaded from build output directory (e.g., `bin\Debug R23\` or `bin\Release R23\`)
- **Native DLL must be in the same build output directory**
- Example: `bin\Debug R23\JSE_RevitAddin_MEP_OPENINGS.dll` and `bin\Debug R23\e_sqlite3.dll` must be together

### Scenario 2: Deployed Add-in (Production)
When add-in is deployed to Revit Addins folder:
- DLL is loaded from `C:\ProgramData\Autodesk\Revit\Addins\{Version}\`
- **Native DLL must be in the same Revit Addins folder**
- Example: Both DLLs in `C:\ProgramData\Autodesk\Revit\Addins\2023\`

## Solution

### For Add-in Manager (Testing) ⚡

**Quick Fix:** Copy `e_sqlite3.dll` to your build output directory:

```powershell
# Find your build output directory (where Add-in Manager loads the DLL from)
# Usually: bin\Debug R23\ or bin\Release R23\

# Option 1: Copy from NuGet cache
$nugetDll = "$env:USERPROFILE\.nuget\packages\sqlitepclraw.bundle_e_sqlite3\2.1.7\runtimes\win-x64\native\e_sqlite3.dll"
$buildOutput = ".\bin\Debug R23"  # Change to your actual build output path
Copy-Item $nugetDll "$buildOutput\e_sqlite3.dll" -Force

# Option 2: If DLL already exists elsewhere, copy it
Copy-Item ".\bin\Release R23\e_sqlite3.dll" ".\bin\Debug R23\e_sqlite3.dll" -Force
```

**Verify:** After copying, check that both files are in the same directory:
```
bin\Debug R23\
├── JSE_RevitAddin_MEP_OPENINGS.dll
└── e_sqlite3.dll  ← Must be here!
```

### For Deployed Add-in (Production)

### Step 1: Verify Native DLL is in Build Output

After building, check if `e_sqlite3.dll` exists in:
```
bin\Release R{Version}\e_sqlite3.dll
```

If missing, the NuGet package may not be copying it. Check:
1. NuGet packages are restored: `dotnet restore`
2. `SQLitePCLRaw.bundle_e_sqlite3` version 2.1.7 is installed
3. Build output includes the DLL

### Step 2: Ensure Native DLL is Copied to Revit Addins Folder

The PostBuild target should copy `e_sqlite3.dll` to:
```
C:\ProgramData\Autodesk\Revit\Addins\{Version}\e_sqlite3.dll
```

**Critical:** The native DLL must be in the **same directory** as `JSE_RevitAddin_MEP_OPENINGS.dll`.

### Step 3: Manual Copy (If Needed)

If the PostBuild target doesn't copy it automatically:

1. **Find the DLL:**
   - Check build output: `bin\Release R{Version}\e_sqlite3.dll`
   - Or NuGet cache: `%USERPROFILE%\.nuget\packages\sqlitepclraw.bundle_e_sqlite3\2.1.7\runtimes\win-x64\native\e_sqlite3.dll`

2. **Copy to Revit Addins folder:**
   ```powershell
   Copy-Item "bin\Release R23\e_sqlite3.dll" "C:\ProgramData\Autodesk\Revit\Addins\2023\e_sqlite3.dll"
   ```

3. **Verify:**
   ```powershell
   .\check-sqlite-deployment.ps1 -RevitVersion "2023"
   ```

## Diagnostics Added

### 1. Enhanced Logging (`SleeveDbContext.cs`)

The code now logs:
- Assembly location and directory
- Native DLL expected path
- Whether DLL exists before initialization
- Alternative locations checked
- Clear error message if DLL is missing

### 2. Build Diagnostics (`JSE_RevitAddin_MEP_OPENINGS.csproj`)

PostBuild target now:
- Logs how many SQLite DLLs were found
- Logs how many were copied
- Warns if `e_sqlite3.dll` is missing from build output

### 3. Deployment Checker Script (`check-sqlite-deployment.ps1`)

Run this script to verify deployment:
```powershell
.\check-sqlite-deployment.ps1 -RevitVersion "2023"
```

It checks:
- All common Revit Addins locations
- Build output directories
- Lists all files in add-in directory
- Provides clear fix instructions if DLL is missing

## Expected File Structure

After deployment, the Revit Addins folder should contain:

```
C:\ProgramData\Autodesk\Revit\Addins\2023\
├── JSE_RevitAddin_MEP_OPENINGS.dll          (main add-in)
├── JSE_RevitAddin_MEP_OPENINGS.addin        (manifest)
├── e_sqlite3.dll                            ← CRITICAL: Native SQLite DLL
├── SQLitePCLRaw.core.dll
├── SQLitePCLRaw.provider.e_sqlite3.dll
└── Microsoft.Data.Sqlite.dll
```

## Troubleshooting

### If DLL is Missing from Build Output

1. **Restore NuGet packages:**
   ```powershell
   dotnet restore
   ```

2. **Clean and rebuild:**
   ```powershell
   dotnet clean
   dotnet build
   ```

3. **Check NuGet package:**
   - Verify `SQLitePCLRaw.bundle_e_sqlite3` is in `packages.config` or `.csproj`
   - Version should be 2.1.7

4. **Manual copy from NuGet cache:**
   ```powershell
   $nugetCache = "$env:USERPROFILE\.nuget\packages\sqlitepclraw.bundle_e_sqlite3\2.1.7\runtimes\win-x64\native\e_sqlite3.dll"
   Copy-Item $nugetCache "bin\Release R23\e_sqlite3.dll"
   ```

### If DLL is in Build Output but Not Copied to Addins Folder

1. **Check PostBuild target logs:**
   - Look for `[PostBuild]` messages in build output
   - Should show "Found X SQLite DLL(s) to copy"

2. **Manual copy:**
   ```powershell
   Copy-Item "bin\Release R23\e_sqlite3.dll" "C:\ProgramData\Autodesk\Revit\Addins\2023\"
   ```

3. **Use deployment script:**
   ```powershell
   .\deploy-addin.ps1 -RevitVersion "2023"
   ```

### If DLL is Copied but Still Not Found

1. **Check Revit journal files:**
   - Look for detailed error messages
   - Check which directory Revit is loading the add-in from

2. **Verify DLL architecture:**
   - Must be x64 (Revit is 64-bit only)
   - Check file properties → Details → Architecture

3. **Check DLL isn't blocked:**
   - Right-click DLL → Properties
   - If "Unblock" button exists, click it

## Quick Fix Commands

### For Add-in Manager (Testing)

```powershell
# Get the build output directory where Add-in Manager loads from
# Check Revit Add-in Manager to see which DLL path it's using
$buildOutput = ".\bin\Debug R23"  # Adjust to your actual path

# Copy native DLL from NuGet cache
$nugetDll = "$env:USERPROFILE\.nuget\packages\sqlitepclraw.bundle_e_sqlite3\2.1.7\runtimes\win-x64\native\e_sqlite3.dll"
if (Test-Path $nugetDll) {
    Copy-Item $nugetDll "$buildOutput\e_sqlite3.dll" -Force
    Write-Host "✅ Copied e_sqlite3.dll to: $buildOutput" -ForegroundColor Green
} else {
    Write-Host "❌ NuGet DLL not found. Restore packages first: dotnet restore" -ForegroundColor Red
}

# Verify both DLLs are together
$mainDll = Join-Path $buildOutput "JSE_RevitAddin_MEP_OPENINGS.dll"
$nativeDll = Join-Path $buildOutput "e_sqlite3.dll"
if ((Test-Path $mainDll) -and (Test-Path $nativeDll)) {
    Write-Host "✅ Both DLLs are in the same directory!" -ForegroundColor Green
} else {
    Write-Host "❌ DLLs are not together. Check paths above." -ForegroundColor Red
}
```

### For Deployed Add-in (Production)

```powershell
# Find the DLL
$dll = Get-ChildItem -Path ".\bin\Release R23" -Filter "e_sqlite3.dll" -Recurse | Select-Object -First 1

if ($dll) {
    # Copy to Revit Addins folder
    $target = "C:\ProgramData\Autodesk\Revit\Addins\2023\e_sqlite3.dll"
    Copy-Item $dll.FullName $target -Force
    Write-Host "✅ Copied e_sqlite3.dll to: $target" -ForegroundColor Green
} else {
    Write-Host "❌ e_sqlite3.dll not found in build output!" -ForegroundColor Red
    Write-Host "Check NuGet package restore and rebuild the project." -ForegroundColor Yellow
}
```

## Summary

**The native DLL (`e_sqlite3.dll`) MUST be in the same directory as the add-in DLL.** 

- ✅ Build output should contain `e_sqlite3.dll`
- ✅ PostBuild target should copy it to Revit Addins folder
- ✅ Both DLLs must be in the same directory when Revit loads the add-in
- ✅ Use `check-sqlite-deployment.ps1` to verify deployment

If all else fails, manually copy `e_sqlite3.dll` from build output to the Revit Addins folder.


# Multi-Targeting Build Fix

## Problem

When building for multiple target frameworks (net48 for R23/R24 and net8.0-windows for R25/R26), Visual Studio / MSBuild shares a single `obj\project.assets.json` file. This causes:

```
Error NETSDK1005: Assets file doesn't have a target for 'net8.0-windows'.
```

**Root Cause:**
1. Build R23 (net48) → `obj\project.assets.json` contains net48 assets
2. Build R25 (net8.0) → MSBuild looks for net8.0-windows in the same file, but it's not there
3. Error occurs because the assets file wasn't regenerated for the new target framework

## Solution

### 1. Directory.Build.props (CRITICAL)

Created `Directory.Build.props` in the project root. This file is loaded **before** NuGet restore and sets:

```xml
<BaseIntermediateOutputPath>obj\$(Configuration)\</BaseIntermediateOutputPath>
<MSBuildProjectExtensionsPath>$(BaseIntermediateOutputPath)</MSBuildProjectExtensionsPath>
```

This ensures:
- R23 uses `obj\Debug R23\project.assets.json`
- R24 uses `obj\Debug R24\project.assets.json`
- R25 uses `obj\Debug R25\project.assets.json`
- R26 uses `obj\Debug R26\project.assets.json`

Each configuration gets its own isolated assets file!

### 2. Removed Duplicate Setting

Removed `BaseIntermediateOutputPath` from `JSE_RevitAddin_MEP_OPENINGS.csproj` since it's now in `Directory.Build.props`.

## How to Build

### Option 1: Batch Script (Simple)

```batch
build-all-versions.bat
```

This builds all 4 versions in sequence with proper restore for each.

### Option 2: PowerShell Script (Advanced)

```powershell
.\Build-AllVersions.ps1

# With clean build
.\Build-AllVersions.ps1 -Clean

# With publish for .NET 8 versions
.\Build-AllVersions.ps1 -Publish
```

### Option 3: Visual Studio

1. **First time only:** Run `dotnet restore` from command line to ensure Directory.Build.props is processed
2. Build each configuration individually:
   - Select "Debug R23" → Build
   - Select "Debug R24" → Build
   - Select "Debug R25" → Build
   - Select "Debug R26" → Build

## Folder Structure After Fix

```
JSE_MEPOPENING_23/
├── Directory.Build.props          ← NEW: Ensures isolated obj folders
├── build-all-versions.bat         ← NEW: Batch build script
├── Build-AllVersions.ps1          ← NEW: PowerShell build script
├── JSE_RevitAddin_MEP_OPENINGS.csproj
├── obj/
│   ├── Debug R23/                 ← Isolated assets for R23
│   │   └── project.assets.json    ← Contains net48 assets
│   ├── Debug R24/                 ← Isolated assets for R24
│   │   └── project.assets.json    ← Contains net48 assets
│   ├── Debug R25/                 ← Isolated assets for R25
│   │   └── project.assets.json    ← Contains net8.0 assets
│   └── Debug R26/                 ← Isolated assets for R26
│       └── project.assets.json    ← Contains net8.0 assets
└── bin/
    ├── Debug R23/
    ├── Debug R24/
    ├── Debug R25/
    └── Debug R26/
```

## Troubleshooting

### Still getting NETSDK1005?

1. **Close Visual Studio** (locks files)
2. Delete all obj folders:
   ```powershell
   Remove-Item -Path "obj" -Recurse -Force
   ```
3. Verify `Directory.Build.props` exists in project root
4. Rebuild

### Building in VS still fails?

1. Close VS
2. Run from command line first:
   ```batch
   dotnet restore
   dotnet build -c "Debug R23"
   dotnet build -c "Debug R25"
   ```
3. Then open VS and build

### One configuration works but others fail?

The obj folder might be corrupted. Clean specific configuration:
```batch
dotnet clean -c "Debug R25"
dotnet build -c "Debug R25"
```

## Technical Details

### Why Directory.Build.props?

MSBuild loads files in this order:
1. `Directory.Build.props` (if exists)
2. `.csproj` file
3. NuGet restore (uses `BaseIntermediateOutputPath`)
4. Build

Setting `BaseIntermediateOutputPath` in `.csproj` is too late - NuGet restore has already happened with the default path (`obj\`).

### Why not just use `<TargetFramework>` instead of `<TargetFrameworks>`?

The project needs to define all frameworks in `TargetFrameworks` for NuGet to restore packages for all of them. But during build, we override with a single framework based on configuration to avoid MSBuild confusion.

```xml
<!-- Define all frameworks for NuGet -->
<TargetFrameworks>net48;net8.0-windows</TargetFrameworks>

<!-- Override during build based on configuration -->
<TargetFrameworks Condition="$(Configuration.Contains('R23'))">net48</TargetFrameworks>
<TargetFrameworks Condition="$(Configuration.Contains('R25'))">net8.0-windows</TargetFrameworks>
```

This hybrid approach allows:
- Single NuGet restore for all frameworks
- Single framework per build configuration
- No conflicts between net48 and net8.0

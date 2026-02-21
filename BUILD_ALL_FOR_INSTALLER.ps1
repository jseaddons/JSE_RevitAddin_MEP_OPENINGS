#requires -Version 5.1
<#
.SYNOPSIS
    Master build script - Builds both projects and prepares for ISS installer
.DESCRIPTION
    1. Builds JSE_Parameter_Service (all 4 versions)
    2. Copies Parameter Service DLLs to main project
    3. Builds JSE_MEPOPENING_23 (all 4 versions)
    4. Ready for ISS installer build
#>

[CmdletBinding()]
param(
    [switch]$Clean,
    [switch]$SkipParameterService,
    [switch]$SkipMainProject,
    [switch]$SkipBuild
)

$ErrorActionPreference = 'Stop'

$PSServicePath = 'C:\Jse_Developments\JSE_Parameter_Service'
$MainProjectPath = 'C:\Jse_Developments\JSE_MEPOPENING_23'

function Write-Step {
    param([int]$Step, [int]$Total, [string]$Message)
    Write-Host ('[' + $Step + '/' + $Total + '] ' + $Message) -ForegroundColor Cyan
}

function Write-Success { param([string]$M) Write-Host ('  SUCCESS: ' + $M) -ForegroundColor Green }
function Write-Fail { param([string]$M) Write-Host ('  FAILED: ' + $M) -ForegroundColor Red }

Clear-Host
Write-Host '======================================================' -ForegroundColor Cyan
Write-Host '  JSE MEP Openings - Complete Build for Installer' -ForegroundColor Cyan
Write-Host '======================================================' -ForegroundColor Cyan
Write-Host ''

# 0. Global Cleanup
if ($Clean) {
    Write-Host '  Performing global cleanup...' -ForegroundColor Yellow
    
    # Clean deploy folder first so it doesn't delete files copied later
    $deployPath = Join-Path $MainProjectPath 'deploy'
    if (Test-Path $deployPath) {
        Write-Host '    Cleaning deploy folder...'
        Remove-Item -Path $deployPath -Recurse -Force -ErrorAction SilentlyContinue
    }
    
    if (-not $SkipParameterService) {
        Write-Host '    Cleaning Parameter Service bin/obj...'
        $psBin = Join-Path $PSServicePath 'bin'
        $psObj = Join-Path $PSServicePath 'obj'
        if (Test-Path $psBin) { Remove-Item $psBin -Recurse -Force -ErrorAction SilentlyContinue }
        if (Test-Path $psObj) { Remove-Item $psObj -Recurse -Force -ErrorAction SilentlyContinue }
    }
    
    if (-not $SkipMainProject) {
        Write-Host '    Cleaning Main Project bin/obj...'
        $mainBin = Join-Path $MainProjectPath 'bin'
        $mainObj = Join-Path $MainProjectPath 'obj'
        if (Test-Path $mainBin) { Remove-Item $mainBin -Recurse -Force -ErrorAction SilentlyContinue }
        if (Test-Path $mainObj) { Remove-Item $mainObj -Recurse -Force -ErrorAction SilentlyContinue }
    }
    Write-Host ''
}

$TotalSteps = 4
if ($SkipParameterService) { $TotalSteps-- }
if ($SkipMainProject) { $TotalSteps-- }

$CurrentStep = 0

# ============================================================================
# Helper to get configs
# ============================================================================
$Configs = @(
    @{ Name = 'Debug R23'; Framework = 'net48'; Out = 'net48'; Revit = '2023' },
    @{ Name = 'Debug R24'; Framework = 'net48'; Out = 'net48'; Revit = '2024' },
    @{ Name = 'Debug R25'; Framework = 'net8.0-windows'; Out = 'net8.0-windows'; Revit = '2025' },
    @{ Name = 'Debug R26'; Framework = 'net8.0-windows'; Out = 'net8.0-windows'; Revit = '2026' }
)

# ============================================================================
# STEP 1: Build Parameter Service
# ============================================================================
if (-not $SkipParameterService) {
    $CurrentStep++
    Write-Step -Step $CurrentStep -Total $TotalSteps -Message 'Building Parameter Service'
    Write-Host ''
    
    Set-Location $PSServicePath
    
    foreach ($c in $Configs) {
        if ($SkipBuild) {
            Write-Host ('  Skipping build for ' + $c.Name) -ForegroundColor Gray
            continue
        }
        
        Write-Host ('  Building ' + $c.Name + '...') -NoNewline
        try {
            $null = dotnet restore JSE_Parameter_Service.csproj -p:Configuration=($c.Name) -v q 2>&1
            if ($LASTEXITCODE -ne 0) { throw 'Restore failed' }
            
            $null = dotnet build JSE_Parameter_Service.csproj -c ($c.Name) --no-restore -v q 2>&1
            if ($LASTEXITCODE -ne 0) { throw 'Build failed' }
            
            Write-Host ' DONE' -ForegroundColor Green
        }
        catch {
            Write-Host ' FAILED' -ForegroundColor Red
            Write-Warning ('Build failed for ' + $c.Name + '. Will attempt to continue if DLL exists.')
        }
    }
    Write-Host ''
}

# ============================================================================
# STEP 2: Copy Parameter Service DLLs
# ============================================================================
if (-not $SkipParameterService) {
    $CurrentStep++
    Write-Step -Step $CurrentStep -Total $TotalSteps -Message 'Copying Parameter Service DLLs'
    Write-Host ''
    
    Set-Location $MainProjectPath
    
    foreach ($c in $Configs) {
        $config = $c.Name
        $out = $c.Out
        $revitVer = $c.Revit
        
        # Try both root and framework subfolder for source
        $sources = @(
            ($PSServicePath + '\bin\' + $config + '\JSE_Parameter_Service.dll'),
            ($PSServicePath + '\bin\' + $config + '\' + $out + '\JSE_Parameter_Service.dll')
        )
        
        $binDestFolder = ($MainProjectPath + '\bin\' + $config + '\' + $config)
        $deployDestFolder = ($MainProjectPath + '\deploy\' + $revitVer + '\JSE_MEP_OPENINGS')
        
        Write-Host ('  Processing ' + $config + ' (' + $revitVer + ')...') -ForegroundColor White
        
        $found = $false
        foreach ($s in $sources) {
            if (Test-Path -Path $s) {
                if (-not (Test-Path -Path $binDestFolder)) { New-Item -ItemType Directory -Path $binDestFolder -Force | Out-Null }
                Copy-Item -Path $s -Destination (Join-Path $binDestFolder 'JSE_Parameter_Service.dll') -Force
                
                if (-not (Test-Path -Path $deployDestFolder)) { New-Item -ItemType Directory -Path $deployDestFolder -Force | Out-Null }
                Copy-Item -Path $s -Destination (Join-Path $deployDestFolder 'JSE_Parameter_Service.dll') -Force
                
                Write-Host ('    ✓ Copied from: ' + $s) -ForegroundColor Green
                $found = $true
                break
            }
        }
        
        if (-not $found) {
            Write-Host '    ❌ SOURCE NOT FOUND! Checked:' -ForegroundColor Red
            foreach ($s in $sources) { Write-Host ('      - ' + $s) -ForegroundColor Gray }
        }
    }
    Write-Host ''
}

# ============================================================================
# STEP 3: Build Main Project
# ============================================================================
if (-not $SkipMainProject) {
    $CurrentStep++
    Write-Step -Step $CurrentStep -Total $TotalSteps -Message 'Building Main Project (MEP OPENINGS)'
    Write-Host ''
    
    Set-Location $MainProjectPath
    
    if ($SkipBuild) {
        Write-Host '  Skipping main project build as requested.' -ForegroundColor Gray
    }
    else {
        # Use the existing build script (Clean handled in Step 0)
        & ($MainProjectPath + '\Build-AllVersions.ps1')
        
        if ($LASTEXITCODE -ne 0) {
            Write-Fail 'Main project build failed'
            exit 1
        }
    }
}

# ============================================================================
# STEP 4: Verification
# ============================================================================
$CurrentStep++
Write-Step -Step $CurrentStep -Total $TotalSteps -Message 'Verification'
Write-Host ''

$VerifyPaths = @(
    'deploy\2023\JSE_RevitAddin_MEP_OPENINGS.dll',
    'deploy\2023\JSE_MEP_OPENINGS\JSE_Parameter_Service.dll',
    'deploy\2023\JSE_MEP_OPENINGS\System.Data.SQLite.dll',
    'deploy\2023\JSE_MEP_OPENINGS\x64\SQLite.Interop.dll',
    
    'deploy\2024\JSE_RevitAddin_MEP_OPENINGS.dll',
    'deploy\2024\JSE_MEP_OPENINGS\JSE_Parameter_Service.dll',
    'deploy\2024\JSE_MEP_OPENINGS\System.Data.SQLite.dll',
    'deploy\2024\JSE_MEP_OPENINGS\x64\SQLite.Interop.dll',
    
    'deploy\2025\JSE_RevitAddin_MEP_OPENINGS.dll',
    'deploy\2025\JSE_MEP_OPENINGS\JSE_Parameter_Service.dll',
    'deploy\2025\JSE_MEP_OPENINGS\Microsoft.Data.Sqlite.dll',
    'deploy\2025\JSE_MEP_OPENINGS\SQLitePCLRaw.core.dll',
    'deploy\2025\JSE_MEP_OPENINGS\e_sqlite3.dll',
    'deploy\2025\JSE_MEP_OPENINGS\Serilog.dll',
    'deploy\2025\JSE_MEP_OPENINGS\CommunityToolkit.Mvvm.dll',
    'deploy\2025\JSE_MEP_OPENINGS\Nice3point.Revit.Toolkit.dll',
    
    'deploy\2026\JSE_RevitAddin_MEP_OPENINGS.dll',
    'deploy\2026\JSE_MEP_OPENINGS\JSE_Parameter_Service.dll',
    'deploy\2026\JSE_MEP_OPENINGS\Microsoft.Data.Sqlite.dll',
    'deploy\2026\JSE_MEP_OPENINGS\SQLitePCLRaw.core.dll',
    'deploy\2026\JSE_MEP_OPENINGS\e_sqlite3.dll',
    'deploy\2026\JSE_MEP_OPENINGS\Serilog.dll',
    'deploy\2026\JSE_MEP_OPENINGS\CommunityToolkit.Mvvm.dll',
    'deploy\2026\JSE_MEP_OPENINGS\Nice3point.Revit.Toolkit.dll'
)

$AllFound = $true
foreach ($path in $VerifyPaths) {
    $fullPath = Join-Path $MainProjectPath $path
    
    # Get a nice descriptive path for display
    $displayPath = $path
    if ($path.StartsWith('deploy\')) { $displayPath = $path.Substring(7) }

    Write-Host ('  Checking ' + $displayPath + '...') -NoNewline
    
    if (Test-Path -Path $fullPath) {
        $fileObj = Get-Item -Path $fullPath
        $sizeKB = [Math]::Round($fileObj.Length / 1KB)
        Write-Host (' FOUND (' + $sizeKB + ' KB)') -ForegroundColor Green
    }
    else {
        Write-Host ' MISSING' -ForegroundColor Red
        $AllFound = $false
    }
}

Write-Host ''
Write-Host '======================================================' -ForegroundColor Cyan

if ($AllFound) {
    Write-Success 'ALL FILES READY FOR INSTALLER!'
    Write-Host ''
    Write-Host 'Next step: Build the ISS installer in Inno Setup' -ForegroundColor Yellow
    Write-Host 'File: JSE_RevitAddin_MEP_OPENINGS.iss' -ForegroundColor Gray
    Write-Host ''
    exit 0
}
else {
    Write-Fail 'SOME FILES ARE MISSING!'
    Write-Host ''
    exit 1
}

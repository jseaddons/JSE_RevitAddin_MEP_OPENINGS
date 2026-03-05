#requires -Version 5.1
<#
.SYNOPSIS
    Master build script - Builds all projects and prepares for ISS installer
.DESCRIPTION
    1. Builds JSE_Parameter_Service (all 4 versions) and copies DLL to deploy folder
    2. Builds JSE_MEPOPENING_23 (all 4 versions)
    3. Verification - confirms all deploy files are ready for installer
    NOTE: JSE_T1 is a separate prerequisite installed independently.
#>

[CmdletBinding()]
param(
    [switch]$Clean,
    [switch]$SkipParameterService,
    [switch]$SkipMainProject
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

# ============================================================================
# Interactive build selection (skipped if switches already passed on command line)
# ============================================================================
if (-not $SkipParameterService.IsPresent -and -not $SkipMainProject.IsPresent) {
    Write-Host '  Which projects need a new build?' -ForegroundColor Yellow
    Write-Host ''
    Write-Host '    [1] Both  - JSE_Parameter_Service + JSE_MEP_OPENINGS' -ForegroundColor White
    Write-Host '    [2] PS    - JSE_Parameter_Service only' -ForegroundColor White
    Write-Host '    [3] MEP   - JSE_MEP_OPENINGS only' -ForegroundColor White
    Write-Host '    [4] None  - Skip all builds, verify deploy files only' -ForegroundColor White
    Write-Host ''

    $choice = Read-Host '  Enter choice (1/2/3/4)'

    switch ($choice.Trim()) {
        '1' {
            # both - nothing to skip
        }
        '2' { $SkipMainProject = $true }
        '3' { $SkipParameterService = $true }
        '4' { $SkipParameterService = $true; $SkipMainProject = $true }
        default {
            Write-Host '  Invalid choice. Defaulting to build both.' -ForegroundColor Yellow
        }
    }
    Write-Host ''
}

# ============================================================================
# Version selection (Interactive)
# ============================================================================
$VersionsToBuild = @("2023", "2024", "2025", "2026") # Default: all

if (-not $SkipParameterService.IsPresent -or -not $SkipMainProject.IsPresent) {
    Write-Host '  Which Revit versions do you want to build?' -ForegroundColor Yellow
    Write-Host '  (Press ENTER for all, or type specific years like: 2026 or 2023,2024)' -ForegroundColor Gray
    Write-Host ''
    $verInput = Read-Host '  Enter versions'
    
    if (-not [string]::IsNullOrWhiteSpace($verInput)) {
        $selectedVers = $verInput.Split(',') | ForEach-Object { $_.Trim() }
        # Validate - only allow 2023-2026
        $validVers = $selectedVers | Where-Object { $_ -match "^202(3|4|5|6)$" }
        
        if ($validVers.Count -gt 0) {
            $VersionsToBuild = $validVers
            Write-Host "  Building versions: $($VersionsToBuild -join ', ')" -ForegroundColor Green
        }
        else {
            Write-Host "  Invalid versions entered. Building ALL versions." -ForegroundColor Yellow
        }
        Write-Host ''
    }
}

# 0. Global Cleanup
if ($Clean) {
    Write-Host '  Performing global cleanup...' -ForegroundColor Yellow

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

$TotalSteps = 3
if ($SkipParameterService) { $TotalSteps-- }
if ($SkipMainProject) { $TotalSteps-- }

$CurrentStep = 0

# ============================================================================
# Helper configs
# ============================================================================
$Configs = @(
    @{ Name = 'Debug R23'; Out = 'net48'; Revit = '2023' },
    @{ Name = 'Debug R24'; Out = 'net48'; Revit = '2024' },
    @{ Name = 'Debug R25'; Out = 'net8.0-windows'; Revit = '2025' },
    @{ Name = 'Debug R26'; Out = 'net8.0-windows'; Revit = '2026' }
)

# Apply version filter
if ($VersionsToBuild -and $VersionsToBuild.Count -lt 4) {
    $Configs = $Configs | Where-Object { $VersionsToBuild -contains $_.Revit }
}

# ============================================================================
# STEP 1: Build JSE_Parameter_Service + copy to deploy
# ============================================================================
if (-not $SkipParameterService) {
    $CurrentStep++
    Write-Step -Step $CurrentStep -Total $TotalSteps -Message 'Building & Deploying Parameter Service'
    Write-Host ''

    Set-Location $PSServicePath

    foreach ($c in $Configs) {
        $config = $c.Name
        $out = $c.Out
        $revitVer = $c.Revit

        Write-Host ('  Building ' + $config + ' (' + $revitVer + ')...') -NoNewline
        try {
            $null = dotnet restore JSE_Parameter_Service.csproj -p:Configuration="$config" -v q 2>&1
            if ($LASTEXITCODE -ne 0) { throw 'Restore failed' }

            $null = dotnet build JSE_Parameter_Service.csproj -c "$config" --no-restore -v q 2>&1
            if ($LASTEXITCODE -ne 0) { throw 'Build failed' }

            Write-Host ' DONE' -ForegroundColor Green
        }
        catch {
            Write-Host (' FAILED (' + $_ + ')') -ForegroundColor Red
        }

        # Copy DLL to deploy folder - try nested config folder first, then RID subfolder, then framework subfolder, then config root
        $sources = @(
            ($PSServicePath + '\bin\' + $config + '\' + $config + '\JSE_Parameter_Service.dll'),
            ($PSServicePath + '\bin\' + $config + '\' + $out + '\win-x64\JSE_Parameter_Service.dll'),
            ($PSServicePath + '\bin\' + $config + '\win-x64\' + '\JSE_Parameter_Service.dll'),
            ($PSServicePath + '\bin\' + $config + '\' + $out + '\JSE_Parameter_Service.dll'),
            ($PSServicePath + '\bin\' + $config + '\JSE_Parameter_Service.dll')
        )

        $deployDestFolder = ($MainProjectPath + '\deploy\' + $revitVer + '\JSE_MEP_OPENINGS')

        $found = $false
        foreach ($s in $sources) {
            if (Test-Path -Path $s) {
                if (-not (Test-Path -Path $deployDestFolder)) { New-Item -ItemType Directory -Path $deployDestFolder -Force | Out-Null }
                Copy-Item -Path $s -Destination (Join-Path $deployDestFolder 'JSE_Parameter_Service.dll') -Force
                $sz = [Math]::Round((Get-Item $s).Length / 1KB)
                Write-Host ('    v Copied (' + $sz + ' KB): ' + $s) -ForegroundColor Green
                $found = $true
                break
            }
        }

        if (-not $found) {
            Write-Host '    x DLL NOT FOUND after build. Checked:' -ForegroundColor Red
            foreach ($s in $sources) { Write-Host ('      - ' + $s) -ForegroundColor Gray }
        }
    }
    Write-Host ''
}

# ============================================================================
# STEP 2: Build Main Project (MEP OPENINGS)
# ============================================================================
if (-not $SkipMainProject) {
    $CurrentStep++
    Write-Step -Step $CurrentStep -Total $TotalSteps -Message 'Building Main Project (MEP OPENINGS)'
    Write-Host ''

    Set-Location $MainProjectPath
    
    $buildParams = @{}
    if ($Clean) { $buildParams.Clean = $true }
    if ($VersionsToBuild -and $VersionsToBuild.Count -lt 4) {
        $buildParams.Versions = $VersionsToBuild
    }

    & ($MainProjectPath + '\Build-AllVersions.ps1') @buildParams

    if ($LASTEXITCODE -ne 0) {
        Write-Fail 'Main project build failed'
        exit 1
    }
}

# ============================================================================
# STEP 3: Verification
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

# Only verify versions we actually targeted/built
if ($VersionsToBuild -and $VersionsToBuild.Count -lt 4) {
    $VerifyPaths = $VerifyPaths | Where-Object { 
        $path = $_
        $match = $false
        foreach ($v in $VersionsToBuild) {
            if ($path.Contains("deploy\$v\")) { $match = $true; break }
        }
        $match
    }
}

$AllFound = $true
foreach ($path in $VerifyPaths) {
    $fullPath = Join-Path $MainProjectPath $path
    $displayPath = if ($path.StartsWith('deploy\')) { $path.Substring(7) } else { $path }

    Write-Host ('  Checking ' + $displayPath + '...') -NoNewline

    if (Test-Path -Path $fullPath) {
        $sizeKB = [Math]::Round((Get-Item -Path $fullPath).Length / 1KB)
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

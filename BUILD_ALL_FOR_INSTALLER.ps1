#requires -Version 5.1
<#
.SYNOPSIS
    Master build script - Builds all projects and prepares for ISS installer
.DESCRIPTION
    1. Builds JSE_Parameter_Service (all versions) and copies DLL to JSE_Parameter_Service\deploy\<year>\
    2. Builds JSE_MEPOPENING_23 (all 4 versions)
    3. Verification - confirms all deploy files are ready for installer
    NOTE: JSE_T1 is a separate prerequisite installed independently.
#>

[CmdletBinding()]
param(
    [switch]$Clean,
    [switch]$SkipParameterService,
    [switch]$SkipMainProject,
    [Parameter(Mandatory=$false)]
    [ValidateSet("Both", "PS", "MEP", "None")]
    [string]$Project,
    [Parameter(Mandatory=$false)]
    [string[]]$Versions
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
# Project selection logic
# ============================================================================
if ($Project) {
    switch ($Project) {
        'PS'   { $SkipMainProject = $true }
        'MEP'  { $SkipParameterService = $true }
        'None' { $SkipParameterService = $true; $SkipMainProject = $true }
        'Both' { # Default behavior
        }
    }
}
elseif (-not $SkipParameterService.IsPresent -and -not $SkipMainProject.IsPresent) {
    Write-Host '  Which projects need a new build?' -ForegroundColor Yellow
    Write-Host ''
    Write-Host '    [1] Both  - JSE_Parameter_Service + JSE_MEP_OPENINGS' -ForegroundColor White
    Write-Host '    [2] PS    - JSE_Parameter_Service only' -ForegroundColor White
    Write-Host '    [3] MEP   - JSE_MEP_OPENINGS only' -ForegroundColor White
    Write-Host '    [4] None  - Skip all builds, verify deploy files only' -ForegroundColor White
    Write-Host ''

    $choice = Read-Host '  Enter choice (1/2/3/4)'

    switch ($choice.Trim()) {
        '1' { }
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
# Version selection logic
# ============================================================================
$VersionsToBuild = if ($Versions) { $Versions } else { @("2022", "2023", "2024", "2025", "2026") }

if (-not $Project -and -not $Versions -and (-not $SkipParameterService.IsPresent -or -not $SkipMainProject.IsPresent)) {
    Write-Host '  Which Revit versions do you want to build?' -ForegroundColor Yellow
    Write-Host '  (Press ENTER for all, or type specific years like: 2026 or 2023,2024)' -ForegroundColor Gray
    Write-Host ''
    $verInput = Read-Host '  Enter versions'
    
    if (-not [string]::IsNullOrWhiteSpace($verInput)) {
        $selectedVers = $verInput.Split(',') | ForEach-Object { $_.Trim() }
        # Validate - only allow 2023-2026
        $validVers = $selectedVers | Where-Object { $_ -match "^202(2|3|4|5|6)$" }
        
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
    @{ Name = 'Debug R22'; Out = 'net48'; Revit = '2022' },
    @{ Name = 'Debug R23'; Out = 'net48'; Revit = '2023' },
    @{ Name = 'Debug R24'; Out = 'net48'; Revit = '2024' },
    @{ Name = 'Debug R25'; Out = 'net8.0-windows'; Revit = '2025' },
    @{ Name = 'Debug R26'; Out = 'net8.0-windows'; Revit = '2026' }
)

# Apply version filter
if ($VersionsToBuild -and $VersionsToBuild.Count -lt 5) {
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

        $deployDestFolder = ($PSServicePath + '\deploy\' + $revitVer)

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
    if ($VersionsToBuild -and $VersionsToBuild.Count -lt 5) {
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

$MEPPaths = @(
    'deploy\{0}\JSE_RevitAddin_MEP_OPENINGS.dll',
    'deploy\{0}\JSE_MEP_OPENINGS\System.Data.SQLite.dll',
    'deploy\{0}\JSE_MEP_OPENINGS\x64\SQLite.Interop.dll',
    'deploy\{0}\JSE_MEP_OPENINGS\Microsoft.Data.Sqlite.dll',
    'deploy\{0}\JSE_MEP_OPENINGS\SQLitePCLRaw.core.dll',
    'deploy\{0}\JSE_MEP_OPENINGS\e_sqlite3.dll',
    'deploy\{0}\JSE_MEP_OPENINGS\Serilog.dll',
    'deploy\{0}\JSE_MEP_OPENINGS\CommunityToolkit.Mvvm.dll',
    'deploy\{0}\JSE_MEP_OPENINGS\Nice3point.Revit.Toolkit.dll'
)

$PSPaths = @(
    '..\JSE_Parameter_Service\deploy\{0}\JSE_Parameter_Service.dll'
)

$VerifyPaths = @()
foreach ($v in $VersionsToBuild) {
    if (-not $SkipMainProject) {
        foreach ($p in $MEPPaths) { $VerifyPaths += ($p -f $v) }
    }
    if (-not $SkipParameterService) {
        foreach ($p in $PSPaths) { $VerifyPaths += ($p -f $v) }
    }
}

$AllFound = $true
foreach ($path in $VerifyPaths) {
    $fullPath = Join-Path $MainProjectPath $path
    
    # Simple check for path existence (some dependencies are version-specific)
    if ($path.Contains("System.Data.SQLite") -and $fullPath.Contains("2025")) { continue }
    if ($path.Contains("System.Data.SQLite") -and $fullPath.Contains("2026")) { continue }
    if ($path.Contains("SQLite.Interop.dll") -and $fullPath.Contains("2025")) { continue }
    if ($path.Contains("SQLite.Interop.dll") -and $fullPath.Contains("2026")) { continue }
    if ($path.Contains("Microsoft.Data.Sqlite") -and ($fullPath.Contains("2022") -or $fullPath.Contains("2023") -or $fullPath.Contains("2024"))) { continue }
    if ($path.Contains("SQLitePCLRaw") -and ($fullPath.Contains("2022") -or $fullPath.Contains("2023") -or $fullPath.Contains("2024"))) { continue }
    if ($path.Contains("e_sqlite3.dll") -and ($fullPath.Contains("2022") -or $fullPath.Contains("2023") -or $fullPath.Contains("2024"))) { continue }

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
    Set-Location $MainProjectPath
    exit 0
}
else {
    Write-Fail 'SOME FILES ARE MISSING!'
    Write-Host ''
    Set-Location $MainProjectPath
    exit 1
}

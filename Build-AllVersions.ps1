#requires -Version 5.1
<#
.SYNOPSIS
    Builds JSE MEP Openings for all Revit versions (2023-2026)
.DESCRIPTION
    Handles the multi-targeting build process correctly by:
    1. Using isolated obj folders per configuration (via Directory.Build.props)
    2. Restoring packages for each target framework separately
    3. Building in framework groups (net48 first, then net8.0)
#>

[CmdletBinding()]
param(
    [switch]$Clean,
    [switch]$Publish,
    [string[]]$Versions # Optional: specify which versions to build (e.g., "2023", "2026")
)

$ErrorActionPreference = "Stop"

# Configuration definitions
$Configurations = @(
    @{ Name = "Debug R23"; Framework = "net48"; RevitVersion = "2023" },
    @{ Name = "Debug R24"; Framework = "net48"; RevitVersion = "2024" },
    @{ Name = "Debug R25"; Framework = "net8.0-windows"; RevitVersion = "2025" },
    @{ Name = "Debug R26"; Framework = "net8.0-windows"; RevitVersion = "2026" }
)

# Apply version filter if specified
if ($Versions -and $Versions.Count -gt 0) {
    $Configurations = $Configurations | Where-Object { $Versions -contains $_.RevitVersion }
    if ($Configurations.Count -eq 0) {
        Write-Warning "No valid configurations found for versions: $($Versions -join ', ')"
    }
}

function Write-Step {
    param([int]$Step, [int]$Total, [string]$Message)
    Write-Host "[$Step/$Total] $Message" -ForegroundColor Cyan
}

function Write-Success {
    param([string]$Message)
    Write-Host "  SUCCESS: $Message" -ForegroundColor Green
}

function Write-Fail {
    param([string]$Message)
    Write-Host "  FAILED: $Message" -ForegroundColor Red
}

Write-Host "======================================================" -ForegroundColor Cyan
Write-Host "JSE MEP Openings - Multi-Version Build Script" -ForegroundColor Cyan
Write-Host "======================================================" -ForegroundColor Cyan
Write-Host ""

$TotalSteps = 5

# Step 1: Pre-build cleanup
Write-Step -Step 1 -Total $TotalSteps -Message "Pre-build cleanup"

if ($Clean) {
    Write-Host "  Cleaning all bin and obj folders..."
    Remove-Item -Path "bin" -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item -Path "obj" -Recurse -Force -ErrorAction SilentlyContinue
    Write-Success -Message "Cleaned all build artifacts"
}
else {
    $sharedFiles = @("obj\project.assets.json", "obj\project.nuget.cache")
    foreach ($file in $sharedFiles) {
        if (Test-Path $file) {
            Remove-Item $file -Force
            Write-Host "  Removed stale: $file"
        }
    }
}
Write-Host ""

# Step 2: Verify Directory.Build.props exists
Write-Step -Step 2 -Total $TotalSteps -Message "Verifying build configuration"
if (-not (Test-Path "Directory.Build.props")) {
    Write-Fail "Directory.Build.props not found! This file is required for multi-targeting."
    exit 1
}
Write-Success -Message "Directory.Build.props found"
Write-Host ""

# Step 3: Build all configurations
Write-Step -Step 3 -Total $TotalSteps -Message "Building all configurations"
Write-Host ""

$BuildResults = @()

foreach ($config in $Configurations) {
    $configName = $config.Name
    $framework = $config.Framework
    
    Write-Host "----------------------------------------" -ForegroundColor Yellow
    Write-Host "Building: $configName ($framework)" -ForegroundColor Yellow
    Write-Host "----------------------------------------" -ForegroundColor Yellow
    
    try {
        Write-Host "  Restoring packages..." -NoNewline
        $null = dotnet restore JSE_RevitAddin_MEP_OPENINGS.csproj -p:Configuration="$configName" -v q 2>&1
        if ($LASTEXITCODE -ne 0) { throw "Restore failed" }
        Write-Host " DONE" -ForegroundColor Green
        
        Write-Host "  Building..." -NoNewline
        $null = dotnet build JSE_RevitAddin_MEP_OPENINGS.csproj -c "$configName" --no-restore -v q 2>&1
        if ($LASTEXITCODE -ne 0) { throw "Build failed" }
        Write-Host " DONE" -ForegroundColor Green
        
        if ($Publish -and $framework -eq "net8.0-windows") {
            Write-Host "  Publishing..." -NoNewline
            $null = dotnet publish JSE_RevitAddin_MEP_OPENINGS.csproj -c "$configName" -f "$framework" --no-build -v q 2>&1
            if ($LASTEXITCODE -ne 0) { throw "Publish failed" }
            Write-Host " DONE" -ForegroundColor Green
        }
        
        $BuildResults += [PSCustomObject]@{
            Configuration = $configName
            Framework     = $framework
            Status        = "Success"
            OutputPath    = "bin\$configName\$configName\JSE_RevitAddin_MEP_OPENINGS.dll"
        }
    }
    catch {
        Write-Fail -Message "$_"
        $BuildResults += [PSCustomObject]@{
            Configuration = $configName
            Framework     = $framework
            Status        = "Failed"
            Error         = $_.Exception.Message
        }
    }
    Write-Host ""
}

# Step 4: Verify outputs
Write-Step -Step 4 -Total $TotalSteps -Message "Verifying build outputs"
Write-Host ""

$AllSuccess = $true
foreach ($result in $BuildResults) {
    if ($result.Status -eq "Success") {
        $dllPath = $result.OutputPath
        $configName = $result.Configuration
        $rootDllPath = "bin\$configName\JSE_RevitAddin_MEP_OPENINGS.dll"
        $ridDllPath = "bin\$configName\$($result.Framework)\win-x64\JSE_RevitAddin_MEP_OPENINGS.dll"
        $ridRootPath = "bin\$configName\win-x64\JSE_RevitAddin_MEP_OPENINGS.dll"

        $foundPath = $null
        if (Test-Path $dllPath) { $foundPath = $dllPath }
        elseif (Test-Path $rootDllPath) { $foundPath = $rootDllPath }
        elseif (Test-Path $ridDllPath) { $foundPath = $ridDllPath }
        elseif (Test-Path $ridRootPath) { $foundPath = $ridRootPath }

        if ($foundPath) {
            $fileInfo = Get-Item $foundPath
            Write-Success -Message "$($result.Configuration) - $($fileInfo.Length) bytes"
        }
        else {
            Write-Fail -Message "$($result.Configuration) - DLL not found"
            $AllSuccess = $false
        }
    }
    else {
        Write-Fail -Message "$($result.Configuration) - $($result.Error)"
        $AllSuccess = $false
    }
}
Write-Host ""

# Step 5: Summary
Write-Step -Step 5 -Total $TotalSteps -Message "Build Summary"
Write-Host "======================================================" -ForegroundColor Cyan

if ($AllSuccess) {
    Write-Host "Status: ALL BUILDS SUCCESSFUL" -ForegroundColor Green
    Write-Host ""
    Write-Host "Output locations:" -ForegroundColor White
    foreach ($result in $BuildResults) {
        Write-Host "  - $($result.OutputPath)" -ForegroundColor Gray
    }
    Write-Host ""
    Write-Host "Next steps:" -ForegroundColor Yellow
    Write-Host "  1. Copy JSE_Parameter_Service.dll to each output folder"
    Write-Host "  2. Run copy-parameter-service.bat"
    Write-Host "  3. Build the installer with Inno Setup"
    Write-Host ""
    exit 0
}
else {
    Write-Host "Status: BUILD FAILED" -ForegroundColor Red
    Write-Host ""
    Write-Host "Troubleshooting:" -ForegroundColor Yellow
    Write-Host "  - Check individual configuration build in Visual Studio"
    Write-Host "  - Ensure all NuGet sources are accessible"
    Write-Host "  - Delete bin and obj folders and try again"
    Write-Host "  - Verify Directory.Build.props is in the project root"
    Write-Host ""
    exit 1
}

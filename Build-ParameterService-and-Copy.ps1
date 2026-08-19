#requires -Version 5.1
<#
.SYNOPSIS
    Builds JSE_Parameter_Service and copies DLLs to MEP_OPENING project
.DESCRIPTION
    1. Builds JSE_Parameter_Service for all Revit versions
    2. Copies JSE_Parameter_Service.dll to each Debug R## folder
    3. Ready for ISS installer build
#>

[CmdletBinding()]
param([switch]$Clean)

$ErrorActionPreference = "Stop"

$ParameterServicePath = "C:\Jse_Developments\JSE_Parameter_Service"
$MainProjectPath = "C:\Jse_Developments\JSE_MEPOPENING_23"

Write-Host "======================================================" -ForegroundColor Cyan
Write-Host "Build Parameter Service + Copy to Main Project" -ForegroundColor Cyan
Write-Host "======================================================" -ForegroundColor Cyan
Write-Host ""

# Step 1: Build Parameter Service
Write-Host "[1/3] Building Parameter Service..." -ForegroundColor Yellow
Write-Host ""

Set-Location $ParameterServicePath

if ($Clean) {
    Write-Host "  Cleaning Parameter Service..."
    Remove-Item -Path "bin", "obj" -Recurse -Force -ErrorAction SilentlyContinue
}

$Configs = @(
    @{ Name = "Debug R22"; Framework = "net48"; OutputFolder = "net48" },
    @{ Name = "Debug R23"; Framework = "net48"; OutputFolder = "net48" },
    @{ Name = "Debug R24"; Framework = "net48"; OutputFolder = "net48" },
    @{ Name = "Debug R25"; Framework = "net8.0-windows"; OutputFolder = "net8.0-windows" },
    @{ Name = "Debug R26"; Framework = "net8.0-windows"; OutputFolder = "net8.0-windows" }
)

$BuildResults = @()

foreach ($config in $Configs) {
    $configName = $config.Name
    Write-Host "  Building $configName..." -NoNewline
    
    try {
        $null = dotnet restore JSE_Parameter_Service.csproj -p:Configuration="$configName" -v q 2>&1
        if ($LASTEXITCODE -ne 0) { throw "Restore failed" }
        
        $null = dotnet build JSE_Parameter_Service.csproj -c "$configName" --no-restore -v q 2>&1
        if ($LASTEXITCODE -ne 0) { throw "Build failed" }
        
        Write-Host " DONE" -ForegroundColor Green
        $BuildResults += @{ Config = $configName; Status = "Success"; OutputFolder = $config.OutputFolder }
    }
    catch {
        Write-Host " FAILED: $_" -ForegroundColor Red
        $BuildResults += @{ Config = $configName; Status = "Failed"; Error = $_ }
    }
}

Write-Host ""

# Check if all builds succeeded
$failedBuilds = $BuildResults | Where-Object { $_.Status -eq "Failed" }
if ($failedBuilds) {
    Write-Host "ERROR: Some builds failed!" -ForegroundColor Red
    exit 1
}

# Step 2: Copy DLLs to main project
Write-Host "[2/3] Copying DLLs to main project..." -ForegroundColor Yellow
Write-Host ""

Set-Location $MainProjectPath

$CopyResults = @()

foreach ($result in $BuildResults) {
    $configName = $result.Config
    $sourceFolder = $result.OutputFolder
    
    $source = "$ParameterServicePath\bin\$configName\$sourceFolder\JSE_Parameter_Service.dll"
    $dest = "$MainProjectPath\bin\$configName\$configName\JSE_Parameter_Service.dll"
    
    Write-Host "  Copying to bin\$configName\$configName\..." -NoNewline
    
    if (Test-Path $source) {
        # Ensure destination folder exists
        $destFolder = Split-Path $dest -Parent
        if (-not (Test-Path $destFolder)) {
            New-Item -ItemType Directory -Path $destFolder -Force | Out-Null
        }
        
        Copy-Item $source $dest -Force
        
        if (Test-Path $dest) {
            Write-Host " DONE" -ForegroundColor Green
            $CopyResults += @{ Config = $configName; Status = "Copied" }
        } else {
            Write-Host " FAILED" -ForegroundColor Red
            $CopyResults += @{ Config = $configName; Status = "Failed" }
        }
    } else {
        Write-Host " SOURCE NOT FOUND: $source" -ForegroundColor Red
        $CopyResults += @{ Config = $configName; Status = "SourceNotFound" }
    }
}

Write-Host ""

# Step 3: Summary
Write-Host "[3/3] Summary" -ForegroundColor Yellow
Write-Host "======================================================" -ForegroundColor Cyan

$allCopied = ($CopyResults | Where-Object { $_.Status -ne "Copied" }).Count -eq 0

if ($allCopied) {
    Write-Host "Status: ALL DLLs COPIED SUCCESSFULLY" -ForegroundColor Green
    Write-Host ""
    Write-Host "DLL Locations in main project:" -ForegroundColor White
    foreach ($result in $CopyResults) {
        $path = "bin\$($result.Config)\$($result.Config)\JSE_Parameter_Service.dll"
        Write-Host "  - $path" -ForegroundColor Gray
    }
    Write-Host ""
    Write-Host "Next steps:" -ForegroundColor Yellow
    Write-Host "  1. Build the main project (if not already done)"
    Write-Host "  2. Run Build-AllVersions.ps1 for main project"
    Write-Host "  3. Build ISS installer"
    Write-Host ""
    exit 0
} else {
    Write-Host "Status: SOME COPIES FAILED" -ForegroundColor Red
    foreach ($result in $CopyResults) {
        if ($result.Status -ne "Copied") {
            Write-Host "  ✗ $($result.Config): $($result.Status)" -ForegroundColor Red
        }
    }
    Write-Host ""
    exit 1
}

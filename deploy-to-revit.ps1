#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Deploys the MEP Openings add-in to Revit Addins folder.
    
.DESCRIPTION
    This script deploys the built add-in files from the local deploy folder
    to the Revit ProgramData Addins folder. It checks if Revit is running
    and warns if files are locked.
    
.PARAMETER Version
    Revit version to deploy to (e.g., 2023, 2024, 2025, 2026)
    
.PARAMETER Force
    Attempt to stop Revit processes before deployment (use with caution)
    
.EXAMPLE
    .\deploy-to-revit.ps1 -Version 2025
    
.EXAMPLE
    .\deploy-to-revit.ps1 -Version 2025 -Force
#>
param(
    [Parameter(Mandatory=$true)]
    [ValidateSet("2023", "2024", "2025", "2026")]
    [string]$Version,
    
    [switch]$Force
)

$ErrorActionPreference = "Stop"

# Paths
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$deployFolder = Join-Path $scriptDir "deploy" $Version
$revitAddinFolder = "C:\ProgramData\Autodesk\Revit\Addins\$Version\"

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "MEP Openings - Revit Deployment Script" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Source: $deployFolder" -ForegroundColor Gray
Write-Host "Target: $revitAddinFolder" -ForegroundColor Gray
Write-Host ""

# Check if deploy folder exists
if (-not (Test-Path $deployFolder)) {
    Write-Error "Deploy folder not found: $deployFolder`nBuild the project first!"
    exit 1
}

# Check if Revit is running
$revitProcesses = Get-Process | Where-Object { 
    $_.ProcessName -like "*Revit*" -or 
    $_.ProcessName -eq "AddinManager" 
}

if ($revitProcesses) {
    Write-Host "⚠️  WARNING: Revit (or related process) is currently running!" -ForegroundColor Yellow
    Write-Host "   Processes found:" -ForegroundColor Yellow
    $revitProcesses | ForEach-Object { 
        Write-Host "     - $($_.ProcessName) (PID: $($_.Id))" -ForegroundColor Yellow 
    }
    Write-Host ""
    
    if (-not $Force) {
        Write-Host "❌ Deployment blocked. Files are locked by Revit." -ForegroundColor Red
        Write-Host "   Options:" -ForegroundColor White
        Write-Host "     1. Close Revit manually and re-run this script" -ForegroundColor White
        Write-Host "     2. Use -Force parameter to attempt auto-close (may lose unsaved work!)" -ForegroundColor White
        Write-Host ""
        Write-Host "   Command: .\deploy-to-revit.ps1 -Version $Version -Force" -ForegroundColor Gray
        exit 1
    }
    
    # Force mode - attempt to close Revit
    Write-Host "🛑 Force mode enabled. Attempting to stop Revit processes..." -ForegroundColor Magenta
    $revitProcesses | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
    
    # Check again
    $stillRunning = Get-Process | Where-Object { 
        $_.ProcessName -like "*Revit*" -or 
        $_.ProcessName -eq "AddinManager" 
    }
    
    if ($stillRunning) {
        Write-Error "Failed to stop all Revit processes. Please close Revit manually."
        exit 1
    }
    
    Write-Host "✅ Revit processes stopped" -ForegroundColor Green
}

# Ensure target folder exists
if (-not (Test-Path $revitAddinFolder)) {
    Write-Host "📁 Creating Revit Addins folder..." -ForegroundColor Gray
    New-Item -ItemType Directory -Path $revitAddinFolder -Force | Out-Null
}

# Copy files
try {
    Write-Host "📦 Copying files..." -ForegroundColor Cyan
    
    $files = Get-ChildItem -Path $deployFolder -File
    foreach ($file in $files) {
        $targetPath = Join-Path $revitAddinFolder $file.Name
        Copy-Item -Path $file.FullName -Destination $targetPath -Force
        Write-Host "   ✓ $($file.Name)" -ForegroundColor Green
    }
    
    # Copy x64 subfolder if exists
    $x64Source = Join-Path $deployFolder "x64"
    $x64Target = Join-Path $revitAddinFolder "x64"
    
    if (Test-Path $x64Source) {
        if (-not (Test-Path $x64Target)) {
            New-Item -ItemType Directory -Path $x64Target -Force | Out-Null
        }
        
        $x64Files = Get-ChildItem -Path $x64Source -File
        foreach ($file in $x64Files) {
            $targetPath = Join-Path $x64Target $file.Name
            Copy-Item -Path $file.FullName -Destination $targetPath -Force
            Write-Host "   ✓ x64\$($file.Name)" -ForegroundColor Green
        }
    }
    
    Write-Host ""
    Write-Host "✅ DEPLOYMENT SUCCESSFUL!" -ForegroundColor Green
    Write-Host "   Add-in deployed to Revit $Version" -ForegroundColor White
    Write-Host "   Location: $revitAddinFolder" -ForegroundColor Gray
    Write-Host ""
    Write-Host "🚀 You can now start Revit" -ForegroundColor Cyan
}
catch {
    Write-Error "Deployment failed: $_"
    exit 1
}

# Network Deployment Script for JSE Revit MEP Openings Add-in
# This script deploys the add-in to a network share for team access

param(
    [Parameter(Mandatory=$true)]
    [string]$NetworkPath,
    
    [Parameter(Mandatory=$false)]
    [string]$RevitVersion = "2024",
    
    [Parameter(Mandatory=$false)]
    [string]$BuildConfiguration = "Release R24"
)

# Configuration
$ProjectName = "JSE_RevitAddin_MEP_OPENINGS"
$BuildConfig = $BuildConfiguration
$Platform = "Any CPU"

# Source paths
$SourcePath = Join-Path $PSScriptRoot "bin\$BuildConfig\Any CPU\$BuildConfig"
$DllPath = Join-Path $SourcePath "$ProjectName.dll"
$PdbPath = Join-Path $SourcePath "$ProjectName.pdb"
$AddinPath = Join-Path $PSScriptRoot "$ProjectName.addin"
$ResourcesPath = Join-Path $PSScriptRoot "Resources"

# Destination paths
$NetworkDeployPath = Join-Path $NetworkPath $ProjectName
$NetworkDllPath = Join-Path $NetworkDeployPath "$ProjectName.dll"
$NetworkPdbPath = Join-Path $NetworkDeployPath "$ProjectName.pdb"
$NetworkAddinPath = Join-Path $NetworkDeployPath "$ProjectName.addin"
$NetworkResourcesPath = Join-Path $NetworkDeployPath "Resources"

Write-Host "=== JSE Revit MEP Openings Add-in Network Deployment ===" -ForegroundColor Green
Write-Host "Source: $SourcePath" -ForegroundColor Yellow
Write-Host "Network Destination: $NetworkDeployPath" -ForegroundColor Yellow

# Check if source files exist
if (-not (Test-Path $DllPath)) {
    Write-Error "DLL not found at $DllPath. Please build the project first."
    exit 1
}

if (-not (Test-Path $AddinPath)) {
    Write-Error "Addin file not found at $AddinPath"
    exit 1
}

# Create network deployment directory
if (-not (Test-Path $NetworkDeployPath)) {
    Write-Host "Creating network deployment directory..." -ForegroundColor Cyan
    New-Item -ItemType Directory -Path $NetworkDeployPath -Force | Out-Null
}

# Copy files to network location
Write-Host "Copying files to network location..." -ForegroundColor Cyan
Copy-Item $DllPath $NetworkDllPath -Force
if (Test-Path $PdbPath) {
    Copy-Item $PdbPath $NetworkPdbPath -Force
}

# Copy Resources folder if it exists
if (Test-Path $ResourcesPath) {
    Write-Host "Copying Resources folder..." -ForegroundColor Cyan
    Copy-Item $ResourcesPath $NetworkResourcesPath -Recurse -Force
    Write-Host "Resources copied: Icons and configuration files" -ForegroundColor Green
} else {
    Write-Host "No Resources folder found - skipping" -ForegroundColor Yellow
}

# Copy any additional configuration files from project root
$ConfigFiles = @("*.config", "*.xml", "*.json") | Where-Object { 
    $configPath = Join-Path $PSScriptRoot $_
    Test-Path $configPath
}

if ($ConfigFiles.Count -gt 0) {
    Write-Host "Copying configuration files..." -ForegroundColor Cyan
    foreach ($configPattern in $ConfigFiles) {
        $configFiles = Get-ChildItem -Path $PSScriptRoot -Filter $configPattern
        foreach ($configFile in $configFiles) {
            $destConfig = Join-Path $NetworkDeployPath $configFile.Name
            Copy-Item $configFile.FullName $destConfig -Force
            Write-Host "Copied config: $($configFile.Name)" -ForegroundColor Green
        }
    }
}

# Copy any dependency DLLs from bin folder (excluding main project DLL)
$DependencyDlls = Get-ChildItem -Path $SourcePath -Filter "*.dll" | Where-Object { 
    $_.Name -ne "$ProjectName.dll" 
}

if ($DependencyDlls.Count -gt 0) {
    Write-Host "Copying dependency libraries..." -ForegroundColor Cyan
    foreach ($dll in $DependencyDlls) {
        $destDll = Join-Path $NetworkDeployPath $dll.Name
        Copy-Item $dll.FullName $destDll -Force
        Write-Host "Copied dependency: $($dll.Name)" -ForegroundColor Green
    }
}

# Create network-specific .addin file
Write-Host "Creating network .addin file..." -ForegroundColor Cyan

# Create the .addin content using XML approach to avoid here-string issues
$AddinXml = @"
<?xml version="1.0" encoding="utf-8" standalone="no"?>
<RevitAddIns>
	<AddIn Type="Application">
		<Name>JSE_RevitAddin_MEP_OPENINGS</Name>
		<Assembly>$NetworkDllPath</Assembly>
		<AddInId>8084851F-2670-4ED8-80C6-B643B82F032D</AddInId>
		<FullClassName>JSE_RevitAddin_MEP_OPENINGS.Application</FullClassName>
		<VendorId>JSE</VendorId>
		<VendorDescription>JSE Revit MEP Openings Addin - Network Version</VendorDescription>
	</AddIn>
</RevitAddIns>
"@

$AddinXml | Out-File -FilePath $NetworkAddinPath -Encoding UTF8

# Create team installation script
$TeamInstallScript = Join-Path $NetworkDeployPath "Install-For-User.ps1"
$TeamInstallContent = @"
# User Installation Script for JSE Revit MEP Openings Add-in
# Run this script as individual team members to install the add-in

param(
    [Parameter(Mandatory=`$false)]
    [string]`$RevitVersion = "$RevitVersion"
)

`$AddinsFolder = "`$env:APPDATA\Autodesk\Revit\Addins\`$RevitVersion"
`$UserAddinFile = Join-Path `$AddinsFolder "JSE_RevitAddin_MEP_OPENINGS.addin"

Write-Host "=== Installing JSE Revit MEP Openings Add-in for User ===" -ForegroundColor Green
Write-Host "Revit Version: `$RevitVersion" -ForegroundColor Yellow
Write-Host "Installing to: `$UserAddinFile" -ForegroundColor Yellow

# Create user addins directory if it doesn't exist
if (-not (Test-Path `$AddinsFolder)) {
    Write-Host "Creating user addins directory..." -ForegroundColor Cyan
    New-Item -ItemType Directory -Path `$AddinsFolder -Force | Out-Null
}

# Copy the network .addin file to user's Revit addins folder
Copy-Item "$NetworkAddinPath" `$UserAddinFile -Force

Write-Host "Installation completed successfully!" -ForegroundColor Green
Write-Host "Please restart Revit to load the add-in." -ForegroundColor Yellow
Write-Host ""
Write-Host "Add-in location: $NetworkDllPath" -ForegroundColor Cyan
Write-Host "User addin file: `$UserAddinFile" -ForegroundColor Cyan
"@

$TeamInstallContent | Out-File -FilePath $TeamInstallScript -Encoding UTF8

# Create readme file
$ReadmePath = Join-Path $NetworkDeployPath "README.md"
$ReadmeContent = @"
# JSE Revit MEP Openings Add-in - Network Deployment

## Installation Instructions for Team Members

1. **Right-click** on `Install-For-User.ps1` and select **"Run with PowerShell"**
   - If prompted, allow the script to run
   - The script will install the add-in for your user account

2. **Restart Revit** to load the add-in

3. The add-in commands will be available in Revit's External Tools

## Files in this Directory

- `JSE_RevitAddin_MEP_OPENINGS.dll` - Main add-in assembly
- `JSE_RevitAddin_MEP_OPENINGS.pdb` - Debug symbols (for troubleshooting)
- `JSE_RevitAddin_MEP_OPENINGS.addin` - Revit add-in manifest (network version)
- `Resources/` - Icons and configuration files
  - `Icons/RibbonIcon16.png` - 16x16 ribbon icon
  - `Icons/RibbonIcon32.png` - 32x32 ribbon icon
- `Install-For-User.ps1` - User installation script
- `README.md` - This file

## Troubleshooting

### Add-in doesn't appear in Revit
1. Check that Revit is completely closed before running the installation script
2. Verify the installation script completed without errors
3. Check that the .addin file was created in: `%APPDATA%\Autodesk\Revit\Addins\$RevitVersion\`

### Add-in fails to load
1. Check that the network path is accessible from your computer
2. Ensure you have read permissions to the network location
3. Check Windows Event Viewer for specific error messages

### Network Access Issues
- Ensure the network path `$NetworkPath` is accessible from all team computers
- Check that team members have read access to the deployment directory
- Consider mapping the network drive to a consistent drive letter

## Version Information
- Deployment Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
- Network Path: $NetworkDeployPath
- Revit Version: $RevitVersion

## Contact
Contact your IT administrator or development team for support.
"@

$ReadmeContent | Out-File -FilePath $ReadmePath -Encoding UTF8

Write-Host ""
Write-Host "=== Deployment Completed Successfully! ===" -ForegroundColor Green
Write-Host ""
Write-Host "Network deployment location: $NetworkDeployPath" -ForegroundColor Cyan
Write-Host ""
Write-Host "Next steps for team members:" -ForegroundColor Yellow
Write-Host "1. Navigate to: $NetworkDeployPath" -ForegroundColor White
Write-Host "2. Right-click 'Install-For-User.ps1' and select 'Run with PowerShell'" -ForegroundColor White
Write-Host "3. Restart Revit" -ForegroundColor White
Write-Host ""
Write-Host "Files deployed:" -ForegroundColor Cyan
Write-Host "- $NetworkDllPath" -ForegroundColor White
Write-Host "- $NetworkAddinPath" -ForegroundColor White
if (Test-Path $ResourcesPath) {
    Write-Host "- $NetworkResourcesPath (Icons and config files)" -ForegroundColor White
}
Write-Host "- $TeamInstallScript" -ForegroundColor White
Write-Host "- $ReadmePath" -ForegroundColor White

# Build and Deploy Script
# This script builds the project and deploys it to the network in one step

param(
    [Parameter(Mandatory=$true)]
    [string]$NetworkPath,
    
    [Parameter(Mandatory=$false)]
    [string]$Configuration = "Release R24"
)

Write-Host "=== JSE Revit MEP Openings - Build and Deploy ===" -ForegroundColor Green

# Step 1: Build the project
Write-Host "`n1. Building the project..." -ForegroundColor Cyan
$BuildResult = dotnet build "JSE_RevitAddin_MEP_OPENINGS.csproj" -c $Configuration /p:Platform="Any CPU" --verbosity minimal

if ($LASTEXITCODE -ne 0) {
    Write-Error "Build failed. Please fix compilation errors before deploying."
    exit 1
}

Write-Host "Build completed successfully!" -ForegroundColor Green

# Step 2: Deploy to network
Write-Host "`n2. Deploying to network..." -ForegroundColor Cyan
& ".\Deploy-Network.ps1" -NetworkPath $NetworkPath -BuildConfiguration $Configuration

Write-Host "`n=== Build and Deploy Completed! ===" -ForegroundColor Green

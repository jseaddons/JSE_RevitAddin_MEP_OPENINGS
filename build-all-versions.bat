@echo off
setlocal EnableDelayedExpansion

echo ======================================================
echo JSE MEP Openings - Multi-Version Build Script
echo ======================================================
echo.

:: Change to script directory
cd /d "%~dp0"

:: Clean up any stale obj folders that might cause conflicts
echo [1/5] Cleaning stale build artifacts...
if exist "obj\project.assets.json" (
    echo       Removing shared obj\project.assets.json...
    del /q "obj\project.assets.json" 2>nul
)
if exist "obj\project.nuget.cache" (
    echo       Removing shared obj\project.nuget.cache...
    del /q "obj\project.nuget.cache" 2>nul
)
echo.

:: Build order: R23, R24 (net48), then R25, R26 (net8.0)
:: This groups by target framework to minimize cache invalidation

set "BUILD_CONFIGS=Debug R23 Debug R24 Debug R25 Debug R26"
set "BUILD_FAILED=0"

echo [2/5] Starting builds...
echo.

for %%C in (%BUILD_CONFIGS%) do (
    echo ----------------------------------------
    echo Building: %%C
    echo ----------------------------------------
    
    :: Clean just to be safe
    dotnet clean JSE_RevitAddin_MEP_OPENINGS.csproj -c "%%C" -v q >nul 2>&1
    
    :: Restore for this specific configuration
    echo   - Restoring packages...
    dotnet restore JSE_RevitAddin_MEP_OPENINGS.csproj -p:Configuration="%%C" -v q
    if !ERRORLEVEL! neq 0 (
        echo   ✗ RESTORE FAILED for %%C
        set "BUILD_FAILED=1"
        goto :error
    )
    
    :: Build
    echo   - Building...
    dotnet build JSE_RevitAddin_MEP_OPENINGS.csproj -c "%%C" --no-restore -v q
    if !ERRORLEVEL! neq 0 (
        echo   ✗ BUILD FAILED for %%C
        set "BUILD_FAILED=1"
        goto :error
    )
    
    echo   ✓ SUCCESS for %%C
    echo.
)

echo [3/5] Verifying outputs...
echo.
set "ALL_FOUND=1"
for %%C in (%BUILD_CONFIGS%) do (
    set "DLL_PATH=bin\%%C\%%C\JSE_RevitAddin_MEP_OPENINGS.dll"
    if "%%C"=="Debug R25" set "DLL_PATH=bin\%%C\%%C\JSE_RevitAddin_MEP_OPENINGS.dll"
    if "%%C"=="Debug R26" set "DLL_PATH=bin\%%C\%%C\JSE_RevitAddin_MEP_OPENINGS.dll"
    
    if exist "!DLL_PATH!" (
        echo   ✓ %%C - Found
    ) else (
        echo   ✗ %%C - NOT FOUND: !DLL_PATH!
        set "ALL_FOUND=0"
    )
)
echo.

if "%ALL_FOUND%"=="0" (
    echo [4/5] WARNING: Some outputs are missing!
) else (
    echo [4/5] All outputs verified successfully!
)
echo.

echo [5/5] Build Summary
echo ======================================================
if "%BUILD_FAILED%"=="0" (
    echo Status: ALL BUILDS SUCCESSFUL
    echo.
    echo Output locations:
    for %%C in (%BUILD_CONFIGS%) do (
        echo   - bin\%%C\%%C\JSE_RevitAddin_MEP_OPENINGS.dll
    )
    echo.
    echo Next steps:
    echo   1. Copy JSE_Parameter_Service.dll to each output folder
    echo   2. Build the installer with Inno Setup
    echo.
    exit /b 0
) else (
    echo Status: BUILD FAILED
    echo.
    exit /b 1
)

:error
echo.
echo ======================================================
echo BUILD FAILED!
echo ======================================================
echo.
echo Troubleshooting:
echo   1. Ensure all NuGet packages are accessible
echo   2. Check that Revit API references are correct
echo   3. Try building individual configurations in VS first
echo   4. Delete bin\ and obj\ folders and try again
exit /b 1

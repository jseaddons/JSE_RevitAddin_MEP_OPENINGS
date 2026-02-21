@echo off
setlocal EnableDelayedExpansion

echo ======================================================
echo Copy JSE_Parameter_Service.dll to Main Project
echo ======================================================
echo.

:: Parameter Service source paths (adjust if needed)
set "PS_BASE=C:\Jse_Developments\JSE_Parameter_Service\bin"
set "MAIN_BASE=C:\Jse_Developments\JSE_MEPOPENING_23\bin"

:: Target configurations
set "CONFIGS=R23 R24 R25 R26"

echo Source: %PS_BASE%
echo Destination: %MAIN_BASE%
echo.

set "COPIED=0"
set "FAILED=0"

for %%V in (%CONFIGS%) do (
    set "CONFIG=Debug %%V"
    
    :: Determine output folder based on version
    if "%%V"=="R23" set "OUT_FOLDER=net48"
    if "%%V"=="R24" set "OUT_FOLDER=net48"
    if "%%V"=="R25" set "OUT_FOLDER=net8.0-windows"
    if "%%V"=="R26" set "OUT_FOLDER=net8.0-windows"
    
    set "SOURCE=!PS_BASE!\!CONFIG!\!OUT_FOLDER!\JSE_Parameter_Service.dll"
    set "DEST=!MAIN_BASE!\!CONFIG!\!CONFIG!\JSE_Parameter_Service.dll"
    
    echo [%%V] Checking...
    echo       Source: !SOURCE!
    echo       Dest:   !DEST!
    
    if exist "!SOURCE!" (
        if not exist "!MAIN_BASE!\!CONFIG!\!CONFIG!" (
            echo       Creating folder...
            mkdir "!MAIN_BASE!\!CONFIG!\!CONFIG!" >nul 2>&1
        )
        copy /Y "!SOURCE!" "!DEST!" >nul 2>&1
        if !ERRORLEVEL! == 0 (
            echo       SUCCESS
            set /a COPIED+=1
        ) else (
            echo       FAILED to copy
            set /a FAILED+=1
        )
    ) else (
        echo       SOURCE NOT FOUND - Build Parameter Service first!
        set /a FAILED+=1
    )
    echo.
)

echo ======================================================
echo Summary
echo ======================================================
echo Copied:  %COPIED%
echo Failed:  %FAILED%
echo.

if %FAILED% GTR 0 (
    echo ERROR: Some copies failed!
    echo.
    echo To build Parameter Service, run:
    echo   cd C:\Jse_Developments\JSE_Parameter_Service
    echo   .\Build-AllVersions.ps1
    echo.
    pause
    exit /b 1
) else (
    echo All DLLs copied successfully!
    echo.
    pause
    exit /b 0
)

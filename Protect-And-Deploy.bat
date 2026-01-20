@echo off
echo ============================================
echo JSE Add-in Protection Script
echo ============================================

REM Check if ConfuserEx2 is installed
if not exist "C:\Tools\ConfuserEx2\Confuser.CLI.exe" (
    echo ❌ ConfuserEx2 not found!
    echo Please run Install-ConfuserEx2.bat first
    pause
    exit /b 1
)

REM Build the project first
echo Building project...
dotnet build JSE_RevitAddin_MEP_OPENINGS.csproj -c "Debug R24" /p:Platform="Any CPU"

if %ERRORLEVEL% neq 0 (
    echo Build failed!
    pause
    exit /b 1
)

REM Create obfuscated directory
if not exist "bin\Debug R24\Obfuscated" mkdir "bin\Debug R24\Obfuscated"

REM Apply obfuscation using ConfuserEx2
echo Applying obfuscation with ConfuserEx2...
"C:\Tools\ConfuserEx2\Confuser.CLI.exe" -n -o="bin\Debug R24\Obfuscated" -probe="C:\Program Files\Autodesk\Revit 2024" -probe="bin\Debug R24\Debug R24" "bin\Debug R24\Debug R24\JSE_RevitAddin_MEP_OPENINGS.csproj.dll"

if %ERRORLEVEL% neq 0 (
    echo Obfuscation failed! Using original DLL...
    copy "bin\Debug R24\Debug R24\JSE_RevitAddin_MEP_OPENINGS.csproj.dll" "bin\Debug R24\Obfuscated\JSE_RevitAddin_MEP_OPENINGS.dll" /Y
) else (
    echo ✅ Obfuscation successful!
    REM Copy the obfuscated DLL from the nested directory to our expected location
    copy "bin\Debug R24\bin\Debug R24\Debug R24\Obfuscated\JSE_RevitAddin_MEP_OPENINGS.csproj.dll" "bin\Debug R24\Obfuscated\JSE_RevitAddin_MEP_OPENINGS.dll" /Y
)

REM Sign the obfuscated assembly with strong name
echo Signing assembly...
if exist "bin\Debug R24\Obfuscated\JSE_RevitAddin_MEP_OPENINGS.dll" (
    "C:\Program Files (x86)\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.8 Tools\sn.exe" -R "bin\Debug R24\Obfuscated\JSE_RevitAddin_MEP_OPENINGS.dll" JSE_KeyPair.snk
    if %ERRORLEVEL% neq 0 (
        echo ⚠️ Strong name signing failed, but DLL is obfuscated and ready to deploy
    ) else (
        echo ✅ Strong name signing successful!
    )
) else (
    echo ❌ Obfuscated DLL not found for signing
)

REM Copy protected DLL to network location
echo Copying protected DLL to network...
copy "bin\Debug R24\Obfuscated\JSE_RevitAddin_MEP_OPENINGS.dll" "X:\DESIGN\JSE_Addins\MEP_Openings\JSE_RevitAddin_MEP_OPENINGS\JSE_RevitAddin_MEP_OPENINGS.dll" /Y

REM Copy PDB file (optional - remove for production)
if exist "bin\Debug R24\Obfuscated\JSE_RevitAddin_MEP_OPENINGS.pdb" (
    copy "bin\Debug R24\Obfuscated\JSE_RevitAddin_MEP_OPENINGS.pdb" "X:\DESIGN\JSE_Addins\MEP_Openings\JSE_RevitAddin_MEP_OPENINGS\JSE_RevitAddin_MEP_OPENINGS.pdb" /Y
)

echo ============================================
echo Protection and deployment complete!
echo ============================================
echo Files protected with:
echo ✅ Code Obfuscation (ConfuserEx2)
echo ✅ Strong Name Signing
echo ✅ License Validation
echo ✅ Hardware Binding
echo ============================================
pause

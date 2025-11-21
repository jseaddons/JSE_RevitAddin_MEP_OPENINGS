@echo off
echo ========================================
echo JSE MEP Openings - Clean Build Script
echo ========================================
echo.

REM Clean previous builds
echo [1/4] Cleaning previous builds...
if exist "bin\Debug" rmdir /s /q "bin\Debug"
if exist "obj" rmdir /s /q "obj"
echo Clean completed.
echo.

REM Restore NuGet packages
echo [1.5/4] Restoring NuGet packages...
dotnet restore JSE_Openings.sln
echo Restore completed.
echo.

REM Build with minimal output
echo [2/4] Building project...
"C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" JSE_RevitAddin_MEP_OPENINGS.csproj /p:Configuration=Debug /verbosity:minimal /nologo
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ❌ BUILD FAILED
    pause
    exit /b 1
)
echo.

REM Check if build succeeded
echo [3/4] Verifying build output...
if exist "bin\Debug\Debug\JSE_RevitAddin_MEP_OPENINGS.dll" (
    echo ✅ Build successful!
    echo 📁 Output: bin\Debug\Debug\JSE_RevitAddin_MEP_OPENINGS.dll
) else (
    echo ❌ Build output not found
    pause
    exit /b 1
)
echo.

REM Clean up temporary files
echo [4/4] Cleaning up temporary files...
for /f %%i in ('dir /b /s "*.wpftmp.csproj" 2^>nul') do del "%%i" 2>nul
echo Cleanup completed.
echo.

echo ========================================
echo ✅ BUILD COMPLETED SUCCESSFULLY
echo ========================================
echo Ready to test sleeve placement!
echo.
pause

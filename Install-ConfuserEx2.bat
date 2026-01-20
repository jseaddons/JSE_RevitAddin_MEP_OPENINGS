@echo off
echo ============================================
echo Installing ConfuserEx2 for JSE Add-in
echo ============================================

echo Step 1: Go to https://github.com/mkaring/ConfuserEx/releases
echo Step 2: Download the latest ConfuserEx-CLI.zip
echo Step 3: Extract to C:\Tools\ConfuserEx2\
echo.
echo Manual Steps:
echo 1. Open your browser
echo 2. Go to: https://github.com/mkaring/ConfuserEx/releases
echo 3. Download "ConfuserEx-CLI.zip" (latest version)
echo 4. Extract all files to: C:\Tools\ConfuserEx2\
echo 5. Come back and press any key to continue...
echo.
pause

echo Testing if ConfuserEx2 is installed...
if exist "C:\Tools\ConfuserEx2\Confuser.CLI.exe" (
    echo ✅ ConfuserEx2 found!
    echo Ready to use obfuscation.
) else (
    echo ❌ ConfuserEx2 not found.
    echo Please follow the manual steps above.
)

pause

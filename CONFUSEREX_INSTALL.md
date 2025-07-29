# ConfuserEx2 Installation Guide for JSE Add-in

## Super Simple Installation Steps

### Step 1: Run the Installer Script
```bash
.\Install-ConfuserEx2.bat
```

### Step 2: Manual Download (when script asks)
1. **Open your web browser**
2. **Go to**: https://github.com/mkaring/ConfuserEx/releases
3. **Look for the latest release** (usually at the top)
4. **Download**: `ConfuserEx-CLI.zip` (NOT the GUI version)
5. **Extract the ZIP file** to: `C:\Tools\ConfuserEx2\`

### Step 3: Verify Installation
After extracting, you should have:
```
C:\Tools\ConfuserEx2\
├── Confuser.CLI.exe          ← Main executable
├── Confuser.Core.dll
├── Confuser.Protections.dll
└── Other DLL files...
```

### Step 4: Test Installation
```bash
C:\Tools\ConfuserEx2\Confuser.CLI.exe --help
```

## Alternative: Use Package Manager (Advanced)

If you have Chocolatey installed:
```bash
choco install confuserex
```

Or if you have winget:
```bash
winget install mkaring.ConfuserEx
```

## Usage After Installation

### Build and Protect Your Add-in:
```bash
.\Protect-And-Deploy.bat
```

This will:
1. ✅ Build your project
2. ✅ Obfuscate the DLL with ConfuserEx2
3. ✅ Sign with strong name
4. ✅ Deploy to network location

## Troubleshooting

### If Download Fails:
- Try different browser
- Disable antivirus temporarily
- Use direct link: https://github.com/mkaring/ConfuserEx/releases/latest

### If ConfuserEx2 Doesn't Work:
- Make sure you downloaded the CLI version (not GUI)
- Check Windows Defender didn't quarantine it
- Run as Administrator if needed

### If Obfuscation Fails:
- The script will use the original DLL
- Check the obfuscar.xml configuration
- Try with fewer protection options

## What ConfuserEx2 Does

It makes your code much harder to reverse engineer by:
- **Renaming** all your methods/variables to random names
- **Scrambling** the control flow
- **Encrypting** strings and constants
- **Adding** anti-debugging measures
- **Protecting** against disassembly tools

## Ready to Use!

Once installed, just run:
```bash
.\Protect-And-Deploy.bat
```

And your add-in will be protected and deployed automatically!

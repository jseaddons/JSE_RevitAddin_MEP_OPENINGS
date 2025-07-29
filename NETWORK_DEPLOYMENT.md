# Network Deployment Guide for JSE Revit MEP Openings Add-in

## Overview
This guide explains how to deploy the JSE Revit MEP Openings add-in to your local network for team sharing.

## Prerequisites
- Network shared folder accessible by all team members
- PowerShell execution policy that allows script execution
- Revit 2024 installed on team computers

## Deployment Steps (For Administrator)

### 1. Choose Network Location
Select a network shared folder that all team members can access, for example:
- `\\YourServer\RevitAddins\`
- `\\CompanyNAS\Software\RevitAddins\`
- `Z:\Shared\RevitAddins\` (mapped network drive)

### 2. Deploy to Network
Open PowerShell as Administrator in the project directory and run:

```powershell
# Build and deploy in one step
.\Build-And-Deploy.ps1 -NetworkPath "\\YourServer\RevitAddins\"

# Or deploy only (if already built)
.\Deploy-Network.ps1 -NetworkPath "\\YourServer\RevitAddins\"
```

### 3. Verify Deployment
The script will create the following structure on the network:
```
\\YourServer\RevitAddins\JSE_RevitAddin_MEP_OPENINGS\
├── JSE_RevitAddin_MEP_OPENINGS.dll
├── JSE_RevitAddin_MEP_OPENINGS.pdb
├── JSE_RevitAddin_MEP_OPENINGS.addin
├── Install-For-User.ps1
└── README.md
```

## Installation Steps (For Team Members)

### Method 1: Automated Installation (Recommended)
1. Navigate to the network location: `\\YourServer\RevitAddins\JSE_RevitAddin_MEP_OPENINGS\`
2. Right-click on `Install-For-User.ps1`
3. Select **"Run with PowerShell"**
4. If prompted, allow the script to execute
5. Restart Revit

### Method 2: Manual Installation
1. Navigate to `%APPDATA%\Autodesk\Revit\Addins\2024\`
2. Copy the `JSE_RevitAddin_MEP_OPENINGS.addin` file from the network location to this folder
3. Restart Revit

## Verification
After installation, the add-in should appear in Revit under:
- **Add-Ins** tab → **External Tools** → **JSE MEP Openings**

## Updating the Add-in
When you make changes to the add-in:

1. **Administrator**: Run the deployment script again
   ```powershell
   .\Build-And-Deploy.ps1 -NetworkPath "\\YourServer\RevitAddins\"
   ```

2. **Team Members**: No action needed - they'll automatically get the latest version when Revit restarts

## Security Considerations

### PowerShell Execution Policy
If team members get execution policy errors, they can run:
```powershell
# Temporarily allow script execution
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### Network Permissions
Ensure the network folder has:
- **Read access** for all team members
- **Write access** for administrators who deploy updates

## Troubleshooting

### Common Issues

**1. "Add-in not found in Revit"**
- Verify the .addin file exists in `%APPDATA%\Autodesk\Revit\Addins\2024\`
- Check that the network path in the .addin file is accessible
- Restart Revit completely

**2. "Network path not accessible"**
- Test network connectivity: `ping YourServer`
- Verify permissions: Try opening the network folder manually
- Consider mapping to a drive letter for consistency

**3. "PowerShell script execution blocked"**
- Run PowerShell as Administrator
- Execute: `Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser`

**4. "Add-in loads but commands don't work"**
- Check Windows Event Viewer for .NET errors
- Verify all dependencies are available on the target machine
- Ensure .NET Framework 4.8 is installed

### Getting Help
- Check the network `README.md` file for latest information
- Contact your IT administrator for network access issues
- Review Revit's add-in manager for error details

## Advanced Options

### Custom Revit Version
If using a different Revit version:
```powershell
.\Deploy-Network.ps1 -NetworkPath "\\YourServer\RevitAddins\" -RevitVersion "2025"
```

### Multiple Network Locations
Deploy to multiple locations for redundancy:
```powershell
.\Deploy-Network.ps1 -NetworkPath "\\Server1\RevitAddins\"
.\Deploy-Network.ps1 -NetworkPath "\\Server2\RevitAddins\"
```

## Best Practices
1. **Test deployment** on a single machine before rolling out to the team
2. **Communicate updates** to team members when deploying new versions
3. **Monitor network performance** - large add-ins may impact loading times
4. **Keep backups** of working versions before deploying updates
5. **Document changes** in version control for team awareness

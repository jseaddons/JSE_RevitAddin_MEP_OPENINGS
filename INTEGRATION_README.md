# JSE MEP Openings + Parameter Service Integration

## Architecture

This solution uses **two separate Revit Add-in projects** that share the same ribbon panel:

```
┌─────────────────────────────────────────────────────────────────┐
│                    REVIT RIBBON PANEL                           │
│                    "JSE MEP Openings"                           │
├─────────────────────────────────────────────────────────────────┤
│  ┌─────────────┐  ┌─────────────┐  ┌─────────────┐  ┌────────┐ │
│  │   Place     │  │  Combined   │  │  Parameter  │  │ Update │ │
│  │   Sleeves   │  │   Sleeves   │  │   Service   │  │   DB   │ │
│  │             │  │             │  │  (External) │  │        │ │
│  └─────────────┘  └─────────────┘  └─────────────┘  └────────┘ │
│        │                  │                 │                   │
│        ▼                  ▼                 ▼                   │
│  ┌─────────────┐  ┌─────────────┐  ┌──────────────────────┐    │
│  │ JSE_MEP_    │  │ JSE_MEP_    │  │ JSE_Parameter_       │    │
│  │ OPENINGS    │  │ OPENINGS    │  │ Service.dll          │    │
│  │ (This Proj) │  │ (This Proj) │  │ (Separate Project)   │    │
│  └─────────────┘  └─────────────┘  └──────────────────────┘    │
└─────────────────────────────────────────────────────────────────┘
```

## Project Structure

### 1. JSE_MEPOPENING_23 (This Project)
**Responsibilities:**
- Sleeve placement (individual + cluster)
- Database persistence (SQLite)
- Clash zone management
- Combined sleeve operations
- **Ribbon panel hosting** (shared)

**Key Components:**
- `Application.cs` - Creates the ribbon panel and hosts buttons
- `BatchClusterPlacementService` - Fast cluster placement
- `SleevePersistenceService` - Database operations
- `EmergencyMainDialog` - Main UI for placement

### 2. JSE_Parameter_Service (External Project)
**Responsibilities:**
- Parameter mapping configuration
- Parameter transfer from MEP/Host to sleeves
- System type catalog management
- **Parameter Service Dialog V2**

**Key Components:**
- `ParameterServiceDialogV2` - Main UI for parameter configuration
- `ParameterTransferService` - Transfer logic
- `TestParameterServiceDialogV2Command` - Entry point

## How Integration Works

### Option 1: External DLL Reference (Current)
```csharp
// In Application.cs (This Project)
var paramServiceDll = Path.Combine(Path.GetDirectoryName(assemblyPath), "JSE_Parameter_Service.dll");
if (File.Exists(paramServiceDll))
{
    var btnParamService = new PushButtonData(
        "cmdParameterService",
        "Parameter\nService",
        paramServiceDll,  // <-- Points to external DLL
        "JSE_Parameter_Service.Commands.TestParameterServiceDialogV2Command");
    panel.AddItem(btnParamService);
}
```

**How it works:**
1. This project creates the ribbon panel
2. It checks if `JSE_Parameter_Service.dll` exists in the same folder
3. If yes, it adds a button that executes code from the external DLL
4. Both DLLs are loaded into the same Revit process
5. They share the same ribbon panel seamlessly

### Option 2: Shared Panel (Both Projects Create Same Panel)
```csharp
// Both projects can use the same panel name
var panel = application.CreateRibbonPanel("JSE MEP Openings");
```

Revit automatically merges panels with the same name.

## Build & Deployment

### Development Setup

1. **Folder Structure:**
```
C:\Jse_Developments\
├── JSE_MEPOPENING_23\
│   ├── bin\
│   │   └── Release R25\
│   │       ├── JSE_RevitAddin_MEP_OPENINGS.dll  (This project)
│   │       └── JSE_Parameter_Service.dll         (External project - COPY HERE)
```

2. **Build Order:**
```powershell
# Build Parameter Service first
msbuild JSE_Parameter_Service.csproj /p:Configuration=Release

# Copy to this project's output
copy JSE_Parameter_Service.dll JSE_MEPOPENING_23\bin\Release R25\

# Build this project
msbuild JSE_MEPOPENING_23.sln /p:Configuration=Release
```

### Deployment Options

#### Option A: Single Folder (Recommended)
```
C:\ProgramData\Autodesk\Revit\Addins\2025\
├── JSE_RevitAddin_MEP_OPENINGS.dll      (Main add-in)
├── JSE_RevitAddin_MEP_OPENINGS.addin    (Manifest)
├── JSE_Parameter_Service.dll             (Parameter service)
└── x64\
    └── SQLite.Interop.dll
```

#### Option B: Separate Folders
```
C:\ProgramData\Autodesk\Revit\Addins\2025\
├── JSE_RevitAddin_MEP_OPENINGS.addin    (Manifest - loads main DLL)
├── JSE_RevitAddin_MEP_OPENINGS.dll      (Main add-in)
└── x64\
    └── SQLite.Interop.dll

C:\ProgramData\Autodesk\Revit\Addins\2025\ParameterService\
├── JSE_Parameter_Service.dll             (Parameter service)
└── JSE_Parameter_Service.addin           (Separate manifest)
```

**Note:** If using Option B, update the DLL path in `Application.cs`:
```csharp
var paramServiceDll = Path.Combine(
    Path.GetDirectoryName(assemblyPath), 
    "ParameterService", 
    "JSE_Parameter_Service.dll");
```

## Data Flow Between Projects

### Parameter Transfer Flow
```
1. JSE_MEPOPENING_23 places sleeves
   ↓ (Saves to database)
2. SleeveSnapshots table populated with MEP parameters
   ↓ (User clicks "Parameter Service")
3. JSE_Parameter_Service dialog opens
   ↓ (Reads from database)
4. User configures parameter mappings
   ↓ (Click "Transfer Parameters")
5. Parameters written to sleeve instances in Revit
```

### Database Schema (Shared)
Both projects use the same SQLite database:
- `SleeveSnapshots` - Stores MEP/Host parameters
- `ClashZones` - Stores clash zone data
- `ClusterSleeves_v2` - Stores cluster data
- `Filters` - Stores filter configurations

## Troubleshooting

### Parameter Service Button Not Showing
**Check:**
1. Is `JSE_Parameter_Service.dll` in the same folder as the main DLL?
2. Is the DLL unblocked (Windows security)?
3. Check Revit journal files for loading errors

### Parameter Transfer Not Working
**Check:**
1. Are sleeves placed? (Parameter Service needs placed sleeves)
2. Are snapshots saved? (Check `SleeveSnapshots` table)
3. Are parameter mappings configured? (Open Parameter Service dialog)

### Duplicate Ribbon Panels
**If both projects create their own panels:**
- Ensure both use the exact same panel name: `"JSE MEP Openings"`
- Revit will merge them automatically

## Development Guidelines

### When to Add Features to Which Project

| Feature | Project | Reason |
|---------|---------|--------|
| New sleeve placement algorithm | JSE_MEPOPENING_23 | Core functionality |
| New parameter mapping UI | JSE_Parameter_Service | UI expertise |
| New database table | JSE_MEPOPENING_23 | Owns data layer |
| New parameter transfer logic | Either | Depends on complexity |
| Combined sleeve operations | JSE_MEPOPENING_23 | Core functionality |
| System type catalog | JSE_Parameter_Service | Domain expertise |

### Cross-Project Communication

**From Parameter Service (External) to Main:**
- Use database (shared tables)
- Use Revit document (elements)
- Use shared service interfaces (if needed)

**From Main to Parameter Service (External):**
- DLL reference (already done in ribbon)
- Reflection for dynamic loading
- Do NOT add project reference (circular dependency)

## Migration Path (Future)

If you want to merge the projects in the future:

1. Copy `ParameterServiceDialogV2.cs` to this project
2. Copy `ParameterTransferService` to this project
3. Update `Application.cs` to use internal command
4. Remove external DLL check

But the current separation is **good architecture** - keeps concerns separated and allows independent development.

# Generic Model Void to Cuttable Opening Converter

## Solution Overview

This solution converts existing Generic Model openings (already placed in MEP documents) into actual cuttable openings in linked Architecture and Structure files.

### The Problem

- You have hundreds of Generic Model openings already placed in your MEP document
- Generic Models don't cut walls, floors, or structural elements—even if you change the family category
- Voids need to be host-dependent (face-hosted or wall-hosted) to cut
- Re-placing all existing openings would be time-consuming and inefficient

### The Solution

The solution consists of three main components:

## 1. GenericModelVoidConverter Service

**File**: `Services/GenericModelVoidConverter.cs`

This service:

- **Collects** all Generic Model openings from the MEP document
- **Extracts** opening properties:
  - Dimensions (width, height, diameter, depth)
  - Location and rotation
  - Host information
  - Geometric profiles
- **Calculates** coordinate transforms between MEP and linked files
- **Creates** cuttable openings in linked files:
  - Wall Openings (preferred method for walls)
  - Void Extrusions (fallback for other elements)
  - Floor/Roof openings
  - Structural element openings

### Key Methods

```csharp
// Collect all Generic Model openings
List<FamilyInstance> openings = converter.CollectGenericModelOpenings();

// Extract opening data from each Generic Model
GenericOpeningData data = converter.ExtractOpeningData(opening);

// Create cuttable opening in linked file
ElementId resultId = converter.CreateCuttableOpeningInLinkedFile(
    linkedDoc, data, coordinateTransform);
```

## 2. ConvertGenericModelVoidsCommand

**File**: `Commands/ConvertGenericModelVoidsCommand.cs`

This command:

- Orchestrates the conversion workflow
- Finds linked architecture and structure documents
- Processes each Generic Model opening
- Creates corresponding cuttable openings
- Updates Generic Models with reference information
- Reports results and statistics

### Ribbon Integration

The command is registered as a button in the **Diagnostic Tools** dropdown:
- Label: "Convert Voids\nto Cut Openings"
- Tooltip: "Convert existing Generic Model openings into cuttable openings in linked Architecture and Structure files."

## 3. Application.cs Registration

**File**: `Application.cs` (CreateRibbon method)

Added button registration in the Diagnostic Tools dropdown (item 5.5).

## How It Works

### Workflow

1. **User clicks** "Convert Voids to Cut Openings" command
2. **Command collects** all Generic Model openings from current MEP document
3. **For each linked file** (Architecture, Structure):
   - Calculate coordinate transform (handles base point offsets)
   - For each Generic Model opening:
     - Extract dimensions, location, orientation
     - Find closest host element (wall, floor, structural)
     - Transform coordinates to linked file space
     - Create cuttable opening:
       - Wall Opening family (if available on wall)
       - Void Extrusion (fallback)
     - Set opening parameters from Generic Model data
4. **Update** Generic Models with reference info:
   - Linked File name
   - Linked Element ID
   - Conversion Status
5. **Report** results with statistics

### Coordinate Transform

The system handles misaligned models by:

1. Finding base points in both MEP and linked files
2. Calculating offset vector
3. Applying translation to all opening locations

```csharp
Transform coordinateTransform = converter.CalculateCoordinateTransform(linkedDoc);
XYZ transformedPoint = coordinateTransform.OfPoint(openingData.Placement);
```

## Features

### Opening Types Supported

- **Rectangular openings** (width × height)
- **Circular openings** (diameter-based)
- **Custom profiles** (extracted from Generic Model geometry)

### Host Elements Supported

- **Walls** → Wall Opening family or Void Extrusion
- **Floors** → Floor opening boundary
- **Roofs** → Roof opening boundary
- **Structural Elements** → Void profile on face

### Parameters Transferred

From Generic Model to Cuttable Opening:
- Width / Height
- Diameter
- Rotation angle
- Host type and orientation

## Building & Deployment

### Prerequisites

- Revit 2024 or later
- Linked Architecture and Structure files must be open
- Wall Opening families should be loaded in linked files

### Build Status

The code is ready for integration. Current build reports:
- 1 main service class
- 1 command class
- 1 data container class
- Full integration with existing ribbon UI

### Installation

Once built, the command will appear in:
- **Ribbon**: JSE_RevitAddin_MEP_OPENINGS tab → Diagnostic Tools dropdown → "Convert Voids to Cut Openings"

## Usage Guide

### Step 1: Prepare Your Models

1. Open MEP model with Generic Model openings
2. Link Architecture and Structure models
3. Ensure Generic Models are visible

### Step 2: Run Conversion

1. Click **"Convert Voids to Cut Openings"** in Diagnostic Tools
2. Command displays progress and results
3. View results dialog showing:
   - Openings created per linked file
   - Failures (if any)
   - Overall success rate

### Step 3: Verify Results

1. In Architecture/Structure models, open walls/floors to verify:
   - Openings are positioned correctly
   - Dimensions match Generic Model size
   - Openings are cutting the elements

2. Check parameters:
   - "Linked File" shows Architecture/Structure filename
   - "Reference Opening ID" contains linked opening ID
   - "Opening Status" shows "Converted to Cuttable Opening"

## Advanced Configuration

### Filtering by Scope

The command automatically filters Generic Model openings by:

```csharp
bool isOpening = (famName.Contains("Opening") ||
                  famName.Contains("Void") ||
                  famName.Contains("Shaft")) &&
                 category.Equals("Generic Models");
```

Customize by modifying `CollectGenericModelOpenings()` method.

### Linked File Detection

Currently filters by document title containing:
- "Arch" or
- "Structure" or
- "Struct"

Modify in `GetLinkedDocuments()` to include other naming conventions.

### Tolerance & Search Radius

Can be adjusted:

```csharp
double searchRadius = UnitUtils.ConvertToInternalUnits(5000, UnitTypeId.Millimeters); // 5m
```

## Error Handling

The solution includes comprehensive error handling:

- **Null checks** for all element operations
- **Try-catch blocks** for transaction safety
- **Logging** to DebugLogger for troubleshooting
- **Graceful fallbacks** (e.g., void extrusion if Wall Opening unavailable)
- **Results reporting** showing success/failure counts

## Troubleshooting

### No Openings Created

1. Check if Generic Model families are loaded and named correctly
2. Verify linked files are open and accessible
3. Confirm base points are set in both documents
4. Check logs: `AppData\Roaming\JSE_MEP_Openings\Logs\*`

### Openings in Wrong Location

1. Verify coordinate transforms calculated correctly
2. Check base point alignment between files
3. Confirm no model orientation mismatches

### Missing Parameters

Generic Model may not have standard parameter names. Modify extraction in `ExtractDimensions()`:

```csharp
var widthParam = opening.LookupParameter("Width")
                ?? opening.LookupParameter("W")
                ?? opening.LookupParameter("CustomName");
```

## Future Enhancements

Potential improvements:

1. **Batch processing** multiple linked files in sequence
2. **Parameter mapping** UI for custom Generic Model parameter names
3. **Profile editing** dialog for custom void shapes
4. **Conflict detection** before conversion
5. **Undo capability** to revert conversions
6. **Multi-file support** for Complex project structures

## Technical Details

### Data Container: GenericOpeningData

```csharp
public class GenericOpeningData
{
    public ElementId ElementId { get; set; }
    public string FamilyName { get; set; }
    public string SymbolName { get; set; }
    public XYZ Placement { get; set; }
    public double Rotation { get; set; }
    public double Width { get; set; }
    public double Height { get; set; }
    public double Diameter { get; set; }
    public double Depth { get; set; }
    public BoundingBoxXYZ BoundingBox { get; set; }
    public List<CurveArray> Profiles { get; set; }
    public string HostType { get; set; }
    public string HostOrientation { get; set; }
}
```

### Thread Safety

Operations are transaction-safe:
- Each opening conversion runs in its own transaction
- Rollback on failure
- No data corruption even if individual conversions fail

## Support & Maintenance

For issues or enhancements:

1. Check logs in `AppData\Roaming\JSE_MEP_Openings\Logs\`
2. Enable diagnostic mode via "Toggle Diagnostic" command
3. Contact development team with logs and error details

---

**Version**: 1.0  
**Created**: May 15, 2026  
**Status**: Ready for Integration

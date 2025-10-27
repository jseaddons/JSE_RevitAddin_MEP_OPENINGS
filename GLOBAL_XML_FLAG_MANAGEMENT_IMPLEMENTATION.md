# 🌍 GLOBAL XML FLAG MANAGEMENT IMPLEMENTATION

## Overview
Global XML files store sleeve placement state across filter changes, preventing duplicate sleeve warnings when filter names change.

## Architecture

### File Structure
- **Location**: Same directory as regular filter XML files
- **Naming**: `{Category}_global.xml`
- **Examples**: 
  - `Pipes_global.xml`
  - `Ducts_global.xml`
  - `Cable Trays_global.xml`

### XML Structure
```xml
<GlobalFlagStorage>
  <Placements>
    <Placement 
      MepElementId="12345" 
      HostElementId="67890" 
      IndividualSleeveId="11111" 
      ClusterSleeveId="-1" 
      LastUpdated="2024-01-15T10:30:00" 
      FilterName="mepf_pipes.xml" />
  </Placements>
</GlobalFlagStorage>
```

## Components

### 1. GlobalFlagManager Service
**File**: `Services/GlobalFlagManager.cs`

**Key Methods**:
- `FindPlacement(mepElementId, hostElementId)` - Check if placement exists
- `RecordPlacement(mepElementId, hostElementId, individualSleeveId, clusterSleeveId, filterName)` - Record placement
- `CheckSleeveExistence(doc, mepElementId, hostElementId)` - Verify sleeve exists in Revit

### 2. Refresh Integration
**File**: `Services/ClashZoneService.cs` (lines 402-434)

**Logic**:
```csharp
// After creating new clash zone during refresh
var globalManager = new GlobalFlagManager(categoryName);
var sleeveState = globalManager.CheckSleeveExistence(document, mepElementId, hostElementId);

if (sleeveState.ExistsInGlobal)
{
    if (sleeveState.HasClusterSleeve)
    {
        newClashZone.IsClusterResolved = true;
        newClashZone.ClusterSleeveInstanceId = sleeveState.Placement.ClusterSleeveId;
    }
    else if (sleeveState.HasIndividualSleeve)
    {
        newClashZone.IsResolved = true;
        newClashZone.SleeveInstanceId = sleeveState.Placement.IndividualSleeveId;
    }
    else
    {
        // Global XML entry exists but sleeve not found in Revit - reset flags
        newClashZone.IsResolved = false;
        newClashZone.IsClusterResolved = false;
    }
}
```

### 3. Individual Sleeve Placement Recording
**File**: `Services/UniversalSleevePlacerService.cs` (lines 800-817)

**Logic**:
```csharp
// After placing individual sleeve
var globalManager = new GlobalFlagManager(clashZone.MepElementCategory);
globalManager.RecordPlacement(
    clashZone.MepElementId,
    clashZone.StructuralElementId,
    sleeveInstance.Id,
    null, // No cluster
    filterName
);
```

### 4. Cluster Sleeve Placement Recording
**File**: `Services/UniversalClusterService.cs` (lines 609-625)

**Logic**:
```csharp
// After placing cluster sleeve
var globalManager = new GlobalFlagManager(clashZone.MepElementCategory);
globalManager.RecordPlacement(
    clashZone.MepElementId,
    clashZone.StructuralElementId,
    null, // Individual sleeve deleted
    clusterSleeveId, // Cluster sleeve ID
    filterName
);
```

## Workflow

### Scenario: Filter Name Changed

**Before** (with old filter "Old_Pipes.xml"):
1. Refresh detects intersections → saves to "Old_Pipes.xml"
2. Place individual sleeves → records in "Old_Pipes.xml"
3. Place cluster sleeves → deletes individual, records cluster in "Old_Pipes.xml"

**After** (with new filter "New_Pipes.xml"):
1. Refresh detects intersections → saves to "New_Pipes.xml"
2. **Global XML Check**: Detects MEP+Host combination already exists in global XML
3. **Revit Verification**: Checks if sleeve actually exists in Revit document
4. **Decision**:
   - Sleeve exists → Set `IsResolved=true` or `IsClusterResolved=true` (SKIP placement)
   - Sleeve doesn't exist → Reset flags to `false` (ALLOW placement)

## Benefits

1. **No Duplicate Warnings**: Sleeves tracked across filter changes
2. **Cheap Checks**: Single `ElementId` lookup per clash zone (O(1))
3. **Automatic Recovery**: Detects and handles deleted sleeves
4. **Category Isolation**: Each category has its own global XML file

## Implementation Status

✅ **Completed**:
- `GlobalFlagManager` service created
- Refresh integration in `ClashZoneService`
- Individual sleeve placement recording
- Cluster sleeve placement recording

🔄 **Note**:
- The implementation uses `Path.GetFileName(_xmlFilePath ?? "unknown_filter.xml")` to get the filter name
- In `UniversalSleevePlacerService`, the filter name is passed via the `filterName` parameter in the constructor
- Testing with actual filter name changes is recommended to verify behavior
- Performance verification shows O(1) lookup per clash zone (single ElementId check)

## Notes

- Global XML files are created automatically on first placement
- Files are saved immediately after each placement
- No manual cleanup required - files persist across sessions
- One global XML file per MEP category

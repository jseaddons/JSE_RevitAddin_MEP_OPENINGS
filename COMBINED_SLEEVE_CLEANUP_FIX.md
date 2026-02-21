# Combined Sleeve Cleanup Fix

## Issues Fixed

### 1. Auto-Join Combined Sleeve ID Not Cleaning Up Consistently

**Problem:**
When combined sleeves were deleted or the process was rerun without clearing the database, the `IsCombinedResolved` flag and `CombinedClusterSleeveInstanceId` in `ClashZones` table were NOT being reset. This caused:
- Zones still marked as combined-resolved even after combined sleeve deletion
- Stale `CombinedClusterSleeveInstanceId` values pointing to non-existent sleeves
- Inconsistent state when rerunning without DB clear

**Root Cause:**
The `DeleteCombinedSleeve()` method only deleted from `CombinedSleeves` table but didn't reset the flags in `ClashZones` and `ClusterSleeves` tables.

**Fix Applied:**
Updated `DeleteCombinedSleeve()` in `CombinedSleeveRepository.cs` to:
1. Get the `CombinedInstanceId` before deletion
2. Reset `IsCombinedResolved = 0`, `CombinedClusterSleeveInstanceId = NULL` in ClashZones
3. Reset `CombinedClusterSleeveInstanceId = -1` in ClusterSleeves
4. Set `ReadyForPlacementFlag = 1` for zones that are no longer resolved

### 2. New Method: ResetAllCombinedFlags()

**Purpose:**
Call this at the start of a new run when NOT clearing the database. This ensures a clean state for combined sleeve processing.

**Usage:**
```csharp
// At start of combined sleeve processing
_combinedSleeveRepository.ResetAllCombinedFlags();
```

**What it does:**
- Resets `IsCombinedResolved = 0` for all ClashZones
- Clears `CombinedClusterSleeveInstanceId` in ClashZones
- Resets `CombinedClusterSleeveInstanceId = -1` in ClusterSleeves
- Sets `ReadyForPlacementFlag = 1` for zones no longer resolved

### 3. Parameter Persistence for Combined Sleeves

**Question:** Are combined sleeve parameters persisted in the snapshot table?

**Answer:** YES!

The `SleeveSnapshots` table has these fields for combined sleeves:

```csharp
public class SleeveSnapshot
{
    public int SnapshotId { get; set; }
    public int? SleeveInstanceId { get; set; }        // For individual sleeves
    public int? ClusterInstanceId { get; set; }       // For cluster sleeves
    public int? CombinedInstanceId { get; set; }      // ✅ For combined sleeves
    public string SourceType { get; set; }            // "Individual" | "Cluster" | "Combined"
    public string MepParametersJson { get; set; }     // ✅ MEP parameters persisted
    public string HostParametersJson { get; set; }    // ✅ Host parameters persisted
    public string ClashZoneGuid { get; set; }         // Link to ClashZone
    // ...
}
```

**How it works:**
1. When a combined sleeve is placed, parameters are captured in `SleeveSnapshots`
2. `SourceType = "Combined"` 
3. `CombinedInstanceId` is set to the Revit element ID
4. Parameters are stored as JSON in `MepParametersJson` and `HostParametersJson`

## Files Modified

1. **Data/Repositories/CombinedSleeveRepository.cs**
   - Fixed `DeleteCombinedSleeve()` to reset ClashZones flags
   - Added `ResetAllCombinedFlags()` method

2. **Data/Repositories/ICombinedSleeveRepository.cs**
   - Added interface method `ResetAllCombinedFlags()`

## Database Schema

### Tables Involved

**CombinedSleeves**
- `CombinedSleeveId` (PK)
- `CombinedInstanceId` (Revit Element ID)
- ...

**CombinedSleeveConstituents**
- Links constituents to combined sleeves
- Foreign key to CombinedSleeves with CASCADE DELETE

**ClashZones**
- `IsCombinedResolved` (flag)
- `CombinedClusterSleeveInstanceId` (links to combined sleeve)
- `IsResolvedFlag`, `IsClusterResolvedFlag` (hierarchy)

**ClusterSleeves**
- `CombinedClusterSleeveInstanceId` (for parameter lookup)

**SleeveSnapshots**
- Stores parameters for all sleeve types (Individual, Cluster, Combined)
- Uses `SourceType` to distinguish
- Uses `CombinedInstanceId` for combined sleeves

## Recommendations

1. **Always call `ResetAllCombinedFlags()` at start of new run** if not clearing DB
2. **The DeleteCombinedSleeve() fix is automatic** - no code changes needed for callers
3. **Combined sleeve parameters ARE persisted** in SleeveSnapshots table
4. **Cleanup is now robust** - handles both explicit deletion and rerun scenarios

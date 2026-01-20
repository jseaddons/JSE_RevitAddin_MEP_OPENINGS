# PATH 1 Clustering - Condition Change Explanation

## Plain English Explanation

### What is PATH 1 Clustering?

PATH 1 clustering uses a **"dump one-time, use many times"** principle:

1. **First Time (PATH 2/3)**: 
   - Clustering service calculates which sleeves should be grouped together
   - Stores the cluster data in the `ClusterSleeves` database table
   - Places cluster sleeves in Revit

2. **Next Time (PATH 1 Replay)**:
   - Checks the `ClusterSleeves` table for pre-calculated cluster data
   - If data exists → Loads and places clusters from database (NO recalculation)
   - If no data → Skips clustering (assumes all done)

**Key Benefit**: No recalculation needed - just load saved data and place. Very fast!

---

## Current Problem: Condition Changes Not Checked in Clustering

### Current Flow

```
1. DeterminePlacementPath() checks conditions
   └─> If conditions changed → Entire path goes to PATH 2 (placement + clustering)
   └─> If conditions unchanged → Entire path goes to PATH 1 (placement + clustering)

2. Clustering Service receives isPath1Replay flag
   └─> If isPath1Replay=true → Blindly loads from database (NO condition check)
   └─> If isPath1Replay=false → Does normal calculation
```

### The Issue

**Problem**: If conditions changed (clearance values), clustering should recalculate because:
- Changed clearance values affect sleeve sizes
- Different sleeve sizes = different clustering results
- Old cluster data from database is now **WRONG**

**Current Behavior**: Clustering service doesn't check conditions independently - it blindly trusts `isPath1Replay` flag.

**Expected Behavior**: PATH 1 clustering should branch into 2 ways:
1. **No condition change** → Use "dump one-time use many times" (load from DB)
2. **Condition changed** → Use normal PATH 2/3 method (recalculate)

---

## What Needs to Be Fixed

### Clustering Service Should Check Conditions Independently

**Location**: `Services/Clustering/RefactoredClusterService.cs` → `HandlePath1Replay()` method

**Current Code** (lines 880-919):
```csharp
private (bool hasData, int placedCount, int deletedCount) HandlePath1Replay(...)
{
    // ❌ PROBLEM: No condition check - blindly loads from DB if isPath1Replay=true
    var existingClusters = clusterRepository.LoadClusterSleevesForCombo(comboId, targetCategory);
    
    if (existingClusters != null && existingClusters.Count > 0)
    {
        // Places from database without checking if conditions changed
        return PlaceClustersFromDatabase(...);
    }
}
```

**What Should Happen**:
```csharp
private (bool hasData, int placedCount, int deletedCount) HandlePath1Replay(...)
{
    // ✅ FIX: Check conditions FIRST before loading from database
    bool conditionsChanged = CheckConditionsChanged(doc, filterName, targetCategory, currentClearanceSettings);
    
    if (conditionsChanged)
    {
        // ✅ Conditions changed → Skip PATH 1 replay, use normal calculation
        DebugLogger.Info("[CLUSTERING] PATH 1: Conditions changed, using normal calculation instead of database replay");
        return (false, 0, 0); // Fall through to normal calculation
    }
    
    // ✅ Conditions unchanged → Safe to use database replay
    var existingClusters = clusterRepository.LoadClusterSleevesForCombo(comboId, targetCategory);
    ...
}
```

---

## Implementation Details

### 1. Add Condition Check to Clustering Service

**File**: `Services/Clustering/RefactoredClusterService.cs`

**Method**: `HandlePath1Replay()` or `ClusterSleeves()` entry point

**Requirements**:
- Check if clearance settings changed (use `CheckConditionsChanged()` method from `RefreshPathDeterminer`)
- Need to pass `currentClearanceSettings` to clustering service
- Need to pass `filterName` to clustering service (currently only gets `targetCategory`)

### 2. Route to Normal Calculation if Conditions Changed

If conditions changed:
- Return `(hasData: false, ...)` from `HandlePath1Replay()`
- This causes clustering to fall through to normal PATH 2/3 calculation
- Old cluster data in database is ignored (will be recalculated and saved fresh)

### 3. Exception: AdoptToDocument

**Note**: User mentioned "if conditions are changed apart from AdoptToDocument"

- `AdoptToDocument` is a different flag that triggers PATH 3
- Condition changes should still route to recalculation (PATH 2/3 method)
- These are independent checks

---

## Summary

### Current State
- ❌ Clustering blindly uses database replay if `isPath1Replay=true`
- ❌ Does not check if clearance values changed
- ❌ Old cluster data may be wrong if conditions changed

### Expected State
- ✅ PATH 1 clustering checks conditions independently
- ✅ If conditions changed → Recalculate (ignore old database data)
- ✅ If conditions unchanged → Use database replay ("dump one-time use many times")
- ✅ Clearance value changes properly trigger recalculation

### Impact
- **Performance**: Normal calculation when conditions unchanged (fast database replay)
- **Correctness**: Recalculation when conditions changed (accurate cluster results)
- **Safety**: No wrong clusters from outdated clearance values


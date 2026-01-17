# 🐛 BUG REPORT: Clustering Not Saving to Database After Enabling UseBatchedParameterWrites

**Date:** January 16, 2026  
**Severity:** CRITICAL 🔴  
**Status:** REPRODUCED  
**Root Cause:** Parameter batching is interfering with cluster sleeve persistence to database

---

## THE PROBLEM

**Symptom:**
```
✅ Before enabling UseBatchedParameterWrites:
   - Cluster sleeves save to ClusterSleeves table ✅
   - MarkedForCluster flag works ✅
   - Everything saves correctly ✅

❌ After enabling UseBatchedParameterWrites = true:
   - Cluster sleeves NOT saving to ClusterSleeves table ❌
   - MarkedForCluster flag not set ❌
   - Data disappears ❌
```

**Affected Components:**
1. `SleeveParameterService.FlushDeferredParameters()` - Parameter flush uses batched dictionary
2. `ClusterSleeveRepository.BatchSaveClusterSleeves()` - Cluster database persistence
3. `RefactoredClusterService` or cluster placement caller - Missing database save call

---

## ROOT CAUSE ANALYSIS

### Issue #1: Deferred Parameters Override Cluster Parameters

When `UseBatchedParameterWrites = true`:

```csharp
// SleeveParameterService.cs - Line 145
if (OptimizationFlags.UseBatchedParameterWrites && !forceImmediate)
{
    var targetDict = ActiveBatchDictionary;  // ← DEFERRED dictionary
    // Parameters are accumulated HERE, not written immediately
    targetDict[currentSleeveId]["Width"] = roundedWidth;
    targetDict[currentSleeveId]["Height"] = roundedHeight;
    targetDict[currentSleeveId]["Depth"] = roundedDepth;
    // ... returns without saving to database
}
```

**Problem:** 
- Individual sleeve parameters are deferred (batched)
- But cluster sleeve persistence to `ClusterSleeves` table is NOT being called
- The batch flush only updates Revit parameters, NOT database records

### Issue #2: Missing Cluster Database Save After Batch Flush

Looking at `ClusterPlacementService.PlaceClusterSleeve()` (line ~145):

```csharp
// ❌ MISSING: After setting parameters, cluster data should be saved to database
placedClusterSleeve = inst;  // ← Instance created
SetSizeParameters(...);      // ← Parameters set (batched)
SetMetadata(...);            // ← Metadata set (batched)
SetScheduleLevelAndElevationForCluster(...); // ← Schedule level set

// ❌ BUT: NO DATABASE SAVE!
// ClusterSleeveRepository.SaveClusterSleeve() never called!
// ClusterSleeveRepository.BatchSaveClusterSleeves() never called!

return true;  // ← Returns success WITHOUT saving to database
```

**The Real Problem:**
- Cluster sleeves are created in Revit ✅
- Parameters are batched (deferred) ✅
- But **cluster metadata is NEVER persisted to `ClusterSleeves` table** ❌

### Issue #3: Database Save Not Happening for Cluster Groups

When cluster sleeves are placed:

1. ❌ Individual parameters are deferred (batched)
2. ❌ Cluster sleeve data is NOT saved to database
3. ❌ When `FlushDeferredParameters()` is called, it only flushes Revit parameters
4. ❌ Database save for cluster sleeves never happens

**Expected Flow:**
```
PlaceClusterSleeve()
  ├─ Create FamilyInstance in Revit ✅
  ├─ Set parameters (batched/deferred) ✅
  ├─ Save cluster metadata to ClusterSleeves table ❌ MISSING!
  └─ Return instance ID

FlushDeferredParameters()
  ├─ Flush all batched parameters to Revit instances ✅
  └─ (Does NOT flush database - wrong place for that)

SaveClusterSleeveDataToDB()  ❌ NEVER CALLED!
  └─ Should save cluster sleeve records to ClusterSleeves table
```

---

## THE FIX

### Solution: Separate Cluster Database Save from Parameter Batching

**Root Cause:** Cluster sleeve database persistence is tied to parameter batching logic

**Fix Location:** `Services/Placement/ClusterPlacementService.cs` - `PlaceClusterSleeve()` method

#### Step 1: Add Cluster Data Collection

After line 145 (after cluster instance is created):

```csharp
// ✅ FIX: Collect cluster data for later database save
var clusterSaveData = new ClusterSaveData
{
    ClusterInstanceId = capturedClusterSleeveId.Value,
    ComboId = groupKey.comboId,  // From SleeveGroupKey
    FilterId = groupKey.filterId,
    Category = targetCategory,
    BoundingBoxMinX = /* calculate from cluster */,
    BoundingBoxMinY = /* calculate from cluster */,
    BoundingBoxMinZ = /* calculate from cluster */,
    BoundingBoxMaxX = /* calculate from cluster */,
    BoundingBoxMaxY = /* calculate from cluster */,
    BoundingBoxMaxZ = /* calculate from cluster */,
    ClusterWidth = width,
    ClusterHeight = height,
    ClusterDepth = depth,
    RotationAngleDeg = rotationAngle * 180 / Math.PI,
    IsRotated = Math.Abs(rotationAngle) > 1e-6,
    PlacementX = placementPoint.X,
    PlacementY = placementPoint.Y,
    PlacementZ = placementPoint.Z,
    HostType = groupKey.hostType,
    HostOrientation = hostOrientation ?? groupKey.orientation,
    ClashZoneIds = cluster
        .Select(c => c is ClashZone cz ? cz.Id : (c?.ClashZone as ClashZone)?.Id)
        .Where(id => id != Guid.Empty)
        .ToList(),
    SleeveFamilyName = sleeveFamilyName ?? familyName,
    Corner1X = /* from cluster */,
    Corner1Y = /* from cluster */,
    // ... all 4 corners
};
```

#### Step 2: Store Cluster Data in Context

Pass cluster data back to caller via out parameter or context:

```csharp
public bool PlaceClusterSleeve(
    Document doc,
    List<dynamic> cluster,
    SleeveGroupKey groupKey,
    string targetCategory,
    XYZ placementPoint,
    double width,
    double height,
    double depth,
    double rotationAngle,
    string? xmlFilePath,
    string? sleeveFamilyName,
    out FamilyInstance? placedClusterSleeve,
    out int? capturedClusterSleeveId,
    out XYZ? actualPlacementPoint,
    out ClusterSaveData? clusterSaveDataOut,  // ✅ NEW: Return cluster data
    Dictionary<ElementId, Dictionary<string, object>>? deferredParameters = null,
    string? hostOrientation = null,
    double mepRotationAngle = 0.0)
{
    // ... existing code ...
    
    // ✅ After creating cluster instance and setting parameters:
    clusterSaveDataOut = clusterSaveData;  // Return cluster data
    return true;
}
```

#### Step 3: Call Database Save After Batch Flush

In the caller (likely `RefactoredClusterService` or orchestrator):

```csharp
// After all cluster sleeves are placed and parameters are batched:

// ✅ FIX: Flush all batched parameters FIRST
sleeveParameterService.FlushDeferredParameters(doc);

// ✅ FIX: THEN save all cluster sleeves to database
if (clusterSleevesDataToSave.Count > 0)
{
    var clusterRepository = new ClusterSleeveRepository(dbContext);
    clusterRepository.BatchSaveClusterSleeves(clusterSleevesDataToSave);
    
    if (!DeploymentConfiguration.DeploymentMode)
    {
        DebugLogger.Info($"✅ Saved {clusterSleevesDataToSave.Count} cluster sleeves to database");
    }
}
```

---

## SPECIFIC CODE CHANGES

### Change #1: ClusterPlacementService.PlaceClusterSleeve() - Add Cluster Data Tracking

**File:** `Services/Clustering/Placement/ClusterPlacementService.cs`

**Location:** After line ~145 (after cluster instance created and parameters set)

**Add:**
```csharp
// ✅ FIX: Collect cluster sleeve data for database save (separate from parameter batching)
// This ensures cluster sleeves are saved to ClusterSleeves table even when parameter batching is enabled
var clusterSaveData = new ClusterSaveData
{
    ClusterInstanceId = capturedClusterSleeveId ?? -1,
    ComboId = groupKey.comboId,
    FilterId = groupKey.filterId,
    Category = targetCategory,
    // Bounding box - calculate from cluster corners or use extremities
    BoundingBoxMinX = /* min X from cluster corners */,
    BoundingBoxMinY = /* min Y from cluster corners */,
    BoundingBoxMinZ = /* min Z from cluster corners */,
    BoundingBoxMaxX = /* max X from cluster corners */,
    BoundingBoxMaxY = /* max Y from cluster corners */,
    BoundingBoxMaxZ = /* max Z from cluster corners */,
    ClusterWidth = width,
    ClusterHeight = height,
    ClusterDepth = depth,
    RotationAngleDeg = rotationAngle * 180.0 / Math.PI,
    IsRotated = Math.Abs(rotationAngle) > 1e-6,
    PlacementX = placementPoint.X,
    PlacementY = placementPoint.Y,
    PlacementZ = placementPoint.Z,
    HostType = groupKey.hostType,
    HostOrientation = hostOrientation ?? groupKey.orientation,
    ClashZoneIds = ExtractClashZoneIds(cluster),
    SleeveFamilyName = sleeveFamilyName ?? familyName,
    // Corner coordinates
    Corner1X = 0.0,
    Corner1Y = 0.0,
    Corner1Z = 0.0,
    Corner2X = 0.0,
    Corner2Y = 0.0,
    Corner2Z = 0.0,
    Corner3X = 0.0,
    Corner3Y = 0.0,
    Corner3Z = 0.0,
    Corner4X = 0.0,
    Corner4Y = 0.0,
    Corner4Z = 0.0
};
```

### Change #2: Add Out Parameter to Return Cluster Data

**File:** `Services/Clustering/Placement/ClusterPlacementService.cs`

**Change Method Signature:**
```csharp
// ❌ BEFORE:
public bool PlaceClusterSleeve(
    Document doc,
    List<dynamic> cluster,
    SleeveGroupKey groupKey,
    string targetCategory,
    XYZ placementPoint,
    double width,
    double height,
    double depth,
    double rotationAngle,
    string? xmlFilePath,
    string? sleeveFamilyName,
    out FamilyInstance? placedClusterSleeve,
    out int? capturedClusterSleeveId,
    out XYZ? actualPlacementPoint,
    Dictionary<ElementId, Dictionary<string, object>>? deferredParameters = null,
    string? hostOrientation = null,
    double mepRotationAngle = 0.0)

// ✅ AFTER:
public bool PlaceClusterSleeve(
    Document doc,
    List<dynamic> cluster,
    SleeveGroupKey groupKey,
    string targetCategory,
    XYZ placementPoint,
    double width,
    double height,
    double depth,
    double rotationAngle,
    string? xmlFilePath,
    string? sleeveFamilyName,
    out FamilyInstance? placedClusterSleeve,
    out int? capturedClusterSleeveId,
    out XYZ? actualPlacementPoint,
    out ClusterSaveData? clusterSaveData,  // ✅ NEW
    Dictionary<ElementId, Dictionary<string, object>>? deferredParameters = null,
    string? hostOrientation = null,
    double mepRotationAngle = 0.0)
```

### Change #3: Caller - Save Cluster Data to Database

**File:** `Services/Clustering/RefactoredClusterService.cs` (or wherever clusters are placed in groups)

**Location:** After `PlaceClusterSleeve()` is called for all sleeves in a group

**Add:**
```csharp
// ✅ FIX: Collect cluster data from all placed sleeves
var clusterSleevesToSave = new List<ClusterSaveData>();

foreach (var clusterData in placedClusterDataList)  // List from PlaceClusterSleeve calls
{
    if (clusterData != null)
    {
        clusterSleevesToSave.Add(clusterData);
    }
}

// ✅ FIX: After all cluster sleeves are placed and BEFORE flushing parameters
// (or immediately after flushing if parameters are batched):

if (clusterSleevesToSave.Count > 0)
{
    // Create repository
    var clusterRepository = new ClusterSleeveRepository(dbContext);
    
    // Batch save to database
    try
    {
        clusterRepository.BatchSaveClusterSleeves(clusterSleevesToSave);
        
        if (!DeploymentConfiguration.DeploymentMode)
        {
            DebugLogger.Info($"[RefactoredClusterService] ✅ Saved {clusterSleevesToSave.Count} cluster sleeves to database");
            SafeFileLogger.SafeAppendText("cluster_debug.log",
                $"[{DateTime.Now:HH:mm:ss.fff}] ✅ Database save complete for {clusterSleevesToSave.Count} cluster sleeves\n");
        }
    }
    catch (Exception ex)
    {
        SafeFileLogger.SafeAppendText("placement_errors.log",
            $"[{DateTime.Now:HH:mm:ss.fff}] ❌ ERROR: Failed to save cluster sleeves to database: {ex.Message}\n");
    }
}
```

---

## VERIFICATION CHECKLIST

After applying fixes, verify:

- [ ] Cluster sleeves are created in Revit ✅
- [ ] Parameters are set correctly (width, height, depth, orientation) ✅
- [ ] **ClusterSleeves table is populated** ✅ (THE KEY FIX)
- [ ] MarkedForCluster flag works (depends on database records existing)
- [ ] No duplicate cluster entries
- [ ] Corner coordinates are saved correctly
- [ ] Batch write flag can be enabled without breaking clustering
- [ ] Test with `UseBatchedParameterWrites = true` and `UseBatchedParameterWrites = false`

---

## TESTING SCENARIO

**Test Case: Cluster Placement with Batch Writing Enabled**

```
Setup:
  1. OptimizationFlags.UseBatchedParameterWrites = true
  2. OptimizationFlags.UseBulkClusterSave = true (also enabled)
  3. Place 6 sleeves that should cluster

Expected Result:
  ✅ 6 sleeves created in Revit
  ✅ 1 cluster sleeve created in Revit (group of 6)
  ✅ 6 individual sleeve records in ClashZones table
  ✅ 1 cluster sleeve record in ClusterSleeves table  ← THE FIX TARGETS THIS
  ✅ MarkedForCluster flag can be read from database

Before Fix:
  ✅ 6 sleeves in Revit
  ✅ 1 cluster in Revit
  ✅ 6 ClashZones records
  ❌ 0 ClusterSleeves records (BUG!)

After Fix:
  ✅ 6 sleeves in Revit
  ✅ 1 cluster in Revit
  ✅ 6 ClashZones records
  ✅ 1 ClusterSleeves records (FIXED!)
```

---

## QUICK SUMMARY

| Aspect | Details |
|--------|---------|
| **Problem** | Cluster sleeves not saving to database when `UseBatchedParameterWrites = true` |
| **Root Cause** | Database save not being called after parameter batching |
| **Solution** | Separate cluster data save from parameter batching - call database save after batch flush |
| **Effort** | ~30 minutes implementation |
| **Risk** | LOW - isolated to cluster placement flow |
| **Testing** | Verify ClusterSleeves table has records after placement |

---

## FILES TO MODIFY

1. ✏️ `Services/Clustering/Placement/ClusterPlacementService.cs`
   - Add cluster data collection
   - Add out parameter for cluster data

2. ✏️ `Services/Clustering/RefactoredClusterService.cs` (or cluster orchestrator)
   - Call `ClusterSleeveRepository.BatchSaveClusterSleeves()` after placement

3. ✏️ `Data/Repositories/ClusterSleeveRepository.cs` (no changes needed - already has BatchSaveClusterSleeves)

---

**Status:** Ready to implement 🚀

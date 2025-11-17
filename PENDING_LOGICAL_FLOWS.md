# Pending Logical Flows - Implementation Checklist

**Document Version:** 1.0  
**Date:** December 2025  
**Based On:** `SLEEVE_PLACEMENT_METHODOLOGY_REFACTORED_FLOWCHART_ORGANIZED.mmd`

---

## Overview

This document lists all pending logical flows that need to be implemented based on the refactored flowchart. These flows are **distinct from error handling** (which is documented separately in `SLEEVE_PLACEMENT_IMPLEMENTATION_PLAN.md`).

---

## 1. PATH 1 Placement - Condition Change Detection

### 1.1 Status: ⚠️ **PENDING**

### 1.2 Flow Description

**Location in Flowchart:** After `P1_Placement` → `P1_CheckConditions`

**Logic:**
- Before placing sleeves in PATH 1 (Replay Mode), check if `OpeningSettings` have changed in the UI
- Compare current UI `OpeningSettings` with saved `OpeningSettings` in database (`Filters` table)
- If conditions changed → Route to PATH 2 placement logic (sizing and recalculation)
- If conditions NOT changed → Continue with PATH 1 replay logic

### 1.3 Implementation Requirements

**Files to Modify:**
- `refresh refactor/refresh_path_strategy.cs` (Path1Strategy)
- `Services/UniversalSleevePlacerService.cs` (placement entry point)
- `Data/Repositories/FilterRepository.cs` (load OpeningSettings)

**Key Methods Needed:**
```csharp
// In Path1Strategy or RefreshPathDeterminer
private bool CheckConditionsChanged(RefreshContext context, string filterName)
{
    // 1. Load current UI OpeningSettings
    var uiSettings = context.ClearanceSettings; // From UI
    
    // 2. Load saved OpeningSettings from DB
    var dbSettings = _filterRepository.LoadOpeningSettings(filterName);
    
    // 3. Compare (structure and values)
    return !AreOpeningSettingsEqual(uiSettings, dbSettings);
}

// Route to PATH 2 placement if changed
if (CheckConditionsChanged(context, filterName))
{
    // Route to PATH 2 sizing logic
    return ExecutePath2Placement(context);
}
```

**Database Query:**
```sql
SELECT OpeningSettings FROM Filters WHERE FilterName = @FilterName
```

**Comparison Logic:**
- Compare clearance settings (RectNormal, RectInsulated, RoundNormal, etc.)
- Compare opening type preferences
- Compare any other settings in `OpeningSettings` JSON

### 1.4 Integration Points

- **Entry Point:** `P1_Placement` node in flowchart
- **Decision Point:** `P1_CheckConditions` diamond
- **Routing:** If changed → `P2_Placement` (PATH 2 sizing)
- **Routing:** If not changed → `P1_LoadPlaceZones` (PATH 1 replay)

---

## 2. PATH 1 Clustering - ClusterSleeves Table Integration

### 2.1 Status: ⚠️ **PENDING**

### 2.2 Flow Description

**Location in Flowchart:** `P1_ClusterDecision` → `P1_CheckClusterDB` → `P1_LoadCluster` or `P1_CalcCluster`

**Logic:**
- After individual sleeve placement in PATH 1, check `ClusterSleeves` table for pre-calculated cluster data
- If cluster data exists → Load from database, skip calculation, place cluster sleeves directly
- If no cluster data exists → **Skip clustering entirely** (do not run calculation)
- Cluster data includes: bounding box, dimensions, rotation angle, placement point, ClashZoneIds

### 2.3 Implementation Requirements

**Files to Modify:**
- `Services/UniversalClusterService.cs` (clustering entry point)
- `Data/Repositories/ClusterSleeveRepository.cs` (already exists, needs integration)

**Key Methods Needed:**
```csharp
// In UniversalClusterService.ClusterSleeves()
public (int placedCount, int deletedCount) ClusterSleeves(
    Document doc, 
    string targetCategory, 
    UIDocument uiDoc = null, 
    string xmlFilePath = null, 
    string filterName = null,
    bool isPath1Replay = false, // NEW PARAMETER
    int? comboId = null, // NEW PARAMETER
    int? filterId = null) // NEW PARAMETER
{
    if (isPath1Replay)
    {
        // PATH 1: Check ClusterSleeves table
        var clusterRepository = new ClusterSleeveRepository();
        var existingClusters = clusterRepository.LoadClusterDataForCombo(comboId.Value, filterId.Value, targetCategory);
        
        if (existingClusters != null && existingClusters.Count > 0)
        {
            // Load and place from database
            return PlaceClustersFromDatabase(doc, existingClusters, uiDoc);
        }
        else
        {
            // No data → Skip clustering
            DebugLogger.Info($"[CLUSTERING] PATH 1: No cluster data found, skipping clustering");
            return (0, 0);
        }
    }
    else
    {
        // PATH 2/3: Always calculate
        return CalculateAndPlaceClusters(doc, targetCategory, uiDoc, xmlFilePath, filterName);
    }
}
```

**Database Query:**
```sql
SELECT 
    ClusterSleeveId,
    ClusterInstanceId,
    BoundingBoxMinX, BoundingBoxMinY, BoundingBoxMinZ,
    BoundingBoxMaxX, BoundingBoxMaxY, BoundingBoxMaxZ,
    ClusterWidth, ClusterHeight, ClusterDepth,
    RotationAngleDeg, IsRotated,
    PlacementX, PlacementY, PlacementZ,
    HostType, HostOrientation,
    ClashZoneIdsJson
FROM ClusterSleeves
WHERE ComboId = @ComboId 
  AND FilterId = @FilterId 
  AND Category = @Category
```

**Placement Logic:**
- Load cluster data from database
- Parse `ClashZoneIdsJson` to get list of ClashZone GUIDs
- Place cluster sleeve at `(PlacementX, PlacementY, PlacementZ)`
- Use `ClusterWidth`, `ClusterHeight`, `ClusterDepth` for dimensions
- Apply rotation if `IsRotated = 1` and `RotationAngleDeg != 0`
- Delete individual sleeves within cluster
- Update flags: `IsClusterResolved = true`, `ClusterSleeveInstanceId`

### 2.4 Integration Points

- **Entry Point:** `P1_ClusterDecision` node
- **Decision Point:** `P1_CheckClusterDB` diamond
- **Load Path:** `P1_LoadCluster` → `P1_PlaceCluster`
- **Skip Path:** `P1_CalcCluster` → Skip (no calculation, just skip)

---

## 3. PATH 2 Clustering - Save to ClusterSleeves Table

### 3.1 Status: ⚠️ **PENDING**

### 3.2 Flow Description

**Location in Flowchart:** `P2_ClusterDecision` → `P2_AlwaysCluster` → `ClusterCalc2` → `SaveCluster2`

**Logic:**
- PATH 2 always runs full clustering calculation (fresh detection)
- After calculation, save cluster data to `ClusterSleeves` table
- This enables PATH 1 replay to use the data later

### 3.3 Implementation Requirements

**Files to Modify:**
- `Services/UniversalClusterService.cs` (after cluster calculation)
- `Data/Repositories/ClusterSleeveRepository.cs` (save method)

**Key Methods Needed:**
```csharp
// In UniversalClusterService, after cluster calculation
private void SaveClusterDataToDatabase(
    List<ClusterData> clusters,
    int comboId,
    int filterId,
    string category)
{
    var repository = new ClusterSleeveRepository();
    
    foreach (var cluster in clusters)
    {
        repository.SaveClusterData(
            comboId: comboId,
            filterId: filterId,
            category: category,
            clusterInstanceId: cluster.ClusterInstanceId,
            boundingBox: cluster.BoundingBox,
            dimensions: cluster.Dimensions,
            rotationAngle: cluster.RotationAngleDeg,
            isRotated: cluster.IsRotated,
            placementPoint: cluster.PlacementPoint,
            hostType: cluster.HostType,
            hostOrientation: cluster.HostOrientation,
            clashZoneIds: cluster.ClashZoneIds);
    }
}
```

**Database Insert:**
```sql
INSERT INTO ClusterSleeves (
    ClusterInstanceId, ComboId, FilterId, Category,
    BoundingBoxMinX, BoundingBoxMinY, BoundingBoxMinZ,
    BoundingBoxMaxX, BoundingBoxMaxY, BoundingBoxMaxZ,
    ClusterWidth, ClusterHeight, ClusterDepth,
    RotationAngleDeg, IsRotated,
    PlacementX, PlacementY, PlacementZ,
    HostType, HostOrientation,
    ClashZoneIdsJson
) VALUES (
    @ClusterInstanceId, @ComboId, @FilterId, @Category,
    @BoundingBoxMinX, @BoundingBoxMinY, @BoundingBoxMinZ,
    @BoundingBoxMaxX, @BoundingBoxMaxY, @BoundingBoxMaxZ,
    @ClusterWidth, @ClusterHeight, @ClusterDepth,
    @RotationAngleDeg, @IsRotated,
    @PlacementX, @PlacementY, @PlacementZ,
    @HostType, @HostOrientation,
    @ClashZoneIdsJson
)
```

**Integration Points:**
- **After:** `ClusterCalc2` (clustering calculation completes)
- **Before:** `SaveCluster2` → `TryDelIndiv2` (delete individual sleeves)

---

## 4. PATH 3 Validated Zones - Always Recalculate Clusters

### 4.1 Status: ⚠️ **PENDING**

### 4.2 Flow Description

**Location in Flowchart:** `P3_Validated` → `P1_Placement` → `P3_ClusterDecision` → `P3_CheckType` → `P3_Val_CheckDB` → `P3_Val_Calc`

**Logic:**
- PATH 3 validated zones route to PATH 1 placement logic (replay)
- After PATH 1 placement, clustering decision is made
- **Even if cluster data exists in ClusterSleeves table, always recalculate**
- Reason: Nearby zone changes (invalidated/new zones) can affect cluster formation for validated zones

### 4.3 Implementation Requirements

**Files to Modify:**
- `Services/UniversalClusterService.cs` (clustering entry point)
- `refresh refactor/refresh_path_strategy.cs` (Path3Strategy)

**Key Methods Needed:**
```csharp
// In UniversalClusterService.ClusterSleeves()
// For PATH 3 validated zones
if (isPath3Validated)
{
    // Always recalculate, ignore existing cluster data
    DebugLogger.Info($"[CLUSTERING] PATH 3 Validated: Always recalculating clusters (nearby changes may affect formation)");
    return CalculateAndPlaceClusters(doc, targetCategory, uiDoc, xmlFilePath, filterName);
}
```

**Integration Points:**
- **Entry Point:** `P3_Validated` → routes to `P1_Placement`
- **After Placement:** `P3_ClusterDecision` → `P3_CheckType` → `P3_Val_CheckDB`
- **Decision:** Even if data exists → `P3_Val_Calc` (always calculate)
- **Save:** After calculation → `SaveCluster3V` → save to `ClusterSleeves` table

---

## 5. PATH 3 Invalidated Zones - Distinct Placement Flow

### 5.1 Status: ⚠️ **PENDING**

### 5.2 Flow Description

**Location in Flowchart:** `P3_Invalidated` → `P3_Commit` → `P3_LoadCond` → `P3_DetectMoved` → `P3_Delete` → `P3_ResetDel` → `P3_CalcSize` → `P3_PlaceNew` → `P3_UpdatePlaced` → `P3_ClusterDecision`

**Logic:**
1. Load conditions from database (`Filters.OpeningSettings`)
2. Run intersection detection for moved elements (MEP or host)
3. Delete affected placed sleeves (individual sleeves at old intersection points)
4. Reset flags for deleted sleeves (`IsResolved = true` to prevent re-placement)
5. Calculate new sleeve size (apply clearance settings)
6. Place new sleeves at new intersection points
7. Recalculate clusters (always calculate, geometry changed)
8. Handle cluster removal/addition based on cluster need

### 5.3 Implementation Requirements

**Files to Modify:**
- `refresh refactor/refresh_path_strategy.cs` (Path3Strategy)
- `Services/UniversalSleevePlacerService.cs` (placement logic)
- `Services/UniversalClusterService.cs` (clustering logic)

**Key Methods Needed:**
```csharp
// In Path3Strategy or new Path3InvalidatedPlacementService
public void ExecutePath3InvalidatedPlacement(
    RefreshContext context,
    List<ClashZone> invalidatedZones)
{
    // 1. Load conditions
    var conditions = _filterRepository.LoadOpeningSettings(context.FilterName);
    
    // 2. Detect moved elements
    var movedZones = DetectMovedElements(context, invalidatedZones);
    
    // 3. Delete affected sleeves
    foreach (var zone in movedZones)
    {
        if (zone.SleeveInstanceId > 0)
        {
            DeleteSleeve(context.Document, zone.SleeveInstanceId);
            // Reset flag to prevent re-placement
            zone.IsResolved = true;
            zone.SleeveInstanceId = 0;
        }
    }
    
    // 4. Calculate new sizes
    foreach (var zone in movedZones)
    {
        var newSize = CalculateSleeveSize(zone, conditions);
        zone.SleeveWidth = newSize.Width;
        zone.SleeveHeight = newSize.Height;
        zone.SleeveDepth = newSize.Depth;
    }
    
    // 5. Place new sleeves
    foreach (var zone in movedZones)
    {
        var sleeve = PlaceSleeve(context.Document, zone);
        zone.SleeveInstanceId = sleeve.Id.IntegerValue;
        zone.IsResolved = true;
    }
    
    // 6. Recalculate clusters (always)
    var clusterService = new UniversalClusterService();
    clusterService.ClusterSleeves(
        context.Document,
        targetCategory: zone.MepElementCategory,
        isPath3Invalidated: true); // Flag to always calculate
}
```

**Cluster Handling Logic:**
```csharp
// After clustering calculation for PATH 3 Invalidated
foreach (var cluster in calculatedClusters)
{
    if (cluster.ClashZoneIds.Count > 1)
    {
        // Cluster needed → Add cluster sleeve
        var clusterSleeve = PlaceClusterSleeve(cluster);
        // Delete individual sleeves within cluster
        DeleteIndividualSleevesInCluster(cluster.ClashZoneIds);
        // Update flags
        UpdateClusterFlags(cluster.ClashZoneIds, clusterSleeve.Id);
    }
    else
    {
        // Cluster not needed → Remove cluster sleeve if exists
        var existingCluster = FindExistingClusterSleeve(cluster.ClashZoneIds[0]);
        if (existingCluster != null)
        {
            DeleteClusterSleeve(existingCluster);
            // Restore individual sleeve if needed
            RestoreIndividualSleeve(cluster.ClashZoneIds[0]);
        }
    }
}
```

**Integration Points:**
- **Entry Point:** `P3_Invalidated` → `P3_Commit` (after saving invalidated zones)
- **Flow:** `P3_LoadCond` → `P3_DetectMoved` → `P3_Delete` → `P3_ResetDel` → `P3_CalcSize` → `P3_PlaceNew` → `P3_UpdatePlaced`
- **Clustering:** `P3_ClusterDecision` → `P3_Inval_Calc` → `CheckNeed3I` → `SaveCluster3I_Add` or `SaveCluster3I_Rem`

---

## 6. PATH 3 New Zones - Routing to PATH 2 Then PATH 3 Clustering

### 6.1 Status: ⚠️ **PENDING**

### 6.2 Flow Description

**Location in Flowchart:** `P3_New` → `P2_Placement` → `CheckNewFromPath3` → `P3_New_Calc`

**Logic:**
- PATH 3 new zones route to PATH 2 placement logic (sizing and placement)
- After PATH 2 placement completes, check if zones came from PATH 3
- If from PATH 3 → Route to PATH 3 new zones clustering (always calculate)
- If from PATH 2 only → Route to PATH 2 clustering

### 6.3 Implementation Requirements

**Files to Modify:**
- `Services/UniversalSleevePlacerService.cs` (placement routing)
- `Services/UniversalClusterService.cs` (clustering routing)

**Key Methods Needed:**
```csharp
// In UniversalSleevePlacerService or RefreshServiceRefactored
public void ExecutePlacement(
    RefreshContext context,
    SleevePlacementPath placementPath)
{
    if (placementPath == SleevePlacementPath.Path3New)
    {
        // Route to PATH 2 placement logic
        ExecutePath2Placement(context);
        
        // After PATH 2 placement, route to PATH 3 new clustering
        ExecutePath3NewClustering(context);
    }
    else if (placementPath == SleevePlacementPath.Path2)
    {
        // Standard PATH 2 placement
        ExecutePath2Placement(context);
        
        // Route to PATH 2 clustering
        ExecutePath2Clustering(context);
    }
}

private void ExecutePath3NewClustering(RefreshContext context)
{
    var clusterService = new UniversalClusterService();
    clusterService.ClusterSleeves(
        context.Document,
        targetCategory: context.SelectedMepCategories[0],
        isPath3New: true); // Flag to always calculate
}
```

**Integration Points:**
- **Entry Point:** `P3_New` → routes to `P2_Placement`
- **After Placement:** `P2_UpdateFlags` → `CheckNewFromPath3` decision
- **Routing:** If from PATH 3 → `P3_New_Calc` (PATH 3 new clustering)
- **Routing:** If from PATH 2 only → `P2_ClusterDecision` (PATH 2 clustering)

---

## 7. IsFilterComboNew Flag Reset - After Cluster Placement

### 7.1 Status: ⚠️ **PENDING**

### 7.2 Flow Description

**Location in Flowchart:** `Update1_Flags`, `Update2_Flags`, `Update3I_Flags`, `Update3N_Flags`, `Update3V_Flags` → `TryResetFlag1-5` → `ResetFlag`

**Logic:**
- `IsFilterComboNew` flag should be reset to `0` **after cluster sleeve placement completes**
- Reset happens via `ClashZoneRepository.ResetFileComboFlag()`
- This ensures PATH 2 (Sizing) gets a chance to run before flag is reset
- Flag reset happens for all paths: PATH 1, PATH 2, PATH 3 (all variants)

### 7.3 Implementation Requirements

**Files to Modify:**
- `Services/UniversalClusterService.cs` (after cluster placement)
- `Data/Repositories/ClashZoneRepository.cs` (reset flag method)

**Key Methods Needed:**
```csharp
// In UniversalClusterService, after cluster placement completes
private void ResetFileComboFlagAfterClusterPlacement(
    Document doc,
    int comboId,
    string filterName)
{
    try
    {
        var repository = new ClashZoneRepository();
        repository.ResetFileComboFlag(comboId);
        
        DebugLogger.Info($"[CLUSTERING] ✅ Reset IsFilterComboNew=0 for ComboId={comboId} after cluster placement");
    }
    catch (Exception ex)
    {
        DebugLogger.Warning($"[CLUSTERING] ❌ Error resetting IsFilterComboNew flag: {ex.Message}");
    }
}
```

**Database Update:**
```sql
UPDATE FileCombos
SET IsFilterComboNew = 0,
    UpdatedAt = datetime('now', '+5 hours', '+30 minutes')
WHERE ComboId = @ComboId
```

**Integration Points:**
- **After:** `Update1_Flags`, `Update2_Flags`, `Update3I_Flags`, `Update3N_Flags`, `Update3V_Flags` (all cluster flag updates)
- **Before:** `ResetFlag` node (final step)
- **Timing:** Only after cluster placement completes (not after individual sleeve placement)

---

## 8. PATH 1 Clustering - Skip if No Data

### 8.1 Status: ⚠️ **PENDING**

### 8.2 Flow Description

**Location in Flowchart:** `P1_CheckClusterDB` → `P1_CalcCluster` → `SkipClustering1`

**Logic:**
- If no cluster data exists in `ClusterSleeves` table for PATH 1
- **Skip clustering entirely** (do not run calculation)
- User should check "Adopt to Modified Document" to trigger PATH 3, which will find and cluster eventually

### 8.3 Implementation Requirements

**Files to Modify:**
- `Services/UniversalClusterService.cs` (PATH 1 clustering logic)

**Key Methods Needed:**
```csharp
// In UniversalClusterService.ClusterSleeves()
if (isPath1Replay)
{
    var clusterRepository = new ClusterSleeveRepository();
    var existingClusters = clusterRepository.LoadClusterDataForCombo(comboId.Value, filterId.Value, targetCategory);
    
    if (existingClusters == null || existingClusters.Count == 0)
    {
        // No data → Skip clustering
        DebugLogger.Info($"[CLUSTERING] PATH 1: No cluster data found, skipping clustering. User should check 'Adopt to Modified Document' to trigger PATH 3.");
        return (0, 0); // Skip
    }
    else
    {
        // Data exists → Load and place
        return PlaceClustersFromDatabase(doc, existingClusters, uiDoc);
    }
}
```

**Integration Points:**
- **Decision Point:** `P1_CheckClusterDB` → if no data
- **Action:** `P1_CalcCluster` → `SkipClustering1` → `UpdateClusterFlags1` (no flags updated, just skip)
- **Note:** This is different from PATH 2/3, which always calculate

---

## 9. PATH 3 Invalidated Zones - Cluster Need Check

### 9.1 Status: ⚠️ **PENDING**

### 9.2 Flow Description

**Location in Flowchart:** `TryCalcCluster3I` → `CheckNeed3I` → `SaveCluster3I_Add` or `SaveCluster3I_Rem`

**Logic:**
- After clustering calculation for PATH 3 Invalidated zones
- Check if cluster is needed (more than 1 sleeve in cluster)
- If cluster needed → Add cluster sleeve, delete individual sleeves
- If cluster not needed → Remove existing cluster sleeve (if exists), restore individual sleeves

### 9.3 Implementation Requirements

**Files to Modify:**
- `Services/UniversalClusterService.cs` (PATH 3 Invalidated clustering)

**Key Methods Needed:**
```csharp
// In UniversalClusterService, after PATH 3 Invalidated clustering calculation
private void HandlePath3InvalidatedClusters(
    Document doc,
    List<ClusterData> calculatedClusters,
    List<ClashZone> invalidatedZones)
{
    foreach (var cluster in calculatedClusters)
    {
        if (cluster.ClashZoneIds.Count > 1)
        {
            // Cluster needed → Add cluster sleeve
            var clusterSleeve = PlaceClusterSleeve(doc, cluster);
            
            // Delete individual sleeves within cluster
            foreach (var zoneId in cluster.ClashZoneIds)
            {
                var zone = invalidatedZones.FirstOrDefault(z => z.Id == zoneId);
                if (zone != null && zone.SleeveInstanceId > 0)
                {
                    DeleteSleeve(doc, zone.SleeveInstanceId);
                    zone.SleeveInstanceId = 0;
                }
            }
            
            // Update flags
            UpdateClusterFlags(cluster.ClashZoneIds, clusterSleeve.Id.IntegerValue);
            
            // Save to ClusterSleeves table
            SaveClusterDataToDatabase(cluster, comboId, filterId, category);
        }
        else
        {
            // Cluster not needed → Remove existing cluster sleeve if exists
            var zoneId = cluster.ClashZoneIds[0];
            var zone = invalidatedZones.FirstOrDefault(z => z.Id == zoneId);
            
            if (zone != null && zone.ClusterSleeveInstanceId > 0)
            {
                // Delete existing cluster sleeve
                DeleteClusterSleeve(doc, zone.ClusterSleeveInstanceId);
                zone.ClusterSleeveInstanceId = 0;
                zone.IsClusterResolved = false;
                
                // Individual sleeve should already be placed (from PATH 3 Invalidated placement)
                // No need to restore
            }
        }
    }
}
```

**Integration Points:**
- **After:** `TryCalcCluster3I` (clustering calculation completes)
- **Decision:** `CheckNeed3I` diamond
- **Actions:** `SaveCluster3I_Add` (add cluster) or `SaveCluster3I_Rem` (remove cluster)

---

## 10. Summary Table

| # | Flow Name | Status | Priority | Files to Modify |
|---|-----------|--------|----------|----------------|
| 1 | PATH 1 Condition Change Detection | ⚠️ PENDING | HIGH | `refresh_path_strategy.cs`, `UniversalSleevePlacerService.cs`, `FilterRepository.cs` |
| 2 | PATH 1 Clustering - ClusterSleeves Table | ⚠️ PENDING | HIGH | `UniversalClusterService.cs`, `ClusterSleeveRepository.cs` |
| 3 | PATH 2 Clustering - Save to DB | ⚠️ PENDING | HIGH | `UniversalClusterService.cs`, `ClusterSleeveRepository.cs` |
| 4 | PATH 3 Validated - Always Recalculate | ⚠️ PENDING | MEDIUM | `UniversalClusterService.cs`, `refresh_path_strategy.cs` |
| 5 | PATH 3 Invalidated - Distinct Placement | ⚠️ PENDING | HIGH | `refresh_path_strategy.cs`, `UniversalSleevePlacerService.cs`, `UniversalClusterService.cs` |
| 6 | PATH 3 New - Routing Logic | ⚠️ PENDING | MEDIUM | `UniversalSleevePlacerService.cs`, `UniversalClusterService.cs` |
| 7 | IsFilterComboNew Flag Reset | ⚠️ PENDING | HIGH | `UniversalClusterService.cs`, `ClashZoneRepository.cs` |
| 8 | PATH 1 Clustering - Skip if No Data | ⚠️ PENDING | MEDIUM | `UniversalClusterService.cs` |
| 9 | PATH 3 Invalidated - Cluster Need Check | ⚠️ PENDING | MEDIUM | `UniversalClusterService.cs` |

---

## 11. Implementation Order Recommendation

### Phase 1: Critical Flows (High Priority)
1. **IsFilterComboNew Flag Reset** (#7) - Required for correct path routing
2. **PATH 2 Clustering - Save to DB** (#3) - Required for PATH 1 replay to work
3. **PATH 1 Clustering - ClusterSleeves Table** (#2) - Enables PATH 1 replay optimization

### Phase 2: Placement Flows (High Priority)
4. **PATH 1 Condition Change Detection** (#1) - User experience improvement
5. **PATH 3 Invalidated - Distinct Placement** (#5) - Core functionality for moved elements

### Phase 3: Clustering Refinements (Medium Priority)
6. **PATH 3 Validated - Always Recalculate** (#4) - Ensures accuracy
7. **PATH 3 Invalidated - Cluster Need Check** (#9) - Handles cluster removal/addition
8. **PATH 3 New - Routing Logic** (#6) - Completes PATH 3 flow
9. **PATH 1 Clustering - Skip if No Data** (#8) - Edge case handling

---

## 12. Testing Checklist

For each implemented flow, verify:

- [ ] Flow matches flowchart logic exactly
- [ ] Database operations are atomic (transactions)
- [ ] Flags are updated correctly
- [ ] Cluster data is saved/loaded correctly
- [ ] PATH routing works as expected
- [ ] Error handling is in place (see `SLEEVE_PLACEMENT_IMPLEMENTATION_PLAN.md`)
- [ ] Logging is comprehensive
- [ ] UI state persistence works (SelectedHostCategories, OpeningSettings)

---

## 13. Notes

- **Error Handling:** All flows should include error handling as documented in `SLEEVE_PLACEMENT_IMPLEMENTATION_PLAN.md`
- **Database Transactions:** All database operations should be wrapped in transactions
- **Logging:** All flows should log to appropriate log files (refresh log, placement log, cluster log)
- **Flag Management:** All flag updates should go through `FlagManager` or `ClashZoneRepository`
- **XML Fallback:** Maintain XML fallback for reading (not writing) during transition period

---

**End of Document**


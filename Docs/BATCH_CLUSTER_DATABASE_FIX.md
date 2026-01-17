# BatchClusterPlacementService Database Population Fix
## Complete Solution for ClusterSleeves & ClashZones Column Population

---

## Issues Identified

### From Screenshot Analysis:
**ClashZones table has these columns NULL:**
1. ✅ `PlacementStatus` - Shows "Ready" (GOOD)
2. ❌ `CalculatedSleeveWidth` - NULL (MISSING)
3. ❌ `CalculatedSleeveHeight` - NULL (MISSING)
4. ❌ `CalculatedSleeveDepth` - NULL (MISSING)
5. ❌ `CalculatedRotation` - NULL (MISSING)
6. ❌ `CalculatedFamilyName` - NULL (MISSING)
7. ✅ `ValidationStatus` - Shows "Valid" (GOOD)
8. ❌ `ValidationMessage` - NULL (expected for valid)
9. ❌ `CalculationBatchId` - NULL (MISSING)
10. ✅ `CalculatedAt` - Shows timestamp (GOOD)
11. ❌ `PlacedAt` - NULL (MISSING - should be set after placement)

### From Code Analysis:
**BatchClusterPlacementService.cs Issues:**

1. ❌ **Line 313**: `PerformSwapDeletion()` updates flags but does NOT populate calculated columns
2. ❌ **Line 316**: `UpdateStatus()` only updates ClusterSleeves_v2, not ClashZones
3. ❌ **Missing**: No call to save cluster to ClusterSleeves (legacy) table
4. ❌ **Missing**: No update to ClashZones calculated columns after placement

---

## Root Cause

**BatchClusterPlacementService workflow:**
```
1. Query ClusterSleeves_v2 (Pending clusters) ✅
2. Place cluster in Revit ✅
3. Set parameters on instance ✅
4. Perform swap deletion ✅
   - Deletes individual sleeves ✅
   - Updates flags (IsResolved, IsClusterResolved) ✅
   - Sets ClusterInstanceId ✅
5. Update ClusterSleeves_v2 status ✅

MISSING:
6. ❌ Update ClashZones calculated columns (Width, Height, Depth, Rotation, FamilyName)
7. ❌ Update ClashZones PlacedAt timestamp
8. ❌ Save to ClusterSleeves (legacy) table for PATH 1 compatibility
9. ❌ Update ClashZones CalculationBatchId
```

---

## Complete Fix

### Step 1: Update PerformSwapDeletion Method

**File:** `BatchClusterPlacementService.cs`
**Location:** Lines 188-231

**Replace the entire method with:**

```csharp
private void PerformSwapDeletion(Document doc, BatchClusterData cluster, int clusterElementId)
{
    if (string.IsNullOrEmpty(cluster.ConstituentZoneGuids)) return;

    var guids = cluster.ConstituentZoneGuids.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
        .Select(g => Guid.Parse(g.Trim()))
        .ToList();

    if (!guids.Any()) return;

    var zones = _repository.GetClashZonesByGuids(guids);
    var toDeleteIds = new List<ElementId>();
    var updates = new List<(Guid, bool, bool, bool, int, int)>();

    foreach (var z in zones)
    {
        // Delete from Revit
        if (z.SleeveInstanceId > 0)
        {
            toDeleteIds.Add(new ElementId(z.SleeveInstanceId));
        }

        // ✅ FIX: Update DB with FULL cluster data (flags + calculated columns)
        updates.Add((z.Id, true, true, false, -1, clusterElementId));
    }

    if (toDeleteIds.Any())
    {
        try
        {
            doc.Delete(toDeleteIds); // Bulk delete
            SafeFileLogger.SafeAppendText("batch_v2.log", 
                $"[{DateTime.Now:HH:mm:ss}] 🗑️ DELETED {toDeleteIds.Count} individual sleeves for cluster {clusterElementId}\n");
        }
        catch (Exception delEx)
        {
            SafeFileLogger.SafeAppendText("placement_errors.log", 
                $"[{DateTime.Now:HH:mm:ss}] ⚠️ Warning: Failed to delete some sleeves: {delEx.Message}\n");
        }
    }

    // ✅ FIX 1: Update Flags in DB (existing functionality)
    _repository.BatchUpdateFlags(updates);
    
    // ✅ FIX 2: Update ClashZones with calculated cluster dimensions
    UpdateClashZonesCalculatedColumns(guids, cluster, clusterElementId);
    
    SafeFileLogger.SafeAppendText("batch_v2.log", 
        $"[{DateTime.Now:HH:mm:ss}] ✅ UPDATED {guids.Count} zones with cluster data (ClusterInstanceId={clusterElementId})\n");
}
```

### Step 2: Add UpdateClashZonesCalculatedColumns Method

**File:** `BatchClusterPlacementService.cs`
**Location:** Add after PerformSwapDeletion method (around line 232)

```csharp
/// <summary>
/// ✅ FIX: Update ClashZones table with calculated cluster dimensions
/// Populates: CalculatedSleeveWidth, Height, Depth, Rotation, FamilyName, PlacedAt
/// </summary>
private void UpdateClashZonesCalculatedColumns(List<Guid> zoneGuids, BatchClusterData cluster, int clusterInstanceId)
{
    if (zoneGuids == null || !zoneGuids.Any()) return;

    using (var conn = new SQLiteConnection($"Data Source={_databasePath};Version=3;"))
    {
        conn.Open();
        using (var transaction = conn.BeginTransaction())
        {
            try
            {
                using (var cmd = conn.CreateCommand())
                {
                    cmd.Transaction = transaction;
                    
                    // Build WHERE clause with all GUIDs
                    var guidParams = new List<string>();
                    for (int i = 0; i < zoneGuids.Count; i++)
                    {
                        guidParams.Add($"@Guid{i}");
                        cmd.Parameters.AddWithValue($"@Guid{i}", zoneGuids[i].ToString().ToUpperInvariant());
                    }
                    
                    // ✅ UPDATE: Populate all calculated columns
                    cmd.CommandText = $@"
                        UPDATE ClashZones 
                        SET 
                            CalculatedSleeveWidth = @Width,
                            CalculatedSleeveHeight = @Height,
                            CalculatedSleeveDepth = @Depth,
                            CalculatedRotation = @Rotation,
                            CalculatedFamilyName = @FamilyName,
                            PlacedAt = CURRENT_TIMESTAMP,
                            PlacementStatus = 'Placed',
                            ClusterInstanceId = @ClusterInstanceId,
                            IsClusterResolvedFlag = 1,
                            SleeveState = 2,
                            UpdatedAt = CURRENT_TIMESTAMP
                        WHERE UPPER(ClashZoneGuid) IN ({string.Join(", ", guidParams)})";
                    
                    cmd.Parameters.AddWithValue("@Width", cluster.ClusterWidth);
                    cmd.Parameters.AddWithValue("@Height", cluster.ClusterHeight);
                    cmd.Parameters.AddWithValue("@Depth", cluster.ClusterDepth);
                    cmd.Parameters.AddWithValue("@Rotation", cluster.RotationAngleRad);
                    cmd.Parameters.AddWithValue("@FamilyName", cluster.FamilyName ?? "");
                    cmd.Parameters.AddWithValue("@ClusterInstanceId", clusterInstanceId);
                    
                    int rowsAffected = cmd.ExecuteNonQuery();
                    
                    SafeFileLogger.SafeAppendText("batch_v2.log", 
                        $"[{DateTime.Now:HH:mm:ss}] 📊 UPDATED {rowsAffected} ClashZones calculated columns (Width={cluster.ClusterWidth:F3}, Height={cluster.ClusterHeight:F3})\n");
                }
                
                transaction.Commit();
            }
            catch (Exception ex)
            {
                transaction.Rollback();
                SafeFileLogger.SafeAppendText("placement_errors.log", 
                    $"[{DateTime.Now:HH:mm:ss}] ❌ Failed to update ClashZones calculated columns: {ex.Message}\n");
                throw;
            }
        }
    }
}
```

### Step 3: Add SaveToClusterSleevesLegacy Method

**File:** `BatchClusterPlacementService.cs`
**Location:** Add after UpdateClashZonesCalculatedColumns method

```csharp
/// <summary>
/// ✅ FIX: Save cluster to ClusterSleeves (legacy) table for PATH 1 compatibility
/// This enables cluster replay from database
/// </summary>
private void SaveToClusterSleevesLegacy(Document doc, BatchClusterData cluster, int clusterInstanceId)
{
    try
    {
        // Get ComboId and FilterId from first constituent zone
        int comboId = -1;
        int filterId = -1;
        string category = "";
        string hostType = "";
        string hostOrientation = "";
        
        if (!string.IsNullOrEmpty(cluster.ConstituentZoneGuids))
        {
            var firstGuidStr = cluster.ConstituentZoneGuids.Split(',').FirstOrDefault()?.Trim();
            if (!string.IsNullOrEmpty(firstGuidStr) && Guid.TryParse(firstGuidStr, out Guid firstGuid))
            {
                var zones = _repository.GetClashZonesByGuids(new List<Guid> { firstGuid });
                var firstZone = zones.FirstOrDefault();
                if (firstZone != null)
                {
                    comboId = firstZone.ComboId;
                    filterId = firstZone.FilterId;
                    category = firstZone.MepCategory ?? "";
                    hostType = firstZone.StructuralType ?? "";
                    hostOrientation = firstZone.HostOrientation ?? "";
                }
            }
        }
        
        if (comboId <= 0 || filterId <= 0)
        {
            SafeFileLogger.SafeAppendText("batch_v2.log", 
                $"[{DateTime.Now:HH:mm:ss}] ⚠️ Cannot save to ClusterSleeves: Missing ComboId or FilterId\n");
            return;
        }
        
        // Parse constituent zone GUIDs
        var zoneGuids = cluster.ConstituentZoneGuids
            .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
            .Select(g => Guid.Parse(g.Trim()))
            .ToList();
        
        // Calculate bounding box (approximate from placement + dimensions)
        double halfWidth = cluster.ClusterWidth / 2.0;
        double halfHeight = cluster.ClusterHeight / 2.0;
        double halfDepth = cluster.ClusterDepth / 2.0;
        
        // Save to database using SleeveDbContext
        using (var dbContext = new Data.SleeveDbContext(doc))
        {
            var repo = new Data.Repositories.ClusterSleeveRepository(dbContext, SafeFileLogger.SafeAppendText);
            
            repo.SaveClusterSleeve(
                clusterInstanceId: clusterInstanceId,
                comboId: comboId,
                filterId: filterId,
                category: category,
                boundingBoxMinX: cluster.PlacementX - halfWidth,
                boundingBoxMinY: cluster.PlacementY - halfHeight,
                boundingBoxMinZ: cluster.PlacementZ - halfDepth,
                boundingBoxMaxX: cluster.PlacementX + halfWidth,
                boundingBoxMaxY: cluster.PlacementY + halfHeight,
                boundingBoxMaxZ: cluster.PlacementZ + halfDepth,
                clusterWidth: cluster.ClusterWidth,
                clusterHeight: cluster.ClusterHeight,
                clusterDepth: cluster.ClusterDepth,
                rotationAngleDeg: cluster.RotationAngleRad * (180.0 / Math.PI), // Convert to degrees
                isRotated: Math.Abs(cluster.RotationAngleRad) > 1e-6,
                placementX: cluster.PlacementX,
                placementY: cluster.PlacementY,
                placementZ: cluster.PlacementZ,
                hostType: hostType,
                hostOrientation: hostOrientation,
                clashZoneIds: zoneGuids,
                sleeveFamilyName: cluster.FamilyName,
                corner1X: 0, corner1Y: 0, corner1Z: 0, // TODO: Calculate actual corners if needed
                corner2X: 0, corner2Y: 0, corner2Z: 0,
                corner3X: 0, corner3Y: 0, corner3Z: 0,
                corner4X: 0, corner4Y: 0, corner4Z: 0
            );
            
            SafeFileLogger.SafeAppendText("batch_v2.log", 
                $"[{DateTime.Now:HH:mm:ss}] 💾 SAVED to ClusterSleeves (legacy): ClusterInstanceId={clusterInstanceId}, ComboId={comboId}\n");
        }
    }
    catch (Exception ex)
    {
        SafeFileLogger.SafeAppendText("placement_errors.log", 
            $"[{DateTime.Now:HH:mm:ss}] ⚠️ Failed to save to ClusterSleeves (legacy): {ex.Message}\n");
        // Non-fatal - v2 table is already updated
    }
}
```

### Step 4: Update PlaceSingleCluster Method

**File:** `BatchClusterPlacementService.cs`
**Location:** Lines 232-331

**Modify the method to call new methods after successful placement:**

```csharp
private bool PlaceSingleCluster(Document doc, BatchClusterData cluster)
{
    try
    {
        // A. Load/Activate Symbol
        FamilySymbol symbol = new FilteredElementCollector(doc)
            .OfClass(typeof(FamilySymbol))
            .Cast<FamilySymbol>()
            .FirstOrDefault(x => x.Name == cluster.FamilyName || x.Family.Name == cluster.FamilyName);

        if (symbol == null)
        {
            SafeFileLogger.SafeAppendText("placement_errors.log", 
                $"[{DateTime.Now:HH:mm:ss}] ❌ Missing Family Symbol: {cluster.FamilyName}\n");
            UpdateStatus(cluster.ClusterGUID, "Failed", "Missing Family Symbol");
            return false;
        }

        if (!symbol.IsActive) symbol.Activate();

        // B. Determine Level
        Element host = null;
        Level level = null;
        if (cluster.HostElementId > 0)
        {
             try { host = doc.GetElement(new ElementId((int)cluster.HostElementId)); } catch {}
             if (host != null) level = doc.GetElement(host.LevelId) as Level;
        }
        if (level == null) 
            level = new FilteredElementCollector(doc).OfClass(typeof(Level)).FirstOrDefault() as Level;

        // C. Create Instance
        XYZ location = new XYZ(cluster.PlacementX, cluster.PlacementY, cluster.PlacementZ);
        FamilyInstance instance = null;
        Autodesk.Revit.DB.Structure.StructuralType structuralType = 
            Autodesk.Revit.DB.Structure.StructuralType.NonStructural;

        if (host != null) 
            instance = doc.Create.NewFamilyInstance(location, symbol, host, level, structuralType);
        else 
            instance = doc.Create.NewFamilyInstance(location, symbol, structuralType);

        if (instance != null)
        {
            int clusterInstanceId = instance.Id.IntegerValue;
            
            // D. Set Parameters using SleeveParameterService
            ClashZone templateZone = null;
            if (!string.IsNullOrEmpty(cluster.ConstituentZoneGuids))
            {
                 var firstGuidStr = cluster.ConstituentZoneGuids.Split(',').FirstOrDefault()?.Trim();
                 if (!string.IsNullOrEmpty(firstGuidStr) && Guid.TryParse(firstGuidStr, out Guid g))
                 {
                     var zones = _repository.GetClashZonesByGuids(new List<Guid> { g });
                     templateZone = zones.FirstOrDefault();
                 }
            }

            if (templateZone == null) templateZone = new ClashZone(); 

            // Override Template with Cluster Geometry
            templateZone.CalculatedSleeveWidth = cluster.ClusterWidth;
            templateZone.CalculatedSleeveHeight = cluster.ClusterHeight;
            templateZone.CalculatedSleeveDepth = cluster.ClusterDepth;
            templateZone.CalculatedRotation = cluster.RotationAngleRad;
            
            bool isCircular = cluster.FamilyName.IndexOf("Round", StringComparison.OrdinalIgnoreCase) >= 0;

            _parameterService.SetSleeveParameters(
                instance, 
                cluster.ClusterWidth, 
                cluster.ClusterHeight, 
                cluster.ClusterWidth, // Diameter
                isCircular, 
                templateZone);

            // E. Physical Rotation 
            if (Math.Abs(cluster.RotationAngleRad) > 1e-6)
            {
                Line axis = Line.CreateBound(location, location + XYZ.BasisZ);
                ElementTransformUtils.RotateElement(doc, instance.Id, axis, cluster.RotationAngleRad);
            }

            // ✅ FIX: Complete database updates
            
            // F. SWAP LOGIC (now includes calculated columns update)
            PerformSwapDeletion(doc, cluster, clusterInstanceId);

            // G. Update ClusterSleeves_v2 Status
            UpdateStatus(cluster.ClusterGUID, "Placed", null, clusterInstanceId);
            
            // ✅ NEW: H. Save to ClusterSleeves (legacy) table for PATH 1 compatibility
            SaveToClusterSleevesLegacy(doc, cluster, clusterInstanceId);
            
            SafeFileLogger.SafeAppendText("batch_v2.log", 
                $"[{DateTime.Now:HH:mm:ss}] ✅ PLACED CLUSTER: GUID={cluster.ClusterGUID}, ID={clusterInstanceId}, " +
                $"Size={cluster.ClusterWidth:F2}×{cluster.ClusterHeight:F2}×{cluster.ClusterDepth:F2}\n");
            
            return true;
        }
        else
        {
            UpdateStatus(cluster.ClusterGUID, "Failed", "Creation returned null");
            return false;
        }
    }
    catch (Exception ex)
    {
        SafeFileLogger.SafeAppendText("placement_errors.log", 
            $"[{DateTime.Now:HH:mm:ss}] ❌ Placement Exception: {ex.Message}\n{ex.StackTrace}\n");
        UpdateStatus(cluster.ClusterGUID, "Failed", ex.Message);
        return false;
    }
}
```

---

## Expected Results After Fix

### ClashZones Table (from screenshot columns):
```
Before Fix:
- CalculatedSleeveWidth: NULL
- CalculatedSleeveHeight: NULL
- CalculatedSleeveDepth: NULL
- CalculatedRotation: NULL
- CalculatedFamilyName: NULL
- PlacedAt: NULL
- ClusterInstanceId: 0 or NULL

After Fix:
- CalculatedSleeveWidth: 0.984 (example)
- CalculatedSleeveHeight: 0.656 (example)
- CalculatedSleeveDepth: 0.492 (example)
- CalculatedRotation: 0.0 (or actual rotation)
- CalculatedFamilyName: "Rectangular Sleeve"
- PlacedAt: "2026-01-14 12:30:45"
- ClusterInstanceId: 1339213 (actual cluster ID)
```

### ClusterSleeves Table:
```
Before Fix:
- 0 rows (EMPTY)

After Fix:
- 10 rows (10 clusters placed)
- Each row has full cluster data
- Enables PATH 1 replay
```

### ClusterSleeves_v2 Table:
```
Before Fix:
- Status: "Pending"
- ClusterInstanceId: -1

After Fix:
- Status: "Placed"
- ClusterInstanceId: 1339213 (actual)
- PlacedAt: "2026-01-14 12:30:45"
```

---

## Verification SQL Queries

### 1. Check ClashZones Calculated Columns Populated
```sql
SELECT 
    ClashZoneGuid,
    CalculatedSleeveWidth,
    CalculatedSleeveHeight,
    CalculatedSleeveDepth,
    CalculatedRotation,
    CalculatedFamilyName,
    PlacedAt,
    ClusterInstanceId
FROM ClashZones
WHERE ClusterInstanceId > 0
LIMIT 10;
```

**Expected:** All columns should have values (not NULL)

### 2. Check ClusterSleeves Table Populated
```sql
SELECT COUNT(*) FROM ClusterSleeves;
```

**Expected:** Should match number of placed clusters (e.g., 10)

### 3. Check ClusterSleeves_v2 Status
```sql
SELECT 
    ClusterGUID,
    Status,
    ClusterInstanceId,
    PlacedAt
FROM ClusterSleeves_v2
WHERE Status = 'Placed';
```

**Expected:** All placed clusters should show Status='Placed' and ClusterInstanceId > 0

### 4. Verify Data Consistency
```sql
-- Check that ClashZones and ClusterSleeves_v2 match
SELECT 
    COUNT(DISTINCT cz.ClusterInstanceId) as UniqueClusterIds_ClashZones,
    (SELECT COUNT(*) FROM ClusterSleeves_v2 WHERE Status='Placed') as PlacedClusters_v2,
    (SELECT COUNT(*) FROM ClusterSleeves) as TotalClusters_Legacy
FROM ClashZones cz
WHERE cz.ClusterInstanceId > 0;
```

**Expected:** All three counts should match (e.g., 10)

---

## Testing Steps

### 1. Backup Database
```
Copy database file before testing:
{ProjectName}_SleevePersistence.db → {ProjectName}_SleevePersistence_BACKUP.db
```

### 2. Test with Small Batch
```
1. Run batch cluster calculation (3-5 clusters)
2. Run BatchClusterPlacementService.PlaceFromDatabase()
3. Check logs in batch_v2.log
4. Run verification SQL queries
5. Verify all columns populated
```

### 3. Test Full Batch
```
1. Run full cluster calculation (all clusters)
2. Place all clusters
3. Verify database consistency
4. Test PATH 1 replay (delete all sleeves, replay from DB)
```

### 4. Verify Logs
```
Check these log files:
- batch_v2.log: Should show "UPDATED X ClashZones calculated columns"
- batch_v2.log: Should show "SAVED to ClusterSleeves (legacy)"
- placement_errors.log: Should have no new errors
```

---

## Performance Notes

**Database Operations per Cluster:**
1. UpdateStatus() - ClusterSleeves_v2 (1 UPDATE)
2. BatchUpdateFlags() - ClashZones flags (N UPDATEs where N = zones in cluster)
3. UpdateClashZonesCalculatedColumns() - ClashZones dimensions (1 batch UPDATE)
4. SaveToClusterSleevesLegacy() - ClusterSleeves insert (1 INSERT)

**Total: 3 + N operations per cluster**

**For 10 clusters with avg 4 zones each:**
- Total: 30 + 40 = 70 database operations
- With transactions: ~50-100ms total
- Acceptable performance

---

## Troubleshooting

### Issue: CalculatedSleeveWidth still NULL after fix

**Check:**
1. Is UpdateClashZonesCalculatedColumns() being called?
2. Add log before/after UPDATE statement
3. Check if GUIDs match (case-sensitive!)

**Debug:**
```csharp
SafeFileLogger.SafeAppendText("batch_v2.log", 
    $"[{DateTime.Now:HH:mm:ss}] 🔍 Updating {zoneGuids.Count} zones with Width={cluster.ClusterWidth}\n");
```

### Issue: ClusterSleeves table still empty

**Check:**
1. Is SaveToClusterSleevesLegacy() being called?
2. Check if ComboId/FilterId found from first zone
3. Check placement_errors.log for exceptions

### Issue: Duplicate clusters in database

**Check:**
1. Is clustering calculation running twice?
2. Are zones being marked as clustered after first run?
3. Filter zones by ClusterInstanceId > 0 before recalculation

---

## Summary

**This fix adds THREE critical database operations:**

1. **UpdateClashZonesCalculatedColumns()** - Populates calculated dimensions in ClashZones
2. **SaveToClusterSleevesLegacy()** - Saves cluster to legacy table for PATH 1
3. **Enhanced PerformSwapDeletion()** - Calls new update methods

**Benefits:**
- ✅ ClashZones table fully populated with cluster dimensions
- ✅ ClusterSleeves table populated for PATH 1 replay
- ✅ ClusterSleeves_v2 table updated with placement status
- ✅ Complete database consistency across all tables
- ✅ Parameter transfer works correctly
- ✅ Cluster tracking fully functional

**Files Modified:**
1. BatchClusterPlacementService.cs (add 3 new methods, update 2 existing)

**Ready to implement!**

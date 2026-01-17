# 🔍 Diagnostic: ClusterSleeves Not Populating Despite Valid ComboId=1

**Given:** ComboId = 1 is valid in both ClashZones and FileCombos  
**Issue:** ClusterSleeves table is still EMPTY

---

## THE NARROWED-DOWN PROBLEM

If ComboId and FilterId are valid, then the issue is ONE of these:

### Option A: SaveToClusterSleevesLegacy() is NEVER CALLED
```
Possibility: Cluster placement fails before reaching this method
Check: Does batch_v2.log show "SaveToClusterSleevesLegacy" message at all?
```

### Option B: SaveClusterSleeve() Method Fails Silently
```
Possibility: The actual database insert fails
Check: Does the insert statement have correct column names?
Check: Do all required columns have values?
```

### Option C: Exception is Caught and Logged as "Non-fatal"
```
Possibility: Error occurs but is swallowed
Check: Look in placement_errors.log for "Failed to save to ClusterSleeves"
```

---

## IMMEDIATE DEBUGGING STEPS

### Step 1: Verify SaveToClusterSleevesLegacy() is Called

**Add this logging to the method:**

```csharp
private void SaveToClusterSleevesLegacy(Document doc, BatchClusterData cluster, int clusterInstanceId)
{
    // ✅ ADD AT VERY START:
    SafeFileLogger.SafeAppendText("batch_v2.log",
        $"\n[{DateTime.Now:HH:mm:ss}] 🔴 SaveToClusterSleevesLegacy CALLED!\n" +
        $"    ClusterInstanceId: {clusterInstanceId}\n" +
        $"    ConstituentZoneGuids: {cluster.ConstituentZoneGuids}\n");

    try
    {
        // ... existing code ...
        
        // ✅ ADD BEFORE THE CRITICAL CHECK:
        SafeFileLogger.SafeAppendText("batch_v2.log",
            $"[{DateTime.Now:HH:mm:ss}]     ComboId: {comboId}, FilterId: {filterId}\n");
        
        if (comboId <= 0 || filterId <= 0)
        {
            SafeFileLogger.SafeAppendText("batch_v2.log",
                $"[{DateTime.Now:HH:mm:ss}]     ❌ EARLY RETURN: Invalid IDs\n");
            return;
        }
        
        // ✅ ADD BEFORE SAVEDB CALL:
        SafeFileLogger.SafeAppendText("batch_v2.log",
            $"[{DateTime.Now:HH:mm:ss}]     ✅ IDs valid, about to call SaveClusterSleeve\n");
        
        repo.SaveClusterSleeve(
            clusterInstanceId: clusterInstanceId,
            // ... parameters ...
        );
        
        // ✅ ADD AFTER SAVE:
        SafeFileLogger.SafeAppendText("batch_v2.log",
            $"[{DateTime.Now:HH:mm:ss}]     ✅ SaveClusterSleeve RETURNED\n");
    }
    catch (Exception ex)
    {
        SafeFileLogger.SafeAppendText("batch_v2.log",
            $"[{DateTime.Now:HH:mm:ss}]     ❌ EXCEPTION: {ex.GetType().Name}: {ex.Message}\n");
        SafeFileLogger.SafeAppendText("placement_errors.log",
            $"[{DateTime.Now:HH:mm:ss}] ❌ SaveToClusterSleevesLegacy FAILED: {ex.Message}\n{ex.StackTrace}\n");
    }
}
```

### Step 2: Check the Method is Called

After rebuilding and testing:
- Open `batch_v2.log`
- Search for "SaveToClusterSleevesLegacy CALLED"
  - **If found:** Method is being called ✓
  - **If NOT found:** Method is never called ✗

---

## IF METHOD IS CALLED BUT CLUSTERSLEEVES IS STILL EMPTY

Then the problem is in `SaveClusterSleeve()` method itself.

**Location:** `Data/Repositories/ClusterSleeveRepository.cs`

### Check the SaveClusterSleeve Method

Find and examine this method:

```csharp
public void SaveClusterSleeve(
    int clusterInstanceId,
    int comboId,
    int filterId,
    string category,
    double boundingBoxMinX,
    // ... all other parameters ...
)
{
    // ❌ CHECK: Is this using the correct table name?
    cmd.CommandText = "INSERT INTO ClusterSleeves (...)";  // Should be ClusterSleeves, not ClusterSleeves_v2
    
    // ❌ CHECK: Are all column names correct?
    // ❌ CHECK: Are parameter names matching @paramName correctly?
    // ❌ CHECK: Is the transaction committed?
}
```

### Add Logging to SaveClusterSleeve

```csharp
public void SaveClusterSleeve(
    int clusterInstanceId,
    int comboId,
    int filterId,
    // ... parameters ...
)
{
    // ✅ ADD AT START:
    SafeFileLogger.SafeAppendText("batch_v2.log",
        $"[{DateTime.Now:HH:mm:ss}]       SaveClusterSleeve CALLED\n" +
        $"           clusterInstanceId={clusterInstanceId}, comboId={comboId}, filterId={filterId}\n");
    
    try
    {
        using (var conn = new SQLiteConnection($"Data Source={_databasePath};Version=3;"))
        {
            conn.Open();
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = @"
                    INSERT INTO ClusterSleeves (
                        ClusterInstanceId, ComboId, FilterId, Category, 
                        BoundingBoxMinX, BoundingBoxMinY, BoundingBoxMinZ,
                        BoundingBoxMaxX, BoundingBoxMaxY, BoundingBoxMaxZ,
                        ClusterWidth, ClusterHeight, ClusterDepth,
                        RotationAngleDeg, IsRotated,
                        PlacementX, PlacementY, PlacementZ,
                        HostType, HostOrientation, FamilyName
                    ) VALUES (
                        @id, @combo, @filter, @cat,
                        @minX, @minY, @minZ, @maxX, @maxY, @maxZ,
                        @w, @h, @d, @rot, @isRot,
                        @px, @py, @pz, @htype, @horient, @fam
                    )";
                
                // Add parameters...
                cmd.Parameters.AddWithValue("@id", clusterInstanceId);
                cmd.Parameters.AddWithValue("@combo", comboId);
                cmd.Parameters.AddWithValue("@filter", filterId);
                // ... rest of parameters ...
                
                // ✅ ADD BEFORE EXECUTE:
                SafeFileLogger.SafeAppendText("batch_v2.log",
                    $"[{DateTime.Now:HH:mm:ss}]           About to execute INSERT\n");
                
                int rows = cmd.ExecuteNonQuery();
                
                // ✅ ADD AFTER EXECUTE:
                SafeFileLogger.SafeAppendText("batch_v2.log",
                    $"[{DateTime.Now:HH:mm:ss}]           ✅ INSERT executed: {rows} row(s) affected\n");
                
                if (rows == 0)
                {
                    SafeFileLogger.SafeAppendText("batch_v2.log",
                        $"[{DateTime.Now:HH:mm:ss}]           ⚠️ WARNING: No rows inserted!\n");
                }
            }
        }
    }
    catch (Exception ex)
    {
        SafeFileLogger.SafeAppendText("batch_v2.log",
            $"[{DateTime.Now:HH:mm:ss}]           ❌ Exception: {ex.Message}\n");
        throw;
    }
}
```

---

## SQL VERIFICATION

Run this to verify the table structure:

```sql
-- Check ClusterSleeves table exists and has correct columns
PRAGMA table_info(ClusterSleeves);

-- Expected columns:
-- ClusterInstanceId
-- ComboId
-- FilterId
-- Category
-- BoundingBoxMinX, Y, Z
-- BoundingBoxMaxX, Y, Z
-- ClusterWidth, Height, Depth
-- RotationAngleDeg
-- IsRotated
-- PlacementX, Y, Z
-- HostType
-- HostOrientation
-- FamilyName
```

If columns are missing → Insert fails!

---

## THREE POSSIBLE DISCOVERIES

### Discovery 1: "SaveToClusterSleevesLegacy CALLED" NOT in logs
**Conclusion:** Method is never called  
**Next Step:** Method is skipped, exception earlier, or execution stops before it  
**Action:** Check batch_v2.log for errors before this point

### Discovery 2: "SaveToClusterSleevesLegacy CALLED" in logs, but "SaveClusterSleeve CALLED" NOT in logs
**Conclusion:** Method is called but fails at ID validation or before save  
**Next Step:** Check the ID validation logic  
**Action:** Share the exact ComboId and FilterId values from logs

### Discovery 3: All logging appears, but "INSERT executed: 0 rows affected"
**Conclusion:** INSERT statement failed silently  
**Next Step:** Check table structure or column names  
**Action:** Run PRAGMA table_info(ClusterSleeves) to verify columns

---

## NEXT STEPS (ORDERED)

1. **Add the logging code above** to both methods
2. **Rebuild the solution**
3. **Place 6 sleeves again**
4. **Check batch_v2.log** for:
   - "SaveToClusterSleevesLegacy CALLED" - Is it there?
   - "SaveClusterSleeve CALLED" - Is it there?
   - "INSERT executed: X rows" - What does it say?
   - Any exceptions?

5. **Share the log output** from batch_v2.log showing the SaveToClusterSleevesLegacy section

6. **Run PRAGMA table_info(ClusterSleeves)** and share the column names

---

## QUICK SANITY CHECK

Run this query to manually test the insert:

```sql
-- Test if we can manually insert a cluster
INSERT INTO ClusterSleeves (
    ClusterInstanceId, ComboId, FilterId, Category,
    BoundingBoxMinX, BoundingBoxMinY, BoundingBoxMinZ,
    BoundingBoxMaxX, BoundingBoxMaxY, BoundingBoxMaxZ,
    ClusterWidth, ClusterHeight, ClusterDepth,
    RotationAngleDeg, IsRotated,
    PlacementX, PlacementY, PlacementZ,
    HostType, HostOrientation, FamilyName
) VALUES (
    999999, 1, 1, 'Pipes',
    0, 0, 0, 10, 10, 10,
    5, 5, 0.5,
    0, 0,
    5, 5, 5,
    'Wall', 'X', 'RectangularOpeningOnWall'
);

-- Check if it inserted
SELECT * FROM ClusterSleeves WHERE ClusterInstanceId = 999999;
```

If manual insert works → Problem is in code logic  
If manual insert fails → Problem is table structure

---

## SUMMARY

**With valid ComboId = 1, the remaining issues are:**

1. Method not being called (execution stops before)
2. SaveClusterSleeve method doesn't exist or has wrong signature
3. Table column mismatch
4. Database connection fails
5. Silent exception being caught

**To find out which:** Add the logging and check what messages appear in logs.

---

**Status:** Awaiting logging output to pinpoint exact failure ⏳

**Next Action:** Add logging code, rebuild, test, and share batch_v2.log output ➡️

# Database Population Debug Guide
## Systematic Diagnosis for BatchClusterPlacementService

---

## Step 1: Verify Code Changes Were Applied

### Check 1: Are the new methods present?

Open `BatchClusterPlacementService.cs` and verify these methods exist:

```csharp
// Should be around line 232
private void UpdateClashZonesCalculatedColumns(List<Guid> zoneGuids, BatchClusterData cluster, int clusterInstanceId)

// Should be around line 280
private void SaveToClusterSleevesLegacy(Document doc, BatchClusterData cluster, int clusterInstanceId)
```

**If NOT found:** The code changes weren't applied. Re-apply the fixes.

**If FOUND:** Continue to Check 2.

---

### Check 2: Are the methods being called?

Look at `PlaceSingleCluster()` method around line 315:

```csharp
// F. SWAP LOGIC (should call new methods)
PerformSwapDeletion(doc, cluster, clusterInstanceId);

// G. Update ClusterSleeves_v2 Status
UpdateStatus(cluster.ClusterGUID, "Placed", null, clusterInstanceId);

// H. NEW - Should be here!
SaveToClusterSleevesLegacy(doc, cluster, clusterInstanceId);
```

**If SaveToClusterSleevesLegacy() call is MISSING:** Add it!

**If FOUND:** Continue to Check 3.

---

### Check 3: Is PerformSwapDeletion calling UpdateClashZonesCalculatedColumns?

Look at `PerformSwapDeletion()` method around line 230:

```csharp
// Should be near the end
_repository.BatchUpdateFlags(updates);

// NEW - Should be here!
UpdateClashZonesCalculatedColumns(guids, cluster, clusterElementId);
```

**If UpdateClashZonesCalculatedColumns() call is MISSING:** Add it!

**If FOUND:** Code changes are applied. Continue to Step 2.

---

## Step 2: Add Diagnostic Logging

### Add logs to verify execution flow

**File:** `BatchClusterPlacementService.cs`

**Location 1: Start of UpdateClashZonesCalculatedColumns (line ~235)**

```csharp
private void UpdateClashZonesCalculatedColumns(List<Guid> zoneGuids, BatchClusterData cluster, int clusterInstanceId)
{
    SafeFileLogger.SafeAppendText("debug_database.log", 
        $"\n[{DateTime.Now:HH:mm:ss}] ═══════════════════════════════════════\n");
    SafeFileLogger.SafeAppendText("debug_database.log", 
        $"[{DateTime.Now:HH:mm:ss}] 🔍 UpdateClashZonesCalculatedColumns CALLED\n");
    SafeFileLogger.SafeAppendText("debug_database.log", 
        $"[{DateTime.Now:HH:mm:ss}]    ZoneGuids Count: {zoneGuids?.Count ?? 0}\n");
    SafeFileLogger.SafeAppendText("debug_database.log", 
        $"[{DateTime.Now:HH:mm:ss}]    ClusterInstanceId: {clusterInstanceId}\n");
    SafeFileLogger.SafeAppendText("debug_database.log", 
        $"[{DateTime.Now:HH:mm:ss}]    Width: {cluster.ClusterWidth:F4}\n");
    SafeFileLogger.SafeAppendText("debug_database.log", 
        $"[{DateTime.Now:HH:mm:ss}]    Height: {cluster.ClusterHeight:F4}\n");
    SafeFileLogger.SafeAppendText("debug_database.log", 
        $"[{DateTime.Now:HH:mm:ss}]    Depth: {cluster.ClusterDepth:F4}\n");
    
    if (zoneGuids == null || !zoneGuids.Any())
    {
        SafeFileLogger.SafeAppendText("debug_database.log", 
            $"[{DateTime.Now:HH:mm:ss}] ❌ EARLY RETURN: No zone GUIDs\n");
        return;
    }

    // ... rest of method
```

**Location 2: After database update (inside UpdateClashZonesCalculatedColumns, after ExecuteNonQuery)**

```csharp
int rowsAffected = cmd.ExecuteNonQuery();

SafeFileLogger.SafeAppendText("debug_database.log", 
    $"[{DateTime.Now:HH:mm:ss}] ✅ UPDATE EXECUTED: {rowsAffected} rows affected\n");
SafeFileLogger.SafeAppendText("debug_database.log", 
    $"[{DateTime.Now:HH:mm:ss}]    SQL: UPDATE ClashZones SET CalculatedSleeveWidth={cluster.ClusterWidth:F4} ...\n");

// Log the GUIDs being updated
SafeFileLogger.SafeAppendText("debug_database.log", 
    $"[{DateTime.Now:HH:mm:ss}]    Updated GUIDs:\n");
foreach (var guid in zoneGuids.Take(5)) // First 5 GUIDs
{
    SafeFileLogger.SafeAppendText("debug_database.log", 
        $"[{DateTime.Now:HH:mm:ss}]       - {guid}\n");
}
```

**Location 3: Exception handling (inside UpdateClashZonesCalculatedColumns catch block)**

```csharp
catch (Exception ex)
{
    transaction.Rollback();
    SafeFileLogger.SafeAppendText("debug_database.log", 
        $"[{DateTime.Now:HH:mm:ss}] ❌❌❌ EXCEPTION IN UpdateClashZonesCalculatedColumns ❌❌❌\n");
    SafeFileLogger.SafeAppendText("debug_database.log", 
        $"[{DateTime.Now:HH:mm:ss}]    Message: {ex.Message}\n");
    SafeFileLogger.SafeAppendText("debug_database.log", 
        $"[{DateTime.Now:HH:mm:ss}]    StackTrace:\n{ex.StackTrace}\n");
    throw;
}
```

**Location 4: Start of SaveToClusterSleevesLegacy**

```csharp
private void SaveToClusterSleevesLegacy(Document doc, BatchClusterData cluster, int clusterInstanceId)
{
    SafeFileLogger.SafeAppendText("debug_database.log", 
        $"\n[{DateTime.Now:HH:mm:ss}] ═══════════════════════════════════════\n");
    SafeFileLogger.SafeAppendText("debug_database.log", 
        $"[{DateTime.Now:HH:mm:ss}] 💾 SaveToClusterSleevesLegacy CALLED\n");
    SafeFileLogger.SafeAppendText("debug_database.log", 
        $"[{DateTime.Now:HH:mm:ss}]    ClusterInstanceId: {clusterInstanceId}\n");
    SafeFileLogger.SafeAppendText("debug_database.log", 
        $"[{DateTime.Now:HH:mm:ss}]    FamilyName: {cluster.FamilyName}\n");
    
    try
    {
        // ... existing code to get comboId/filterId
        
        SafeFileLogger.SafeAppendText("debug_database.log", 
            $"[{DateTime.Now:HH:mm:ss}]    ComboId: {comboId}\n");
        SafeFileLogger.SafeAppendText("debug_database.log", 
            $"[{DateTime.Now:HH:mm:ss}]    FilterId: {filterId}\n");
        
        if (comboId <= 0 || filterId <= 0)
        {
            SafeFileLogger.SafeAppendText("debug_database.log", 
                $"[{DateTime.Now:HH:mm:ss}] ❌ EARLY RETURN: Missing ComboId or FilterId\n");
            return;
        }
        
        // ... rest of method
```

**Location 5: After successful save to ClusterSleeves**

```csharp
repo.SaveClusterSleeve(/* ... parameters ... */);

SafeFileLogger.SafeAppendText("debug_database.log", 
    $"[{DateTime.Now:HH:mm:ss}] ✅ SaveClusterSleeve COMPLETED\n");
SafeFileLogger.SafeAppendText("debug_database.log", 
    $"[{DateTime.Now:HH:mm:ss}]    Saved to ClusterSleeves table\n");
```

---

## Step 3: Run Test and Check Logs

### 3.1 Run Cluster Placement

1. Delete all existing sleeves from Revit
2. Run batch cluster placement
3. Wait for completion

### 3.2 Check Debug Log

**File location:**
```
%APPDATA%\JSE_MEP_Openings\Logs\R2023\debug_database.log
```

**What to look for:**

#### Scenario A: Methods NOT called at all
```
[12:30:45] 🏗️ STARTING PLACEMENT V2: Batch ABC, Pending=10
[12:30:46] ✅ PLACED CLUSTER: GUID=xxx, ID=1339213
[12:30:46] ✅ BATCH COMPLETED: Placed=10, Failed=0

MISSING:
- "UpdateClashZonesCalculatedColumns CALLED"
- "SaveToClusterSleevesLegacy CALLED"
```

**Diagnosis:** Methods exist but are NOT being called
**Fix:** The method calls weren't added to PlaceSingleCluster or PerformSwapDeletion

---

#### Scenario B: Methods called but early return
```
[12:30:46] 🔍 UpdateClashZonesCalculatedColumns CALLED
[12:30:46]    ZoneGuids Count: 0
[12:30:46] ❌ EARLY RETURN: No zone GUIDs
```

**Diagnosis:** Zone GUIDs not being passed correctly
**Fix:** Check how `cluster.ConstituentZoneGuids` is populated

---

#### Scenario C: Methods called but exception thrown
```
[12:30:46] 🔍 UpdateClashZonesCalculatedColumns CALLED
[12:30:46]    ZoneGuids Count: 4
[12:30:46] ❌❌❌ EXCEPTION IN UpdateClashZonesCalculatedColumns
[12:30:46]    Message: no such column: CalculatedSleeveWidth
```

**Diagnosis:** Database schema missing columns
**Fix:** Run schema migration (see Step 4)

---

#### Scenario D: Methods called, no errors, but 0 rows affected
```
[12:30:46] 🔍 UpdateClashZonesCalculatedColumns CALLED
[12:30:46]    ZoneGuids Count: 4
[12:30:46]    Width: 0.9843
[12:30:46]    Height: 0.6562
[12:30:46] ✅ UPDATE EXECUTED: 0 rows affected
```

**Diagnosis:** GUIDs don't match database records (case mismatch or wrong GUIDs)
**Fix:** Check GUID formatting (see Step 5)

---

#### Scenario E: SaveToClusterSleeves early return
```
[12:30:46] 💾 SaveToClusterSleevesLegacy CALLED
[12:30:46]    ClusterInstanceId: 1339213
[12:30:46]    ComboId: -1
[12:30:46]    FilterId: -1
[12:30:46] ❌ EARLY RETURN: Missing ComboId or FilterId
```

**Diagnosis:** Can't get ComboId/FilterId from constituent zones
**Fix:** Check if constituent zones have ComboId/FilterId populated

---

## Step 4: Database Schema Check

### Verify columns exist in ClashZones table

**Run this SQL query:**

```sql
PRAGMA table_info(ClashZones);
```

**Look for these columns:**
- CalculatedSleeveWidth
- CalculatedSleeveHeight
- CalculatedSleeveDepth
- CalculatedRotation
- CalculatedFamilyName
- PlacedAt
- ClusterInstanceId

**If ANY are missing:**

The schema migration didn't run. Run this SQL:

```sql
-- Add missing columns
ALTER TABLE ClashZones ADD COLUMN CalculatedSleeveWidth REAL;
ALTER TABLE ClashZones ADD COLUMN CalculatedSleeveHeight REAL;
ALTER TABLE ClashZones ADD COLUMN CalculatedSleeveDepth REAL;
ALTER TABLE ClashZones ADD COLUMN CalculatedRotation REAL;
ALTER TABLE ClashZones ADD COLUMN CalculatedFamilyName TEXT;
ALTER TABLE ClashZones ADD COLUMN PlacedAt TEXT;
-- ClusterInstanceId should already exist

SELECT 'Columns added successfully';
```

### Verify ClusterSleeves table exists

```sql
SELECT name FROM sqlite_master WHERE type='table' AND name='ClusterSleeves';
```

**If returns no rows:**

The table doesn't exist! Run schema creation:

```sql
CREATE TABLE IF NOT EXISTS ClusterSleeves (
    ClusterSleeveId      INTEGER PRIMARY KEY AUTOINCREMENT,
    ClusterInstanceId   INTEGER NOT NULL UNIQUE,
    ComboId             INTEGER NOT NULL,
    FilterId            INTEGER NOT NULL,
    Category            TEXT NOT NULL,
    BoundingBoxMinX     REAL NOT NULL,
    BoundingBoxMinY     REAL NOT NULL,
    BoundingBoxMinZ     REAL NOT NULL,
    BoundingBoxMaxX     REAL NOT NULL,
    BoundingBoxMaxY     REAL NOT NULL,
    BoundingBoxMaxZ     REAL NOT NULL,
    ClusterWidth        REAL NOT NULL,
    ClusterHeight       REAL NOT NULL,
    ClusterDepth        REAL NOT NULL,
    RotationAngleDeg    REAL DEFAULT 0.0,
    IsRotated           INTEGER NOT NULL DEFAULT 0,
    PlacementX          REAL NOT NULL,
    PlacementY          REAL NOT NULL,
    PlacementZ          REAL NOT NULL,
    HostType            TEXT,
    HostOrientation     TEXT,
    ClashZoneIdsJson    TEXT NOT NULL,
    ClashZoneGuids      TEXT,
    MepSizes            TEXT,
    MepSystemNames      TEXT,
    MepElementIds       TEXT,
    Corner1X            REAL DEFAULT 0.0,
    Corner1Y            REAL DEFAULT 0.0,
    Corner1Z            REAL DEFAULT 0.0,
    Corner2X            REAL DEFAULT 0.0,
    Corner2Y            REAL DEFAULT 0.0,
    Corner2Z            REAL DEFAULT 0.0,
    Corner3X            REAL DEFAULT 0.0,
    Corner3Y            REAL DEFAULT 0.0,
    Corner3Z            REAL DEFAULT 0.0,
    Corner4X            REAL DEFAULT 0.0,
    Corner4Y            REAL DEFAULT 0.0,
    Corner4Z            REAL DEFAULT 0.0,
    CreatedAt           DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
    UpdatedAt           DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP
);

SELECT 'ClusterSleeves table created';
```

---

## Step 5: GUID Matching Check

### Verify GUIDs in database match what code is looking for

**Run this query:**

```sql
-- Check GUIDs in ClashZones
SELECT 
    ClashZoneGuid,
    LENGTH(ClashZoneGuid) as GuidLength,
    CASE 
        WHEN ClashZoneGuid = UPPER(ClashZoneGuid) THEN 'UPPERCASE'
        WHEN ClashZoneGuid = LOWER(ClashZoneGuid) THEN 'lowercase'
        ELSE 'MixedCase'
    END as GuidCase
FROM ClashZones
WHERE ClashZoneGuid IS NOT NULL AND ClashZoneGuid != ''
LIMIT 5;
```

**Expected:**
- GuidLength: 36 (format: XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX)
- GuidCase: UPPERCASE (most likely)

**If GuidCase is lowercase or MixedCase:**

The UPDATE statement needs to match the case! Change this line in UpdateClashZonesCalculatedColumns:

```csharp
// BEFORE
cmd.Parameters.AddWithValue($"@Guid{i}", zoneGuids[i].ToString().ToUpperInvariant());

// AFTER (match database case)
cmd.Parameters.AddWithValue($"@Guid{i}", zoneGuids[i].ToString()); // No case conversion
```

---

## Step 6: Manual Test of UPDATE Statement

### Test if UPDATE works manually

**Run this SQL directly:**

```sql
-- Pick a real GUID from your database
SELECT ClashZoneGuid FROM ClashZones LIMIT 1;

-- Copy the GUID and test UPDATE
UPDATE ClashZones
SET 
    CalculatedSleeveWidth = 0.9843,
    CalculatedSleeveHeight = 0.6562,
    CalculatedSleeveDepth = 0.4921,
    CalculatedRotation = 0.0,
    CalculatedFamilyName = 'TEST',
    PlacedAt = datetime('now'),
    ClusterInstanceId = 9999999
WHERE ClashZoneGuid = 'YOUR-GUID-HERE';

-- Check if it worked
SELECT 
    ClashZoneGuid,
    CalculatedSleeveWidth,
    CalculatedSleeveHeight,
    ClusterInstanceId
FROM ClashZones
WHERE ClashZoneGuid = 'YOUR-GUID-HERE';
```

**If UPDATE works manually but not from code:**

The SQL syntax is correct, but the GUIDs passed from code don't match. Check GUID formatting.

**If UPDATE doesn't work manually:**

There's a database issue (table locked, permissions, etc.)

---

## Step 7: Check Database File Location

### Verify code is writing to correct database

**Add this log at start of UpdateClashZonesCalculatedColumns:**

```csharp
SafeFileLogger.SafeAppendText("debug_database.log", 
    $"[{DateTime.Now:HH:mm:ss}]    Database Path: {_databasePath}\n");
```

**Then check:**
1. Is this the same database file you're querying in DB Browser?
2. Open this EXACT file in DB Browser
3. Check if updates are there

**Common issue:**
- Code writes to: `C:\Users\...\Filters\Project_SleevePersistence.db`
- You're checking: `C:\Users\...\Desktop\Project_SleevePersistence.db` (wrong file!)

---

## Step 8: Transaction Commit Check

### Verify transaction is committing

**Add this log after transaction.Commit():**

```csharp
transaction.Commit();

SafeFileLogger.SafeAppendText("debug_database.log", 
    $"[{DateTime.Now:HH:mm:ss}] ✅ TRANSACTION COMMITTED\n");

// CRITICAL: Close connection to ensure writes are flushed
conn.Close();

SafeFileLogger.SafeAppendText("debug_database.log", 
    $"[{DateTime.Now:HH:mm:ss}] ✅ CONNECTION CLOSED\n");
```

**If transaction commits but changes don't appear:**

SQLite WAL (Write-Ahead Logging) might be caching writes. Add this:

```csharp
conn.Close();

// Force WAL checkpoint
using (var conn2 = new SQLiteConnection($"Data Source={_databasePath};Version=3;"))
{
    conn2.Open();
    using (var cmd = conn2.CreateCommand())
    {
        cmd.CommandText = "PRAGMA wal_checkpoint(FULL);";
        cmd.ExecuteNonQuery();
    }
}

SafeFileLogger.SafeAppendText("debug_database.log", 
    $"[{DateTime.Now:HH:mm:ss}] ✅ WAL CHECKPOINT FORCED\n");
```

---

## Complete Diagnostic Checklist

Run through this checklist in order:

- [ ] Step 1: Verify code changes applied
- [ ] Step 2: Add all diagnostic logging
- [ ] Step 3: Run test and capture logs
- [ ] Step 4: Check database schema
- [ ] Step 5: Verify GUID matching
- [ ] Step 6: Test manual UPDATE
- [ ] Step 7: Verify database file location
- [ ] Step 8: Check transaction commit

---

## Quick Fix Summary

**Most common issues:**

### Issue #1: Methods not being called
**Fix:** Add method calls to PlaceSingleCluster and PerformSwapDeletion

### Issue #2: Missing database columns
**Fix:** Run ALTER TABLE commands to add columns

### Issue #3: GUID case mismatch
**Fix:** Match GUID case in code to database

### Issue #4: Wrong database file
**Fix:** Verify _databasePath matches the file you're querying

### Issue #5: Transaction not committing
**Fix:** Add connection.Close() and WAL checkpoint

---

## Next Steps

1. Add ALL the diagnostic logging from Step 2
2. Run cluster placement
3. Check `debug_database.log` file
4. Share the log file contents with me
5. I'll tell you exactly what's wrong!

**The log file will show us exactly where the process is failing.**

# Combined Sleeve Database Fix

## ❌ Problem
```
❌ Batch save failed: SQL logic error
no such column: IsCombinedResolved
```

The `IsCombinedResolved` column was missing from the `ClusterSleeves` table.

## ✅ Solution Implemented

### 1. Added Missing Column Migration
**File:** `Data/SleeveDbContext.cs` (lines 1135-1140)

```csharp
// ✅ COMBINED RESOLVED: Flag for Phase 4 combined sleeves (ClusterSleeves table)
if (AddColumnIfMissing("ClusterSleeves", "IsCombinedResolved", "INTEGER NOT NULL DEFAULT 0", transaction))
    _logger("[SQLite] ✅ Added IsCombinedResolved column to ClusterSleeves (for combined sleeves)");

// ✅ COMBINED RESOLVED: Instance ID for the combined sleeve if this cluster is part of one
AddColumnIfMissing("ClusterSleeves", "CombinedClusterSleeveInstanceId", "INTEGER", transaction);
```

### 2. What These Columns Do

**`IsCombinedResolved`** (INTEGER, default 0)
- `0` = This cluster sleeve is NOT part of a combined sleeve
- `1` = This cluster sleeve HAS BEEN combined into a Phase 4 combined sleeve

**`CombinedClusterSleeveInstanceId`** (INTEGER, nullable)
- Stores the Revit ElementId of the combined sleeve that this cluster is part of
- Used to track which combined sleeve "owns" this cluster

### 3. Where They're Used

**In `CombinedSleeveRepository.cs`:**

```csharp
// Line 90: Mark cluster as resolved when creating combined sleeve
cmd.CommandText = @"UPDATE ClusterSleeves 
                    SET IsCombinedResolved = 1, 
                        CombinedClusterSleeveInstanceId = @CId 
                    WHERE ClusterInstanceId = @Id";
```

```csharp
// Line 506: Mark cluster as resolved when flagging constituents
cmd.CommandText = @"UPDATE ClusterSleeves 
                    SET IsCombinedResolved = 1,
                        CombinedClusterSleeveInstanceId = @CombinedInstanceId
                    WHERE ClusterInstanceId = @ClusterInstanceId";
```

## 🔄 How to Apply the Fix

### Option 1: Restart Revit (Recommended)
1. Close Revit completely
2. Reopen Revit
3. The database schema migration will run automatically
4. The `IsCombinedResolved` column will be added to `ClusterSleeves`

### Option 2: Delete Database (If Option 1 Doesn't Work)
1. Close Revit
2. Delete the database file: `[Project]_SleeveData.db`
3. Reopen Revit
4. The database will be recreated with the correct schema

## ✅ Verification

After restarting Revit, check the log for:
```
[SQLite] ✅ Added IsCombinedResolved column to ClusterSleeves (for combined sleeves)
```

Then try placing combined sleeves again. You should see:
```
[SQLite] ✅ Batch saved 2 combined sleeves and updated flags.
```

Instead of:
```
❌ Batch save failed: SQL logic error
no such column: IsCombinedResolved
```

## 📊 Database Schema

### ClusterSleeves Table (After Fix)
```sql
CREATE TABLE ClusterSleeves (
    ClusterSleeveId      INTEGER PRIMARY KEY AUTOINCREMENT,
    ClusterInstanceId    INTEGER NOT NULL UNIQUE,
    ...
    IsCombinedResolved   INTEGER NOT NULL DEFAULT 0,  -- ✅ NEW
    CombinedClusterSleeveInstanceId INTEGER,          -- ✅ NEW
    ...
);
```

### ClashZones Table (Already Has Column)
```sql
CREATE TABLE ClashZones (
    ClashZoneId   INTEGER PRIMARY KEY AUTOINCREMENT,
    ...
    IsCombinedResolved   INTEGER NOT NULL DEFAULT 0,  -- ✅ Already exists
    CombinedClusterSleeveInstanceId INTEGER,          -- ✅ Already exists
    ...
);
```

## 🎯 Expected Behavior After Fix

### Phase 4: Combined Sleeve Placement
1. **Create Combined Sleeve** → Saves to `CombinedSleeves` table
2. **Mark Constituents** → Updates `ClashZones` and `ClusterSleeves`:
   - Sets `IsCombinedResolved = 1`
   - Sets `CombinedClusterSleeveInstanceId = [combined sleeve ID]`
3. **Delete Constituents** → Deletes individual and cluster sleeves from Revit
4. **Result** → One combined sleeve in Revit, constituents marked in database

### Database State After Combined Sleeve Placement

**ClashZones Table:**
```
ClashZoneGuid                          | IsCombinedResolved | CombinedClusterSleeveInstanceId
8abcb5eb-fa43-1dd6-0735-d8cf977ff6cb  | 1                  | 1158140
```

**ClusterSleeves Table:**
```
ClusterInstanceId | IsCombinedResolved | CombinedClusterSleeveInstanceId
1158126          | 1                  | 1158140
1158117          | 1                  | 1158149
```

**CombinedSleeves Table:**
```
CombinedInstanceId | Categories              | CombinedWidth | CombinedHeight
1158140           | Duct Accessories,Pipes  | 2.13         | 1.96
1158149           | Duct Accessories,Pipes  | 3.16         | 0.98
```

## ✅ Summary

- **Problem**: Missing `IsCombinedResolved` column in `ClusterSleeves` table
- **Cause**: Database was deleted/recreated, migration didn't run
- **Fix**: Added column migration to `EnsureClusterSleevesTable()`
- **Action Required**: Restart Revit to apply migration
- **Expected Result**: Combined sleeves save successfully with flags populated

# Combined Sleeve Database Fix - FINAL

## ✅ **CORRECT Solution**

### Problem Analysis
The original error was:
```
❌ Batch save failed: SQL logic error
no such column: IsCombinedResolved
```

### Root Cause
The code was trying to UPDATE **two columns** in `ClusterSleeves` table:
1. `IsCombinedResolved` ❌ **Does NOT exist in ClusterSleeves**
2. `CombinedClusterSleeveInstanceId` ✅ **SHOULD exist in ClusterSleeves**

### Correct Schema

**`ClashZones` Table:**
```sql
IsCombinedResolved              INTEGER NOT NULL DEFAULT 0  ✅
CombinedClusterSleeveInstanceId INTEGER                     ✅
```

**`ClusterSleeves` Table:**
```sql
IsCombinedResolved              ❌ NOT NEEDED (only in ClashZones)
CombinedClusterSleeveInstanceId INTEGER DEFAULT -1          ✅ NEEDED for parameter transfer
```

### Why `CombinedClusterSleeveInstanceId` is Needed in Both Tables

**For Individual Sleeves (ClashZones):**
```
ClashZone → IsCombinedResolved = 1
         → CombinedClusterSleeveInstanceId = 1158140
```
This tracks that this individual sleeve is part of combined sleeve 1158140.

**For Cluster Sleeves (ClusterSleeves):**
```
ClusterSleeve → CombinedClusterSleeveInstanceId = 1158140
```
This tracks that this cluster sleeve is part of combined sleeve 1158140.

**Why this is critical for parameter transfer:**
When transferring parameters to a combined sleeve, the system needs to:
1. Look up which ClashZones are part of the combined sleeve
2. Look up which ClusterSleeves are part of the combined sleeve
3. Aggregate all MEP parameters from both sources
4. Transfer to the combined sleeve

Without `CombinedClusterSleeveInstanceId` in `ClusterSleeves`, parameter transfer cannot find the cluster's MEP data!

### Changes Made

#### 1. Schema Migration (`SleeveDbContext.cs`)
```csharp
// ✅ COMBINED SLEEVE TRACKING: Instance ID for the combined sleeve
// NOTE: IsCombinedResolved is NOT needed here - only in ClashZones
// But CombinedClusterSleeveInstanceId IS needed for parameter transfer lookup
AddColumnIfMissing("ClusterSleeves", "CombinedClusterSleeveInstanceId", "INTEGER DEFAULT -1", transaction);
```

#### 2. Repository Updates (`CombinedSleeveRepository.cs`)

**For Individual Sleeves:**
```csharp
// Line 79: Update ClashZones with BOTH columns
cmd.CommandText = @"UPDATE ClashZones 
                    SET IsCombinedResolved = 1, 
                        CombinedClusterSleeveInstanceId = @CId 
                    WHERE ClashZoneGuid = @Guid";
```

**For Cluster Sleeves:**
```csharp
// Line 92: Update ClusterSleeves with ONLY CombinedClusterSleeveInstanceId
cmd.CommandText = @"UPDATE ClusterSleeves 
                    SET CombinedClusterSleeveInstanceId = @CId 
                    WHERE ClusterInstanceId = @Id";
```

### How to Apply

**Option 1: Restart Revit (Recommended)**
1. Close Revit
2. Reopen Revit
3. Migration runs automatically
4. Column is added to `ClusterSleeves`

**Option 2: Manual SQL (Immediate)**
```sql
ALTER TABLE ClusterSleeves ADD COLUMN CombinedClusterSleeveInstanceId INTEGER DEFAULT -1;
```

### Expected Database State After Fix

**After placing combined sleeve 1158140:**

**ClashZones Table:**
```
ClashZoneGuid                          | IsCombinedResolved | CombinedClusterSleeveInstanceId
8abcb5eb-fa43-1dd6-0735-d8cf977ff6cb  | 1                  | 1158140
```

**ClusterSleeves Table:**
```
ClusterInstanceId | CombinedClusterSleeveInstanceId
1158126          | 1158140
1158117          | 1158149
```

**CombinedSleeves Table:**
```
CombinedInstanceId | Categories              | CombinedWidth | CombinedHeight
1158140           | Duct Accessories,Pipes  | 2.13         | 1.96
1158149           | Duct Accessories,Pipes  | 3.16         | 0.98
```

### Parameter Transfer Flow

1. **User clicks "Transfer Parameters" on combined sleeve 1158140**
2. **System looks up constituents:**
   - Query `ClashZones` WHERE `CombinedClusterSleeveInstanceId = 1158140`
   - Query `ClusterSleeves` WHERE `CombinedClusterSleeveInstanceId = 1158140`
3. **System aggregates MEP parameters from:**
   - Individual sleeve ClashZones
   - Cluster sleeve ClashZones (via ClusterSleeves lookup)
4. **System transfers aggregated parameters to combined sleeve**

### Verification

After restart, check logs for:
```
[SQLite] ✅ Batch saved 2 combined sleeves and updated flags.
[SQLite]   -> CombinedInstanceId set for ClusterSleeve 1158126
[SQLite]   -> CombinedInstanceId set for ClusterSleeve 1158117
```

Then check database:
```sql
SELECT ClusterInstanceId, CombinedClusterSleeveInstanceId 
FROM ClusterSleeves 
WHERE CombinedClusterSleeveInstanceId != -1;
```

Should return:
```
1158126 | 1158140
1158117 | 1158149
```

## ✅ Summary

- **`IsCombinedResolved`**: Only in `ClashZones` table ✅
- **`CombinedClusterSleeveInstanceId`**: In BOTH `ClashZones` AND `ClusterSleeves` tables ✅
- **Reason**: Parameter transfer needs to look up cluster sleeve data via `CombinedClusterSleeveInstanceId`
- **Action**: Restart Revit to apply migration

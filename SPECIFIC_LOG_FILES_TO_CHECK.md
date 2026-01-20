# Specific Log Files to Check for Optimizations

## Location
All logs are in: `C:\Users\jse2084\AppData\Roaming\JSE_MEP_Openings\Logs\R2023\`

---

## 1. ✅ PARALLEL PLANNING

### Log Files:
1. **`placement_debug.log`** (PRIMARY)
   - Contains: `[PLANNING]` entries
   - Look for:
     - `[PLANNING] Starting parallel planning for X zones...`
     - `[PLANNING] Planning completed in X ms`
     - `[PLANNING] Parallel planning succeeded: X zones analyzed`
     - `[PLANNING] Parallel planning disabled` (if not working)

2. **`planning_debug.log`** (SECONDARY)
   - Created by: `ParallelSleevePlacementPlanner.cs`
   - Contains: Detailed planning results per zone

### How to Verify:
```bash
# Search for planning entries
grep -i "\[PLANNING\]" placement_debug.log

# Check if parallel planning ran
grep -i "Parallel planning succeeded" placement_debug.log

# Check if it was disabled
grep -i "Parallel planning disabled" placement_debug.log
```

### Expected Output (if working):
```
[PLANNING] Starting parallel planning for 91 zones...
[PLANNING] Planning completed in 45.23 ms
[PLANNING] Total zones: 91
[PLANNING] Zones to skip: 0
[PLANNING] Parallel planning succeeded: 91 zones analyzed, 0 skipped, 91 eligible for placement
```

### If NOT Working:
```
[PLANNING] Parallel planning disabled (EnableParallelPlanning=true, Planner=available)
```

---

## 2. ✅ R-TREE DATABASE QUERIES

### Log Files:
1. **`Refresh_*.log`** (PRIMARY - from refresh operation)
   - Contains: `[SQLite]` entries with R-tree queries
   - Look for:
     - `[SQLite] Using R-tree query path`
     - `[SQLite] ✅ R-tree query returned X zones`
     - `[SQLite] ⚠️ R-tree query failed, falling back to B-tree` (if not working)

2. **`placement_debug.log`** (SECONDARY - may have SQLite entries)
   - Check for any `[SQLite]` entries during placement

### How to Verify:
```bash
# Search for R-tree entries in refresh logs
grep -i "R-tree" Refresh_*.log

# Check if R-tree was used
grep -i "Using R-tree query path" Refresh_*.log

# Check for fallback (means R-tree failed)
grep -i "falling back to B-tree" Refresh_*.log
```

### Expected Output (if working):
```
[SQLite] [FLAG-RESET] Using R-tree query path for filter 'FilterName', category 'Pipes'
[SQLite] ✅ R-tree query returned 91 zones for filter 'FilterName', category 'Pipes'
```

### If NOT Working:
```
[SQLite] ⚠️ R-tree query failed, falling back to B-tree: [error message]
[SQLite] [FLAG-RESET] Using B-tree fallback path for filter 'FilterName', category 'Pipes'
```

**Note**: R-tree only works when:
- Section box is ACTIVE in 3D view
- `OptimizationFlags.UseRTreeDatabaseIndex = true`
- Database has `ClashZonesRTree` table

---

## 3. ✅ SPATIAL GRID (Two-Tier Filtering)

### Log Files:
1. **`DIAGNOSTIC_TEST.log`** (PRIMARY)
   - Contains: `[SPATIAL_GRID_INIT]` and `[TwoTier]` entries
   - Look for:
     - `[SPATIAL_GRID_INIT] OptimizationFlags.UseSpatialGrid = true`
     - `[SPATIAL_GRID_INIT] _spatialService created successfully`
     - `[TwoTier] TIER 1` and `[TwoTier] TIER 2` messages

2. **`placement_debug.log`** (SECONDARY)
   - May contain spatial filtering messages

### How to Verify:
```bash
# Check if spatial grid is initialized
grep -i "SPATIAL_GRID_INIT" DIAGNOSTIC_TEST.log

# Check for two-tier filtering
grep -i "\[TwoTier\]" DIAGNOSTIC_TEST.log

# Check spatial filtering stats
grep -i "spatially filtered" DIAGNOSTIC_TEST.log
```

### Expected Output (if working):
```
[SPATIAL_GRID_INIT] OptimizationFlags.UseSpatialGrid = true
[SPATIAL_GRID_INIT] _spatialService created successfully
[TwoTier] Summary MEP 12345: Tier1 rejected 500, Tier2 rejected 200, Total rejected 700/1000 (70.0%)
```

### If NOT Working:
```
[SPATIAL_GRID_INIT] Spatial grid DISABLED (flag is false)
[TwoTier] R-tree filtering disabled, using spatial grid results directly
```

---

## 4. ✅ PARAMETER BATCHING (Already Identified Issue)

### Log Files:
1. **`parameter_batching_performance.log`** ✅ (YOU ALREADY HAVE THIS)
   - Shows: Incremental flushing (PROBLEM!)

2. **`placement_debug.log`**
   - Contains: `[BATCH-PARAMS]` entries
   - Look for:
     - `[BATCH-PARAMS] 🔄 ABOUT TO FLUSH: X sleeves`
     - `[BATCH-PARAMS] ✅ FLUSH COMPLETE`

### How to Verify:
```bash
# Check batching entries
grep -i "\[BATCH-PARAMS\]" placement_debug.log

# Count how many flushes occurred
grep -i "ABOUT TO FLUSH" placement_debug.log | wc -l
```

### Expected Output (if working correctly):
```
[BATCH-PARAMS] 🔄 ABOUT TO FLUSH: 91 sleeves with 728 total parameters after regeneration...
[BATCH-PARAMS] ✅ FLUSH COMPLETE: 91 sleeves (728 parameters) in 30ms
```
**Should see ONLY 1 flush, not 91!**

### Current Output (BROKEN):
```
[BATCH-PARAMS] 🔄 ABOUT TO FLUSH: 1 sleeves with 8 total parameters...
[BATCH-PARAMS] 🔄 ABOUT TO FLUSH: 2 sleeves with 16 total parameters...
... (repeated 91 times)
```

---

## 5. 📋 SUMMARY CHECKLIST

### Files to Check (in order of importance):

1. **`placement_debug.log`** ⭐⭐⭐
   - Parallel Planning: `[PLANNING]` entries
   - Parameter Batching: `[BATCH-PARAMS]` entries
   - General placement operations

2. **`Refresh_*.log`** ⭐⭐⭐
   - R-Tree Database: `[SQLite]` entries
   - Latest refresh operation log

3. **`DIAGNOSTIC_TEST.log`** ⭐⭐
   - Spatial Grid: `[SPATIAL_GRID_INIT]` and `[TwoTier]` entries

4. **`planning_debug.log`** ⭐
   - Detailed parallel planning results

5. **`parameter_batching_performance.log`** ⭐⭐⭐
   - Already identified - shows incremental flushing issue

---

## 6. 🔍 QUICK VERIFICATION COMMANDS

### Windows PowerShell (run in log directory):
```powershell
# Check parallel planning
Select-String -Path "placement_debug.log" -Pattern "\[PLANNING\]" | Select-Object -First 10

# Check R-tree
Select-String -Path "Refresh_*.log" -Pattern "R-tree" | Select-Object -First 10

# Check spatial grid
Select-String -Path "DIAGNOSTIC_TEST.log" -Pattern "SPATIAL" | Select-Object -First 10

# Count batching flushes (should be 1, not 91)
(Select-String -Path "placement_debug.log" -Pattern "ABOUT TO FLUSH").Count
```

---

## 7. ⚠️ IF LOGS ARE MISSING

### Possible Reasons:
1. **Deployment Mode Enabled**: `DeploymentConfiguration.DeploymentMode = true`
   - **Fix**: Set to `false` for debugging

2. **Logs Not Generated Yet**: Operations haven't run
   - **Fix**: Run refresh and placement operations

3. **Different Log Location**: Check other subdirectories
   - **Fix**: Search entire `Logs` folder

4. **Optimization Disabled**: Flag is `false`
   - **Fix**: Check `OptimizationFlags.cs` and `DeploymentConfiguration.cs`

---

**Generated**: 2025-12-02  
**Purpose**: Identify exact log files to check for optimization verification


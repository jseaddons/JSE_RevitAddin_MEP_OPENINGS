# Files Affected by Batching Optimization - Restore List

## ✅ **Files to RESTORE from Git** (Batching broke working logic)

### **Primary Files (Critical - Must Restore)**
1. **`Services/UniversalSleevePlacerService.cs`** ⚠️ **CRITICAL**
   - Contains `_deferredParameters` dictionary
   - Contains `TimedSetDouble` method with batching logic
   - Contains `FlushDeferredParameters` method
   - **Issue**: Batching broke Depth/Wall Width setting logic
   - **Lines affected**: ~4600-4700 (TimedSetDouble), ~4400-4600 (FlushDeferredParameters), ~5100-5200 (Depth setting)

2. **`Services/ParameterBatchingService.cs`** ⚠️ **NEW FILE - DELETE**
   - Entirely new file created for batching
   - Can be deleted if not used elsewhere

3. **`Services/OptimizationFlags.cs`** ⚠️ **DO NOT RESTORE**
   - Contains `UseBatchedParameterWrites` flag (line 314)
   - **Action**: Set `UseBatchedParameterWrites = false` to disable batching
   - **Reason**: User has added refresh flags - keep the file but disable batching

### **Secondary Files (May have batching references)**
4. **`Services/Clustering/Placement/ClusterPlacementService.cs`**
   - May have batching logic for cluster sleeves
   - Check if it uses `ParameterBatchingService` or `_deferredParameters`

5. **`Services/Clustering/RefactoredClusterService.cs`**
   - May have batching references
   - Check for `UseBatchedParameterWrites` usage

6. **`Services/Placement/Stages/FlushParametersStage.cs`** (if exists)
   - New file for batching stage
   - May need to delete

7. **`Services/Placement/Stages/ParameterAssignStage.cs`** (if exists)
   - New file for parameter assignment
   - May need to delete

## ✅ **Files to KEEP** (User specified)
- `Services/ClashZoneService_Legacy.cs` - **KEEP**
- `refresh refactor/refresh_service_refactored.cs` - **KEEP**
- `refresh refactor/intersection_processor.cs` - **KEEP**
- `Data/Repositories/ClashZoneRepository.cs` - **KEEP** (has diagnostic logging)

## 📋 **Restore Commands**

### **Option 1: Restore specific files from Git**
```bash
# Restore UniversalSleevePlacerService.cs (CRITICAL)
git checkout HEAD -- Services/UniversalSleevePlacerService.cs

# DO NOT restore OptimizationFlags.cs - user has refresh flags
# ✅ ALREADY DONE: UseBatchedParameterWrites set to false in OptimizationFlags.cs

# Delete new batching files
git rm Services/ParameterBatchingService.cs
git rm Services/Placement/Stages/FlushParametersStage.cs  # if exists
git rm Services/Placement/Stages/ParameterAssignStage.cs   # if exists

# Check cluster services
git checkout HEAD -- Services/Clustering/Placement/ClusterPlacementService.cs
git checkout HEAD -- Services/Clustering/RefactoredClusterService.cs
```

### **Option 2: Restore all except specified files**
```bash
# Restore all modified files
git checkout HEAD -- .

# Then restore the files you want to keep
git checkout HEAD~1 -- Services/ClashZoneService_Legacy.cs
git checkout HEAD~1 -- refresh\ refactor/refresh_service_refactored.cs
git checkout HEAD~1 -- refresh\ refactor/intersection_processor.cs
git checkout HEAD~1 -- Data/Repositories/ClashZoneRepository.cs
```

## 🔍 **What to Check After Restore**

1. **Verify Depth/Wall Width setting works correctly**
   - Check that `param.Set(thickness)` is called directly (not deferred)
   - Verify no `_deferredParameters` dictionary exists
   - Confirm `TimedSetDouble` calls `p.Set(value)` immediately

2. **Verify Width/Height setting works correctly**
   - Check that cable tray dimensions are set correctly
   - Verify clearances are read from database (not XML)

3. **Check for any remaining batching references**
   ```bash
   grep -r "UseBatchedParameterWrites" Services/
   grep -r "_deferredParameters" Services/
   grep -r "ParameterBatchingService" Services/
   ```

## ⚠️ **After Restore - Re-implement Batching Carefully**

When re-implementing batching:
1. **DO NOT modify parameter setting logic** - keep `param.Set(value)` calls
2. **Add batching as a wrapper** - don't change the core logic
3. **Test Depth/Wall Width first** - ensure it works before batching other parameters
4. **Add comprehensive logging** - track what values are being set and flushed
5. **Test with 9 sleeves** - verify all get correct depth values


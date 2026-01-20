# NewSleevePlacerService Update Requirements

**Date:** December 2025  
**Purpose:** Document all updates needed in `NewSleevePlacerService` to match recent improvements in `UniversalSleevePlacerService`

---

## 📋 Overview

`UniversalSleevePlacerService` has received significant updates for:
- Parameter batching with safety mechanisms
- Family symbol cache validation and stale cache handling
- Diagnostic logging and performance monitoring
- Crash-safe error handling
- Insulation-aware sizing service integration

`NewSleevePlacerService` needs these same updates to maintain feature parity and SOLID compliance.

---

## ✅ Already Implemented in NewSleevePlacerService

Based on `REFACTORED_SERVICE_3_DIVISION_WORK_PLAN.md`:

- ✅ Dependency Injection infrastructure (IParameterBatchingService, IPerformanceMonitor)
- ✅ Constructor with DI for all services
- ✅ Document state validation
- ✅ Performance monitoring framework
- ✅ Parallel planning support
- ✅ Basic error handling

---

## 🔴 Missing Updates (Required)

### 1. Parameter Batching with Safety Flags

**Status:** ❌ **MISSING**

**What UniversalSleevePlacerService Has:**
```csharp
// Safety flag to prevent multiple flushes
private bool _hasFlushedParameters = false;

// Deferred parameters dictionary
private Dictionary<ElementId, Dictionary<string, object>> _deferredParameters = 
    new Dictionary<ElementId, Dictionary<string, object>>();

// Reset at placement start
_hasFlushedParameters = false;
_deferredParameters?.Clear();
```

**What NewSleevePlacerService Needs:**
- Add `_hasFlushedParameters` safety flag
- Add `_deferredParameters` dictionary
- Reset flags at start of `PlaceAllSleevesInTransaction`
- Implement `FlushDeferredParameters()` with safety checks
- Update `SetSleeveParameters()` to use deferred batching

**Location:** `Services/NewSleevePlacerService.cs`
- Lines ~429-443: `SetSleeveParameters` method
- Add new private fields after line 38
- Add flush logic at end of `PlaceAllSleevesInTransaction`

---

### 2. Family Symbol Cache Validation

**Status:** ❌ **MISSING**

**What UniversalSleevePlacerService Has:**
```csharp
// Cache validation before use
if (!familySymbol.IsValidObject)
{
    // Remove from cache, log error
    _familySymbolCache.Remove(familyName);
    return null;
}

// Test IsActive access in try-catch (prevents stale reference errors)
try
{
    if (familySymbol != null && familySymbol.IsValidObject)
    {
        shouldActivate = !familySymbol.IsActive;
    }
}
catch (InvalidOperationException)
{
    // Symbol is stale - remove from cache
    _familySymbolCache.Remove(familyName);
    return null;
}

// Clear cache after regeneration
_doc.Regenerate();
_familySymbolCache.Clear();
```

**What NewSleevePlacerService Needs:**
- Add `IsValidObject` check in `LoadFamilySymbol` (line ~366)
- Add try-catch around `IsActive` access
- Remove stale entries from cache
- Clear cache after document regeneration
- Add `PreCacheFamilySymbols` method with validation

**Location:** `Services/NewSleevePlacerService.cs`
- Lines ~366-397: `LoadFamilySymbol` method
- Add cache clearing after placement loop

---

### 3. GetParameterValueWithBatchingSupport Method

**Status:** ❌ **MISSING**

**What UniversalSleevePlacerService Has:**
```csharp
/// <summary>
/// ✅ CRITICAL FIX: Read parameter from deferred cache first, then fallback to Revit element.
/// This prevents stale reads during corner placement calculations when batching is enabled.
/// </summary>
private double GetParameterValueWithBatchingSupport(
    FamilyInstance sleeve, 
    string parameterName, 
    double fallbackValue)
{
    if (OptimizationFlags.UseBatchedParameterWrites && _deferredParameters != null)
    {
        var sleeveId = sleeve.Id;
        if (_deferredParameters.ContainsKey(sleeveId) && 
            _deferredParameters[sleeveId].ContainsKey(parameterName))
        {
            var cachedValue = _deferredParameters[sleeveId][parameterName];
            if (cachedValue is double doubleVal)
            {
                return doubleVal; // Return cached value
            }
        }
    }
    
    // Fallback to Revit element
    var param = sleeve.LookupParameter(parameterName);
    return param?.AsDouble() ?? fallbackValue;
}
```

**What NewSleevePlacerService Needs:**
- Add this method to read from deferred cache
- Use in dimension calculations when batching is enabled
- Prevents reading stale values during placement

**Location:** `Services/NewSleevePlacerService.cs`
- Add new private method after `SetSleeveParameters`
- Use in `CalculateSleeveDimensions` if needed

---

### 4. TimedSetDouble with Diagnostic Logging

**Status:** ❌ **MISSING**

**What UniversalSleevePlacerService Has:**
```csharp
private void TimedSetDouble(Parameter param, double value, string logicalName, ElementId currentSleeveId)
{
    if (param == null || param.IsReadOnly) return;
    
    if (OptimizationFlags.UseBatchedParameterWrites)
    {
        if (!_deferredParameters.ContainsKey(currentSleeveId))
            _deferredParameters[currentSleeveId] = new Dictionary<string, object>();
        
        // ✅ CRITICAL DIAGNOSTIC: Log if parameter is being overwritten
        bool isOverwrite = _deferredParameters[currentSleeveId].ContainsKey(logicalName);
        if (isOverwrite && (logicalName == "Width" || logicalName == "Height" || 
                           logicalName == "Depth" || logicalName == "Wall Width"))
        {
            var oldValue = _deferredParameters[currentSleeveId][logicalName];
            SafeFileLogger.SafeAppendText("parameter_overwrite_debug.log",
                $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] ⚠️ PARAMETER OVERWRITE: Sleeve {currentSleeveId.IntegerValue}, " +
                $"Parameter='{logicalName}', OldValue={oldValue}, NewValue={value}\n");
        }
        
        _deferredParameters[currentSleeveId][logicalName] = value;
    }
    else
    {
        param.Set(value);
    }
}
```

**What NewSleevePlacerService Needs:**
- Replace direct `param.Set()` calls with `TimedSetDouble`
- Add diagnostic logging for parameter overwrites
- Support both batched and immediate parameter setting

**Location:** `Services/NewSleevePlacerService.cs`
- Lines ~429-443: `SetSleeveParameters` method
- Replace direct parameter setting with `TimedSetDouble` calls

---

### 5. FlushDeferredParameters with Comprehensive Safety

**Status:** ❌ **MISSING**

**What UniversalSleevePlacerService Has:**
```csharp
public void FlushDeferredParameters()
{
    // ✅ SAFETY FLAG: Prevent multiple flushes
    if (_hasFlushedParameters)
    {
        DebugLogger.Warning($"[BATCH-PARAMS] ⚠️ SAFETY: FlushDeferredParameters called AGAIN - IGNORING");
        return;
    }
    
    if (_deferredParameters == null || _deferredParameters.Count == 0)
    {
        _hasFlushedParameters = true;
        return;
    }
    
    // Flush all parameters in batch
    int successCount = 0;
    int failCount = 0;
    
    try
    {
        foreach (var kvp in _deferredParameters)
        {
            var sleeveId = kvp.Key;
            var paramValues = kvp.Value;
            
            var sleeve = _doc.GetElement(sleeveId) as FamilyInstance;
            if (sleeve == null) continue;
            
            foreach (var paramKvp in paramValues)
            {
                var param = sleeve.LookupParameter(paramKvp.Key);
                if (param != null && !param.IsReadOnly)
                {
                    try
                    {
                        if (paramKvp.Value is double dVal)
                            param.Set(dVal);
                        else if (paramKvp.Value is string sVal)
                            param.Set(sVal);
                        
                        successCount++;
                    }
                    catch (Exception ex)
                    {
                        failCount++;
                        // Log error
                    }
                }
            }
        }
    }
    finally
    {
        _hasFlushedParameters = true;
        _deferredParameters.Clear();
    }
}
```

**What NewSleevePlacerService Needs:**
- Implement `FlushDeferredParameters()` method
- Add safety flag check
- Add comprehensive error handling
- Add diagnostic logging
- Call at end of `PlaceAllSleevesInTransaction` before transaction commit

**Location:** `Services/NewSleevePlacerService.cs`
- Add new public method after `SetSleeveParameters`
- Call in `PlaceAllSleevesInTransaction` before return statement

---

### 6. PreCacheFamilySymbols with Validation

**Status:** ❌ **MISSING**

**What UniversalSleevePlacerService Has:**
```csharp
private void PreCacheFamilySymbols(List<ClashZone> clashZones)
{
    var uniqueFamilyNames = clashZones
        .Select(cz => GetSleeveFamilyName(cz, /* determine shape */))
        .Distinct()
        .ToList();
    
    foreach (var familyName in uniqueFamilyNames)
    {
        var symbol = LoadFamilySymbol(familyName);
        if (symbol != null && symbol.IsValidObject)
        {
            try
            {
                // Test if symbol is active (will throw if stale)
                var _ = symbol.IsActive;
                _familySymbolCache[familyName] = symbol;
            }
            catch (InvalidOperationException)
            {
                // Symbol is stale - don't cache
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning($"[UniversalSleevePlacer] Stale symbol detected for '{familyName}', skipping cache");
            }
        }
    }
}
```

**What NewSleevePlacerService Needs:**
- Add `PreCacheFamilySymbols` method
- Validate symbols before caching
- Call at start of `PlaceAllSleevesInTransaction`

**Location:** `Services/NewSleevePlacerService.cs`
- Add new private method after `GetSleeveFamilyName`
- Call in `PlaceAllSleevesInTransaction` before placement loop

---

### 7. Diagnostic Logging for Parameter Batching

**Status:** ❌ **MISSING**

**What UniversalSleevePlacerService Has:**
```csharp
// At placement start
if (!DeploymentConfiguration.DeploymentMode)
{
    DebugLogger.Info($"[BATCH-PARAMS] ═══ PLACEMENT START ═══ UseBatchedParameterWrites={OptimizationFlags.UseBatchedParameterWrites}, " +
        $"_deferredParameters initialized={_deferredParameters != null}, FlushFlag reset");
}

// Before flush
if (!DeploymentConfiguration.DeploymentMode)
{
    DebugLogger.Info($"[BATCH-PARAMS] 🔄 Flushing {_deferredParameters.Count} individual sleeves with {totalParams} total parameters...");
}

// After flush
if (!DeploymentConfiguration.DeploymentMode)
{
    DebugLogger.Info($"[BATCH-PARAMS] ✅ Flushed {successCount} individual sleeves ({totalParams} parameters) in batch");
}
```

**What NewSleevePlacerService Needs:**
- Add diagnostic logging at placement start
- Add logging before/after parameter flush
- Add logging for parameter overwrites
- Respect `DeploymentConfiguration.DeploymentMode`

**Location:** `Services/NewSleevePlacerService.cs`
- Add logging in `PlaceAllSleevePlacerService` at start
- Add logging in `FlushDeferredParameters`
- Add logging in `TimedSetDouble`

---

### 8. Cache Clearing After Regeneration

**Status:** ❌ **MISSING**

**What UniversalSleevePlacerService Has:**
```csharp
// After document regeneration
_doc.Regenerate();
_familySymbolCache.Clear(); // Clear stale cache entries
```

**What NewSleevePlacerService Needs:**
- Clear `_familySymbolCache` after any `_doc.Regenerate()` call
- Prevents stale symbol references

**Location:** `Services/NewSleevePlacerService.cs`
- Add cache clearing after regeneration (if regeneration is performed)

---

### 9. Insulation-Aware Sizing Service Integration

**Status:** ✅ **ALREADY IMPLEMENTED**

**What NewSleevePlacerService Has:**
```csharp
// ✅ OOP METHOD: Insulation-aware sizing service (SOLID principles)
private readonly IInsulationAwareSizingService _sizingService;

// In constructor
_sizingService = sizingService ?? new InsulationAwareSizingService();

// In CalculateSleeveDimensions
(double finalWidth, double finalHeight, double finalDiameter) = 
    _sizingService.CalculateFinalDimensionsFromClashZone(
        rawWidth, rawHeight, rawDiameter, zone, clearance);
```

**Status:** ✅ **No changes needed** - Already properly implemented

---

### 10. Performance Monitoring Integration

**Status:** ⚠️ **PARTIALLY IMPLEMENTED**

**What UniversalSleevePlacerService Has:**
- Performance monitoring for all major operations
- Timing for parameter setting, family loading, placement
- Comprehensive metrics logging

**What NewSleevePlacerService Needs:**
- Add performance timing to `SetSleeveParameters`
- Add performance timing to `LoadFamilySymbol`
- Add performance timing to `PlaceSleeveInstance`
- Add performance timing to `CalculateSleeveDimensions`
- Add comprehensive metrics at end of `PlaceAllSleevesInTransaction`

**Location:** `Services/NewSleevePlacerService.cs`
- Add `_performanceMonitor.StartOperation()` / `StopOperation()` calls in each method
- Add summary logging at end of `PlaceAllSleevesInTransaction`

---

## 📝 Implementation Priority

### HIGH PRIORITY (Critical for Functionality)
1. ✅ Parameter Batching with Safety Flags (#1)
2. ✅ Family Symbol Cache Validation (#2)
3. ✅ FlushDeferredParameters Method (#5)

### MEDIUM PRIORITY (Important for Performance)
4. ✅ TimedSetDouble with Diagnostic Logging (#4)
5. ✅ GetParameterValueWithBatchingSupport (#3)
6. ✅ PreCacheFamilySymbols (#6)

### LOW PRIORITY (Nice to Have)
7. ✅ Diagnostic Logging (#7)
8. ✅ Cache Clearing After Regeneration (#8)
9. ✅ Performance Monitoring Integration (#10)

---

## 🔧 Implementation Steps

### Step 1: Add Parameter Batching Infrastructure
1. Add `_hasFlushedParameters` and `_deferredParameters` fields
2. Reset flags at start of `PlaceAllSleevesInTransaction`
3. Implement `FlushDeferredParameters()` method
4. Call flush at end of placement loop

### Step 2: Update Parameter Setting
1. Create `TimedSetDouble()` method
2. Update `SetSleeveParameters()` to use `TimedSetDouble()`
3. Add diagnostic logging for overwrites

### Step 3: Add Family Symbol Cache Validation
1. Update `LoadFamilySymbol()` with validation
2. Add try-catch around `IsActive` access
3. Remove stale entries from cache
4. Add `PreCacheFamilySymbols()` method

### Step 4: Add Supporting Methods
1. Add `GetParameterValueWithBatchingSupport()` method
2. Add diagnostic logging throughout
3. Add performance monitoring calls

### Step 5: Testing
1. Build for R2023 and R2024
2. Test with `UseNewSleevePlacerService = true`
3. Verify parameter batching works
4. Check logs for diagnostic information
5. Verify no performance regression

---

## 📚 Reference Documents

- `SOLID_REFACTORING_TEST_CONFIGURATION.md` - Configuration guide
- `REFACTORED_SERVICE_3_DIVISION_WORK_PLAN.md` - Original implementation plan
- `REFACTORED_SERVICE_VERIFICATION_CHECKLIST.md` - Verification checklist
- `Services/UniversalSleevePlacerService.cs` - Reference implementation

---

## ✅ Success Criteria

- [ ] All HIGH PRIORITY items implemented
- [ ] All MEDIUM PRIORITY items implemented
- [ ] R2023 and R2024 builds succeed (0 errors)
- [ ] Parameter batching provides 4-6× speedup
- [ ] No stale family symbol errors
- [ ] Diagnostic logging works correctly
- [ ] Performance monitoring logs comprehensive metrics
- [ ] Feature parity with UniversalSleevePlacerService achieved

---

**Document Status:** ✅ Complete  
**Last Updated:** December 2025  
**Next Steps:** Implement HIGH PRIORITY items first, then MEDIUM, then LOW


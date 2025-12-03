# Batching Code Comparison: Original vs Current

## Summary
**GOOD NEWS**: The original parameter setting logic (`param.Set(value)`) is **PRESERVED** and still works when batching is disabled.

## Key Finding: Original Logic Intact

### Original Code (Still Present)
```csharp
// Line 4634-4645: IMMEDIATE WRITE PATH (when batching disabled)
if (OptimizationFlags.EnableParameterTimingInstrumentation)
{
    var sw = Stopwatch.StartNew();
    p.Set(value);  // ✅ ORIGINAL CODE - STILL HERE!
    sw.Stop();
    AppendTiming(logicalName, sw.ElapsedTicks, sw.ElapsedMilliseconds, value);
}
else
{
    p.Set(value);  // ✅ ORIGINAL CODE - STILL HERE!
}
```

### New Code (When Batching Enabled)
```csharp
// Line 4609-4631: DEFER PATH (when batching enabled)
if (OptimizationFlags.UseBatchedParameterWrites)
{
    // Store in dictionary for later batch write
    _deferredParameters[sleeveInstance.Id][logicalName] = value;
    return; // Skip immediate write
}
```

## What Changed (From Git Diff)

### 1. Added Safety Flags
- `_hasFlushedParameters` - Prevents duplicate flushes
- `_hasFlushed` in `ParameterBatchingService` - Same purpose

### 2. Added FamilySymbol Validation
- Cache validation to prevent "referenced object is not valid" errors
- Cache clearing after `doc.Regenerate()`

### 3. Added Diagnostic Logging
- Logs when parameters are deferred
- Logs when parameters are flushed
- Logs if parameters are overwritten (NEW - just added)

### 4. **NO CHANGES TO CORE PARAMETER SETTING LOGIC**
- The `p.Set(value)` call is still there
- It's just wrapped in a conditional that checks batching flag

## Potential Issue: Parameter Overwriting

### The Problem
When batching is enabled, parameters are stored in a dictionary:
```csharp
_deferredParameters[sleeveInstance.Id][logicalName] = value;
```

**If the same parameter is set twice for the same sleeve, the second value OVERWRITES the first!**

### Where This Could Happen
1. **Line 4916-4917**: Width/Height set for rectangular sleeves (including cable trays)
2. **Line 4974-4975**: Width/Height set AGAIN for floor ducts (but NOT for cable trays - line 4981-4986)

**For cable trays on walls**: Should only be set ONCE (line 4916-4917)
**For cable trays on floors**: Should only be set ONCE (line 4916-4917, then line 4981-4986 just logs, no swap)

### Diagnostic Logging Added
1. **Overwrite Detection** (line 4615-4623): Logs if Width/Height are overwritten in deferred dictionary
2. **Flush Values** (line 4487-4497): Logs what values are actually being flushed to sleeves
3. **Set Values** (line 4910-4913): Logs what values are being set (already added)

## How to Verify

When you run a placement, check `cabletray_dimension_trace.log` for:

1. `[SET-PARAMS]` - What values are being deferred
2. `[OVERWRITE-WARNING]` - If Width/Height are being set twice (THIS IS THE BUG!)
3. `[FLUSH]` - What values are actually being written to sleeves

If you see `[OVERWRITE-WARNING]`, that means Width/Height are being set multiple times, and only the LAST value is being flushed.

## Conclusion

**The original code is NOT lost** - it's still there and works when batching is disabled.

**The issue is likely**: Parameters being overwritten in the deferred dictionary, or values being calculated incorrectly before being deferred.

The diagnostic logging will show exactly what's happening.


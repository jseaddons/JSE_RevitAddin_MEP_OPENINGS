# Batch Parameter Write Fix Plan - Individual Sleeves

## Executive Summary

The batch parameter write system for **individual sleeve placement** is broken. While cluster batch mode has been wired up correctly (with `DivertedBatchDictionary` and flush calls), individual sleeve placement never calls `FlushDeferredParameters()` after the placement loop, causing all deferred parameters to be lost.

---

## Current State Analysis

### What Works (Clusters)

```
RefactoredClusterService.ClusterSleeves()
    ↓
_parameterService.DivertedBatchDictionary = _deferredClusterParameters  ← WIRED
    ↓
foreach (cluster) { PlaceClusterSleeve() }
    ↓
doc.Regenerate()
    ↓
FlushDeferredClusterParameters()  ← CALLED ✅
    ↓
doc.Regenerate()
    ↓
Cleanup + DB Save
```

### What's Broken (Individual Sleeves)

```
UniversalSleevePlacerService / SleeveProcessingOrchestrator
    ↓
foreach (zone) { PlaceIndividualSleeve() }
    ↓
SleeveParameterService.SetSleeveParameters()
    ↓
Parameters added to _deferredParameters  ← ACCUMULATED
    ↓
... placement loop ends ...
    ↓
❌ FlushDeferredParameters() NEVER CALLED ❌
    ↓
Parameters are LOST, sleeves have default values
```

---

## Root Causes

### 1. Missing Flush Call

`SleeveParameterService.FlushDeferredParameters()` exists and works correctly, but **nobody calls it** after the individual sleeve placement loop completes.

**Location of issue:** The individual sleeve orchestrator (likely `UniversalSleevePlacerService` or `SleeveProcessingOrchestrator`) needs to call flush after all sleeves are placed.

### 2. `UseBatchedParameterWrites` Flag Not Consistently Checked

The flag `OptimizationFlags.UseBatchedParameterWrites` is checked in `SetSleeveParameters()`, but the caller doesn't know to call flush when the flag is ON.

### 3. No Reset Between Batches

`SleeveParameterService` has `_hasFlushedParameters` flag to prevent double-flush, but `ResetFlushFlag()` may not be called at the start of each placement batch.

---

## Files To Modify

| File | Change |
|------|--------|
| **Services/Placement/UniversalSleevePlacerService.cs** | Add flush call after placement loop |
| **Services/SleeveProcessingOrchestrator.cs** | May also need flush call (verify which is the main orchestrator) |
| **Services/Placement/SleeveParameterService.cs** | Already correct, just verify |

---

## Fix Implementation

### Step 1: Find the Individual Sleeve Placement Orchestrator

Look for the method that loops through zones and places individual sleeves. It will have a pattern like:

```csharp
foreach (var zone in zonesToPlace)
{
    // ... place sleeve ...
    _parameterService.SetSleeveParameters(instance, width, height, diameter, isCircular, zone);
}
```

The flush call needs to go **AFTER** this loop.

### Step 2: Add Flush Call After Placement Loop

```csharp
// AFTER the placement loop completes
if (OptimizationFlags.UseBatchedParameterWrites && _parameterService != null)
{
    try
    {
        // Regenerate to ensure geometry is ready for parameter writes
        doc.Regenerate();
        
        // Flush all accumulated parameters in single batch
        int flushedCount = _parameterService.FlushDeferredParameters();
        
        SafeFileLogger.SafeAppendText("placement_debug.log",
            $"[{DateTime.Now:HH:mm:ss}] ✅ BATCH-FLUSH: Flushed {flushedCount} individual sleeve parameters\n");
        
        // Regenerate again to apply parameter changes
        doc.Regenerate();
    }
    catch (Exception ex)
    {
        SafeFileLogger.SafeAppendText("placement_errors.log",
            $"[{DateTime.Now:HH:mm:ss}] ❌ BATCH-FLUSH ERROR: {ex.Message}\n");
    }
}
```

### Step 3: Reset Flush Flag at Start of Batch

At the **START** of the placement method, reset the flag:

```csharp
public void PlaceSleeves(...)
{
    // Reset flush flag for new placement batch
    _parameterService?.ResetFlushFlag();
    
    // ... rest of method ...
}
```

### Step 4: Verify Parameter Flow in SetSleeveParameters

The existing code in `SetSleeveParameters()` already handles batching correctly:

```csharp
if (OptimizationFlags.UseBatchedParameterWrites)
{
    var targetDict = ActiveBatchDictionary;  // Uses _deferredParameters if DivertedBatchDictionary is null
    if (!targetDict.ContainsKey(currentSleeveId))
        targetDict[currentSleeveId] = new Dictionary<string, object>();
    
    targetDict[currentSleeveId]["Width"] = roundedWidth;
    targetDict[currentSleeveId]["Height"] = roundedHeight;
    // ... etc
}
```

**No changes needed here** - just ensure flush is called.

---

## Verification Checklist

After implementing the fix, verify:

- [ ] `OptimizationFlags.UseBatchedParameterWrites` is `true`
- [ ] Parameters are accumulated during placement (check logs for `[BATCH-ADD]` entries)
- [ ] `FlushDeferredParameters()` is called after loop (check logs for `[BATCH-PARAMS]` entries)
- [ ] Sleeves have correct Depth/Width/Height after placement
- [ ] No double-flush (check for `SAFETY: FlushDeferredParameters called AGAIN` warning)

---

## Test Scenarios

### Test 1: Batch Mode ON

1. Set `UseBatchedParameterWrites = true`
2. Place 10 individual sleeves
3. Verify all have correct Depth parameter
4. Check logs show flush happened once at end

### Test 2: Batch Mode OFF (Regression Test)

1. Set `UseBatchedParameterWrites = false`
2. Place same 10 sleeves
3. Verify all have correct Depth parameter
4. Check parameters were set immediately (no flush log)

### Test 3: Mixed Cluster + Individual

1. Place clusters (uses `_deferredClusterParameters`)
2. Place individual sleeves (uses `_deferredParameters`)
3. Verify both have correct parameters
4. No cross-contamination between dictionaries

---

## Code Locations Reference

### SleeveParameterService Key Methods

| Method | Purpose |
|--------|---------|
| `SetSleeveParameters()` | Main entry - accumulates to `_deferredParameters` |
| `SetDepthParameter()` | Sets Depth + Wall Width (calls `SetParameter()`) |
| `SetParameter()` | Adds to `ActiveBatchDictionary` if batching enabled |
| `FlushDeferredParameters()` | Writes all accumulated params to Revit elements |
| `ResetFlushFlag()` | Resets `_hasFlushedParameters` for new batch |
| `ActiveBatchDictionary` | Returns `DivertedBatchDictionary ?? _deferredParameters` |

### Key Properties

```csharp
// Internal dictionary for individual sleeves
private Dictionary<ElementId, Dictionary<string, object>> _deferredParameters

// External dictionary for clusters (when set by RefactoredClusterService)
public Dictionary<ElementId, Dictionary<string, object>> DivertedBatchDictionary

// Returns the active dictionary (external if set, else internal)
private Dictionary<ElementId, Dictionary<string, object>> ActiveBatchDictionary 
    => DivertedBatchDictionary ?? _deferredParameters;
```

---

## Sequence Diagram (Correct Flow)

```
┌─────────────────────────────────┐
│  Start Individual Placement     │
└─────────────────────────────────┘
               │
               ▼
┌─────────────────────────────────┐
│  ResetFlushFlag()               │  ← NEW: Reset at start
└─────────────────────────────────┘
               │
               ▼
┌─────────────────────────────────┐
│  foreach (zone in zones)        │
│  {                              │
│    PlaceSleeve()                │
│    SetSleeveParameters()        │  → Accumulates to _deferredParameters
│  }                              │
└─────────────────────────────────┘
               │
               ▼
┌─────────────────────────────────┐
│  doc.Regenerate()               │  ← NEW: Pre-flush regen
└─────────────────────────────────┘
               │
               ▼
┌─────────────────────────────────┐
│  FlushDeferredParameters()      │  ← NEW: The missing call!
└─────────────────────────────────┘
               │
               ▼
┌─────────────────────────────────┐
│  doc.Regenerate()               │  ← NEW: Post-flush regen
└─────────────────────────────────┘
               │
               ▼
┌─────────────────────────────────┐
│  Continue to DB Save, etc.      │
└─────────────────────────────────┘
```

---

## Summary of Changes

| Change | File | Description |
|--------|------|-------------|
| **Add flush call** | UniversalSleevePlacerService.cs (or orchestrator) | Call `_parameterService.FlushDeferredParameters()` after placement loop |
| **Add reset call** | Same file | Call `_parameterService.ResetFlushFlag()` at start of batch |
| **Add regenerations** | Same file | Add `doc.Regenerate()` before and after flush |
| **Add logging** | Same file | Log flush count for verification |

---

## Expected Outcome

After implementing this fix:

1. **Batch mode ON**: All parameters accumulated during loop, flushed once at end → Faster performance
2. **Batch mode OFF**: Parameters written immediately as before → Same behavior as current
3. **All sleeves have correct Depth**: Parameter is no longer lost
4. **Performance improvement**: Batch mode should be 4-6× faster than sequential writes

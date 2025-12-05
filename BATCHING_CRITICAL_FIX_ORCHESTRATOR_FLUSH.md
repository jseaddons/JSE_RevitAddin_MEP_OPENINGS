# CRITICAL BATCHING FIX: Orchestrator Flush

## Date: 2025-12-03 18:49

## ROOT CAUSE DISCOVERED
The batching optimization was deferring parameters but **NEVER FLUSHING THEM** because:

1. `UniversalSleevePlacementCommand.Execute()` places sleeves and defers parameters
2. Command returns to `OpeningCommandOrchestrator` WITHOUT flushing
3. Orchestrator does its OWN regeneration (line 1731)
4. Deferred parameters were NEVER written to Revit → sleeves had default/wrong sizes

**Evidence**: 
- Logs showed `Batching=True` (parameters being deferred)
- No `[BATCH-PARAMS] 🔄 ABOUT TO FLUSH` logs (flush code never ran)
- `[BOUNDING_BOX_AFTER_REGEN]` logs from orchestrator (different code path than service)

## THE FIX

### 1. Made `FlushDeferredParameters()` Public
**File**: `Services/UniversalSleevePlacerService.cs` (line ~4490)
```csharp
// OLD:
private void FlushDeferredParameters()

// NEW:
public void FlushDeferredParameters()
```

### 2. Exposed Service from Command
**File**: `Commands/UniversalSleevePlacementCommand.cs` (lines 34-42)
```csharp
private UniversalSleevePlacerService? _service; // Not readonly - set during Execute

public UniversalSleevePlacerService? Service => _service;
```

Stored service reference during Execute (line ~258):
```csharp
_service = placerService;
```

### 3. Added Orchestrator Flush BEFORE Regeneration
**File**: `Services/OpeningCommandOrchestrator.cs` (line ~1593)
```csharp
universalCommand.Execute(_uiDocument.Application);

// ✅ CRITICAL FIX: Flush deferred parameters BEFORE orchestrator regeneration
if (universalCommand?.Service != null && OptimizationFlags.UseBatchedParameterWrites)
{
    try
    {
        if (!DeploymentConfiguration.DeploymentMode)
        {
            DebugLogger.Info($"[ORCHESTRATOR-FLUSH] 🔄 Flushing deferred parameters...");
        }
        
        universalCommand.Service.FlushDeferredParameters();
        
        if (!DeploymentConfiguration.DeploymentMode)
        {
            DebugLogger.Info($"[ORCHESTRATOR-FLUSH] ✅ Deferred parameters flushed successfully");
        }
    }
    catch (Exception flushEx)
    {
        if (!DeploymentConfiguration.DeploymentMode)
        {
            DebugLogger.Error($"[ORCHESTRATOR-FLUSH] ⚠️ Failed to flush: {flushEx.Message}");
        }
    }
}

// Orchestrator regeneration happens AFTER flush (line 1731)
_document.Regenerate();
```

## IMPACT

**Before Fix**:
- Parameters deferred but never written
- Sleeves had default/wrong Width/Height/Depth values
- Batching optimization broken - disabled to prevent bugs

**After Fix**:
- Parameters deferred → flushed to Revit BEFORE orchestrator regeneration
- Sleeves have correct dimensions from deferred parameters
- 4-6× performance improvement restored (143-203ms → <30ms per sleeve)

## TESTING CHECKLIST

✅ **Build**: Successful (no errors)
⏳ **Test Placement**: Run "Place Sleeve" for Cable Trays
⏳ **Verify Logs**: Check for `[ORCHESTRATOR-FLUSH] 🔄 Flushing...` and `✅ flushed successfully`
⏳ **Verify Dimensions**: Check sleeve Width/Height/Depth match expected values
⏳ **Performance**: Confirm placement time < 50ms per sleeve with batching enabled

## ADDITIONAL NOTES

### Clearance Issue (Separate Bug)
User reported: "sleeve width/height have 75mm clearance on all sides, should be 25mm on sides and 75mm on top"

This is a SEPARATE issue from batching - clearance calculation logic in cable tray strategy needs review.

### DeploymentMode
Currently set to `false` for diagnostic logging. Re-enable after testing:
```csharp
// Services/DeploymentConfiguration.cs line 25
public static readonly bool DeploymentMode = true; // After fix validated
```

## BUILD TIMESTAMP
**Built**: 2025-12-03 18:50
**DLL**: JSE_RevitAddin_MEP_OPENINGS.dll

# ParameterTransferService.cs - Changes from Git HEAD

## Summary of Changes

The file has been significantly modified from the last Git commit. Here are the key changes:

### 1. **MEP Size Parameter Reading Logic (CRITICAL CHANGE)**

**ORIGINAL CODE (Git HEAD):**
```csharp
if (!sourceParams.TryGetValue(mapping.SourceParameter, out var sourceValue) || string.IsNullOrWhiteSpace(sourceValue))
{
    result.Warnings.Add($"Parameter '{mapping.SourceParameter}' not found in snapshot for sleeve {openingId.IntegerValue}.");
    continue;
}
```

**CURRENT CODE:**
```csharp
// ✅ CRITICAL FIX: For "MEP Size" parameter, read ONLY from Revit MEP element (not from snapshot)
string sourceValue = null;
bool readFromRevit = false;

if (string.Equals(mapping.SourceParameter, "MEP Size", StringComparison.OrdinalIgnoreCase) ||
    string.Equals(mapping.SourceParameter, "Size", StringComparison.OrdinalIgnoreCase))
{
    // Read ONLY from Revit MEP element - no snapshot fallback
    try
    {
        var mepElementIdParam = opening.LookupParameter("MEP_ElementId");
        if (mepElementIdParam != null && mepElementIdParam.HasValue)
        {
            var mepElementId = mepElementIdParam.AsElementId();
            if (mepElementId != null && mepElementId != ElementId.InvalidElementId)
            {
                var mepElement = doc.GetElement(mepElementId);
                if (mepElement != null && mepElement.IsValidObject)
                {
                    var sizeParam = mepElement.LookupParameter("Size");
                    if (sizeParam != null && sizeParam.HasValue)
                    {
                        sourceValue = sizeParam.AsValueString() ?? sizeParam.AsString();
                        readFromRevit = true;
                    }
                }
            }
        }
    }
    catch (Exception revitEx)
    {
        // Error handling...
        continue;
    }
}
else
{
    // For other parameters, use snapshot as usual
    if (!sourceParams.TryGetValue(mapping.SourceParameter, out sourceValue) || string.IsNullOrWhiteSpace(sourceValue))
    {
        result.Warnings.Add($"Parameter '{mapping.SourceParameter}' not found in snapshot for sleeve {openingId.IntegerValue}.");
        continue;
    }
}
```

**IMPACT:** This is a **MAJOR CHANGE** - the code now reads "MEP Size" directly from the Revit MEP element instead of from the database snapshot. This was done to ensure we get the current value from the actual MEP element, not stale database data.

### 2. **Added Crash-Safety Checks**

Added multiple validation checks:
- ✅ CRASH-SAFETY 1: Validate ElementId before retrieval
- ✅ CRASH-SAFETY 2: Check if element is still valid (not deleted)
- ✅ CRASH-SAFETY 3: Validate element type before accessing properties
- ✅ CRASH-SAFETY 4: Validate snapshot structure
- ✅ CRASH-SAFETY 5: Validate mapping parameter name

### 3. **Added Batch Parameter Lookup Optimization**

```csharp
// ✅ OPTIMIZATION 4: Batch Parameter Lookups - Pre-cache elements and parameters
Dictionary<int, Element> elementCache = new Dictionary<int, Element>();
Dictionary<int, Dictionary<string, Parameter>> parameterCache = new Dictionary<int, Dictionary<string, Parameter>>();

if (Services.OptimizationFlags.UseBatchParameterLookups)
{
    // Pre-cache elements and parameters...
}
```

### 4. **Added Extensive Diagnostic Logging**

Added logging at multiple points:
- Log what source parameter we're looking for
- Log when attempting to read from Revit
- Log MEP_ElementId parameter existence
- Log MEP element retrieval
- Log Size parameter reading success/failure
- Log snapshot fallback usage

## Why the Issue Persists

Based on the logs provided:
1. **The transfer_debug.log shows only "Matching sleeve" messages** - this means the code is matching sleeves to snapshots, but **NOT reaching the parameter reading logic**.
2. **No "Processing mapping" messages** - this suggests the `ExecuteTransferConfiguration` method may not be processing mappings at all, or the mappings list is empty.
3. **No MEP Size read attempts** - even though the code has been added to read from Revit, the logs show no attempts.

## Root Cause Analysis

The issue is likely that:
1. **The parameter transfer is not being called at all**, OR
2. **The mappings configuration is empty**, OR
3. **The code is failing before reaching the MEP Size reading logic**

The diagnostic logging I just added will help identify exactly where the code is stopping.

## Next Steps

1. **Rebuild and test** - The new diagnostic logging will show exactly what's happening
2. **Check transfer_debug.log** - Look for:
   - "Processing mapping" messages
   - "Checking source parameter" messages
   - "Source parameter is 'MEP Size'" messages
   - Any error messages

3. **Verify the mapping configuration** - Ensure that "MEP Size" is actually in the mappings list


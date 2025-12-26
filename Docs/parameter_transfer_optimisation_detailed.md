# Parameter Transfer Optimisation: Full Implementation Guide with Code

## Introduction
This guide provides a comprehensive, code-rich walkthrough for refactoring and optimising parameter transfer, mark, and remark operations in the Revit add-in. It covers both `ParameterTransferService.cs` and `SleeveParameterService.cs` with before/after code samples, rationale, and step-by-step instructions.

---

## 1. Remove/Buffer Logging Inside Loops

### Problem
Logging (especially file I/O) inside loops causes massive slowdowns.

### Before
```csharp
foreach (var openingId in openingIds) {
    SafeFileLogger.SafeAppendText("transfer_debug.log", $"Processing {openingId}\n");
    // ...
}
```

### After
```csharp
var logBuffer = new StringBuilder();
foreach (var openingId in openingIds) {
    if (!DeploymentConfiguration.DeploymentMode) {
        logBuffer.AppendLine($"Processing {openingId}");
    }
    // ...
}
if (logBuffer.Length > 0) {
    SafeFileLogger.SafeAppendText("transfer_debug.log", logBuffer.ToString());
}
```

### Steps
1. Search for all `SafeFileLogger.SafeAppendText` and `DebugLogger.Info` calls inside loops.
2. Replace with `StringBuilder` buffer as above.
3. Flush buffer after loop or every N entries for very large batches.
4. Wrap all diagnostic logging with `if (!DeploymentConfiguration.DeploymentMode)`.

---

## 2. Implement Lazy Parameter Cache

### Problem
Eagerly looking up all parameters for all elements upfront is wasteful.

### Before
```csharp
foreach (var openingId in openingIds) {
    var element = doc.GetElement(openingId);
    var param = element.LookupParameter(mapping.TargetParameter);
    // ...
}
```

### After
```csharp
private Dictionary<string, Parameter> parameterCache = new();
private Parameter GetOrCacheParameter(Element element, string paramName) {
    var key = $"{element.Id.IntegerValue}_{paramName}";
    if (parameterCache.TryGetValue(key, out var cached)) return cached;
    var param = element.LookupParameter(paramName);
    if (param != null) parameterCache[key] = param;
    return param;
}
// Usage:
var param = GetOrCacheParameter(element, mapping.TargetParameter);
```

### Steps
1. Add a private cache dictionary and helper method as above.
2. Replace all direct `element.LookupParameter` calls with the cache method.
3. Clear the cache at the start/end of each batch operation.

---

## 3. Struct-Based Deferred Parameters

### Problem
Nested dictionaries with string keys and boxed objects are slow and error-prone.

### Before
```csharp
Dictionary<ElementId, Dictionary<string, object>> _deferredParameters;
// Usage:
_deferredParameters[elementId][paramName] = value;
```

### After
```csharp
struct ParameterValue {
    public string Name;
    public double? DoubleValue;
    public string StringValue;
    public int? IntValue;
    public ElementId ElementIdValue;
    public ParameterType Type;
}
Dictionary<int, List<ParameterValue>> _deferredParameters;
// Usage:
_deferredParameters[elementId.IntegerValue].Add(new ParameterValue {
    Name = paramName,
    StringValue = value,
    Type = ParameterType.String
});
```

### Steps
1. Define the `ParameterValue` struct.
2. Change the deferred parameters dictionary to use `int` keys and lists of `ParameterValue`.
3. Update all logic to use the struct fields directly.

---

## 4. Replace LINQ in Hot Paths

### Problem
LINQ in tight loops creates unnecessary allocations.

### Before
```csharp
var ids = items.Select(x => x.Id).ToList();
```

### After
```csharp
var ids = new List<int>();
foreach (var x in items) ids.Add(x.Id);
```

### Steps
1. Search for `.Select`, `.Where`, `.ToList`, `.Sum`, etc. in main loops.
2. Replace with explicit `foreach` loops.

---

## 5. Batch File I/O

### Problem
Multiple file writes per element processed.

### Before
```csharp
foreach (var item in items) {
    SafeFileLogger.SafeAppendText("transfer_debug.log", $"Log: {item}\n");
}
```

### After
```csharp
var logBuffer = new StringBuilder();
foreach (var item in items) {
    logBuffer.AppendLine($"Log: {item}");
}
SafeFileLogger.SafeAppendText("transfer_debug.log", logBuffer.ToString());
```

### Steps
1. Use a log buffer for all file writes in critical paths.
2. Flush buffer after processing or at regular intervals.

---

## 6. Unified Snapshot Lookup

### Problem
Multiple dictionary lookups per element due to fallback logic.

### Before
```csharp
if (snapshotIndex.TryGetByCombined(...)) { }
else if (clusterInstanceId > 0 && snapshotIndex.TryGetByCluster(...)) { }
else if (sleeveInstanceId > 0 && snapshotIndex.TryGetBySleeve(...)) { }
// ...
```

### After
```csharp
var unifiedLookup = new Dictionary<int, SleeveSnapshotView>();
foreach (var kvp in snapshotIndex.BySleeve) unifiedLookup[kvp.Key] = kvp.Value;
foreach (var kvp in snapshotIndex.ByCluster) unifiedLookup[kvp.Key] = kvp.Value;
// Usage:
if (unifiedLookup.TryGetValue(sleeveInstanceId, out var snapshot)) { ... }
```

### Steps
1. Build the unified lookup dictionary at the start of processing.
2. Replace all fallback chains with a single lookup.

---

## 7. Example: Refactored Main Loop

### Before
```csharp
foreach (var openingId in openingIds) {
    var element = doc.GetElement(openingId);
    var param = element.LookupParameter(mapping.TargetParameter);
    SafeFileLogger.SafeAppendText("transfer_debug.log", $"Processing {openingId}\n");
    // ...
}
```

### After
```csharp
var logBuffer = new StringBuilder();
foreach (var openingId in openingIds) {
    var element = doc.GetElement(openingId);
    var param = GetOrCacheParameter(element, mapping.TargetParameter);
    if (!DeploymentConfiguration.DeploymentMode) {
        logBuffer.AppendLine($"Processing {openingId}");
    }
    // ...
}
if (logBuffer.Length > 0) {
    SafeFileLogger.SafeAppendText("transfer_debug.log", logBuffer.ToString());
}
```

---

## 8. Final Testing and Profiling
1. Run the full parameter transfer process on a large model.
2. Profile CPU, memory, and file I/O before and after refactoring.
3. Compare log file sizes and number of file writes.
4. Document results and update this guide as needed.

---

*Prepared: 2025-12-24*

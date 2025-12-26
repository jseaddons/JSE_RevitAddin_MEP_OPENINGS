# Parameter Transfer Optimisation: Detailed Implementation Steps

## Introduction
This document provides a step-by-step guide for refactoring and optimising the parameter transfer, mark, and remark operations in the Revit add-in. The focus is on `ParameterTransferService.cs` and `SleeveParameterService.cs` to achieve significant performance improvements for large models.

---

## Step 1: Remove/Buffer Logging Inside Loops
1. **Identify all logging calls** (e.g., `SafeFileLogger.SafeAppendText`, `DebugLogger.Info`, etc.) inside main processing loops in both files.
2. **Replace per-iteration file writes** with appending to a `StringBuilder` log buffer.
3. **Flush the buffer** to file only after the loop, or every N entries for very large batches.
4. **Disable all diagnostic logging** in deployment mode (wrap with `if (!DeploymentConfiguration.DeploymentMode)`).
5. **Test**: Run with a large dataset and verify only a few file writes occur.

---

## Step 2: Implement Lazy Parameter Cache
1. **Locate parameter cache logic** in `ParameterTransferService.cs` (typically a dictionary for element/parameter pairs).
2. **Refactor cache to be lazy**:
    - Only perform `element.LookupParameter` when a parameter is first needed.
    - Use a composite key (element ID + parameter name) for the cache.
    - Example:
      ```csharp
      private Parameter GetOrCacheParameter(Element element, string paramName) {
          var key = $"{element.Id.IntegerValue}_{paramName}";
          if (parameterCache.TryGetValue(key, out var cached)) return cached;
          var param = element.LookupParameter(paramName);
          if (param != null) parameterCache[key] = param;
          return param;
      }
      ```
3. **Replace all direct lookups** with calls to the cache method.
4. **Test**: Profile to ensure parameter lookups are not repeated for the same element/parameter.

---

## Step 3: Struct-Based Deferred Parameters
1. **Find all uses of** `Dictionary<ElementId, Dictionary<string, object>>` in `SleeveParameterService.cs`.
2. **Define a struct** for parameter values:
    ```csharp
    struct ParameterValue {
        public string Name;
        public double DoubleValue;
        public string StringValue;
        public int IntValue;
        public ElementId ElementIdValue;
        public ParameterType Type;
    }
    ```
3. **Change the deferred parameters dictionary** to:
    ```csharp
    Dictionary<int, List<ParameterValue>> _deferredParameters;
    ```
4. **Update all logic** to use the struct fields instead of boxing/unboxing objects.
5. **Test**: Ensure all parameter types are handled and values are set correctly.

---

## Step 4: Replace LINQ in Hot Paths
1. **Search for all LINQ usage** in main processing loops (e.g., `.Select`, `.Where`, `.ToList`, `.Sum`).
2. **Replace with explicit loops** for performance-critical sections.
    - Example:
      ```csharp
      // Before
      var ids = items.Select(x => x.Id).ToList();
      // After
      var ids = new List<int>();
      foreach (var x in items) ids.Add(x.Id);
      ```
3. **Test**: Profile to confirm reduced allocations and improved speed.

---

## Step 5: Batch File I/O
1. **Create a log buffer** (e.g., `StringBuilder`) in both services.
2. **Replace all direct file writes** in loops with buffer appends.
3. **Flush the buffer** to file after processing or every N entries.
4. **Test**: Confirm that file I/O is minimized and logs are complete.

---

## Step 6: Unified Snapshot Lookup
1. **Identify all snapshot lookup chains** (multiple dictionary lookups per element).
2. **At the start of processing**, build a unified dictionary:
    ```csharp
    var unifiedLookup = new Dictionary<int, SleeveSnapshotView>();
    foreach (var kvp in snapshotIndex.BySleeve) unifiedLookup[kvp.Key] = kvp.Value;
    foreach (var kvp in snapshotIndex.ByCluster) unifiedLookup[kvp.Key] = kvp.Value;
    ```
3. **Replace all fallback chains** with a single lookup:
    ```csharp
    if (unifiedLookup.TryGetValue(sleeveInstanceId, out var snapshot)) { ... }
    ```
4. **Test**: Ensure all elements are found as expected and performance improves.

---

## Step 7: Final Testing and Profiling
1. **Run the full parameter transfer process** on a large model.
2. **Profile CPU and memory usage** before and after refactoring.
3. **Compare log file sizes and number of file writes**.
4. **Document results** in this folder for future reference.

---

## Notes
- Refactor incrementally and test after each step.
- Use deployment mode to verify minimal logging and maximum speed.
- Update documentation and code comments as you go.

---

*Prepared: 2025-12-24*

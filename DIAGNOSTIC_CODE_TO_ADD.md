# EXACT CODE TO ADD FOR DIAGNOSTICS

## FILE: RefactoredClusterService.cs

### STEP 1: Add this field at the top of the class (around line 100-120)

Find this section:
```csharp
public class RefactoredClusterService
{
    private const bool ALLOW_MODIFICATIONS_TO_PROTECTED_CODE = false;
    
    // Phase 1-5 Services
    private readonly ClusteringStrategyFactory _strategyFactory;
    private readonly IClusterPlacementService _placementService;
```

ADD THIS LINE after the const:
```csharp
private int _clusterPlacementCallCount = 0;  // ✅ ADD THIS FOR DIAGNOSTICS
```

So it looks like:
```csharp
public class RefactoredClusterService
{
    private const bool ALLOW_MODIFICATIONS_TO_PROTECTED_CODE = false;
    private int _clusterPlacementCallCount = 0;  // ✅ DIAGNOSTIC FIELD
    
    // Phase 1-5 Services
    private readonly ClusteringStrategyFactory _strategyFactory;
    private readonly IClusterPlacementService _placementService;
```

### STEP 2: Find ALL calls to PlaceClusterSleeve

Use Ctrl+F to search for: `PlaceClusterSleeve(`

You should find multiple occurrences. For EACH occurrence, add the diagnostic line BEFORE it.

### EXAMPLE 1: If you find this:

```csharp
bool placed = _placementService.PlaceClusterSleeve(
    doc, 
    cluster, 
    groupKey,
    // ... more parameters ...
);
```

CHANGE IT TO:
```csharp
_clusterPlacementCallCount++;
_logger?.Invoke($"[DIAGNOSTIC] 🔥 PLACEMENT CALL #{_clusterPlacementCallCount} in method: {System.Reflection.MethodBase.GetCurrentMethod().Name}");

bool placed = _placementService.PlaceClusterSleeve(
    doc, 
    cluster, 
    groupKey,
    // ... more parameters ...
);
```

### EXAMPLE 2: If you find this inside a loop:

```csharp
foreach (var cluster in clusters)
{
    // ... some code ...
    
    var result = _placementService.PlaceClusterSleeve(
        doc, cluster, groupKey, category, placementPoint,
        width, height, depth, rotationAngle, xmlFilePath, familyName,
        out placedSleeve, out sleeveId, out actualPoint, out saveData,
        deferredParams, hostOrientation, mepAngle);
    
    // ... more code ...
}
```

CHANGE IT TO:
```csharp
foreach (var cluster in clusters)
{
    // ... some code ...
    
    _clusterPlacementCallCount++;  // ✅ ADD THIS
    var methodName = System.Reflection.MethodBase.GetCurrentMethod().Name;  // ✅ ADD THIS
    _logger?.Invoke($"[DIAGNOSTIC] 🔥 PLACEMENT CALL #{_clusterPlacementCallCount} in method: {methodName}");  // ✅ ADD THIS
    
    var result = _placementService.PlaceClusterSleeve(
        doc, cluster, groupKey, category, placementPoint,
        width, height, depth, rotationAngle, xmlFilePath, familyName,
        out placedSleeve, out sleeveId, out actualPoint, out saveData,
        deferredParams, hostOrientation, mepAngle);
    
    // ... more code ...
}
```

### STEP 3: Compile and Run

After adding the diagnostic code:
1. Build the solution (Ctrl+Shift+B)
2. Run the tool in Revit
3. Check the logs

### STEP 4: Read the Diagnostic Output

You'll see output like this in your logs:

```
[DIAGNOSTIC] 🔥 PLACEMENT CALL #1 in method: ProcessClusterGroup
[... cluster placement happens ...]

[DIAGNOSTIC] 🔥 PLACEMENT CALL #2 in method: ProcessClusterGroup
[... cluster placement happens AGAIN for same clusters! ...]
```

OR:

```
[DIAGNOSTIC] 🔥 PLACEMENT CALL #1 in method: PlaceClusterBatch
[... first placement: 413ms ...]

[DIAGNOSTIC] 🔥 PLACEMENT CALL #2 in method: BulkPlacementClusters
[... second placement: 880ms - THIS IS THE DUPLICATE! ...]
```

### STEP 5: Identify the Duplicate

The SECOND call (and any subsequent calls) are the duplicates.

Look at the method name shown in the diagnostic output, then:
1. Find that method in the code
2. Comment out or delete the PlaceClusterSleeve call in that method
3. Recompile and test

### COMMON SCENARIOS

**Scenario A: Both calls in same method**
```csharp
[DIAGNOSTIC] 🔥 PLACEMENT CALL #1 in method: ClusterSleevesV2
[DIAGNOSTIC] 🔥 PLACEMENT CALL #2 in method: ClusterSleevesV2
```
→ The method has TWO loops or TWO calls to PlaceClusterSleeve
→ Remove the second one

**Scenario B: Calls in different methods**
```csharp
[DIAGNOSTIC] 🔥 PLACEMENT CALL #1 in method: ProcessIndividualClusters
[DIAGNOSTIC] 🔥 PLACEMENT CALL #2 in method: BulkProcessClusters
```
→ Two different methods are placing the same clusters
→ Remove the call in method #2 (usually the "bulk" or "batch" method)

---

## QUICK REFERENCE: SEARCH PATTERNS

Use these Ctrl+F searches in `RefactoredClusterService.cs`:

1. Search: `.PlaceClusterSleeve(` - Find all placement calls
2. Search: `"Place Cluster Instances"` - Find first placement log
3. Search: `"BULK PLACEMENT"` - Find second placement log
4. Search: `413ms` - Find first timing reference
5. Search: `880ms` - Find second timing reference

---

## AFTER YOU FIND IT

Once you identify the duplicate:

### To Comment It Out (Safe Testing):
```csharp
// ❌ DUPLICATE PLACEMENT DETECTED AND DISABLED
// This was causing clusters to be placed twice
// Original timing: 880ms overhead per batch
// _logger?.Invoke("Step 8: BULK PLACEMENT - CLUSTERS");
// var result = BulkPlaceClusters(clusters);
```

### To Delete It (Permanent Fix):
Just delete the entire method call and any associated logging.

---

## VERIFICATION

After fixing, your logs should show:
```
✅ [DIAGNOSTIC] 🔥 PLACEMENT CALL #1 in method: ProcessClusterGroup
   (This is the only call - no #2, no #3)
```

And performance should improve:
```
BEFORE: Total cluster time ~1300ms (4.3 sleeves/sec)
AFTER:  Total cluster time ~413ms (12+ sleeves/sec)
        2.8x improvement ✅
```

---

## STILL STUCK?

If you add all the diagnostics and run it, but still can't figure out which call to remove:

1. Copy the ENTIRE section of your log that shows:
   ```
   [DIAGNOSTIC] 🔥 PLACEMENT CALL #1 ...
   [DIAGNOSTIC] 🔥 PLACEMENT CALL #2 ...
   ```

2. Copy the method names shown

3. Share those and I can tell you exactly which lines to comment out or delete.

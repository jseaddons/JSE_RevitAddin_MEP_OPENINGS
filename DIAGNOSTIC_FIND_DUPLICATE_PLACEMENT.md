# DIAGNOSTIC GUIDE: Finding Duplicate Cluster Placement

## ✅ ISSUE #1 STATUS: ALREADY FIXED
The `BulkPlacementService.cs` already has the transaction wrapper in place (lines 91-135).
Your Step 5 issue should already be resolved.

## 🔍 ISSUE #2: Finding the Duplicate Cluster Placement

Based on your logs showing:
- "Place Cluster Instances: 413ms"
- "Step 8: BULK PLACEMENT - CLUSTERS: 880ms"

There are TWO separate placement operations happening.

## WHERE THE PLACEMENT HAPPENS

The actual Revit API call is in `ClusterPlacementService.cs` line ~300:
```csharp
inst = doc.Create.NewFamilyInstance(placementPoint, familySymbol, refLevel, StructuralType.NonStructural);
```

This service is being called TWICE from somewhere in `RefactoredClusterService.cs`.

## HOW TO FIND THE DUPLICATE

### Step 1: Add Diagnostic Logging

Add this field at the top of `RefactoredClusterService.cs` (around line 100):

```csharp
// ADD THIS FIELD
private int _clusterPlacementCallCount = 0;
```

### Step 2: Find All Calls to PlaceClusterSleeve

Search for `PlaceClusterSleeve` in `RefactoredClusterService.cs` and add logging BEFORE each call:

```csharp
// BEFORE THIS LINE:
bool placed = _placementService.PlaceClusterSleeve(

// ADD THIS:
_clusterPlacementCallCount++;
_logger?.Invoke($"[DIAGNOSTIC] 🔥 PLACEMENT CALL #{_clusterPlacementCallCount}: About to call PlaceClusterSleeve");

// THEN THE ORIGINAL CALL:
bool placed = _placementService.PlaceClusterSleeve(
```

### Step 3: Search Patterns to Look For

In `RefactoredClusterService.cs`, search for these patterns:

**Pattern A: Loop calling PlaceClusterSleeve**
```csharp
foreach (var cluster in clusters)
{
    // ... code ...
    _placementService.PlaceClusterSleeve(...)  // First call
}

// Later in the same method...
foreach (var cluster in clusters)  // DUPLICATE LOOP
{
    // ... code ...
    _placementService.PlaceClusterSleeve(...)  // Second call - DELETE THIS
}
```

**Pattern B: Method called twice**
```csharp
// First placement
var result1 = SomeMethod_That_Calls_PlaceClusterSleeve();  // 413ms

// Later...
var result2 = AnotherMethod_That_Calls_PlaceClusterSleeve();  // 880ms - DELETE THIS
```

**Pattern C: Batch methods**
```csharp
// Look for methods named like:
- PlaceClusterBatch()
- BulkPlaceClusters()
- ProcessClusterGroup()
- PlaceClusterGroup()

// One of these is being called twice
```

### Step 4: Specific Places to Check

1. **Check ClusterSleevesV2() method** - Main orchestration method
2. **Search for "Place Cluster Instances"** - This is the log message for the first call (413ms)
3. **Search for "BULK PLACEMENT - CLUSTERS"** - This is the log message for the second call (880ms)

### Step 5: Quick Visual Search

Open `RefactoredClusterService.cs` and use Ctrl+F to search for:
- `PlaceClusterSleeve` - Count how many times this appears
- `413ms` or `Place Cluster Instances` - Find first placement
- `880ms` or `BULK PLACEMENT` - Find second placement

## EXPECTED FINDINGS

You should find TWO distinct code paths that both call `PlaceClusterSleeve`:

1. **First Path (413ms)**: Likely in a method like `PlaceIndividualClusters()`
2. **Second Path (880ms)**: Likely in a method like `BulkPlaceClusters()`

One of these needs to be removed or commented out.

## LIKELY ROOT CAUSE

This is probably a refactoring artifact where:
- Old code: Had a single placement method
- New code: Added a "bulk" placement method for optimization
- Bug: Both methods are still being called

## THE FIX

Once you find the duplicate, you have two options:

### Option A: Comment Out Duplicate (Safe)
```csharp
// ❌ DUPLICATE PLACEMENT - DISABLED
// This is causing clusters to be placed twice
// var result = BulkPlaceClusters(clusters);  // 880ms
```

### Option B: Delete Duplicate (Permanent)
Just delete the redundant method call entirely.

## VERIFICATION AFTER FIX

After removing the duplicate, run the tool and check logs:
- Should see ONE placement log (not two)
- Cluster time should drop from ~1300ms to ~413ms
- Cluster rate should improve from 4.3/sec to 10+/sec

## IF YOU CAN'T FIND IT

If you still can't locate the duplicate, add this comprehensive logging:

```csharp
// At the START of any method that processes clusters:
_logger?.Invoke($"[TRACE] Entering method: {System.Reflection.MethodBase.GetCurrentMethod().Name}");

// BEFORE any call to PlaceClusterSleeve:
_logger?.Invoke($"[TRACE] About to place cluster in method: {System.Reflection.MethodBase.GetCurrentMethod().Name}");
```

This will show you the call stack and which methods are calling PlaceClusterSleeve.

## NEED MORE HELP?

If you run the diagnostic logging and share the output, I can pinpoint exactly which lines to remove.

The key is to find where `_placementService.PlaceClusterSleeve(...)` is being called twice for the same clusters.

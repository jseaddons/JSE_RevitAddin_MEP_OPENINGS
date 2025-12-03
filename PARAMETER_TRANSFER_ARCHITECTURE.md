# Parameter Transfer Service - Architecture & Flow Documentation

## Overview

The `ParameterTransferService` is responsible for transferring MEP element parameters (e.g., "Size", "System Type", "System Abbreviation") to sleeve openings in Revit. It supports both **individual sleeves** and **cluster sleeves**, with different data source strategies for each.

---

## Entry Point: `ExecuteTransferConfigurationInTransaction`

### Purpose
Main entry point for parameter transfer operations. Called from the UI dialog (`ParameterServiceDialogV2`) when the user clicks "Transfer Parameters".

### Signature
```csharp
public ParameterTransferResult ExecuteTransferConfigurationInTransaction(
    Document doc,
    List<ElementId> openingIds,
    ParameterTransferConfiguration config,
    UIDocument uiDoc = null)
```

### Flow

1. **Validation**
   - Checks if `openingIds` is null or empty
   - Returns error if no sleeves found

2. **Load Snapshot Index**
   - Opens SQLite database connection via `SleeveDbContext`
   - Loads `SleeveSnapshotIndex` from `SleeveSnapshotRepository`
   - `SleeveSnapshotIndex` contains:
     - `BySleeve`: Dictionary<int, SleeveSnapshotView> (keyed by `SleeveInstanceId`)
     - `ByCluster`: Dictionary<int, SleeveSnapshotView> (keyed by `ClusterInstanceId`)
   - If no snapshots found, checks if sleeves exist in database and provides helpful error message

3. **Execute Each Mapping**
   - Iterates through `config.Mappings` (list of `ParameterMapping` objects)
   - **✅ SKIP disabled mappings:** Checks `mapping.IsEnabled` before processing
   - For each enabled mapping, routes to appropriate transfer method based on `TransferType`:
     - `TransferType.ReferenceToOpening` → `TransferFromReferenceElementsInTransaction()` → `TransferFromElementsWithSnapshot(useHost: false)`
     - `TransferType.HostToOpening` → `TransferFromHostElementsInTransaction()` → `TransferFromElementsWithSnapshot(useHost: true)`
     - `TransferType.LevelToOpening` → `TransferFromLevelsInTransaction()` (different method, doesn't use snapshots)
   - Aggregates results from all mappings

4. **Return Result**
   - Combines all individual mapping results
   - Returns `ParameterTransferResult` with overall success/failure status

---

## Core Method: `TransferFromElementsWithSnapshot`

### Purpose
Transfers a single parameter mapping from source (MEP elements/snapshots) to target (sleeve openings).

### Signature
```csharp
private ParameterTransferResult TransferFromElementsWithSnapshot(
    Document doc,
    List<ElementId> openingIds,
    ParameterMapping mapping,
    SleeveSnapshotIndex snapshotIndex,
    bool useHost = false)
```

### Key Parameters
- `doc`: Active Revit document (sleeves are always in active document)
- `openingIds`: List of sleeve element IDs to process
- `mapping`: Single parameter mapping (e.g., "Size" → "MEP Size")
- `snapshotIndex`: Pre-loaded snapshot index from database
- `useHost`: If `true`, reads from `HostParameters`; if `false`, reads from `MepParameters`
- `successfullyTransferredSleeveIds`: HashSet to track unique sleeves that had at least one parameter transferred (for final count)

### TransferType Routing

The `ExecuteTransferConfigurationInTransaction` method routes mappings to different transfer methods based on `mapping.TransferType`:

- **`TransferType.ReferenceToOpening`**: Transfers from MEP elements (reference elements)
  - Calls `TransferFromReferenceElementsInTransaction()`
  - Delegates to `TransferFromElementsWithSnapshot(useHost: false)`
  - Uses `MepParameters` from snapshot

- **`TransferType.HostToOpening`**: Transfers from host elements (walls, floors, etc.)
  - Calls `TransferFromHostElementsInTransaction()`
  - Delegates to `TransferFromElementsWithSnapshot(useHost: true)`
  - Uses `HostParameters` from snapshot

- **`TransferType.LevelToOpening`**: Transfers from level elements
  - Calls `TransferFromLevelsInTransaction()`
  - Does NOT use snapshots (reads directly from level elements)
  - Different implementation path

---

## Architecture Flow Diagram

```
┌─────────────────────────────────────────────────────────────────┐
│ ExecuteTransferConfigurationInTransaction                       │
│                                                                 │
│ 1. Validate openingIds                                         │
│ 2. Load SleeveSnapshotIndex from SQLite                        │
│ 3. For each ParameterMapping in config:                        │
│    └─> TransferFromElementsWithSnapshot()                       │
│         │                                                        │
│         ├─> [OPTIONAL] Batch Parameter Cache Init              │
│         │   (if UseBatchParameterLookups = true)                │
│         │                                                        │
│         └─> FOREACH LOOP: Process each sleeve                   │
│              │                                                  │
│              ├─> STEP 1: Validate & Retrieve Opening Element   │
│              │   - Check ElementId validity                     │
│              │   - Check document validity                      │
│              │   - Retrieve element from doc                    │
│              │   - Validate element type (FamilyInstance)      │
│              │                                                  │
│              ├─> STEP 2: Read Sleeve/Cluster IDs                 │
│              │   - Read "Sleeve Instance ID" parameter          │
│              │   - Read "Cluster Sleeve Instance ID" parameter  │
│              │   - Determine if cluster sleeve (clusterId > 0) │
│              │                                                  │
│              ├─> STEP 3: Match Snapshot                         │
│              │   - If clusterId > 0: Try GetByCluster()         │
│              │   - Else if sleeveId > 0: Try GetBySleeve()     │
│              │   - If no snapshot found: SKIP (continue)        │
│              │                                                  │
│              ├─> STEP 4: Load Source Parameters                  │
│              │   - Get sourceParams from snapshot               │
│              │     (useHost ? HostParameters : MepParameters)   │
│              │   - Validate sourceParams not null/empty         │
│              │                                                  │
│              ├─> STEP 5: Read Source Parameter Value            │
│              │   │                                            │
│              │   ├─> IF "MEP Size" or "Size":                  │
│              │   │   │                                        │
│              │   │   ├─> IF Cluster Sleeve:                   │
│              │   │   │   - Read from snapshot (aggregated)   │
│              │   │   │   - Try "Size" or "MEP Size"           │
│              │   │   │                                        │
│              │   │   └─> ELSE (Individual Sleeve):           │
│              │   │       - Read MEP_ElementId from sleeve      │
│              │   │       - Get MEP element from doc/linked     │
│              │   │       - Read "Size" param from MEP element  │
│              │   │       - Fallback to snapshot if MEP fails  │
│              │   │                                            │
│              │   └─> ELSE (Other Parameters):                 │
│              │       - Try exact match in snapshot             │
│              │       - Try variations (space/underscore)       │
│              │       - For Cable Trays: Try "Service Type"     │
│              │       - If not found: Try Revit MEP element     │
│              │         (for individual sleeves only)           │
│              │                                                  │
│              ├─> STEP 6: Validate Source Value                 │
│              │   - Check sourceValue not null/empty           │
│              │   - If empty: SKIP (continue)                   │
│              │                                                  │
│              ├─> STEP 7: Lookup Target Parameter               │
│              │   - If UseBatchParameterLookups: Check cache     │
│              │   - Else: LookupParameter() on opening          │
│              │   - Validate parameter belongs to correct elem  │
│              │                                                  │
│              ├─> STEP 8: Skip Check (Optimization)              │
│              │   - If SkipAlreadyTransferredParameters:        │
│              │     Check if value already matches               │
│              │   - If matches: SKIP (continue)                 │
│              │                                                  │
│              └─> STEP 9: Set Parameter Value                   │
│                  - Call SetParameterValueSafely()               │
│                  - Increment transferredCount or failedCount    │
│                                                                 │
└─────────────────────────────────────────────────────────────────┘
```

---

## Detailed Step-by-Step Flow

### STEP 1: Validate & Retrieve Opening Element

**Validations:**
1. `openingId` is not null and `IntegerValue > 0`
2. `doc` is not null
3. `doc.IsModifiable == true` and `doc.IsReadOnly == false`
4. Element retrieved successfully (`doc.GetElement(openingId)`)
5. Element is valid (`opening.IsValidObject`)
6. Element is a `FamilyInstance`

**Note:** Document validation check was **removed** (see "Common Issues" section for bug history). Since we retrieve the element using `doc.GetElement(openingId)`, it MUST be in `doc`, making the check redundant and harmful.

**Early Exit:** If any validation fails, log error and `continue` to next sleeve.

---

### STEP 2: Read Sleeve/Cluster IDs

**Parameters Read:**
- `"Sleeve Instance ID"` → `sleeveInstanceId` (int)
- `"Cluster Sleeve Instance ID"` → `clusterInstanceId` (int)

**Determination:**
- `isClusterSleeve = (clusterInstanceId > 0)`

**Purpose:** These IDs are used to match the sleeve to its snapshot in the database.

---

### STEP 3: Match Snapshot

**Matching Strategy:**
1. **If `clusterInstanceId > 0`:**
   - Try `snapshotIndex.TryGetByCluster(clusterInstanceId, out snapshot)`
   - Cluster sleeves aggregate multiple MEP elements, so snapshot contains aggregated data

2. **Else if `sleeveInstanceId > 0`:**
   - Try `snapshotIndex.TryGetBySleeve(sleeveInstanceId, out snapshot)`
   - Individual sleeves have a single MEP element

3. **If no snapshot found:**
   - Log warning with available IDs
   - `continue` to next sleeve

**Snapshot Structure:**
- `SleeveSnapshotView` contains:
  - `MepParameters`: Dictionary<string, string> (MEP element parameters)
  - `HostParameters`: Dictionary<string, string> (host element parameters)
  - `SleeveInstanceId`: int (for individual sleeves)
  - `ClusterInstanceId`: int? (for cluster sleeves)

---

### STEP 4: Load Source Parameters

**Source Selection:**
```csharp
var sourceParams = useHost ? snapshot.HostParameters : snapshot.MepParameters;
```

**Validation:**
- `sourceParams` must not be null
- `sourceParams.Count > 0`

**Early Exit:** If validation fails, log warning and `continue`.

---

### STEP 5: Read Source Parameter Value

This is the **most complex step** with different logic for different parameter types and sleeve types.

#### 5A. Special Case: "MEP Size" or "Size"

**For Cluster Sleeves:**
- **Always use snapshot** (aggregated data)
- Try `sourceParams.TryGetValue("Size", out sourceValue)`
- If not found, try `sourceParams.TryGetValue("MEP Size", out sourceValue)`
- If found and not empty: `readFromRevit = false`
- If not found or empty: Log warning and `continue`

**For Individual Sleeves:**
- **Primary source: Revit MEP element**
- Read `MEP_ElementId` parameter from sleeve
- Get MEP element using `ElementRetrievalService.GetElementFromDocumentOrLinked()` (handles linked documents)
- Read `"Size"` parameter from MEP element
- If successful: `readFromRevit = true`
- **Fallback:** If MEP element not found or invalid, try snapshot

#### 5B. Other Parameters (e.g., "System Type", "System Abbreviation")

**Strategy:**
1. **Try exact match** in snapshot (case-insensitive)
2. **Try variations:**
   - `"System Type"` → `"System_Type"` (space to underscore)
   - `"System_Type"` → `"System Type"` (underscore to space)
   - `"System Type"` → `"MEP System Type"` (add "MEP " prefix)
   - `"MEP System Type"` → `"System Type"` (remove "MEP " prefix)
3. **Cable Trays special case:**
   - If category is "Cable Trays" and source is "System Type":
     - Prioritize `"Service Type"` variation
4. **Fallback (individual sleeves only):**
   - If not in snapshot, try reading from Revit MEP element
   - Read `MEP_ElementId` from sleeve
   - Get MEP element and read parameter directly

**For Cluster Sleeves:**
- **No Revit fallback** - snapshot is the only source
- If not found in snapshot: Log warning and `continue`

---

### STEP 6: Validate Source Value

**Validation:**
- `sourceValue` must not be null or empty/whitespace

**Early Exit:** If empty, log warning and `continue`.

---

### STEP 7: Lookup Target Parameter

**Strategy:**

**If `UseBatchParameterLookups = true`:**
1. Check `parameterCache[openingId.IntegerValue][mapping.TargetParameter]`
2. **Validate cached parameter:**
   - Parameter element is not null
   - Parameter element is valid
   - Parameter element ID matches opening ID
   - Parameter definition is valid
   - Parameter name matches target parameter name
3. If validation fails: Clear cache entry and lookup fresh

**Else:**
- Call `opening.LookupParameter(mapping.TargetParameter)`

**Validation:**
- Target parameter must exist
- If missing: Log error (critical) or warning (non-critical) and `continue`

**Critical Parameters:**
- `"MEP_ElementId"`
- `"MEP System Type"`
- `"System Type"`

---

### STEP 8: Skip Check (Optimization)

**If `SkipAlreadyTransferredParameters = true`:**
- Compare current target parameter value with `sourceValue`
- If they match: Skip transfer (log and `continue`)
- **Purpose:** Avoid unnecessary parameter writes

---

### STEP 8.5: Read-Only Parameter Check

**Validation:**
- Check if `targetParam.IsReadOnly == true`
- If read-only: Increment `failedCount`, add error, and `continue`

**Purpose:** Prevent attempting to set read-only parameters (would fail anyway, but this provides clearer error message)

**Early Exit:** If read-only, log error and `continue`.

---

### STEP 8.6: Source Value Empty Check (Safety Net)

**Validation:**
- Check if `sourceValue` is null or empty/whitespace
- If empty: Add warning and `continue`

**Purpose:** Safety net check before parameter setting (duplicate of STEP 6, but ensures value is still valid after all processing)

**Note:** This is a defensive check - `sourceValue` should already be validated in STEP 6, but this ensures it hasn't become empty due to any intermediate processing.

---

### STEP 9: Final Element Validation

**Validation:**
- Check if `opening.IsValidObject == true`
- Check if `targetParam.Element.IsValidObject == true`
- If either is invalid: Increment `failedCount`, add error, and `continue`

**Purpose:** Final safety check before parameter setting to ensure elements haven't been deleted or become invalid during processing

**Early Exit:** If invalid, log error and `continue`.

---

### STEP 9.5: Set Parameter Value

**Method:** `SetParameterValueSafely(targetParam, sourceValue)`

**Implementation:**
- Handles different parameter storage types (String, Integer, Double, ElementId, etc.)
- Uses appropriate `Parameter.Set()` overload
- Returns `true` if successful, `false` otherwise

**Result:**
- If successful: Increment `transferredCount`
- If failed: Increment `failedCount` and add error message

---

## Key Design Decisions

### 1. Cluster vs. Individual Sleeve Handling

**Cluster Sleeves:**
- Always use snapshot (aggregated data from multiple MEP elements)
- No Revit MEP element lookup
- Reason: Cluster sleeves aggregate multiple MEP elements, so there's no single `MEP_ElementId`

**Individual Sleeves:**
- For "MEP Size": Primary source is Revit MEP element (current value)
- For other parameters: Snapshot first, Revit fallback
- Reason: Ensures "MEP Size" is always current, while other parameters can use snapshot for performance

### 2. MEP Size Special Handling

**Why read from Revit?**
- "MEP Size" can change in the MEP element after sleeve placement
- Snapshot may contain stale data
- Reading from Revit ensures current value

**Why snapshot for clusters?**
- Cluster sleeves aggregate multiple MEP elements
- No single `MEP_ElementId` to read from
- Snapshot contains aggregated size data

### 3. Parameter Name Variations

**Why needed?**
- Revit parameter names may use spaces or underscores inconsistently
- Database snapshots may store different variations
- Example: "System Type" vs. "System_Type" vs. "MEP System Type"

**Variation Strategy:**
- Try exact match first
- Try common variations (space ↔ underscore, add/remove "MEP " prefix)
- For Cable Trays: Prioritize "Service Type" → "System Type" mapping

### 4. Batch Parameter Lookup Optimization

**Purpose:** Reduce repeated `LookupParameter()` calls

**Implementation:**
- Pre-cache all target parameters for all sleeves before loop
- Cache validation ensures parameters are still valid
- Falls back to fresh lookup if cache is stale

**When Enabled:**
- `OptimizationFlags.UseBatchParameterLookups = true`

**When Disabled:**
- Each sleeve performs fresh `LookupParameter()` call

### 5. Skip Already Transferred Parameters

**Purpose:** Avoid unnecessary parameter writes

**Implementation:**
- Compare current target parameter value with source value
- If they match: Skip transfer

**When Enabled:**
- `OptimizationFlags.SkipAlreadyTransferredParameters = true`

---

## Error Handling

### Exception Handling

**Outer Try-Catch:**
- Wraps entire `foreach` loop
- Catches any unexpected exceptions
- **Comprehensive Logging:**
  - Logs exception type name (`ex.GetType().Name`)
  - Logs exception message (`ex.Message`)
  - Logs full stack trace (`ex.StackTrace`)
  - All logged to `transfer_debug.log` with timestamps
- Increments `failedCount` and adds error message

**Inner Validations:**
- Each validation step has early exit with `continue`
- Logs specific reason for skip/failure
- Distinguishes between errors (critical) and warnings (non-critical)

### Error Classification

**Errors (Critical):**
- Missing critical target parameters (`MEP_ElementId`, `MEP System Type`)
- Element validation failures
- Document mismatches
- Read-only parameters
- Invalid elements before setting

**Warnings (Non-Critical):**
- Missing optional parameters (`MEP Size` if not critical)
- Parameter not found in snapshot (with fallback available)
- Empty parameter values
- Missing non-critical target parameters

### Success Criteria

**Success Determination:**
- `result.Success = (errors.Count == 0)`
- **NOT** based on `failedCount == 0`
- Warnings do NOT affect success status
- Only critical errors (added to `errors` list) affect success

**Result Counts:**
- `TransferredCount`: Number of **unique sleeves** that had at least one parameter successfully transferred (uses `successfullyTransferredSleeveIds` HashSet)
- `FailedCount`: Number of parameter transfers that failed
- `SkippedCount`: Number of parameters skipped (already transferred or optimization)
- **Note:** `TransferredCount` represents unique sleeves, not total parameter transfers

---

## Data Sources

### 1. Database Snapshots (`SleeveSnapshots` table)

**When Created:**
- During sleeve placement (individual sleeves)
- During cluster placement (cluster sleeves)
- During Refresh operation (recreates snapshots from existing sleeves)

**Contents:**
- `MepParameters`: Dictionary of MEP element parameters (JSON)
- `HostParameters`: Dictionary of host element parameters (JSON)
- `SleeveInstanceId`: For individual sleeves
- `ClusterInstanceId`: For cluster sleeves
- `ClashZoneGuid`: Deterministic GUID for matching

**Indexing:**
- `SleeveSnapshotIndex.BySleeve`: Keyed by `SleeveInstanceId`
- `SleeveSnapshotIndex.ByCluster`: Keyed by `ClusterInstanceId`

### 2. Live Revit MEP Elements

**When Used:**
- Primary source for "MEP Size" on individual sleeves
- Fallback for other parameters if not in snapshot (individual sleeves only)

**Retrieval:**
- Read `MEP_ElementId` parameter from sleeve
- Use `ElementRetrievalService.GetElementFromDocumentOrLinked()` to handle linked documents

**Limitations:**
- Not available for cluster sleeves (no single `MEP_ElementId`)
- May be slower than snapshot (requires Revit API calls)

---

## Performance Optimizations

### 1. Batch Parameter Lookup

**Benefit:** Reduces `LookupParameter()` calls from N (per sleeve) to 1 (pre-cached)

**Trade-off:** Requires memory for cache, but negligible for typical sleeve counts

### 2. Skip Already Transferred Parameters

**Benefit:** Avoids unnecessary parameter writes

**Trade-off:** Requires parameter read to compare values, but saves write operations

### 3. Snapshot Index (Pre-loaded)

**Benefit:** Fast O(1) lookup instead of database query per sleeve

**Trade-off:** Requires memory for index, but much faster than per-sleeve queries

---

## Logging

### Log File: `transfer_debug.log`

**Log Levels:**
- `✅`: Success/confirmation
- `🔍`: Diagnostic/search operation
- `⚠️`: Warning (non-critical)
- `❌`: Error (critical)
- `⏭️`: Skip (optimization)

**Key Log Markers:**
- `METHOD ENTRY`: Method called
- `Loop iteration X/Y`: Processing sleeve X of Y
- `SKIP`: Early exit reason
- `SUCCESS`: Parameter transferred successfully
- `EXCEPTION`: Exception caught and logged

**Build Stamp:**
- Logs DLL build timestamp to verify correct version is running

---

## Transaction Management

### Transaction Scope

**Assumption:** Caller (UI dialog) manages transaction

**Method Signature:**
- `ExecuteTransferConfigurationInTransaction()` assumes transaction is already started
- Does not call `Transaction.Start()` or `Transaction.Commit()`

**UI Dialog Responsibility:**
```csharp
using (var transaction = new Transaction(_document, "Transfer Parameters to Sleeves"))
{
    transaction.Start();
    result = transferService.ExecuteTransferConfigurationInTransaction(_document, openings, config, _uiDocument);
    transaction.Commit();
}
```

---

## Testing & Debugging

### Diagnostic Logging

**Enable Full Logging:**
- Set `DeploymentConfiguration.DeploymentMode = false`
- Set `OptimizationFlags.UseBatchParameterLookups = false` (to see all lookups)

### Common Issues

1. **"No snapshot found"**
   - Cause: Sleeve not in database or snapshot not created
   - Solution: Run Refresh to create snapshots

2. **"MEP element not found"**
   - Cause: `MEP_ElementId` is invalid or element was deleted
   - Solution: Check if MEP element exists, re-run Refresh

3. **"Parameter not found in snapshot"**
   - Cause: Parameter wasn't captured during placement/refresh
   - Solution: Check snapshot contents, verify parameter exists on MEP element

4. **"Document mismatch" (FIXED - 2025-12-03)**
   - **Bug History:** A "PROTECTION 2" crash safety check was added that validated `opening.Document != doc` using reference equality
   - **Problem:** Document reference equality fails even for the same document (different object instances)
   - **Impact:** Caused ALL sleeves to be skipped with false "document mismatch" errors, wasting 1 day of debugging
   - **Root Cause:** The check used `opening.Document != doc` which compares object references, not actual document identity
   - **Fix:** Removed the check entirely because:
     - We retrieve elements using `doc.GetElement(openingId)`, so they MUST be in `doc`
     - Element ID validation later (STEP 7) already ensures parameter belongs to correct element
     - The check was redundant and harmful
   - **Lesson Learned:** Not all "crash safety" checks are beneficial - some can introduce bugs if they use incorrect comparison methods
   - **Protection:** Code now has explicit comment warning against reintroducing this check

---

## Future Enhancements

### Potential Improvements

1. **Parallel Processing:**
   - Process multiple sleeves in parallel (with transaction safety)
   - Benefit: Faster for large batches

2. **Incremental Snapshot Updates:**
   - Update snapshots when MEP elements change
   - Benefit: Always current data without full refresh

3. **Parameter Validation:**
   - Validate parameter types match (source vs. target)
   - Benefit: Catch type mismatches early

4. **Batch Parameter Writes:**
   - Defer parameter writes and flush in batch
   - Benefit: Reduce Revit API overhead

---

## Related Files

- `Services/ParameterTransferService.cs`: Main service implementation
- `Views/ParameterServiceDialogV2.cs`: UI dialog (entry point)
- `Data/Repositories/SleeveSnapshotRepository.cs`: Database snapshot operations
- `Models/ParameterTransferResult.cs`: Result model
- `Models/ParameterMapping.cs`: Mapping configuration model
- `Services/OptimizationFlags.cs`: Feature flags for optimizations
- `Services/DeploymentConfiguration.cs`: Deployment mode configuration

---

## Summary

The parameter transfer service uses a **two-tier data source strategy**:
1. **Database snapshots** (fast, pre-aggregated for clusters)
2. **Live Revit elements** (current, for "MEP Size" on individual sleeves)

It handles **cluster sleeves** and **individual sleeves** differently, with cluster sleeves always using snapshot data and individual sleeves preferring live Revit data for "MEP Size".

The architecture is designed for **performance** (batch lookups, skip checks) and **reliability** (comprehensive validation, error handling, diagnostic logging).

### Key Safety Features

1. **Multiple Validation Layers:**
   - Element validation at multiple points (retrieval, before parameter access, before setting)
   - Parameter validation (existence, read-only status, element ownership)
   - Source value validation (empty checks at multiple points)

2. **Comprehensive Exception Handling:**
   - Full exception logging with stack traces
   - Graceful degradation (warnings vs. errors)
   - Per-sleeve exception handling (one failure doesn't stop others)

3. **Optimization Safeguards:**
   - Skip check prevents unnecessary writes
   - Batch parameter lookup with cache validation
   - Disabled mapping check prevents unnecessary processing (`mapping.IsEnabled`)

4. **Data Source Fallbacks:**
   - Snapshot → Revit element fallback for individual sleeves
   - Parameter name variation matching
   - Category-specific mappings (Cable Trays: "Service Type" → "System Type")


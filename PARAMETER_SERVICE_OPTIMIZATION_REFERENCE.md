# Parameter Service Project - Optimization Reference

Reference document from MEP Opening project optimizations.
Apply these patterns to the Parameter Service app for batch DB writes and Revit parameter flush queuing.

---

## 1. BATCH PARAMETER WRITE TO DB (Deferred Dictionary Pattern)

### Pattern: Accumulate in dictionary, flush once in single transaction

**Source:** `Services/Placement/SleeveParameterService.cs` (line 575+)

```
Architecture:
  Per-sleeve loop:
    targetDict[sleeveElementId]["Width"] = 0.5;
    targetDict[sleeveElementId]["Height"] = 0.3;
    targetDict[sleeveElementId]["Depth"] = 0.15;
    // ... accumulate ALL params for ALL sleeves

  After loop:
    FlushDeferredParameters("Context-Name")
    // Single pass writes everything
```

### Key Implementation Details:

**a) Deferred Dictionary Structure:**
```csharp
// Dictionary<ElementId, Dictionary<string, object>>
private Dictionary<ElementId, Dictionary<string, object>> _deferredBatchParams;
```

**b) Queuing (during placement loop):**
```csharp
if (!_deferredBatchParams.ContainsKey(sleeveId))
    _deferredBatchParams[sleeveId] = new Dictionary<string, object>();
_deferredBatchParams[sleeveId][paramName] = paramValue;
```

**c) Flushing (single pass after loop):**
- Pre-cache all elements via `doc.GetElement()` BEFORE the param-setting loop
- Cache `Definition` objects once from a representative element (avoids N x LookupParameter calls)
- Group sleeves by size key to detect repeated dimension sets
- Use `element.get_Parameter(definition).Set(value)` instead of `LookupParameter(name).Set(value)`

### Performance Numbers (304 sleeves, 2875 params):
- Flush time: 668ms total (0.23ms per param)
- Without batching: ~3,400ms (individual writes)
- Savings: ~80%

---

## 2. REVIT PARAMETER FLUSH - DEFINITION CACHING

**Source:** `Services/Placement/SleeveParameterService.cs` (line 621-640)

### The Problem:
`LookupParameter("Width")` is expensive (~0.5ms per call). For 304 sleeves x 9 params = 2,736 lookups.

### The Solution:
```csharp
// Cache definitions ONCE from first element
var defCache = new Dictionary<string, Definition>();
var firstSleeve = elementCache.Values.First();
foreach (var paramName in allParamNames)
{
    var p = firstSleeve.LookupParameter(paramName);
    if (p != null) defCache[paramName] = p.Definition;
}

// In the per-sleeve loop, use cached definition:
var param = element.get_Parameter(defCache[paramName]);
param.Set(value);
```

### Performance:
- `get_Parameter(Definition)`: ~0.05ms per call
- `LookupParameter(string)`: ~0.5ms per call
- Savings: 10x faster parameter access

---

## 3. BACKGROUND DB PERSISTENCE (Task.Run Pattern)

**Source:** `Services/OpeningCommandOrchestrator.cs`

### Pattern: Overlap DB writes with Revit transaction commit

```
Timeline:
  [Revit Main Thread] Place sleeves + Set params + tx.Commit()
  [Background Thread]                    Task.Run(() => WriteToDB())
                                         |----- DB write overlaps with Revit commit -----|
  [Revit Main Thread] .................. tx.Commit() done ........ await dbTask;
```

### Implementation:
```csharp
// Start DB write in background BEFORE calling tx.Commit()
var dbTask = Task.Run(() => {
    dbContext.BulkInsertOrUpdate(sleeves);
    // SQLite write happens while Revit regenerates
});

transaction.Commit(); // Revit regenerates ~3,300ms

await dbTask; // DB is already done (ran in parallel)
```

### Key Rules:
- NEVER access Revit API from background thread
- Only DB/file operations in Task.Run
- Use `await` after Revit commit to catch any DB errors

---

## 4. SQLITE OPTIMIZATION - Microsoft.Data.Sqlite Migration (R25/R26)

**Source:** `Data/SqliteCompat.cs`

### Architecture:
```
R23/R24 (net48):  System.Data.SQLite.Core 1.0.118.0
R25/R26 (net8.0): Microsoft.Data.Sqlite 8.0.11 (30-50% faster for bulk ops)
```

### Compatibility Layer (zero code changes at call sites):
1. **Global using aliases** for Connection, Command, Transaction, DataReader, Exception
2. **Wrapper class** for SQLiteParameter (handles `DbType` vs `SqliteType` constructor difference)
3. **Extension method** for `cmd.Parameters.Add(name, DbType)` overload
4. **Connection string helper** (`SqliteConnStr.Build()`) strips `Version=3` for NET8

### Key API Differences to Watch:
| Feature | System.Data.SQLite | Microsoft.Data.Sqlite |
|---------|-------------------|----------------------|
| Parameter ctor | `(string, DbType)` | `(string, SqliteType)` |
| Params.Add | `Add(string, DbType)` | `Add(string, SqliteType)` |
| Connection string | `Version=3` required | `Version=3` NOT supported |
| ConnectionStringBuilder | `SQLiteConnectionStringBuilder` | Not available (use plain string) |
| JournalMode enum | `SQLiteJournalModeEnum.Wal` | Set via PRAGMA |
| Error codes | `SQLiteErrorCode.Busy` (enum) | `SqliteErrorCode == 5` (int) |
| PRAGMA foreign_keys | In connection string builder | Must execute as SQL command |

---

## 5. REVIT 2026 ELEMENTID COMPATIBILITY

**Source:** `Helpers/ElementIdCompat.cs`

### The Breaking Change:
- Revit 2023-2025: `ElementId.IntegerValue` (int property)
- Revit 2026: `ElementId.Value` (long property), `IntegerValue` REMOVED

### Solution: Extension method
```csharp
internal static class ElementIdCompat
{
    internal static int GetIntegerValue(this ElementId id)
    {
#if REVIT2026
        return (int)id.Value;
#else
        return id.IntegerValue;
#endif
    }
}
```
Replace ALL `.IntegerValue` with `.GetIntegerValue()` across the codebase.

---

## 6. BATCH ELEMENT CREATION (NewFamilyInstances2)

**Source:** `Services/BulkPlacementService.cs`

### Pattern:
```csharp
// Prepare all placement data
var creationDataList = new List<FamilyInstanceCreationData>();
foreach (var zone in zones)
{
    creationDataList.Add(new FamilyInstanceCreationData(
        point, symbol, level, StructuralType.NonStructural));
}

// Single API call creates ALL elements
var instanceIds = doc.Create.NewFamilyInstances2(creationDataList);
```

### Performance:
- Individual `NewFamilyInstance`: ~15ms per element
- Batch `NewFamilyInstances2`: ~2.5ms per element (6x faster)

---

## 7. DB BULK OPERATIONS - Temp Table Pattern

**Source:** `Data/Repositories/ClashZoneRepository.cs`

### Pattern for 1000+ row updates:
```sql
-- 1. Create temp table
CREATE TEMP TABLE BulkUpdateZones (...);

-- 2. Insert all updates into temp table (parameterized, in transaction)
INSERT INTO BulkUpdateZones VALUES (@p1, @p2, ...);

-- 3. Single JOIN-based UPDATE from temp table
UPDATE ClashZones SET
    Width = t.Width, Height = t.Height
FROM BulkUpdateZones t
WHERE ClashZones.Id = t.Id;

-- 4. Drop temp table
DROP TABLE BulkUpdateZones;
```

### Key: Reuse prepared statement with parameter rebinding
```csharp
// Prepare ONCE outside loop
cmd.CommandText = "INSERT INTO BulkUpdateZones VALUES (@id, @w, @h)";
var pId = cmd.Parameters.Add("@id", DbType.Int32);
var pW = cmd.Parameters.Add("@w", DbType.Double);
cmd.Prepare();

// Rebind values in loop (no re-parsing)
foreach (var zone in zones)
{
    pId.Value = zone.Id;
    pW.Value = zone.Width;
    cmd.ExecuteNonQuery();
}
```

---

## 8. PERFORMANCE LOG PATTERN

**Source:** `Services/Placement/PlacementPerformanceMonitor.cs`

### Pattern: Wrap operations with Stopwatch, log to file
```csharp
var sw = Stopwatch.StartNew();
// ... operation ...
sw.Stop();
SafeFileLogger.SafeAppendText("performance.log",
    $"[FLUSH-PERF] Context={context}, Sleeves={count}, Time={sw.ElapsedMilliseconds}ms\n");
```

---

## SUMMARY OF GAINS (MEP Opening Project)

| Optimization | Before | After | Gain |
|-------------|--------|-------|------|
| Parameter flush (304 sleeves) | 3,435ms | 668ms | 5.1x |
| NewFamilyInstances2 batch | 4,500ms | 799ms | 5.6x |
| DB bulk updates | 517ms | 30ms | 17x |
| Definition caching | 0.5ms/lookup | 0.05ms/lookup | 10x |
| Background DB persistence | Sequential | Parallel | -3,300ms overlap |
| Microsoft.Data.Sqlite (R25+) | System.Data.SQLite | Microsoft.Data.Sqlite | 30-50% bulk |

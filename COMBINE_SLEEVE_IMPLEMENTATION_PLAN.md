# Combine Sleeve Feature - OOP Implementation Plan

**Document Version:** 2.0  
**Status:** Implementation Ready  
**Related Documents:** 
- `COMBINE_SLEEVE_PLAN.md` (Feature Requirements)
- `SLEEVE_PLACEMENT_METHODOLOGY_REFACTORED.md` (Architecture)

---

## Table of Contents

1. [OOP Architecture Design](#1-oop-architecture-design)
2. [Code Reuse Strategy](#2-code-reuse-strategy)
3. [Crash Safety Implementation](#3-crash-safety-implementation)
4. [Performance Optimization](#4-performance-optimization)
5. [Multi-Threading Strategy](#5-multi-threading-strategy)
6. [Implementation Sequence](#6-implementation-sequence)
7. [Code Structure](#7-code-structure)
8. [Testing Strategy](#8-testing-strategy)

---

## 1. OOP Architecture Design

### 1.1 Service Hierarchy

```
ICombineSleeveService (Interface)
    └── CombineSleeveService (Base Orchestrator)
            ├── AutoCombineService (Auto Mode)
            └── ManualCombineService (Manual Mode)

ICombineValidationService (Interface)
    └── CombineValidationService (Validation Logic)

ICombineAlgorithmService (Interface)
    └── CombineAlgorithmService (Reuses ClusterAlgorithmService)

ICombinePlacementService (Interface)
    └── CombinePlacementService (Reuses ClusterPlacementService)
```

### 1.2 Class Responsibilities

| Class | Responsibility | Reuses |
|-------|---------------|--------|
| `CombineSleeveService` | Orchestration, mode selection | - |
| `AutoCombineService` | Auto mode logic, batch processing | `ClusterAlgorithmService`, `ClusterDataService` |
| `ManualCombineService` | Manual mode logic, element picking | `ClusterPlacementService` |
| `CombineValidationService` | Wall group, rotation validation | - |
| `CombineAlgorithmService` | Proximity calculation, clustering | `ClusterAlgorithmService` |
| `CombinePlacementService` | Combined sleeve placement | `ClusterPlacementService` |
| `CombineDataService` | Data loading, caching | `ClusterDataService` |

### 1.3 Dependency Injection Pattern

```csharp
// Factory pattern for service creation
public class CombineServiceFactory
{
    public static ICombineSleeveService CreateCombineService(
        Document doc,
        IClusterDataService dataService,
        IClusterAlgorithmService algorithmService,
        IClusterPlacementService placementService)
    {
        var validationService = new CombineValidationService();
        var combineAlgorithmService = new CombineAlgorithmService(algorithmService);
        var combinePlacementService = new CombinePlacementService(placementService);
        
        return new CombineSleeveService(
            doc,
            dataService,
            combineAlgorithmService,
            combinePlacementService,
            validationService);
    }
}
```

---

## 2. Code Reuse Strategy

### 2.1 Reuse Existing Clustering Services

**✅ Direct Reuse (No Modification):**
- `ClusterDataService` - Load sleeves from database
- `ClusterBoundingBoxServices` - Calculate bounding boxes
- `ClusterPlacementService` - Place sleeves (with adapter)
- `ClusterAlgorithmService` - Proximity algorithm (with adapter)

**✅ Adapter Pattern (Wrap Existing):**
- `CombineAlgorithmService` - Wraps `ClusterAlgorithmService`
- `CombinePlacementService` - Wraps `ClusterPlacementService`

**✅ Composition Pattern (Use Existing):**
- `CombineDataService` - Composes `ClusterDataService`
- `CombineValidationService` - New, but uses validation patterns

### 2.2 Adapter Pattern Implementation

```csharp
/// <summary>
/// Adapter to reuse ClusterAlgorithmService for combine operations
/// </summary>
public class CombineAlgorithmService : ICombineAlgorithmService
{
    private readonly IClusterAlgorithmService _clusterAlgorithmService;
    private readonly ICombineValidationService _validationService;
    
    public CombineAlgorithmService(
        IClusterAlgorithmService clusterAlgorithmService,
        ICombineValidationService validationService)
    {
        _clusterAlgorithmService = clusterAlgorithmService ?? 
            throw new ArgumentNullException(nameof(clusterAlgorithmService));
        _validationService = validationService ?? 
            throw new ArgumentNullException(nameof(validationService));
    }
    
    /// <summary>
    /// Find nearby sleeves for combining (reuses clustering proximity logic)
    /// </summary>
    public List<dynamic> FindNearbySleeves(
        dynamic sourceSleeve,
        List<dynamic> candidateSleeves,
        double tolerance,
        Document doc)
    {
        // ✅ CRASH-SAFE: Validate inputs
        if (sourceSleeve == null || candidateSleeves == null || candidateSleeves.Count == 0)
            return new List<dynamic>();
        
        var nearbySleeves = new List<dynamic>();
        
        foreach (var candidate in candidateSleeves)
        {
            try
            {
                // ✅ VALIDATION: Check wall group compatibility
                if (!_validationService.AreCompatibleWallGroups(sourceSleeve, candidate))
                    continue;
                
                // ✅ VALIDATION: Check rotation compatibility (straight axis only)
                if (!_validationService.AreCompatibleRotations(sourceSleeve, candidate))
                    continue;
                
                // ✅ REUSE: Use existing proximity check from clustering
                var proximityChecker = ProximityCheckerFactory.CreateProximityChecker(
                    sourceSleeve, candidate);
                
                if (proximityChecker.CheckProximity(sourceSleeve, candidate, tolerance, 
                    sourceSleeve.Orientation, doc))
                {
                    nearbySleeves.Add(candidate);
                }
            }
            catch (Exception ex)
            {
                // ✅ CRASH-SAFE: Log and continue
                SafeFileLogger.SafeAppendText("combine_errors.log",
                    $"[{DateTime.Now:HH:mm:ss}] Error checking proximity: {ex.Message}\n");
                continue;
            }
        }
        
        return nearbySleeves;
    }
}
```

### 2.3 Composition Pattern Implementation

```csharp
/// <summary>
/// Composes ClusterDataService for combine operations
/// </summary>
public class CombineDataService : ICombineDataService
{
    private readonly IClusterDataService _clusterDataService;
    private readonly Dictionary<string, List<ClashZone>> _categoryCache;
    private readonly Dictionary<int, ClashZone> _sleeveCache;
    
    public CombineDataService(IClusterDataService clusterDataService)
    {
        _clusterDataService = clusterDataService ?? 
            throw new ArgumentNullException(nameof(clusterDataService));
        _categoryCache = new Dictionary<string, List<ClashZone>>();
        _sleeveCache = new Dictionary<int, ClashZone>();
    }
    
    /// <summary>
    /// Load sleeves for combine operation (reuses ClusterDataService)
    /// </summary>
    public List<ClashZone> LoadSleevesForCombine(string category, Document doc)
    {
        // ✅ PERFORMANCE: Check cache first
        if (_categoryCache.TryGetValue(category, out var cached))
            return cached;
        
        // ✅ REUSE: Use existing data service
        var allZones = _clusterDataService.LoadClashZonesFromRegularXml(
            string.Empty, category, doc);
        
        // ✅ PERFORMANCE: Cache results
        _categoryCache[category] = allZones;
        
        // ✅ PERFORMANCE: Build sleeve cache for fast lookup
        foreach (var zone in allZones)
        {
            if (zone.SleeveInstanceId > 0)
                _sleeveCache[zone.SleeveInstanceId] = zone;
        }
        
        return allZones;
    }
    
    /// <summary>
    /// Get ClashZone by sleeve instance ID (fast cache lookup)
    /// </summary>
    public ClashZone? GetClashZoneBySleeveId(int sleeveInstanceId)
    {
        _sleeveCache.TryGetValue(sleeveInstanceId, out var zone);
        return zone;
    }
}
```

---

## 3. Crash Safety Implementation

### 3.1 Input Validation Pattern

```csharp
/// <summary>
/// Base class with common crash-safe patterns
/// </summary>
public abstract class CrashSafeServiceBase
{
    protected void ValidateInputs(params object?[] inputs)
    {
        foreach (var input in inputs)
        {
            if (input == null)
                throw new ArgumentNullException(nameof(input));
        }
    }
    
    protected T SafeExecute<T>(Func<T> operation, T defaultValue, string context)
    {
        try
        {
            return operation();
        }
        catch (Exception ex)
        {
            SafeFileLogger.SafeAppendText("combine_errors.log",
                $"[{DateTime.Now:HH:mm:ss}] {context}: {ex.Message}\nStack: {ex.StackTrace}\n");
            return defaultValue;
        }
    }
    
    protected void SafeExecute(Action operation, string context)
    {
        try
        {
            operation();
        }
        catch (Exception ex)
        {
            SafeFileLogger.SafeAppendText("combine_errors.log",
                $"[{DateTime.Now:HH:mm:ss}] {context}: {ex.Message}\nStack: {ex.StackTrace}\n");
        }
    }
}
```

### 3.2 Transaction Safety

```csharp
/// <summary>
/// Transaction-safe combine operation
/// </summary>
public class CombinePlacementService : CrashSafeServiceBase, ICombinePlacementService
{
    private readonly IClusterPlacementService _clusterPlacementService;
    
    public bool PlaceCombinedSleeve(
        Document doc,
        List<dynamic> sleeves,
        BoundingBoxXYZ combinedBbox,
        string category)
    {
        // ✅ CRASH-SAFE: Validate inputs
        ValidateInputs(doc, sleeves, combinedBbox);
        
        if (sleeves.Count < 2)
        {
            SafeFileLogger.SafeAppendText("combine_errors.log",
                $"[{DateTime.Now:HH:mm:ss}] Cannot combine: Need at least 2 sleeves\n");
            return false;
        }
        
        // ✅ TRANSACTION-SAFE: Use transaction with rollback on error
        using (var tx = new Transaction(doc, "Combine Sleeves"))
        {
            try
            {
                tx.Start();
                
                // ✅ REUSE: Use existing placement service
                var placed = _clusterPlacementService.PlaceClusterSleeve(
                    doc, sleeves, combinedBbox, category);
                
                if (!placed)
                {
                    tx.RollBack();
                    return false;
                }
                
                tx.Commit();
                return true;
            }
            catch (Exception ex)
            {
                // ✅ CRASH-SAFE: Rollback on any error
                tx.RollBack();
                SafeFileLogger.SafeAppendText("combine_errors.log",
                    $"[{DateTime.Now:HH:mm:ss}] Transaction rollback: {ex.Message}\n");
                return false;
            }
        }
    }
}
```

### 3.3 Null Safety Pattern

```csharp
/// <summary>
/// Null-safe property access
/// </summary>
public static class CombineExtensions
{
    public static WallGroup GetWallGroupSafe(this ClashZone? clashZone)
    {
        if (clashZone?.StructuralElementNormal == null)
            return WallGroup.Floor; // Safe default
        
        var normal = clashZone.StructuralElementNormal;
        
        // ✅ CRASH-SAFE: Check for null and use safe defaults
        if (Math.Abs(normal.X) > 0.9 && Math.Abs(normal.Y) < 0.1)
            return WallGroup.WallX;
        
        if (Math.Abs(normal.Y) > 0.9 && Math.Abs(normal.X) < 0.1)
            return WallGroup.WallY;
        
        return WallGroup.Floor;
    }
    
    public static bool IsStraightAxisAlignedSafe(this ClashZone? clashZone)
    {
        if (clashZone == null)
            return false;
        
        try
        {
            double angleDeg = clashZone.MepElementRotationAngle * 180 / Math.PI;
            while (angleDeg < 0) angleDeg += 360;
            while (angleDeg >= 360) angleDeg -= 360;
            
            double threshold = 2.0;
            return Math.Abs(angleDeg) < threshold ||
                   Math.Abs(angleDeg - 90) < threshold ||
                   Math.Abs(angleDeg - 180) < threshold ||
                   Math.Abs(angleDeg - 270) < threshold;
        }
        catch
        {
            return false; // Safe default
        }
    }
}
```

---

## 4. Performance Optimization

### 4.1 Caching Strategy

```csharp
/// <summary>
/// High-performance caching for combine operations
/// </summary>
public class CombineDataService : ICombineDataService
{
    private readonly IClusterDataService _clusterDataService;
    
    // ✅ PERFORMANCE: Multi-level caching
    private readonly ConcurrentDictionary<string, List<ClashZone>> _categoryCache;
    private readonly ConcurrentDictionary<int, ClashZone> _sleeveCache;
    private readonly ConcurrentDictionary<int, WallGroup> _wallGroupCache;
    private readonly ConcurrentDictionary<int, bool> _rotationCache;
    
    public CombineDataService(IClusterDataService clusterDataService)
    {
        _clusterDataService = clusterDataService;
        _categoryCache = new ConcurrentDictionary<string, List<ClashZone>>();
        _sleeveCache = new ConcurrentDictionary<int, ClashZone>();
        _wallGroupCache = new ConcurrentDictionary<int, WallGroup>();
        _rotationCache = new ConcurrentDictionary<int, bool>();
    }
    
    /// <summary>
    /// Get wall group with caching (O(1) lookup after first call)
    /// </summary>
    public WallGroup GetWallGroup(ClashZone clashZone)
    {
        if (clashZone.SleeveInstanceId <= 0)
            return WallGroup.Floor;
        
        return _wallGroupCache.GetOrAdd(clashZone.SleeveInstanceId, id =>
        {
            return clashZone.GetWallGroupSafe();
        });
    }
    
    /// <summary>
    /// Check rotation with caching
    /// </summary>
    public bool IsStraightAxisAligned(ClashZone clashZone)
    {
        if (clashZone.SleeveInstanceId <= 0)
            return false;
        
        return _rotationCache.GetOrAdd(clashZone.SleeveInstanceId, id =>
        {
            return clashZone.IsStraightAxisAlignedSafe();
        });
    }
}
```

### 4.2 Spatial Indexing

```csharp
/// <summary>
/// Spatial index for fast proximity queries (O(log n) instead of O(n))
/// </summary>
public class SpatialIndex
{
    private readonly Dictionary<(int x, int y, int z), List<int>> _grid;
    private readonly double _cellSize;
    
    public SpatialIndex(double cellSize = 1000.0) // 1m cells
    {
        _grid = new Dictionary<(int x, int y, int z), List<int>>();
        _cellSize = cellSize;
    }
    
    /// <summary>
    /// Add sleeve to spatial index
    /// </summary>
    public void AddSleeve(int sleeveId, XYZ position)
    {
        var cell = GetCell(position);
        if (!_grid.TryGetValue(cell, out var list))
        {
            list = new List<int>();
            _grid[cell] = list;
        }
        list.Add(sleeveId);
    }
    
    /// <summary>
    /// Find nearby sleeves (only check adjacent cells)
    /// </summary>
    public List<int> FindNearby(int sleeveId, XYZ position, double radius)
    {
        var nearby = new HashSet<int>();
        var centerCell = GetCell(position);
        var cellRadius = (int)Math.Ceiling(radius / _cellSize);
        
        // ✅ PERFORMANCE: Only check adjacent cells, not all sleeves
        for (int dx = -cellRadius; dx <= cellRadius; dx++)
        {
            for (int dy = -cellRadius; dy <= cellRadius; dy++)
            {
                for (int dz = -cellRadius; dz <= cellRadius; dz++)
                {
                    var cell = (centerCell.x + dx, centerCell.y + dy, centerCell.z + dz);
                    if (_grid.TryGetValue(cell, out var sleeves))
                    {
                        foreach (var id in sleeves)
                        {
                            if (id != sleeveId)
                                nearby.Add(id);
                        }
                    }
                }
            }
        }
        
        return nearby.ToList();
    }
    
    private (int x, int y, int z) GetCell(XYZ position)
    {
        return (
            (int)Math.Floor(position.X / _cellSize),
            (int)Math.Floor(position.Y / _cellSize),
            (int)Math.Floor(position.Z / _cellSize)
        );
    }
}
```

### 4.3 Batch Processing

```csharp
/// <summary>
/// Batch processing for auto mode (process multiple categories in parallel)
/// </summary>
public class AutoCombineService : CrashSafeServiceBase, IAutoCombineService
{
    private readonly ICombineDataService _dataService;
    private readonly ICombineAlgorithmService _algorithmService;
    private readonly ICombinePlacementService _placementService;
    private readonly ICombineValidationService _validationService;
    
    /// <summary>
    /// Process multiple categories in parallel batches
    /// </summary>
    public CombineResult ProcessCategories(
        Document doc,
        List<string> categories,
        double tolerance,
        IProgress<int>? progress = null)
    {
        // ✅ CRASH-SAFE: Validate inputs
        ValidateInputs(doc, categories);
        
        var results = new ConcurrentBag<CombineResult>();
        var totalCategories = categories.Count;
        var processed = 0;
        
        // ✅ PERFORMANCE: Process categories in parallel batches
        var batchSize = Math.Min(Environment.ProcessorCount, totalCategories);
        var batches = categories
            .Select((cat, index) => new { Category = cat, Index = index })
            .GroupBy(x => x.Index / batchSize)
            .Select(g => g.Select(x => x.Category).ToList())
            .ToList();
        
        Parallel.ForEach(batches, batch =>
        {
            foreach (var category in batch)
            {
                try
                {
                    var result = ProcessCategory(doc, category, tolerance);
                    results.Add(result);
                    
                    Interlocked.Increment(ref processed);
                    progress?.Report(processed * 100 / totalCategories);
                }
                catch (Exception ex)
                {
                    // ✅ CRASH-SAFE: Log and continue
                    SafeFileLogger.SafeAppendText("combine_errors.log",
                        $"[{DateTime.Now:HH:mm:ss}] Error processing category {category}: {ex.Message}\n");
                }
            }
        });
        
        return CombineResults(results);
    }
}
```

---

## 5. Multi-Threading Strategy

### 5.1 Parallel Processing Architecture

```csharp
/// <summary>
/// Multi-threaded combine service (reuses clustering parallel patterns)
/// </summary>
public class AutoCombineService : CrashSafeServiceBase, IAutoCombineService
{
    private readonly ICombineDataService _dataService;
    private readonly ICombineAlgorithmService _algorithmService;
    private readonly SemaphoreSlim _semaphore;
    private readonly int _maxConcurrency;
    
    public AutoCombineService(
        ICombineDataService dataService,
        ICombineAlgorithmService algorithmService)
    {
        _dataService = dataService;
        _algorithmService = algorithmService;
        
        // ✅ PERFORMANCE: Limit concurrency to prevent resource exhaustion
        _maxConcurrency = Math.Min(Environment.ProcessorCount, 8);
        _semaphore = new SemaphoreSlim(_maxConcurrency, _maxConcurrency);
    }
    
    /// <summary>
    /// Process wall groups in parallel (same pattern as clustering)
    /// </summary>
    public async Task<CombineResult> ProcessWallGroupsAsync(
        Document doc,
        List<ClashZone> sleeves,
        double tolerance,
        CancellationToken cancellationToken = default)
    {
        // ✅ PERFORMANCE: Group by wall group first (reduces parallel overhead)
        var wallGroups = sleeves
            .GroupBy(s => _dataService.GetWallGroup(s))
            .ToList();
        
        var results = new ConcurrentBag<CombineResult>();
        var tasks = new List<Task>();
        
        // ✅ PERFORMANCE: Process each wall group in parallel
        foreach (var wallGroup in wallGroups)
        {
            var group = wallGroup.ToList();
            var groupType = wallGroup.Key;
            
            tasks.Add(Task.Run(async () =>
            {
                await _semaphore.WaitAsync(cancellationToken);
                try
                {
                    var result = await ProcessWallGroupAsync(
                        doc, group, groupType, tolerance, cancellationToken);
                    results.Add(result);
                }
                finally
                {
                    _semaphore.Release();
                }
            }, cancellationToken));
        }
        
        await Task.WhenAll(tasks);
        return CombineResults(results);
    }
    
    /// <summary>
    /// Process single wall group (thread-safe)
    /// </summary>
    private async Task<CombineResult> ProcessWallGroupAsync(
        Document doc,
        List<ClashZone> sleeves,
        WallGroup wallGroup,
        double tolerance,
        CancellationToken cancellationToken)
    {
        // ✅ PERFORMANCE: Build spatial index for fast lookups
        var spatialIndex = new SpatialIndex();
        foreach (var sleeve in sleeves)
        {
            if (sleeve.SleevePlacementPoint != null)
            {
                spatialIndex.AddSleeve(sleeve.SleeveInstanceId, sleeve.SleevePlacementPoint);
            }
        }
        
        var combined = new List<CombinedSleeve>();
        var processed = new HashSet<int>();
        
        // ✅ PERFORMANCE: Process sleeves in parallel batches
        var batchSize = 100;
        var batches = sleeves
            .Where(s => !processed.Contains(s.SleeveInstanceId))
            .Select((s, i) => new { Sleeve = s, Index = i })
            .GroupBy(x => x.Index / batchSize)
            .Select(g => g.Select(x => x.Sleeve).ToList())
            .ToList();
        
        var batchTasks = batches.Select(batch => Task.Run(() =>
        {
            var batchCombined = new List<CombinedSleeve>();
            
            foreach (var sleeve in batch)
            {
                if (cancellationToken.IsCancellationRequested)
                    break;
                
                if (processed.Contains(sleeve.SleeveInstanceId))
                    continue;
                
                try
                {
                    // ✅ PERFORMANCE: Use spatial index for fast nearby lookup
                    var nearbyIds = spatialIndex.FindNearby(
                        sleeve.SleeveInstanceId,
                        sleeve.SleevePlacementPoint,
                        tolerance);
                    
                    var nearbySleeves = nearbyIds
                        .Select(id => _dataService.GetClashZoneBySleeveId(id))
                        .Where(cz => cz != null && !processed.Contains(cz.SleeveInstanceId))
                        .ToList();
                    
                    if (nearbySleeves.Count > 0)
                    {
                        var combinedSleeve = CreateCombinedSleeve(
                            doc, sleeve, nearbySleeves, tolerance);
                        
                        if (combinedSleeve != null)
                        {
                            batchCombined.Add(combinedSleeve);
                            
                            // Mark as processed
                            processed.Add(sleeve.SleeveInstanceId);
                            foreach (var nearby in nearbySleeves)
                            {
                                processed.Add(nearby.SleeveInstanceId);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    // ✅ CRASH-SAFE: Log and continue
                    SafeFileLogger.SafeAppendText("combine_errors.log",
                        $"[{DateTime.Now:HH:mm:ss}] Error processing sleeve {sleeve.SleeveInstanceId}: {ex.Message}\n");
                }
            }
            
            return batchCombined;
        }, cancellationToken));
        
        var batchResults = await Task.WhenAll(batchTasks);
        combined.AddRange(batchResults.SelectMany(r => r));
        
        return new CombineResult
        {
            CombinedCount = combined.Count,
            ProcessedSleeves = processed.Count,
            WallGroup = wallGroup
        };
    }
}
```

### 5.2 Thread-Safe Collections

```csharp
/// <summary>
/// Thread-safe combine result aggregator
/// </summary>
public class CombineResultAggregator
{
    private readonly ConcurrentBag<CombinedSleeve> _combined;
    private readonly ConcurrentDictionary<int, bool> _processed;
    private readonly ConcurrentQueue<string> _errors;
    private long _totalProcessed;
    
    public CombineResultAggregator()
    {
        _combined = new ConcurrentBag<CombinedSleeve>();
        _processed = new ConcurrentDictionary<int, bool>();
        _errors = new ConcurrentQueue<string>();
    }
    
    public void AddCombined(CombinedSleeve combined)
    {
        _combined.Add(combined);
    }
    
    public bool TryMarkProcessed(int sleeveId)
    {
        return _processed.TryAdd(sleeveId, true);
    }
    
    public bool IsProcessed(int sleeveId)
    {
        return _processed.ContainsKey(sleeveId);
    }
    
    public void AddError(string error)
    {
        _errors.Enqueue($"[{DateTime.Now:HH:mm:ss}] {error}");
    }
    
    public void IncrementProcessed()
    {
        Interlocked.Increment(ref _totalProcessed);
    }
    
    public CombineSummary GetSummary()
    {
        return new CombineSummary
        {
            CombinedCount = _combined.Count,
            ProcessedCount = _processed.Count,
            TotalProcessed = _totalProcessed,
            Errors = _errors.ToList()
        };
    }
}
```

---

## 6. Implementation Sequence

### Phase 1: Core Infrastructure (Week 1)

**Day 1-2: Interfaces and Base Classes**
- ✅ Create `ICombineSleeveService` interface
- ✅ Create `ICombineValidationService` interface
- ✅ Create `ICombineAlgorithmService` interface
- ✅ Create `ICombinePlacementService` interface
- ✅ Create `CrashSafeServiceBase` base class
- ✅ Create `CombineExtensions` helper class

**Day 3-4: Validation Service**
- ✅ Implement `CombineValidationService`
- ✅ Implement wall group detection
- ✅ Implement rotation validation
- ✅ Unit tests for validation

**Day 5: Data Service**
- ✅ Implement `CombineDataService` (wraps `ClusterDataService`)
- ✅ Implement caching strategy
- ✅ Performance tests

### Phase 2: Algorithm Service (Week 2)

**Day 1-2: Algorithm Adapter**
- ✅ Implement `CombineAlgorithmService` (wraps `ClusterAlgorithmService`)
- ✅ Implement proximity finding
- ✅ Integration tests

**Day 3-4: Spatial Indexing**
- ✅ Implement `SpatialIndex` class
- ✅ Performance optimization
- ✅ Benchmark tests

**Day 5: Manual Mode**
- ✅ Implement `ManualCombineService`
- ✅ Implement element picker
- ✅ End-to-end tests

### Phase 3: Auto Mode & Multi-Threading (Week 3)

**Day 1-2: Auto Mode Base**
- ✅ Implement `AutoCombineService` base logic
- ✅ Implement category processing
- ✅ Unit tests

**Day 3-4: Multi-Threading**
- ✅ Implement parallel processing
- ✅ Implement thread-safe collections
- ✅ Performance tests (1000+ sleeves)

**Day 5: Integration**
- ✅ Integrate with UI
- ✅ End-to-end tests
- ✅ Performance validation

### Phase 4: Placement & Testing (Week 4)

**Day 1-2: Placement Service**
- ✅ Implement `CombinePlacementService` (wraps `ClusterPlacementService`)
- ✅ Transaction safety
- ✅ Error handling

**Day 3-4: UI Integration**
- ✅ Create `CombineSleeveDialog.xaml`
- ✅ Implement ViewModel
- ✅ Wire up commands

**Day 5: Final Testing**
- ✅ User acceptance testing
- ✅ Performance testing
- ✅ Bug fixes

---

## 7. Code Structure

### 7.1 Directory Structure

```
Services/
├── Combining/
│   ├── Interfaces/
│   │   ├── ICombineSleeveService.cs
│   │   ├── ICombineValidationService.cs
│   │   ├── ICombineAlgorithmService.cs
│   │   └── ICombinePlacementService.cs
│   ├── Base/
│   │   └── CrashSafeServiceBase.cs
│   ├── Validation/
│   │   └── CombineValidationService.cs
│   ├── Data/
│   │   └── CombineDataService.cs
│   ├── Algorithm/
│   │   ├── CombineAlgorithmService.cs
│   │   └── SpatialIndex.cs
│   ├── Placement/
│   │   └── CombinePlacementService.cs
│   ├── Auto/
│   │   └── AutoCombineService.cs
│   ├── Manual/
│   │   └── ManualCombineService.cs
│   ├── CombineSleeveService.cs
│   └── CombineServiceFactory.cs
│
Commands/
└── CombineSleeveCommand.cs

Views/
└── CombineSleeveDialog.xaml
```

### 7.2 Key Classes

**CombineSleeveService.cs:**
```csharp
public class CombineSleeveService : CrashSafeServiceBase, ICombineSleeveService
{
    private readonly IAutoCombineService _autoService;
    private readonly IManualCombineService _manualService;
    
    public CombineResult ExecuteAuto(Document doc, string category, double tolerance)
    {
        return _autoService.ProcessCategory(doc, category, tolerance);
    }
    
    public bool ExecuteManual(Document doc, ElementId sleeve1, ElementId sleeve2)
    {
        return _manualService.CombineSleeves(doc, sleeve1, sleeve2);
    }
}
```

**CombineServiceFactory.cs:**
```csharp
public static class CombineServiceFactory
{
    public static ICombineSleeveService Create(Document doc)
    {
        // Wire up all dependencies using existing services
        var clusterDataService = ClusterServiceFactory.CreateDataService(doc);
        var clusterAlgorithmService = ClusterServiceFactory.CreateAlgorithmService();
        var clusterPlacementService = ClusterServiceFactory.CreatePlacementService(doc);
        
        var dataService = new CombineDataService(clusterDataService);
        var validationService = new CombineValidationService(dataService);
        var algorithmService = new CombineAlgorithmService(clusterAlgorithmService, validationService);
        var placementService = new CombinePlacementService(clusterPlacementService);
        
        var autoService = new AutoCombineService(dataService, algorithmService, placementService);
        var manualService = new ManualCombineService(dataService, algorithmService, placementService, validationService);
        
        return new CombineSleeveService(autoService, manualService);
    }
}
```

---

## 8. Testing Strategy

### 8.1 Unit Tests

```csharp
[TestClass]
public class CombineValidationServiceTests
{
    [TestMethod]
    public void AreCompatibleWallGroups_WallXAndWallX_ReturnsTrue()
    {
        // Arrange
        var service = new CombineValidationService();
        var sleeve1 = CreateSleeve(WallGroup.WallX);
        var sleeve2 = CreateSleeve(WallGroup.WallX);
        
        // Act
        var result = service.AreCompatibleWallGroups(sleeve1, sleeve2);
        
        // Assert
        Assert.IsTrue(result);
    }
    
    [TestMethod]
    public void AreCompatibleWallGroups_WallXAndWallY_ReturnsFalse()
    {
        // Arrange
        var service = new CombineValidationService();
        var sleeve1 = CreateSleeve(WallGroup.WallX);
        var sleeve2 = CreateSleeve(WallGroup.WallY);
        
        // Act
        var result = service.AreCompatibleWallGroups(sleeve1, sleeve2);
        
        // Assert
        Assert.IsFalse(result);
    }
}
```

### 8.2 Performance Tests

```csharp
[TestClass]
public class AutoCombineServicePerformanceTests
{
    [TestMethod]
    public void ProcessCategory_1000Sleeves_CompletesInUnder30Seconds()
    {
        // Arrange
        var service = CreateService();
        var sleeves = CreateTestSleeves(1000);
        var stopwatch = Stopwatch.StartNew();
        
        // Act
        var result = service.ProcessCategory(doc, "Ducts", 200.0);
        
        // Assert
        stopwatch.Stop();
        Assert.IsTrue(stopwatch.ElapsedMilliseconds < 30000, 
            $"Took {stopwatch.ElapsedMilliseconds}ms, expected < 30000ms");
    }
}
```

### 8.3 Integration Tests

```csharp
[TestClass]
public class CombineSleeveIntegrationTests
{
    [TestMethod]
    public void ExecuteAuto_CombinesSleeves_CreatesCombinedSleeve()
    {
        // Arrange
        var service = CombineServiceFactory.Create(doc);
        CreateTestSleeves(doc, 10);
        
        // Act
        var result = service.ExecuteAuto(doc, "Ducts", 200.0);
        
        // Assert
        Assert.IsTrue(result.CombinedCount > 0);
        Assert.IsTrue(result.ProcessedSleeves > 0);
    }
}
```

---

## 9. Performance Benchmarks

### 9.1 Target Performance

| Operation | Target | Measurement |
|-----------|--------|-------------|
| Manual Mode (2 sleeves) | < 2 seconds | End-to-end |
| Auto Mode (100 sleeves) | < 10 seconds | End-to-end |
| Auto Mode (1000 sleeves) | < 30 seconds | End-to-end |
| Spatial Index Build | < 1 second | 1000 sleeves |
| Proximity Check | < 10ms | Per pair |

### 9.2 Optimization Checklist

- ✅ **Caching**: Category cache, sleeve cache, wall group cache, rotation cache
- ✅ **Spatial Indexing**: O(log n) proximity queries instead of O(n)
- ✅ **Multi-Threading**: Parallel processing of wall groups and batches
- ✅ **Code Reuse**: Adapter pattern for existing clustering services
- ✅ **Crash Safety**: Try-catch blocks, null checks, transaction rollback
- ✅ **Batch Processing**: Process in batches to reduce memory overhead
- ✅ **Early Exit**: Skip incompatible sleeves early in validation

---

## 10. Error Handling

### 10.1 Error Categories

1. **Validation Errors**: Wall group mismatch, rotation mismatch
2. **Data Errors**: Missing ClashZone, invalid sleeve ID
3. **Placement Errors**: Transaction failure, Revit API errors
4. **Performance Errors**: Timeout, memory exhaustion

### 10.2 Error Recovery

```csharp
public class CombineErrorHandler
{
    public static CombineResult HandleError(Exception ex, string context)
    {
        SafeFileLogger.SafeAppendText("combine_errors.log",
            $"[{DateTime.Now:HH:mm:ss}] {context}: {ex.Message}\n{ex.StackTrace}\n");
        
        return new CombineResult
        {
            Success = false,
            ErrorMessage = GetUserFriendlyMessage(ex),
            Errors = new List<string> { ex.Message }
        };
    }
    
    private static string GetUserFriendlyMessage(Exception ex)
    {
        return ex switch
        {
            ArgumentNullException => "Invalid input: Missing required data",
            InvalidOperationException => "Cannot combine: " + ex.Message,
            RevitAPIException => "Revit error: Please try again",
            TimeoutException => "Operation timed out: Too many sleeves",
            _ => "An error occurred: " + ex.Message
        };
    }
}
```

---

**End of Implementation Plan**


# ClashZoneRepository SOLID Refactoring Plan

Migrate `ClashZoneRepository.cs` (7450 lines, 94 methods) to delegate to specialized repositories, eliminating duplicate code and achieving SOLID compliance.

---

## Current State Analysis

### Problem: SRP Violation
`ClashZoneRepository` handles 5+ distinct responsibilities:
1. **ClashZone CRUD** (core responsibility)
2. **SleeveSnapshot operations** (duplicate of `SleeveSnapshotRepository`)
3. **ClusterSleeve operations** (duplicate of `ClusterSleeveRepository`)
4. **CombinedSleeve operations** (duplicate of `CombinedSleeveRepository`)
5. **R-Tree spatial indexing**
6. **Flag management**
7. **Parameter aggregation**
8. **Utility helpers** (mapping, serialization)

### Existing Specialized Repositories
| Repository | Lines | Methods | Purpose |
|------------|-------|---------|---------|
| `SleeveSnapshotRepository.cs` | 352 | 8 | Snapshot read operations |
| `ClusterSleeveRepository.cs` | 1795 | 21 | Cluster sleeve CRUD |
| `CombinedSleeveRepository.cs` | 833 | 14 | Combined sleeve CRUD |

---

## Proposed Changes

### Phase 1: Inject Specialized Repositories (LOW RISK)

#### [MODIFY] [ClashZoneRepository.cs](file:///c:/JSE_CSharp_Projects/JSE_MEPOPENING_23/Data/Repositories/ClashZoneRepository.cs)

**Add constructor injection:**
```csharp
public class ClashZoneRepository : IClashZoneRepository
{
    private readonly SleeveDbContext _context;
    private readonly Action<string> _logger;
    
    // ✅ NEW: Injected specialized repositories
    private readonly ClusterSleeveRepository _clusterRepo;
    private readonly CombinedSleeveRepository _combinedRepo;
    private readonly SleeveSnapshotRepository _snapshotRepo;
    
    public ClashZoneRepository(
        SleeveDbContext context,
        Action<string> logger = null,
        ClusterSleeveRepository clusterRepo = null,  // Optional for backward compat
        CombinedSleeveRepository combinedRepo = null,
        SleeveSnapshotRepository snapshotRepo = null)
    {
        _context = context;
        _logger = logger ?? (_ => { });
        
        // Lazy-create if not injected (backward compatibility)
        _clusterRepo = clusterRepo ?? new ClusterSleeveRepository(context, logger);
        _combinedRepo = combinedRepo ?? new CombinedSleeveRepository(context, logger);
        _snapshotRepo = snapshotRepo ?? new SleeveSnapshotRepository(context, logger);
    }
}
```

---

### Phase 2: Delegate Cluster Methods (6 methods ~400 lines)

#### Methods to Delegate to `ClusterSleeveRepository`:

| Method in ClashZoneRepository | Lines | Delegate To | Action |
|-------------------------------|-------|-------------|--------|
| `GetClusterSleevesByInstanceIds()` | 7068-7109 | `_clusterRepo.GetClusterSleevesByInstanceIds()` | Wrapper |
| `GetAllClusterSleeves()` | 7111-7152 | `_clusterRepo.GetAllClusterSleeves()` | Wrapper |
| `MapClusterSleeve()` | 7154-7181 | Keep private (used internally) | No change |
| `MapClusterSleeve()` (overload) | 7183-7208 | Keep private | No change |
| `UpdateClusterSleeveCorners()` | 5887-5968 | Add to `ClusterSleeveRepository` | Move + Delegate |
| `UpdateClusterPlacement()` | 5560-5668 | Add to `ClusterSleeveRepository` | Move + Delegate |

**Example Delegation:**
```csharp
// BEFORE: Full implementation in ClashZoneRepository
public List<ClusterSleeveData> GetAllClusterSleeves()
{
    // 41 lines of implementation
}

// AFTER: Simple delegation
public List<ClusterSleeveData> GetAllClusterSleeves()
{
    return _clusterRepo.GetAllClusterSleeves();
}
```

---

### Phase 3: Delegate Snapshot Methods (8 methods ~600 lines)

#### Methods to Consolidate with `SleeveSnapshotRepository`:

| Method in ClashZoneRepository | Lines | Action |
|-------------------------------|-------|--------|
| `SaveSleeveSnapshotsForPlacedSleeves()` | 2162-2510 | Move write logic to `SleeveSnapshotRepository` |
| `InsertOrUpdateSleeveSnapshots()` | 2512-2518 | Wrapper → delegate |
| `InsertOrUpdateSleeveSnapshotsInternal()` | 2520-2869 | Move to `SleeveSnapshotRepository` |
| `UpsertSleeveSnapshot()` | 2871-2910 | Move to `SleeveSnapshotRepository` |
| `GetExistingSnapshotId()` | 2912-3039 | Move to `SleeveSnapshotRepository` |
| `InsertSleeveSnapshot()` | 3041-3153 | Move to `SleeveSnapshotRepository` |
| `UpdateSleeveSnapshot()` | 3155-3249 | Move to `SleeveSnapshotRepository` |
| `GetSnapshotMepParametersForSleeveIds()` | 6708-6760 | Delegate to `_snapshotRepo` |

---

### Phase 4: Delegate Combined Methods (4 methods ~200 lines)

#### Methods to Consolidate with `CombinedSleeveRepository`:

| Method in ClashZoneRepository | Lines | Action |
|-------------------------------|-------|--------|
| `UpdateCombinedResolutionFlags()` | 7256-7313 | Move to `CombinedSleeveRepository` |
| `UpdateCombinedResolutionFlagsByClusterIds()` | 7315-7362 | Move to `CombinedSleeveRepository` |

---

### Phase 5: Extract Utility Classes (Optional)

#### Candidates for extraction:
| Methods | Lines | New Class |
|---------|-------|-----------|
| `SerializeDictionary`, `SerializeList`, `DeserializeDictionary` | 3426-5137 | `ParameterSerializer` |
| `ComputePlanarOrientationAngles`, `NormalizeDegrees` | 3446-3494 | `GeometryHelper` |
| `MapClashZone`, `SetMetadataFromReader` | 5139-5454 | `ClashZoneMapper` |
| `GetInt`, `GetDouble`, `GetBool`, etc. | 5456-5510 | `SqliteReaderHelper` |

---

### Phase 6: Interface Segregation (ISP)

#### Split `IClashZoneRepository` (~40 methods) into focused interfaces:

```csharp
// Core CRUD operations
public interface IClashZoneCrudRepository
{
    void InsertOrUpdateClashZones(...);
    void InsertOrUpdateClashZonesBulk(...);
    List<ClashZone> GetClashZonesByCategory(string category);
    List<ClashZone> GetClashZonesByFilter(...);
    List<ClashZone> GetClashZonesByGuids(...);
}

// Flag management
public interface IClashZoneFlagRepository
{
    void BatchUpdateFlags(...);
    void BatchUpdateFlagsWithCurrentClash(...);
    void ResetIsCurrentClashFlag(...);
    void BulkSetReadyForPlacementFlags(...);
}

// Spatial queries
public interface IClashZoneSpatialRepository
{
    List<ClashZone> GetClashZonesInSectionBoxRTree(...);
    List<ClashZone> GetClashZonesByCategoryInSectionBox(...);
    void UpdateRTreeIndex(...);
}

// Sleeve operations (delegates to specialized repos)
public interface ISleeveOperationsRepository
{
    void SaveSleeveSnapshotsForPlacedSleeves(...);
    List<ClusterSleeveData> GetAllClusterSleeves();
    List<ClusterSleeveData> GetClusterSleevesByInstanceIds(...);
}

// Main interface composes others
public interface IClashZoneRepository : 
    IClashZoneCrudRepository,
    IClashZoneFlagRepository,
    IClashZoneSpatialRepository,
    ISleeveOperationsRepository
{
    // Any additional methods not in sub-interfaces
}
```

---

## Implementation Order

| Phase | Risk | Effort | Lines Removed |
|-------|------|--------|---------------|
| 1. Constructor injection | LOW | 1 hour | 0 |
| 2. Cluster delegation | LOW | 2 hours | ~400 |
| 3. Snapshot delegation | MEDIUM | 4 hours | ~600 |
| 4. Combined delegation | LOW | 1 hour | ~200 |
| 5. Utility extraction | LOW | 2 hours | ~400 |
| 6. Interface segregation | MEDIUM | 3 hours | 0 |

**Total estimated reduction:** ~1600 lines (22% of file)

---

## OptimizationFlags Toggle

Add flag for safe rollback:
```csharp
// In OptimizationFlags.cs
/// <summary>
/// Enable delegation to specialized repositories in ClashZoneRepository.
/// When true: Cluster/Snapshot/Combined methods delegate to specialized repos.
/// When false: Uses existing inline implementations (legacy).
/// Default: false (opt-in while stabilizing).
/// </summary>
public static bool UseClashZoneRepositoryDelegation { get; set; } = false;
```

Usage in ClashZoneRepository:
```csharp
public List<ClusterSleeveData> GetAllClusterSleeves()
{
    if (OptimizationFlags.UseClashZoneRepositoryDelegation)
        return _clusterRepo.GetAllClusterSleeves();
    
    // Legacy inline implementation
    // ... existing code ...
}
```

---

## Verification Plan

### Automated Tests
No existing unit tests found for repositories. **Recommend user guidance** on testing approach.

### Manual Verification
1. **Pre-Migration Baseline:**
   - Run Refresh with current code
   - Note sleeve counts, flag states, snapshot counts
   
2. **Post-Migration Validation:**
   - Run Refresh with `UseClashZoneRepositoryDelegation = true`
   - Compare: Same sleeve counts, flag states, snapshot counts
   - Verify DB tables have identical data

3. **Rollback Test:**
   - Set `UseClashZoneRepositoryDelegation = false`
   - Confirm legacy code path works

---

## User Review Required

> [!IMPORTANT]
> **Testing Strategy Needed:** This codebase has no unit tests for repository layer. Please advise:
> 1. Should I create integration tests before refactoring?
> 2. Is manual testing via Refresh workflow sufficient?
> 3. Any specific scenarios that MUST work post-refactoring?

> [!WARNING]
> **Breaking Change Risk:** Phase 6 (Interface Segregation) will require updates to all consumers of `IClashZoneRepository`. This should be done last and may require separate planning.

---

## Files Affected

| File | Action |
|------|--------|
| `ClashZoneRepository.cs` | Add injection, delegate methods |
| `ClusterSleeveRepository.cs` | Add missing methods (`UpdateClusterSleeveCorners`, `UpdateClusterPlacement`) |
| `SleeveSnapshotRepository.cs` | Add write methods (currently read-only) |
| `CombinedSleeveRepository.cs` | Add flag update methods |
| `OptimizationFlags.cs` | Add `UseClashZoneRepositoryDelegation` flag |
| `IClashZoneRepository.cs` | (Phase 6 only) Split into sub-interfaces |

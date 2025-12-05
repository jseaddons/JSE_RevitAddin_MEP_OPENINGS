# Quick Reference: Architecture Overview

## Three Core Operations at a Glance

```
┌─────────────────────────────────────────────────────────────┐
│  OPERATION 1: DETECT                                        │
│  ├─ Service: ClashZoneService_Legacy                        │
│  ├─ Input: MEP elements + Structural elements               │
│  ├─ Process: Find intersections (ReferenceIntersector)      │
│  ├─ NEW: Flag-gated damper priority (UseSOLIDCompliantDamperFilter) │
│  ├─ Output: ClashZone records in database                   │
│  └─ Key Classes:                                             │
│     ├─ IntersectionDetectionService (geometry)              │
│     ├─ MepIntersectionService (category logic)              │
│     ├─ DamperDetection (priority filtering)                 │
│     └─ ClashZoneService_Legacy (orchestration)              │
└─────────────────────────────────────────────────────────────┘
                            ↓↓↓
┌─────────────────────────────────────────────────────────────┐
│  OPERATION 2: MANAGE FLAGS                                  │
│  ├─ Service: FlagManager_Legacy                             │
│  ├─ Input: ClashZone records                                │
│  ├─ Process:                                                 │
│  │  ├─ Transaction management (all or nothing)              │
│  │  ├─ Crash recovery (XML backup)                          │
│  │  ├─ Deterministic GUID tracking                          │
│  │  └─ Batch updates (100+ zones at once)                   │
│  ├─ Output: Database + XML backup                           │
│  └─ Key Classes:                                             │
│     ├─ FlagManager_Legacy (state tracking)                  │
│     ├─ GuidManager (deterministic GUIDs)                    │
│     ├─ ClashZoneDataService (DB operations)                 │
│     └─ GlobalIndexService (XML persistence)                 │
└─────────────────────────────────────────────────────────────┘
                            ↓↓↓
┌─────────────────────────────────────────────────────────────┐
│  OPERATION 3: PLACE SLEEVES                                 │
│  ├─ Service: NewSleevePlacerService (modern)                │
│  ├─ Legacy: UniversalSleevePlacerService (fallback)         │
│  ├─ Orchestration: OpeningCommandOrchestrator               │
│  ├─ Coordinator: SleevePlacementCoordinator (PATH selection)│
│  ├─ Input: ClashZone records from database                  │
│  ├─ Process:                                                 │
│  │  ├─ PATH 0: REPLAY (fastest - use cached)                │
│  │  ├─ PATH 1: SIZING (standard - calculate)                │
│  │  └─ PATH 2: DETECTION (full - re-scan)                   │
│  ├─ Features:                                                │
│  │  ├─ SOLID strategy pattern for sizing                    │
│  │  ├─ Deferred parameter batching (4-6× faster)            │
│  │  ├─ RCS coordinate transforms (wall-aligned)             │
│  │  ├─ Clustering support                                    │
│  │  └─ Crash-safe timeout protection (5 min)                │
│  ├─ Output: Family instances + updated flags                │
│  └─ Key Classes:                                             │
│     ├─ NewSleevePlacerService (SOLID-compliant)             │
│     ├─ UniversalSleevePlacerService (monolithic)            │
│     ├─ OpeningCommandOrchestrator (multi-filter exec)        │
│     ├─ SleevePlacementCoordinator (path selection)           │
│     ├─ ISleevePlacementStrategy (sizing strategies)          │
│     ├─ ClearanceCalculationService (clearance logic)         │
│     ├─ RcsBoundingBoxService (coordinate transforms)         │
│     ├─ ParameterBatchingService (deferred writes)            │
│     ├─ CrashSafeExecutor (timeout protection)                │
│     └─ ZoneFilterService (pre-filtering)                     │
└─────────────────────────────────────────────────────────────┘
```

---

## Flag Control Panel

### Operational Flags (Enable/Disable Features)

| Flag | Default | Purpose | Impact |
|------|---------|---------|--------|
| `UseSOLIDCompliantDamperFilter` | `true` | ✨ NEW: Damper priority (flag-gated) | MEDIUM (filtering logic) |
| `UseNewSleevePlacerService` | `true` | Use modern service vs. legacy | HIGH (entire pipeline) |
| `UseBatchedParameterWrites` | `true` | Deferred parameters (4-6× faster) | HIGH (performance) |
| `UseRefactoredCommandServices` | `true` | Inject SOLID services | HIGH (architecture) |
| `UseCrashSafeExecution` | `true` | Timeout protection (5 min) | MEDIUM (safety) |

### Database Optimization Flags

| Flag | Default | Purpose | Impact |
|------|---------|---------|--------|
| `ReuseDbContextDuringRefresh` | `true` | Connection pooling | MEDIUM (performance) |
| `UseBatchedSleeveExistenceCheck` | `true` | Batch flag queries | LOW (optimization) |
| `SkipXmlLoadingDuringRefresh` | `true` | Database-first mode | MEDIUM (performance) |
| `UseGlobalCategoryIndexForRefresh` | `true` | Fast index for large projects | LOW (optimization) |

### Safe Rollback Command

If issues occur, disable in order:

```csharp
// Step 1: Disable new damper filter
OptimizationFlags.UseSOLIDCompliantDamperFilter = false;

// Step 2: If still issues, disable new service
OptimizationFlags.UseNewSleevePlacerService = false;

// Step 3: If still issues, disable performance features
OptimizationFlags.UseBatchedParameterWrites = false;
OptimizationFlags.UseRefactoredCommandServices = false;

// Step 4: Reload add-in
// Ctrl+Alt+F12 → Manage Add-ins → Reload
```

---

## Path Selection (SleevePlacementCoordinator)

```
Is zone already resolved?
    ├─ YES + cached dimensions exist
    │  └─ PATH 0: REPLAY
    │     ├─ Speed: 100-200ms per zone
    │     ├─ Action: Use saved dimensions
    │     └─ Best for: Reruns, unchanged zones
    │
    ├─ YES + sleeve exists but dimensions need update
    │  └─ PATH 1: SIZING
    │     ├─ Speed: 500ms per zone
    │     ├─ Action: Recalculate optimal size
    │     └─ Best for: Size adjustments, clearance changes
    │
    └─ NO
       └─ PATH 2: DETECTION
          ├─ Speed: 1000+ms per zone
          ├─ Action: Full re-scan + calculate
          └─ Best for: New detections, fresh start
```

---

## Damper Priority Filter (SOLID-Compliant)

**NEW Feature** gated by `UseSOLIDCompliantDamperFilter`

```
For each Duct element:
    ├─ Check cached HasDamperNearby flag (O(1) - fastest)
    │  └─ If true → SKIP this duct (damper takes priority)
    │
    ├─ Else, run proximity check (O(n) on same wall only)
    │  ├─ Pre-filter dampers by SAME WALL
    │  ├─ For each damper on same wall:
    │  │  ├─ Distance from damper center to intersection ≤ 0.2ft?
    │  │  ├─ Intersection point inside damper bbox?
    │  │  └─ Damper bbox overlap with duct vicinity?
    │  └─ If damper found → Set flag + SKIP this duct
    │
    └─ Else → PASS this duct (proceed to placement)
```

**Why SOLID?**
- ✅ **SRP**: Damper logic is in isolated conditional block
- ✅ **OCP**: Can be extracted to separate filter without modifying ClashZoneService_Legacy
- ✅ **Flag-Gated**: Can be toggled without recompilation
- ✅ **Safe**: Two-tier caching avoids expensive proximity checks on reruns

---

## Performance Optimizations

### Deferred Parameter Batching (4-6× Faster)

```
❌ SLOW (Per-Instance Writes):
    For each sleeve instance:
        └─ Set Width parameter
        └─ Regenerate
        └─ Set Height parameter
        └─ Regenerate
        └─ ... repeat for each parameter
    Result: 50 instances × 4 parameters = 200 regenerations!

✅ FAST (Batched Writes):
    For each sleeve instance:
        └─ Accumulate in _deferredParameters dict
    ONCE after all placements:
        └─ Write ALL parameters
        └─ Regenerate ONCE
        └─ Clear dict
    Result: 50 instances × 4 parameters = 1 regeneration!
```

**Code Location**: `ParameterBatchingService` + flag `UseBatchedParameterWrites`

---

## Crash Recovery (Resilient Pipeline)

```
Exception occurs during placement:
    ├─ ClashZone flags saved to database
    │  ├─ IsResolved = true/false
    │  ├─ SleeveId = instance ID (if placed)
    │  └─ SleevePlacementGuid = deterministic ID
    │
    ├─ Backup saved to XML
    │  └─ GlobalIndexService.BackupToXml()
    │
    └─ On resume:
       ├─ Check flags first (O(log n) database query)
       ├─ If flags indicate already processed → skip
       ├─ If flags indicate partial → resume from checkpoint
       └─ If corrupted → restore from XML backup
```

**Timeout Protection**: `CrashSafeExecutor` (5-minute limit per category)

---

## Database Schema (ClashZone Table)

```sql
CREATE TABLE ClashZones (
    Id GUID PRIMARY KEY,
    MepElementId INTEGER,
    StructuralElementId INTEGER,
    IntersectionPointX REAL,
    IntersectionPointY REAL,
    IntersectionPointZ REAL,
    
    -- ✨ Damper Priority Flag (NEW)
    HasDamperNearby BOOLEAN DEFAULT FALSE,
    
    -- State Tracking
    IsResolved BOOLEAN DEFAULT FALSE,
    SleeveId INTEGER,  -- Family instance ID
    SleevePlacementGuid GUID,
    
    -- Cached Dimensions (for PATH 0 REPLAY)
    SleeveBoundingBoxRCS_MinX REAL,
    SleeveBoundingBoxRCS_MinY REAL,
    SleeveBoundingBoxRCS_MinZ REAL,
    SleeveBoundingBoxRCS_MaxX REAL,
    SleeveBoundingBoxRCS_MaxY REAL,
    SleeveBoundingBoxRCS_MaxZ REAL,
    
    -- Clustering
    IsClusterResolved BOOLEAN DEFAULT FALSE,
    ClusterGroupId GUID,
    
    -- Timestamps
    CreatedAt DATETIME,
    UpdatedAt DATETIME,
    
    -- Category & Filter
    MepElementCategory TEXT,
    FilterName TEXT
);

CREATE INDEX idx_category ON ClashZones(MepElementCategory);
CREATE INDEX idx_filter ON ClashZones(FilterName);
CREATE INDEX idx_resolved ON ClashZones(IsResolved);
```

---

## Common Tasks for Developers

### Add New Sizing Strategy (SOLID Pattern)

1. Create interface implementation:
   ```csharp
   public class CustomSizingStrategy : ISleevePlacementStrategy
   {
       public SleeveDimensions CalculateDimensions(ClashZone zone, OpeningConditions conditions)
       {
           // Your logic here
           return new SleeveDimensions { Width = ..., Height = ... };
       }
   }
   ```

2. Register in factory:
   ```csharp
   public class StrategyFactoryService : IStrategyFactory
   {
       public ISleevePlacementStrategy CreateStrategy(string strategyName)
       {
           return strategyName switch
           {
               "Custom" => new CustomSizingStrategy(),
               ...
           };
       }
   }
   ```

3. Use via dependency injection:
   ```csharp
   var strategy = factory.CreateStrategy("Custom");
   var placer = new NewSleevePlacerService(..., strategy);
   ```

### Add New Clearance Provider

1. Create interface implementation:
   ```csharp
   public class CustomClearanceProvider : IClearanceProvider
   {
       public SleeveClearance GetClearance(string mepCategory, string hostType)
       {
           // Your rules here
           return new SleeveClearance(25, 25, 25, 25);
       }
   }
   ```

2. Inject into clearance service:
   ```csharp
   var provider = new CustomClearanceProvider();
   var clearanceService = new ClearanceCalculationService(provider);
   ```

### Debug Placement Issues

```
1. Check logs:
   C:\Users\[user]\AppData\Roaming\JSE_MEP_Openings\Logs\R2024\
   - placement_debug.log (verbose)
   - placement_errors.log (errors)
   - placement_performance.log (timing)

2. Search for keywords:
   - "PATH 0" / "PATH 1" / "PATH 2" (path selection)
   - "SKIP DAMPER CHECK" (damper filtering)
   - "Deferred" (parameter batching)
   - "RCS" (coordinate transforms)
   - "TIMEOUT" (crash protection)

3. Disable features one by one:
   - UseSOLIDCompliantDamperFilter = false (new filter)
   - UseBatchedParameterWrites = false (performance)
   - UseNewSleevePlacerService = false (modern service)

4. Check database:
   SELECT * FROM ClashZones WHERE MepElementId = ?
   - Verify flags are set correctly
   - Check if SleeveId is populated
   - Review cached dimensions
```

---

## Testing Strategy

### Unit Tests (Per Component)

- `IntersectionDetectionServiceTests` → Geometry accuracy
- `ClearanceCalculationServiceTests` → Clearance values
- `RcsBoundingBoxServiceTests` → Coordinate transforms
- `FlagManagerTests` → State persistence
- `SizingStrategyTests` → Dimension calculations

### Integration Tests (Multi-Component)

- `PlacementPipelineTests` (all 3 operations)
- `DamperFilteringTests` (new SOLID filter)
- `CrashRecoveryTests` (timeout + rollback)
- `ParameterBatchingTests` (performance verification)

### Acceptance Tests (End-to-End)

- User places sleeves on ducts (all paths)
- Damper priority filtering works correctly
- Placement times faster with batching
- Crash recovery works after timeout
- Flags correctly persist state

---

## Performance Targets

| Operation | Current | Target | Path |
|-----------|---------|--------|------|
| Detect (per zone) | 50-100ms | 30-50ms | Use global index |
| Flag Management | 10-20ms | 5-10ms | Batch updates |
| Place (PATH 0 REPLAY) | 100-200ms | 50-100ms | Cached dimensions |
| Place (PATH 1 SIZING) | 500ms | 300-400ms | Batched parameters |
| Place (PATH 2 DETECTION) | 1000+ms | 800-900ms | Optimize intersection |
| Whole Pipeline (50 zones) | ~60s | ~25s | All optimizations |

**Key Driver**: Deferred parameter batching (4-6× faster placement)

---

**Last Updated**: December 4, 2025  
**Document Purpose**: Quick reference for developers  
**See Also**: COMPREHENSIVE_ARCHITECTURE_PLAN.md (detailed guide)

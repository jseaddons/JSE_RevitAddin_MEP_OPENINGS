# Multi-Agent Refactoring Context Document
## UniversalClusterService Refactoring Implementation Guide

**Purpose**: This document provides context and instructions for multi-agent implementation of the `UniversalClusterService.cs` refactoring plan.

**Target File**: `Services/UniversalClusterService.cs` (8,530 lines → ~500 lines orchestrator + 5 focused services)

**Reference Document**: `REFACTORING_PLAN_UniversalClusterService.md` (comprehensive plan with all details)

---

## Multi-Agent Coordination Strategy

### Agent Roles

1. **Architecture Agent**: Creates interfaces, base classes, and coordinates overall structure
2. **Service Extraction Agents**: Each agent handles one service extraction (Placement, Rotation, Cleanup, Algorithm)
3. **Testing Agent**: Creates unit tests and integration tests for each phase
4. **Integration Agent**: Ensures services work together, updates orchestrator
5. **Safety Agent**: Implements timeout, failure recovery, and safety checks

### Workflow

```
Phase 1 → Phase 2 → Phase 3 → Phase 4 → Phase 5 → Phase 6 → Phase 7 → Phase 8 → Phase 9
  ↓         ↓         ↓         ↓         ↓         ↓         ↓         ↓         ↓
Test     Test     Test     Test     Test     Test     Test     Test     Test
```

**Rule**: Each phase must be completed and tested before moving to the next phase.

---

## Phase-by-Phase Implementation Guide

### Phase 1: Extract Geometry Calculations + Crash-Safe Guards

**Agent**: Architecture Agent + Geometry Agent + Safety Agent

**Tasks**:
1. Create `Services/Clustering/Geometry/DistanceCalculator.cs`
   - Extract `CalculateMinimumDistance2D` method
   - Extract `CalculateMinimumDistance3D` method
   - Add null-safety (guard clauses, nullable annotations)
   - **NEW**: Validate geometric inputs (zero-length vectors, NaN values)
   - **NEW**: Return nullable results for invalid calculations

2. Create `Services/Clustering/Geometry/RotationMatrixCalculator.cs`
   - Extract rotation matrix calculation logic
   - Methods: `CreateRotationMatrix`, `ApplyRotation`, `InverseRotation`
   - **NEW**: Validate rotation angles (NaN, Infinity checks)
   - **NEW**: Return nullable results for invalid rotations

3. Create `Services/Clustering/Geometry/CoordinateTransformer.cs`
   - Extract coordinate transformation logic
   - Methods: `TransformToWorld`, `TransformToLocal`, `TransformPoint`
   - **NEW**: Validate coordinate inputs (NaN, Infinity checks)

4. **NEW**: Replace all file logging with `SafeFileLogger.SafeAppendText`
   - Find all `File.AppendAllText` calls
   - Replace with `SafeFileLogger.SafeAppendText(logPath, message)`
   - Ensure no hardcoded file paths

5. Update `UniversalClusterService.cs`
   - Replace inline geometry calculations with calls to new classes
   - Preserve all existing logic (no behavior changes)
   - **NEW**: Add guard clauses for null/invalid inputs (fail-fast)

**Crash-Safe Requirements**:
- ✅ Wrap all file operations in `SafeFileLogger.SafeAppendText`
- ✅ Add guard clauses for null/invalid inputs (fail-fast)
- ✅ Validate geometric inputs (zero-length vectors, NaN values)
- ✅ Use null-conditional operators for all property access
- ✅ Return nullable results for invalid calculations

**Dependencies**: None (can start immediately)

**Success Criteria**:
- ✅ All geometry calculations moved to dedicated classes
- ✅ `UniversalClusterService` uses new classes (no inline calculations)
- ✅ All file logging uses `SafeFileLogger.SafeAppendText`
- ✅ Guard clauses added for all inputs
- ✅ Unit tests pass for all geometry classes (including null/invalid inputs)
- ✅ Integration test: Cluster dimensions match baseline
- ✅ Performance validation: Calculation time ±5%, results exact match
- ✅ 0 new warnings introduced

**Performance Validation**:
1. ✅ Capture baseline metrics before refactoring (geometry calculation time)
2. ✅ Run performance tests after extraction
3. ✅ Compare: calculation time ±5%, results exact match
4. ✅ Fail phase if performance regression detected
5. ✅ Log metrics to `performance_baseline.json`

**Rollback Strategy**:
- ✅ Keep old methods as `[Obsolete]` during migration
- ✅ Feature flag: `USE_REFACTORED_GEOMETRY = false` for instant rollback
- ✅ Remove old code only after successful production validation

**Files to Create**:
- `Services/Clustering/Geometry/DistanceCalculator.cs`
- `Services/Clustering/Geometry/RotationMatrixCalculator.cs`
- `Services/Clustering/Geometry/CoordinateTransformer.cs`

**Files to Modify**:
- `Services/UniversalClusterService.cs` (replace geometry calculations)

---

### Phase 2: Extract Proximity Checking + Timeout Protection

**Agent**: Proximity Agent + Safety Agent

**Tasks**:
1. Create `Services/Clustering/Proximity/IProximityChecker.cs`
   - Interface: `CheckProximity(sleeve1, sleeve2, tolerance): bool`
   - Interface: `CalculateDistance(sleeve1, sleeve2): double`

2. Create `Services/Clustering/Proximity/BoundingBoxProximityChecker.cs`
   - Extract `BoundingBoxesOverlapFromXml` logic
   - Use pre-calculated bounding boxes from database
   - **NEW**: Validate inputs: null checks, IsValidObject for elements
   - **NEW**: Return false (no proximity) for invalid checks

3. Create `Services/Clustering/Proximity/EdgeToEdgeProximityChecker.cs`
   - Extract edge-to-edge distance logic for round pipes/ducts
   - Formula: `centerToCenter - (radius1 + radius2)`
   - Use `SleeveDiameter` from database
   - **NEW**: Validate sleeve diameter > 0 before calculation
   - **NEW**: Check for null placement points

4. Create `Services/Clustering/Proximity/RotatedProximityChecker.cs`
   - Extract `CheckRotatedSleeveProximity` logic
   - Use pre-calculated rotation matrix (cos/sin) from database
   - **NEW**: Null-safe parameter access with `??` operators

5. Create `Services/Clustering/Proximity/ProximityCheckerFactory.cs`
   - Decision logic: round pipes/ducts → rotated → bounding box
   - Use `SleeveGroupKey` to determine checker type

6. **NEW**: Create `Services/Clustering/Safety/TimeoutMonitor.cs`
   - Track elapsed time per group
   - 5-minute timeout per group (configurable)
   - Check timeout every 100 iterations
   - Log warning and skip remaining if timeout exceeded

7. Update `UniversalClusterService.cs`
   - Replace proximity checking with factory pattern
   - **NEW**: Add timeout monitoring to proximity loops
   - **NEW**: Wrap file logging in `SafeFileLogger.SafeAppendText`

**Crash-Safe Requirements**:
- ✅ Add timeout checks every 100 proximity checks (not every check)
- ✅ Wrap file logging in `SafeFileLogger.SafeAppendText`
- ✅ Validate inputs: null checks, IsValidObject for elements
- ✅ Null-safe parameter access with `??` operators
- ✅ Return safe defaults (false) for invalid checks

**Timeout Protection**:
- ✅ `TimeoutMonitor` tracks elapsed time per group
- ✅ 5-minute timeout per group (configurable)
- ✅ Check timeout every 100 iterations
- ✅ Log warning and skip remaining if timeout exceeded
- ✅ Allow partial success (process completed groups)

**Dependencies**: Phase 1 (uses `DistanceCalculator`)

**Success Criteria**:
- ✅ All proximity checking moved to strategy pattern
- ✅ Factory correctly selects checker based on sleeve type
- ✅ Timeout monitoring integrated into proximity loops
- ✅ Unit tests pass for all checkers (including timeout scenarios)
- ✅ Integration test: Proximity detection matches baseline
- ✅ Pre-calculated data used (no redundant calculations)
- ✅ Performance: Checking time ±5%, timeout overhead <1%
- ✅ Timeout test: Verify timeout triggers and partial results work

**Performance Validation**:
1. ✅ Baseline: proximity checking time for 100 sleeves
2. ✅ Compare: checking time ±5%, timeout overhead <1%
3. ✅ Verify timeout triggers correctly (test with 10,000 sleeves)

**Rollback Strategy**:
- ✅ Keep old methods as `[Obsolete]`
- ✅ Feature flag: `USE_REFACTORED_PROXIMITY = false`

**Files to Create**:
- `Services/Clustering/Proximity/IProximityChecker.cs`
- `Services/Clustering/Proximity/BoundingBoxProximityChecker.cs`
- `Services/Clustering/Proximity/EdgeToEdgeProximityChecker.cs`
- `Services/Clustering/Proximity/RotatedProximityChecker.cs`
- `Services/Clustering/Proximity/ProximityCheckerFactory.cs`

**Files to Modify**:
- `Services/UniversalClusterService.cs` (replace proximity checking)

---

### Phase 3: Extract Bounding Box Calculations

**Agent**: BoundingBox Agent

**Tasks**:
1. Create `Services/Clustering/Geometry/IBoundingBoxCalculator.cs`
   - Interface: `Calculate(cluster, rotationAngle): BoundingBoxResult`
   - Interface: `CalculateFromCorners(corners): BoundingBoxResult`

2. Create `Services/Clustering/Geometry/AxisAlignedBoundingBoxCalculator.cs`
   - Extract `GetClusterBoundingBoxFromXml` logic
   - Use pre-calculated bounding boxes from database

3. Create `Services/Clustering/Geometry/RotatedBoundingBoxCalculator.cs`
   - Extract `GetClusterBoundingBoxWithRotatedCoordinates` logic
   - Use corner-based algorithm (4 corners per sleeve)

4. Create `Services/Clustering/Geometry/CornerBasedBoundingBoxCalculator.cs`
   - Extract corner-based watertight algorithm
   - Transform all corners to rotated coordinate system
   - Find min/max extents

5. Update `UniversalClusterService.cs`
   - Replace bounding box calculations with calculator pattern

**Dependencies**: Phase 1 (uses `RotationMatrixCalculator`, `CoordinateTransformer`)

**Success Criteria**:
- ✅ All bounding box calculations moved to calculator pattern
- ✅ Corner-based algorithm preserved (watertight)
- ✅ Unit tests pass for all calculators
- ✅ Integration test: Cluster dimensions match baseline exactly
- ✅ Pre-calculated corners used from database

**Files to Create**:
- `Services/Clustering/Geometry/IBoundingBoxCalculator.cs`
- `Services/Clustering/Geometry/AxisAlignedBoundingBoxCalculator.cs`
- `Services/Clustering/Geometry/RotatedBoundingBoxCalculator.cs`
- `Services/Clustering/Geometry/CornerBasedBoundingBoxCalculator.cs`

**Files to Modify**:
- `Services/UniversalClusterService.cs` (replace bounding box calculations)

---

### Phase 4A: Extract Host-Type Clustering Strategies

**Agent**: Strategy Agent

**Tasks**:
1. Create `Services/Clustering/Strategy/IClusteringStrategy.cs`
   - Interface: `CalculateProximity(sleeve1, sleeve2, tolerance): double`
   - Interface: `FormClusters(sleeves, tolerance): List<List<dynamic>>`
   - Interface: `CanHandle(groupKey): bool`
   - **NEW**: Interface: `GetStrategyName(): string` (for logging/diagnostics)

2. Create `Services/Clustering/Strategy/FloorCircularClusteringStrategy.cs`
   - Extract floor circular sleeve logic (2D X-Y plane)
   - Edge-to-edge calculation for round pipes/ducts
   - Formula: `centerToCenter - (radius1 + radius2)`
   - **NEW**: Validate sleeve diameter > 0 before calculation
   - **NEW**: Check for null placement points
   - **NEW**: Return false (no proximity) if data is invalid

3. Create `Services/Clustering/Strategy/FloorRectangularClusteringStrategy.cs`
   - Extract floor rectangular sleeve logic (2D X-Y plane)
   - Bounding box minimum distance with `CalculateMinimumDistance2D`
   - **NEW**: Uses pre-calculated bounding boxes from database

4. Create `Services/Clustering/Strategy/WallAxisAlignedStrategy.cs`
   - Extract X-oriented wall logic (2D X-Z plane, ignore Y)
   - Extract Y-oriented wall logic (2D Y-Z plane, ignore X)
   - **NEW**: Same-wall validation (StructuralElementIdValue matching)
   - **NEW**: **CRITICAL**: Prevent cross-wall clustering
   - **NEW**: Check wall hosts BEFORE bounding box overlap

5. Create `Services/Clustering/Strategy/WallRotatedClusteringStrategy.cs`
   - Extract rotated sleeve logic (angle detection)
   - Same-axis validation with 180° modulo
   - Rotation transforms for proximity checking
   - **NEW**: Uses pre-calculated rotation matrix (cos/sin) from database

6. **NEW**: Create `Services/Clustering/Strategy/StructuralFramingStrategy.cs`
   - Extract Structural Framing-specific logic
   - Similar to wall strategies but with framing-specific rules

7. Create `Services/Clustering/Strategy/ClusteringStrategyFactory.cs`
   - **ENHANCED**: Decision tree with logging
   - Check if sleeves are circular or rectangular
   - Check if sleeves are rotated
   - Log strategy selection for diagnostics using `SafeFileLogger.SafeAppendText`
   - Handle hybrid clusters (mixed rotated/axis-aligned)

8. Update `UniversalClusterService.cs`
   - Replace clustering logic with strategy pattern
   - **NEW**: Add strategy selection logging

**Hybrid Cluster Support**:
- ✅ Check rotation per sleeve pair (not per group)
- ✅ Fallback: rotated → axis-aligned if angle mismatch
- ✅ If one sleeve is circular and one rectangular, use rectangular strategy
- ✅ Log mixed-shape clusters for diagnostics

**Cross-Wall Prevention** (WallAxisAlignedStrategy):
```csharp
// CRITICAL FIX: Check wall hosts BEFORE bounding box overlap
if (sleeve1.HostType == "Wall" && sleeve2.HostType == "Wall")
{
    var host1Id = sleeve1.ClashZone?.StructuralElementIdValue ?? -1;
    var host2Id = sleeve2.ClashZone?.StructuralElementIdValue ?? -1;
    
    if (host1Id != host2Id && host1Id != -1 && host2Id != -1)
    {
        // Different walls - do NOT cluster
        SafeFileLogger.SafeAppendText("wall_clustering.log",
            $"Rejected clustering: sleeves on different walls ({host1Id} vs {host2Id})");
        return false;
    }
}
```

**Dependencies**: Phase 2 (uses `IProximityChecker`), Phase 3 (uses `IBoundingBoxCalculator`)

**Success Criteria**:
- ✅ All clustering strategies extracted (Floor Circular, Floor Rectangular, Wall Axis-Aligned, Wall Rotated, Structural Framing)
- ✅ Factory correctly selects strategy based on `SleeveGroupKey` with logging
- ✅ Hybrid clusters handled (pre-filtering approach)
- ✅ Cross-wall clustering prevented
- ✅ Unit tests pass for each strategy (including hybrid scenarios)
- ✅ Integration test: Clusters match baseline
- ✅ Spatial grid and multi-threading preserved
- ✅ Strategy selection logged for diagnostics

**Files to Create**:
- `Services/Clustering/Strategy/IClusteringStrategy.cs`
- `Services/Clustering/Strategy/FloorCircularClusteringStrategy.cs`
- `Services/Clustering/Strategy/FloorRectangularClusteringStrategy.cs`
- `Services/Clustering/Strategy/WallAxisAlignedStrategy.cs`
- `Services/Clustering/Strategy/WallRotatedClusteringStrategy.cs`
- `Services/Clustering/Strategy/ClusteringStrategyFactory.cs`

**Files to Modify**:
- `Services/UniversalClusterService.cs` (replace clustering logic)

---

### Phase 5: Extract ClusterPlacementService

**Agent**: Placement Agent

**Tasks**:
1. Create `Services/Clustering/Placement/IClusterPlacementService.cs`
   - Interface: `PlaceClusterSleeve(...): (placed, deleted, placedSleeve)`
   - Interface: `SetSizeParameters(...): void`
   - Interface: `SetMetadata(...): void`

2. Create `Services/Clustering/Placement/ClusterPlacementService.cs`
   - Extract `PlaceClusterSleeve` method (8 methods total)
   - Extract `SetClusterSizeParameters` method
   - Extract `SetClusterSleeveMetadata` method
   - Move caches: `_mepElementCache`, `_bboxCache`, `_parameterCache`

3. Create `Services/Clustering/Placement/ClusterSleeveBuilder.cs`
   - Extract cluster sleeve creation logic
   - Handle family loading, parameter setting

4. Update `UniversalClusterService.cs`
   - Replace placement logic with service calls
   - Inject `IClusterPlacementService` via constructor

**Dependencies**: Phase 3 (uses `IBoundingBoxCalculator`), Phase 4 (uses `IClusteringStrategy`)

**Success Criteria**:
- ✅ All placement methods moved to service
- ✅ Caches moved with service
- ✅ Unit tests pass for placement service
- ✅ Integration test: Cluster sleeves created correctly
- ✅ Performance: Caching still works (no regression)
- ✅ ~30 warnings reduced (null-safety layer added)

**Files to Create**:
- `Services/Clustering/Placement/IClusterPlacementService.cs`
- `Services/Clustering/Placement/ClusterPlacementService.cs`
- `Services/Clustering/Placement/ClusterSleeveBuilder.cs`

**Files to Modify**:
- `Services/UniversalClusterService.cs` (replace placement logic, add dependency injection)

---

### Phase 6: Extract ClusterRotationService

**Agent**: Rotation Agent

**Tasks**:
1. Create `Services/Clustering/Rotation/IClusterRotationService.cs`
   - Interface: `DetermineAngle(cluster): double`
   - Interface: `CalculateRotatedBoundingBox(cluster, angle): BoundingBoxResult`

2. Create `Services/Clustering/Rotation/ClusterRotationService.cs`
   - Extract `DetermineDominantRotationAngle` method (9 methods total)
   - Extract `CalculateRotatedBoundingBox` method
   - Move `_clusterRotationData` dictionary

3. Integrate `RotationMatrixCalculator` from Phase 1
   - Use existing calculator (don't duplicate)

4. Update `UniversalClusterService.cs`
   - Replace rotation logic with service calls
   - Inject `IClusterRotationService` via constructor

**Dependencies**: Phase 1 (uses `RotationMatrixCalculator`), Phase 3 (uses `IBoundingBoxCalculator`)

**Success Criteria**:
- ✅ All rotation methods moved to service
- ✅ Rotation data dictionary moved with service
- ✅ Unit tests pass for rotation service
- ✅ Integration test: Rotated cluster dimensions match baseline
- ✅ Performance: Pre-calculated rotation matrix used
- ✅ ~20 warnings reduced (null-safety layer added)

**Files to Create**:
- `Services/Clustering/Rotation/IClusterRotationService.cs`
- `Services/Clustering/Rotation/ClusterRotationService.cs`

**Files to Modify**:
- `Services/UniversalClusterService.cs` (replace rotation logic, add dependency injection)

---

### Phase 7: Extract ClusterCleanupService

**Agent**: Cleanup Agent

**Tasks**:
1. Create `Services/Clustering/Cleanup/IClusterCleanupService.cs`
   - Interface: `DeleteSleevesInCluster(...): int`
   - Interface: `CleanupSleeves(...): int`
   - Interface: `ResetFlags(...): void`

2. Create `Services/Clustering/Cleanup/ClusterCleanupService.cs`
   - Extract `DeleteIndividualSleevesInCluster` method
   - Extract `CleanupSleevesWithinClusters` method
   - Extract `ResetClusterFlagsForDeletedSleeves` method
   - **CRITICAL**: Preserve cluster sleeve protection (HashSet-based)

3. Create `Services/Clustering/Cleanup/ClusterSleeveProtectionValidator.cs`
   - Extract validation logic for cluster sleeve protection
   - Multiple safety checks (primary, secondary, tertiary)

4. Update `UniversalClusterService.cs`
   - Replace cleanup logic with service calls
   - Inject `IClusterCleanupService` via constructor

**Dependencies**: Phase 5 (uses `IClusterPlacementService` for cluster IDs)

**Success Criteria**:
- ✅ All cleanup methods moved to service
- ✅ Cluster sleeve protection preserved (HashSet-based)
- ✅ Unit tests pass for cleanup service
- ✅ Integration test: Cluster sleeves NOT deleted, individual sleeves ARE deleted
- ✅ ~25 warnings reduced (null-safety layer added)

**Files to Create**:
- `Services/Clustering/Cleanup/IClusterCleanupService.cs`
- `Services/Clustering/Cleanup/ClusterCleanupService.cs`
- `Services/Clustering/Cleanup/ClusterSleeveProtectionValidator.cs`

**Files to Modify**:
- `Services/UniversalClusterService.cs` (replace cleanup logic, add dependency injection)

---

### Phase 8: Extract ClusterAlgorithmService

**Agent**: Algorithm Agent

**Tasks**:
1. Create `Services/Clustering/Algorithm/IClusterAlgorithmService.cs`
   - Interface: `FormClusters(clashZones, category): Dictionary<SleeveGroupKey, List<List<dynamic>>>`
   - **Pure algorithms**: No I/O dependencies (no database, no XML, no Revit API)

2. Create `Services/Clustering/Algorithm/ClusterAlgorithmService.cs`
   - Extract `FormClusters` method (15 methods total)
   - Extract `CalculateClustersUsingXmlData` method
   - **CRITICAL**: Preserve spatial grid logic (O(n×k) complexity)
   - **CRITICAL**: Preserve multi-threading (group-level + within-group)

3. Create `Services/Clustering/Algorithm/FloodFillClusteringAlgorithm.cs`
   - Extract flood-fill clustering logic
   - BFS-based clustering

4. Create `Services/Clustering/Algorithm/SpatialGridClusteringAlgorithm.cs`
   - Extract `BuildSpatialGrid` method
   - Extract `FormClustersFromGrid` method
   - Preserve O(n×k) complexity

5. Create `Services/Clustering/Algorithm/WithinGroupParallelProcessor.cs`
   - Extract within-group parallelization logic
   - Split sleeves into chunks, process in parallel
   - Maximize CPU usage even for single group

6. Update `UniversalClusterService.cs`
   - Replace algorithm logic with service calls
   - Inject `IClusterAlgorithmService` via constructor

**Dependencies**: Phase 2 (uses `IProximityChecker`), Phase 4 (uses `IClusteringStrategy`)

**Success Criteria**:
- ✅ All algorithm methods moved to service
- ✅ Pure algorithms (no I/O dependencies)
- ✅ Spatial grid preserved (O(n×k) complexity)
- ✅ Multi-threading preserved (group-level + within-group)
- ✅ Unit tests pass for algorithm service
- ✅ Integration test: Clusters match baseline
- ✅ Performance: Clustering time matches baseline (2-4× speedup)
- ✅ ~30 warnings reduced (null-safety layer added)

**Files to Create**:
- `Services/Clustering/Algorithm/IClusterAlgorithmService.cs`
- `Services/Clustering/Algorithm/ClusterAlgorithmService.cs`
- `Services/Clustering/Algorithm/FloodFillClusteringAlgorithm.cs`
- `Services/Clustering/Algorithm/SpatialGridClusteringAlgorithm.cs`
- `Services/Clustering/Algorithm/WithinGroupParallelProcessor.cs`

**Files to Modify**:
- `Services/UniversalClusterService.cs` (replace algorithm logic, add dependency injection)

---

### Phase 9: Extract Data Access + Add Recovery System

**Agent**: Integration Agent + Recovery Agent

**Tasks**:
1. Create `Services/Clustering/Data/IClusterDataRepository.cs`
   - Interface: `LoadExistingClusters(...): List<Cluster>`
   - Interface: `SaveClusterData(...): void`
   - Interface: `UpdateClusterFlags(...): void`

2. Create `Services/Clustering/Data/ClusterDataRepository.cs`
   - Extract `LoadClashZonesFromXml` method
   - Extract `SaveClusterDataToDatabase` method
   - Extract `MarkClusterResolved` method
   - Set `CommandTimeout = 60` for batch operations
   - **NEW**: Add transaction safety (try-finally with rollback)

3. Create `Services/Clustering/Caching/IClusteringCache.cs`
   - Interface: `GetClashZones(...): List<ClashZone>`
   - Interface: `Invalidate(): void`

4. Create `Services/Clustering/Caching/ClusteringCache.cs`
   - Extract caching logic
   - Cache invalidation strategy

5. **NEW**: Create `Services/Clustering/Recovery/ClusteringProgressTracker.cs`
   - Track successful placements, failures, skipped clusters
   - Record success: `RecordSuccess(clusterIndex, sleeveCount, clusterSleeveId)`
   - Record failure: `RecordFailure(clusterIndex, sleeveCount, error)`
   - Generate detailed report: `GenerateReport()`
   - Save checkpoint every 10 successful placements

6. **NEW**: Create `Services/Clustering/Recovery/ClusteringCheckpoint.cs`
   - Store checkpoint data in database: `ClusteringCheckpoints` table
   - Support resume from last checkpoint on crash/timeout
   - Allow retry of failed clusters only (skip successful ones)

7. Refactor `UniversalClusterService` as Orchestrator (~500 lines)
   - Inject all services via constructor (dependency injection)
   - Implement PATH 1/2/3 coordination logic
   - Manage transactions
   - Coordinate between services
   - **NEW**: Add progress tracking to orchestrator
   - **NEW**: Integrate checkpoint/resume functionality

8. Update all callers
   - Update `OpeningCommandOrchestrator` to use orchestrator
   - Update other callers if any

**Partial Failure Recovery Features**:
```csharp
public class ClusteringProgressTracker
{
    public List<ClusterPlacementResult> SuccessfulPlacements { get; } = new();
    public List<ClusterPlacementFailure> FailedPlacements { get; } = new();
    public List<ClusterPlacementSkipped> SkippedPlacements { get; } = new();
    
    public void RecordSuccess(int clusterIndex, int sleeveCount, int clusterSleeveId)
    {
        SuccessfulPlacements.Add(new ClusterPlacementResult 
        { 
            ClusterIndex = clusterIndex, 
            SleeveCount = sleeveCount,
            ClusterSleeveId = clusterSleeveId,
            Timestamp = DateTime.Now
        });
        
        // Save checkpoint every 10 successful placements
        if (SuccessfulPlacements.Count % 10 == 0)
        {
            SaveCheckpoint();
        }
    }
    
    public void RecordFailure(int clusterIndex, int sleeveCount, Exception error)
    {
        FailedPlacements.Add(new ClusterPlacementFailure 
        { 
            ClusterIndex = clusterIndex, 
            SleeveCount = sleeveCount,
            Error = error.Message,
            StackTrace = error.StackTrace,
            Timestamp = DateTime.Now
        });
        
        SafeFileLogger.SafeAppendText("clustering_failures.log",
            $"Cluster {clusterIndex} failed ({sleeveCount} sleeves): {error.Message}");
    }
    
    public string GenerateReport()
    {
        return $"Clustering Report:\n" +
               $"  Placed: {SuccessfulPlacements.Count} clusters\n" +
               $"  Failed: {FailedPlacements.Count} clusters\n" +
               $"  Skipped: {SkippedPlacements.Count} clusters\n" +
               $"  Total Sleeves Clustered: {SuccessfulPlacements.Sum(p => p.SleeveCount)}";
    }
}
```

**Checkpoint/Resume Support**:
- ✅ Save progress every 10 clusters
- ✅ Resume from last checkpoint on crash/timeout
- ✅ Allow retry of failed clusters only (skip successful ones)
- ✅ Store checkpoint in database: `ClusteringCheckpoints` table

**Dependencies**: All previous phases (uses all extracted services)

**Success Criteria**:
- ✅ `UniversalClusterService` reduced to ~500 lines (orchestrator)
- ✅ All services injected via constructor
- ✅ PATH 1/2/3 logic implemented
- ✅ Transaction management working with rollback safety
- ✅ Progress tracking integrated
- ✅ Checkpoint/resume functionality working
- ✅ Unit tests pass for repository, cache, and recovery system
- ✅ Integration test: All paths work correctly
- ✅ **NEW**: Test checkpoint save/resume
- ✅ **NEW**: Test failure recovery (place 42/50, retry failed 8)
- ✅ **NEW**: Test timeout recovery (partial success)
- ✅ ~15 warnings reduced (null-safety layer added)

**Files to Create**:
- `Services/Clustering/Data/IClusterDataRepository.cs`
- `Services/Clustering/Data/ClusterDataRepository.cs`
- `Services/Clustering/Data/ClusterData.cs`
- `Services/Clustering/Caching/IClusteringCache.cs`
- `Services/Clustering/Caching/ClusteringCache.cs`

**Files to Modify**:
- `Services/UniversalClusterService.cs` (refactor to orchestrator, add dependency injection)
- `Services/OpeningCommandOrchestrator.cs` (update to use orchestrator)

---

## Critical Implementation Rules

### 1. Preserve All Optimizations

**DO**:
- ✅ Use pre-calculated data from database (corners, rotation matrix, bounding boxes)
- ✅ Preserve spatial grid (O(n×k) complexity)
- ✅ Preserve multi-threading (group-level + within-group)
- ✅ Preserve caching (move caches with their consuming services)
- ✅ Preserve "dump once, use many times" principle

**DON'T**:
- ❌ Recalculate data that's already in database
- ❌ Remove spatial grid optimization
- ❌ Remove multi-threading
- ❌ Invalidate caches unnecessarily

### 2. Null-Safety Requirements

**Each service must**:
- ✅ Add nullable annotations to all method parameters
- ✅ Add guard clauses for null checks (fail-fast)
- ✅ Use null-conditional operators (`?.`) for property access
- ✅ Use null-coalescing operators (`??`) for default values
- ✅ Validate Revit elements with `IsValidObject` before access
- ✅ Return nullable results for invalid operations
- ✅ Target: Reduce warnings by ~20-30 per service

### 3. Testing Requirements

**Each phase must**:
- ✅ Create unit tests for new classes (including null/invalid inputs)
- ✅ Create integration test to verify behavior matches baseline
- ✅ **NEW**: Create regression test suite (prevent future breakage)
- ✅ **NEW**: Test timeout scenarios (verify timeout triggers correctly)
- ✅ **NEW**: Test failure recovery (partial success scenarios)
- ✅ Run all tests before moving to next phase
- ✅ Fix any failing tests before proceeding
- ✅ **NEW**: Capture baseline performance metrics before refactoring
- ✅ **NEW**: Verify performance matches baseline (±5% tolerance)

### 4. Performance Requirements

**Each phase must**:
- ✅ **NEW**: Capture baseline metrics BEFORE refactoring
- ✅ Verify no performance regression (±5% tolerance)
- ✅ Baseline: 1000 sleeves in 2.5s
- ✅ Target: Same or better performance
- ✅ Measure and log performance metrics to `performance_baseline.json`
- ✅ **NEW**: Fail phase if performance regression detected (>5%)
- ✅ **NEW**: Verify timeout overhead <1% (for timeout-protected phases)

### 5. Code Quality Requirements

**Each phase must**:
- ✅ Follow existing code style
- ✅ Add XML documentation comments
- ✅ Use meaningful variable names
- ✅ Keep methods focused (single responsibility)
- ✅ Avoid code duplication

### 6. Rollback Strategy (NEW - CRITICAL)

**Each phase must**:
- ✅ Keep old methods as `[Obsolete]` during migration
- ✅ Add feature flag: `USE_REFACTORED_[PHASE] = false` for instant rollback
- ✅ Remove old code only after successful production validation
- ✅ Document rollback procedure in phase notes

---

## Multi-Agent Coordination

### Communication Protocol

1. **Before Starting Phase**:
   - Read `REFACTORING_PLAN_UniversalClusterService.md` for full context
   - Check dependencies (previous phases must be complete)
   - Review success criteria for the phase

2. **During Implementation**:
   - Create files in correct namespace structure
   - Follow existing code patterns
   - Preserve all optimizations
   - Add null-safety layer

3. **After Completing Phase**:
   - Run all tests (unit + integration)
   - Verify no warnings introduced
   - Verify performance matches baseline
   - Document any deviations from plan

### Conflict Resolution

**If multiple agents work on same file**:
- Use file locking or sequential processing
- Agent 1 completes → Agent 2 reviews → Agent 2 continues
- Never edit same file simultaneously

**If dependencies conflict**:
- Complete dependencies first
- Wait for previous phase to be fully tested
- Don't proceed until dependencies are stable

---

## Success Metrics

### Overall Success Criteria

- ✅ `UniversalClusterService` reduced from 8,530 lines to ~500 lines
- ✅ 5 focused services created (~2,000 total lines)
- ✅ 0 critical warnings (CS8602, CS8604, CS8600, CS8625, CS8603)
- ✅ < 50 total warnings (down from 2,172)
- ✅ All unit tests pass
- ✅ All integration tests pass
- ✅ Performance matches or improves baseline
- ✅ All optimizations preserved
- ✅ Revit 2024 compatibility maintained

### Phase-Specific Success Criteria

See individual phase sections above.

---

## File Structure Reference

```
Services/
├── Clustering/
│   ├── Geometry/
│   │   ├── DistanceCalculator.cs
│   │   ├── RotationMatrixCalculator.cs
│   │   ├── CoordinateTransformer.cs
│   │   ├── IBoundingBoxCalculator.cs
│   │   ├── AxisAlignedBoundingBoxCalculator.cs
│   │   ├── RotatedBoundingBoxCalculator.cs
│   │   └── CornerBasedBoundingBoxCalculator.cs
│   │
│   ├── Proximity/
│   │   ├── IProximityChecker.cs
│   │   ├── BoundingBoxProximityChecker.cs
│   │   ├── EdgeToEdgeProximityChecker.cs
│   │   ├── RotatedProximityChecker.cs
│   │   └── ProximityCheckerFactory.cs
│   │
│   ├── Strategy/
│   │   ├── IClusteringStrategy.cs
│   │   ├── FloorCircularClusteringStrategy.cs
│   │   ├── FloorRectangularClusteringStrategy.cs
│   │   ├── WallAxisAlignedStrategy.cs
│   │   ├── WallRotatedClusteringStrategy.cs
│   │   └── ClusteringStrategyFactory.cs
│   │
│   ├── Placement/
│   │   ├── IClusterPlacementService.cs
│   │   ├── ClusterPlacementService.cs
│   │   └── ClusterSleeveBuilder.cs
│   │
│   ├── Rotation/
│   │   ├── IClusterRotationService.cs
│   │   └── ClusterRotationService.cs
│   │
│   ├── Cleanup/
│   │   ├── IClusterCleanupService.cs
│   │   ├── ClusterCleanupService.cs
│   │   └── ClusterSleeveProtectionValidator.cs
│   │
│   ├── Algorithm/
│   │   ├── IClusterAlgorithmService.cs
│   │   ├── ClusterAlgorithmService.cs
│   │   ├── FloodFillClusteringAlgorithm.cs
│   │   ├── SpatialGridClusteringAlgorithm.cs
│   │   └── WithinGroupParallelProcessor.cs
│   │
│   ├── Data/
│   │   ├── IClusterDataRepository.cs
│   │   ├── ClusterDataRepository.cs
│   │   └── ClusterData.cs
│   │
│   └── Caching/
│       ├── IClusteringCache.cs
│       └── ClusteringCache.cs
│
└── UniversalClusterService.cs (refactored to ~500 lines orchestrator)
```

---

## Quick Reference: Key Methods to Extract

### From UniversalClusterService.cs

**Geometry** (Phase 1):
- `CalculateMinimumDistance2D` → `DistanceCalculator`
- `CalculateMinimumDistance3D` → `DistanceCalculator`
- Rotation matrix logic → `RotationMatrixCalculator`
- Coordinate transformation → `CoordinateTransformer`

**Proximity** (Phase 2):
- `BoundingBoxesOverlapFromXml` → `BoundingBoxProximityChecker`
- Edge-to-edge logic → `EdgeToEdgeProximityChecker`
- `CheckRotatedSleeveProximity` → `RotatedProximityChecker`

**Bounding Box** (Phase 3):
- `GetClusterBoundingBoxFromXml` → `AxisAlignedBoundingBoxCalculator`
- `GetClusterBoundingBoxWithRotatedCoordinates` → `RotatedBoundingBoxCalculator`
- Corner-based algorithm → `CornerBasedBoundingBoxCalculator`

**Clustering** (Phase 4):
- `FormClusters` → Strategy implementations
- `CalculateClustersUsingXmlData` → Strategy implementations
- Floor/wall specific logic → Respective strategies

**Placement** (Phase 5):
- `PlaceClusterSleeve` → `ClusterPlacementService`
- `SetClusterSizeParameters` → `ClusterPlacementService`
- `SetClusterSleeveMetadata` → `ClusterPlacementService`

**Rotation** (Phase 6):
- `DetermineDominantRotationAngle` → `ClusterRotationService`
- `CalculateRotatedBoundingBox` → `ClusterRotationService`

**Cleanup** (Phase 7):
- `DeleteIndividualSleevesInCluster` → `ClusterCleanupService`
- `CleanupSleevesWithinClusters` → `ClusterCleanupService`
- `ResetClusterFlagsForDeletedSleeves` → `ClusterCleanupService`

**Algorithm** (Phase 8):
- `FormClusters` → `ClusterAlgorithmService`
- `BuildSpatialGrid` → `SpatialGridClusteringAlgorithm`
- `FormClustersFromGrid` → `SpatialGridClusteringAlgorithm`

**Data** (Phase 9):
- `LoadClashZonesFromXml` → `ClusterDataRepository`
- `SaveClusterDataToDatabase` → `ClusterDataRepository`
- Caching logic → `ClusteringCache`

---

## Notes for Multi-Agent Implementation

1. **Cursor Multi-Agent Support**: Yes, Cursor supports multi-agent workflows. Agents can work on different files/phases simultaneously, but must coordinate on shared files.

2. **Sequential Phases**: Phases 1-9 must be completed sequentially (each phase depends on previous).

3. **Parallel Work Within Phase**: Within a phase, multiple agents can work on different files (e.g., one agent creates interface, another creates implementation).

4. **Testing**: Each phase must be tested before proceeding. Use `Testing Agent` to create tests after each phase.

5. **Code Review**: After each phase, review code for:
   - Correctness (matches baseline behavior)
   - Performance (no regression)
   - Warnings (reduced, not increased)
   - Code quality (follows patterns)

6. **Documentation**: Update `REFACTORING_PLAN_UniversalClusterService.md` with any deviations or learnings.

---

**Last Updated**: [Current Date]  
**Version**: 1.0  
**Status**: Ready for Multi-Agent Implementation


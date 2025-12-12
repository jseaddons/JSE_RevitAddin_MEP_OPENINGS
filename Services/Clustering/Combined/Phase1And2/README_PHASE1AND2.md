# Phase 1-2: Discovery & Formation Infrastructure (Agent 1)

**Agent**: This Agent  
**Status**: 📋 Ready for Implementation  
**Scope**: Data Model Extensions + Multi-Threaded Discovery Services  
**Files**: See file list below  
**Estimated Duration**: 3-4 days  
**Flag**: `OptimizationFlags.UseCombinedClusteringPhase1And2`

---

## Overview

Phase 1-2 handles the **discovery and grouping** of existing cluster sleeves from the database. This is **purely CPU-bound database operations** - no Revit API calls. It's fully thread-safe and can be heavily parallelized.

### What Phase 1-2 Does:

1. **Phase 1: Infrastructure**
   - Extend `ClashZone` model with combined cluster fields
   - Define SOLID service interfaces
   - Create DTOs (data transfer objects)
   - Build repository for database queries

2. **Phase 2: Multi-Threaded Discovery**
   - Load all cluster sleeves from database (parallel by category)
   - Group by host type + orientation (for combining)
   - Calculate proximity matrices (CPU-intensive, parallelizable)
   - Form proximity-based cluster candidates

### Output for Agent 2:

Returns `List<CombinedClusterCandidate>` ready for family instance creation.

---

## Files to Create

### 1. Interfaces (SOLID Contracts)

**File**: `Interfaces/ICombinedClusterDiscovery.cs`
- `DiscoverClusterSleevesAsync()` - Load clusters from DB
- `GroupByHostTypeAndOrientation()` - Group for combining

**File**: `Interfaces/ICombinedClusterFormation.cs`
- `FormCombinedClustersAsync()` - Find nearby clusters
- `FindIndividualSleevesNearCombinedCluster()` - Find sleeves to incorporate

---

### 2. Data Models (DTOs)

**File**: `Models/ClusterSleeveInfo.cs`
- Immutable DTO representing a single cluster sleeve
- Contains ID, category, host type, orientation, bounding box, parameters

**File**: `Models/CombinedClusterCandidate.cs`
- Represents a candidate for combining
- Contains list of member clusters + individual sleeves to incorporate
- Aggregated geometry + categories

**File**: `Models/AggregatedParameterSnapshot.cs`
- Stores appended parameter snapshots from all contributing sleeves
- Per-category breakdowns + statistics

---

### 3. Repository (Database Access)

**File**: `Repository/CombinedClusterRepository.cs`
- `GetAllClusterSleeves()` - Query all cluster sleeves
- `FindIndividualSleevesNearBoundingBox()` - Query individual sleeves
- Helper methods for spatial calculations

---

### 4. Services (Business Logic)

**File**: `Services/CombinedClusterDiscoveryService.cs`
- Implements `ICombinedClusterDiscovery`
- Multi-threaded cluster discovery
- Grouping logic

**File**: `Services/CombinedClusterFormationService.cs`
- Implements `ICombinedClusterFormation`
- Proximity matrix calculation (parallelized)
- Cluster formation algorithm
- Individual sleeve incorporation

---

## Data Model Extensions Required

### Extend ClashZone Model

Add these fields to `Data/Models/ClashZone.cs`:

```csharp
// NEW: Combined Cluster Tracking (Phase 1 addition)
public int CombinedClusterSleeveInstanceId { get; set; } = -1;
public string CategoriesInCombinedCluster { get; set; } = string.Empty;
public double CombinedClusterSleeveBoundingBoxMinX { get; set; } = 0.0;
public double CombinedClusterSleeveBoundingBoxMinY { get; set; } = 0.0;
public double CombinedClusterSleeveBoundingBoxMinZ { get; set; } = 0.0;
public double CombinedClusterSleeveBoundingBoxMaxX { get; set; } = 0.0;
public double CombinedClusterSleeveBoundingBoxMaxY { get; set; } = 0.0;
public double CombinedClusterSleeveBoundingBoxMaxZ { get; set; } = 0.0;
public bool IsIncorporatedInCombinedCluster { get; set; } = false;
public string CombinedClusterParameterSnapshot { get; set; } = string.Empty;
```

### Extend OptimizationFlags

Add to `Services/OptimizationFlags.cs`:

```csharp
#region Combined Clustering - Agent 1 (Phase 1-2)

public static bool UseCombinedClusteringPhase1And2 { get; set; } = false;
public static double CombinedClusteringProximityTolerance { get; set; } = 100.0;  // mm
public static int CombinedClusteringThreadCount { get; set; } = 
    Environment.ProcessorCount / 2;

#endregion
```

---

## Code Architecture

### Discovery Workflow:

```
DiscoverClusterSleevesAsync()
  ├─ Check phase flag
  ├─ Partition categories across threads (e.g., 4 categories → 2 threads)
  ├─ For each thread:
  │  └─ CombinedClusterRepository.GetAllClusterSleeves(category_batch)
  ├─ Merge results from all threads
  └─ Return: List<ClusterSleeveInfo>

GroupByHostTypeAndOrientation()
  └─ Group by "HostType_Orientation" key
     └─ Return: Dictionary<string, List<ClusterSleeveInfo>>

FormCombinedClustersAsync()
  ├─ For each group:
  │  ├─ CalculateProximityMatrixAsync() [parallelized]
  │  ├─ FormClustersFromProximityMatrix() [greedy algorithm]
  │  └─ Filter for multi-category candidates
  └─ Return: List<CombinedClusterCandidate>
```

### Multi-Threading Safety:

- ✅ **Safe**: Database queries only (no Revit API)
- ✅ **Safe**: CPU calculations (proximity, grouping)
- ✅ **Thread-safe**: No shared state
- ✅ **Parallelizable**: Embarrassingly parallel workloads

---

## Testing Strategy

### Unit Tests:

1. **DiscoveryService Tests**
   - `DiscoverClusterSleevesAsync_WithValidData_ReturnsAllClusters()`
   - `DiscoverClusterSleevesAsync_WithMultipleCategories_ParallelWorks()`
   - `GroupByHostTypeAndOrientation_WithMixedData_GroupsCorrectly()`

2. **FormationService Tests**
   - `FormCombinedClustersAsync_WithNearClusters_CombinesThem()`
   - `CalculateProximityMatrix_Parallelized_ReturnsCorrectDistances()`
   - `FindIndividualSleevesNearCombinedCluster_WithNearSleeves_ReturnsAll()`

3. **Repository Tests**
   - `GetAllClusterSleeves_WithValidCategories_ReturnsAll()`
   - `FindIndividualSleevesNearBoundingBox_WithNearSleeves_ReturnsAll()`

### Mocking:

- Mock `SleeveDbContext` for unit tests
- Create test data: ClusterSleeveInfo, ClashZone objects
- No database dependency for tests

---

## Performance Targets

| Operation | Single-Threaded | Multi-Threaded | Target Gain |
|-----------|-----------------|----------------|-----------|
| Discover 100 clusters | 100ms | 50ms | 50% faster |
| Proximity matrix 100 clusters | 200ms | 100ms | 50% faster |
| Formation algorithm | 50ms | 50ms | 0% (not parallelizable) |
| **Total** | **350ms** | **200ms** | **43% faster** |

---

## Integration Points (Agent 1 → Agent 2)

**Output Interface**:
```csharp
List<CombinedClusterCandidate> candidates = new();

// For each candidate, Agent 2 will:
foreach (var candidate in candidates)
{
    // Phase 3: Aggregate parameters (Agent 2)
    candidate.AggregatedSnapshot = aggregator.AggregateParameters(candidate, individual sleeves);
    
    // Phase 4: Create family + persist (Agent 2)
    executor.CreateFamilyInstance(candidate);
}
```

Agent 1 doesn't need to know what Agent 2 does - just return clean candidates.

---

## Acceptance Criteria

- ✅ All 6 files created (2 interfaces, 3 models, 1 repository, 2 services)
- ✅ All methods implemented per specification
- ✅ Multi-threading working (4+ threads detected, parallel execution confirmed)
- ✅ Feature flag functional: `UseCombinedClusteringPhase1And2`
- ✅ Unit tests >80% coverage
- ✅ No breaking changes to existing code
- ✅ Performance: 40%+ faster than single-threaded version
- ✅ Returns `List<CombinedClusterCandidate>` ready for Agent 2

---

## Dependencies

**Required (must exist):**
- ✅ `SleeveDbContext` (existing)
- ✅ `ClashZone` model (need to extend)
- ✅ `OptimizationFlags` (need to add flags)

**Provided (Agent 1 provides):**
- ✅ All service interfaces
- ✅ All DTOs
- ✅ All implementations

**Will be used by:**
- ✅ Agent 2 (takes candidates as input)
- ✅ Integration layer (orchestrates workflow)

---

## Document Versions

- **Phase 1-2 README**: This document
- **Master Plan**: COMBINED_SLEEVES_AGENT_SPLIT_TASK.md
- **Full Architecture**: COMBINED_SLEEVES_IMPLEMENTATION_PLAN_*.md

---

## Next Steps

1. ✅ Agent 1 creates Phase 1-2 code (this agent)
2. ⏳ Agent 2 creates Phase 3-4 code (other agent)
3. ✅ Both meet acceptance criteria
4. ✅ Integration layer wires both together
5. ✅ End-to-end testing
6. ✅ UI button activation


# Phase 3-4: Creation & Persistence Infrastructure (Agent 2)

**Agent**: Other Agent  
**Status**: 📋 Ready for Implementation  
**Scope**: Parameter Aggregation + Batch Writing + Transaction Safety  
**Files**: See file list below  
**Estimated Duration**: 3-4 days  
**Flag**: `OptimizationFlags.UseCombinedClusteringPhase3And4`

---

## Overview

Phase 3-4 handles **creation and persistence** of combined cluster sleeves. This uses **Revit API** for family instance creation and is **single-threaded** due to Revit API constraints.

### What Phase 3-4 Does:

1. **Phase 3: Parameter Aggregation**
   - Collect parameter snapshots from all contributing sleeves (append model)
   - Aggregate statistics (count, totals, averages)
   - Create combined metadata
   - Prepare parameter set for family instance

2. **Phase 4: Batch Writing + Transaction Safety**
   - Create family instances (Revit API)
   - Set parameters on instances
   - Bulk write to database (single transaction)
   - Update XML files
   - Handle timeout + rollback

### Input from Agent 1:

Receives `List<CombinedClusterCandidate>` from Phase 1-2.

---

## Files to Create

### 1. Interfaces (SOLID Contracts)

**File**: `Interfaces/IParameterAggregator.cs`
- `AggregateParameters()` - Collect & append snapshots
- `CreateCombinedSleeveParameterSet()` - Build parameter set

**File**: `Interfaces/ICombinedClusterPersistence.cs`
- `QueueDatabaseUpdates()` - Queue (don't write yet)
- `UpdateXmlWithCombinedClusterInfo()` - Update XML

---

### 2. Services (Business Logic)

**File**: `Services/ParameterAggregatorService.cs`
- Implements `IParameterAggregator`
- Append model implementation
- Parameter collection from clusters + individual sleeves
- Statistics calculation

**File**: `Services/CombinedClusterPersistenceService.cs`
- Implements `ICombinedClusterPersistence`
- Database update queueing
- XML update logic
- Global Index + Filter XML updates

**File**: `Services/CombinedClusterBatchExecutor.cs`
- Execute combined cluster creation
- Single transaction wrapper
- Timeout protection (5 minutes)
- Batch processing (e.g., 50 clusters per transaction)
- Crash-safe rollback

---

## Append Model (Parameter Strategy)

### What "Append" Means:

Instead of transferring individual parameters one-by-one, **collect all snapshots and store as historical record**:

```
Combined Sleeve = Historical Record of All Components

├─ Snapshot from Ducts Cluster:
│  ├─ Duct 1: {Diameter=100mm, Velocity=3.5m/s, ...}
│  ├─ Duct 2: {Diameter=150mm, Velocity=2.8m/s, ...}
│  └─ Aggregated: {Count=2, AvgDiameter=125mm, ...}
│
├─ Snapshot from Pipes Cluster:
│  ├─ Pipe 1: {Diameter=50mm, Pressure=5bar, ...}
│  └─ Aggregated: {Count=1, ...}
│
└─ Snapshot from Individual CableTray (if incorporated):
   └─ CableTray: {Width=300mm, ...}

Result: Combined sleeve contains complete history
```

### Implementation:

```csharp
public class AggregatedParameterSnapshot
{
    // Per-category snapshots (APPEND MODEL)
    public Dictionary<string, List<Dictionary<string, object>>> 
        CategoryParameterSnapshots { get; set; } = new();
    
    // "Ducts" -> [param1, param2, ...]
    // "Pipes" -> [param1, ...]
    // Each category has list of parameter sets (append)
}
```

---

## Crash-Safe Transaction Strategy

### Goal: All-or-Nothing Execution

```csharp
using (var txn = new Transaction(doc, "Create Combined Clusters"))
{
    txn.Start();
    
    try
    {
        // Step 1: Create all family instances
        // Step 2: Set all parameters
        // Step 3: Bulk write to database
        // (If ANY step fails, entire transaction rolls back)
        
        txn.Commit();  // All-or-nothing
    }
    catch
    {
        txn.RollBack();  // Automatic rollback on error
        throw;
    }
}
```

### Timeout Protection:

```csharp
using (var cts = new CancellationTokenSource(300000))  // 5 minutes
{
    // If execution takes >5 minutes, throw OperationCanceledException
    // Transaction automatically rolls back
}
```

---

## Batch Writing Strategy

**Why Batch?**
- 50 combined clusters per transaction (configurable)
- Single write to database (not 50 separate writes)
- 4-6× faster than individual writes

**How Batch Works:**

```
Batch 1 (50 clusters):
  ├─ Create 50 family instances
  ├─ Set parameters on all 50
  ├─ Collect 50 × N zone updates (where N = zones per cluster)
  ├─ Bulk write to DB (1 SaveChanges call)
  └─ Transaction.Commit()

Batch 2 (50 clusters):
  ├─ [same as Batch 1]
  └─ [repeat until all processed]
```

---

## Testing Strategy

### Unit Tests:

1. **ParameterAggregator Tests**
   - `AggregateParameters_WithValidCandidates_ReturnsAggregatedSnapshot()`
   - `CreateCombinedSleeveParameterSet_WithAggregatedSnapshot_ReturnsDictionary()`
   - `AppendModel_MultipleCategories_MaintainsHistory()`

2. **PersistenceService Tests**
   - `QueueDatabaseUpdates_WithCombinedCluster_ReturnsUpdates()`
   - `UpdateXmlWithCombinedClusterInfo_WithValidData_UpdatesCorrectly()`

3. **BatchExecutor Tests**
   - `ExecuteCombinedClusterBatchAsync_WithValidCandidates_CreatesInstances()`
   - `Timeout_AfterFiveMinutes_RollsBack()`
   - `SingleTransaction_AllOrNothing_RollsBackOnError()`

### Mocking:

- Mock `Document` (Revit API)
- Mock `SleeveDbContext`
- Mock `Transaction` (use real or mock)
- Create test `CombinedClusterCandidate` objects

---

## Performance Targets

| Operation | Single-Threaded | Batch (50/txn) | Target Gain |
|-----------|-----------------|----------------|-----------|
| Parameter agg (100 clusters) | 3000ms | 3000ms | 0% (sequential) |
| Create 100 instances | 2000ms | 2000ms | 0% (Revit constraint) |
| DB writes (100 updates) | 5000ms | 1000ms | **80% faster** |
| **Total** | **10000ms** | **6000ms** | **40% faster** |

---

## Integration Points (Agent 1 → Agent 2)

**Input from Agent 1**:
```csharp
List<CombinedClusterCandidate> candidates = await discoveryService
    .DiscoverClusterSleevesAsync(...);

// Agent 2 receives these candidates:
foreach (var candidate in candidates)
{
    // candidate.MemberClusters - clusters to combine
    // candidate.IncorporatedIndividualSleeves - individual sleeves to add
    // candidate.CombinedBoundingBox - where to place
    // candidate.CategoriesInvolved - what categories
}
```

Agent 2 processes them and returns:
```csharp
(int CreatedCount, int ErrorCount) result = await batchExecutor
    .ExecuteCombinedClusterBatchAsync(candidates, ...);
```

---

## Acceptance Criteria

- ✅ All 3 files created (2 interfaces, 3 services)
- ✅ All methods implemented per specification
- ✅ Parameter aggregation working (append model verified)
- ✅ Family instances created (Revit API calls working)
- ✅ Single transaction wrapper functional
- ✅ Timeout protection working (5 minute limit enforced)
- ✅ Database bulk writes working (verified faster than individual)
- ✅ XML updates working (Global Index + Filter XML)
- ✅ Feature flag functional: `UseCombinedClusteringPhase3And4`
- ✅ Unit tests >80% coverage
- ✅ No breaking changes to existing code
- ✅ Crash-safe rollback working

---

## Dependencies

**Required (must exist):**
- ✅ `SleeveDbContext` (existing)
- ✅ `Document` object (passed in)
- ✅ `Transaction` API (Revit)
- ✅ Agent 1 output: `List<CombinedClusterCandidate>`

**Provided by Agent 1:**
- ✅ `CombinedClusterCandidate` DTO
- ✅ `AggregatedParameterSnapshot` DTO (or define here)
- ✅ `ClusterSleeveInfo` DTO

**Provided (Agent 2 provides):**
- ✅ `ParameterAggregatorService`
- ✅ `CombinedClusterPersistenceService`
- ✅ `CombinedClusterBatchExecutor`

---

## Integration Points (After Both Agents Complete)

```csharp
// In CombinedClusteringIntegration.cs:

// Agent 1 discovery
var candidates = await discoveryService.DiscoverClusterSleevesAsync(...);

// Agent 2 creation
var (created, errors) = await batchExecutor
    .ExecuteCombinedClusterBatchAsync(candidates, ...);
```

---

## Document Versions

- **Phase 3-4 README**: This document
- **Master Plan**: COMBINED_SLEEVES_AGENT_SPLIT_TASK.md
- **Full Architecture**: COMBINED_SLEEVES_IMPLEMENTATION_PLAN_*.md

---

## Next Steps

1. ⏳ Agent 1 creates Phase 1-2 code (other agent)
2. ✅ Agent 2 creates Phase 3-4 code (this agent)
3. ✅ Both meet acceptance criteria
4. ✅ Integration layer wires both together
5. ✅ End-to-end testing
6. ✅ UI button activation


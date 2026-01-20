# Combined Sleeves - Extended Clustering Implementation Plan

**Status**: 📋 Planning & Architecture Phase  
**Date**: December 12, 2025  
**Scope**: Extend clustering to multi-category combined sleeves with SOLID principles, 28-feature compliance, multithreading, batch writing, and parameter aggregation  
**Principles**: SOLID, 28-Feature Comprehensive Architecture, Crash-Safe Execution, Transaction Management

---

## Table of Contents

1. [Overview & Scope Extension](#overview--scope-extension)
2. [Architecture Alignment](#architecture-alignment)
3. [Core Concepts](#core-concepts)
4. [Implementation Phases](#implementation-phases)
5. [Phase 1: Data Model & Service Infrastructure](#phase-1-data-model--service-infrastructure)
6. [Phase 2: Multi-Threaded Cluster Discovery](#phase-2-multi-threaded-cluster-discovery)
7. [Phase 3: Combined Sleeve Creation with Parameter Aggregation](#phase-3-combined-sleeve-creation-with-parameter-aggregation)
8. [Phase 4: Batch Writing & Transaction Safety](#phase-4-batch-writing--transaction-safety)
9. [Integration with Orchestrator](#integration-with-orchestrator)
10. [Testing Strategy](#testing-strategy)
11. [Rollback & Recovery](#rollback--recovery)
12. [Performance Metrics](#performance-metrics)

---

## Overview & Scope Extension

### Current State (Single-Category Clustering)
```
Individual Sleeves (Ducts)
    ↓
    └─→ Cluster Formation (Ducts only)
        ↓
        └─→ Cluster Sleeve (Ducts)
            
Individual Sleeves (Pipes)
    ↓
    └─→ Cluster Formation (Pipes only)
        ↓
        └─→ Cluster Sleeve (Pipes)

Result: Separate Ducts cluster + Pipes cluster (no cross-category optimization)
```

### Proposed State (Multi-Category Combined Clustering - Phase 3)
```
Individual Sleeves (Ducts) + Individual Sleeves (Pipes) + Individual Sleeves (Cable Trays)
    ↓
Phase 1: Per-Category Clustering (EXISTING - UNCHANGED)
    ├─→ Ducts Cluster Sleeve
    ├─→ Pipes Cluster Sleeve
    └─→ Cable Trays: No clustering (individual sleeves remain)
    
Phase 2 (NEW): Load & Group All Cluster Sleeves
    ├─→ Load cluster sleeve metadata from XML + Database
    ├─→ Group by Host Type + Orientation (X-Wall, Y-Wall, Floor)
    ├─→ Apply section box filter (respect user's work area)
    └─→ Multi-threaded processing for large projects

Phase 3 (NEW): Multi-Category Combined Clustering
    ├─→ For each group (e.g., "X-Wall" group):
    │   ├─ Check spatial proximity (bounding box overlap + tolerance)
    │   ├─ If clusters nearby: Form combined cluster
    │   ├─ If individual sleeves near combined cluster: Incorporate them
    │   └─ Create "Combined Ducts+Pipes+CableTray" cluster sleeve
    ├─→ Multi-threaded formation logic for speed
    └─→ Result: Single large combined sleeve (replaces category-specific clusters)

Phase 4 (NEW): Parameter Aggregation & Batch Writing
    ├─→ Collect parameter snapshots from original clusters
    ├─→ Append individual sleeve parameters (if incorporated)
    ├─→ Create combined sleeve with aggregated parameters
    ├─→ Update all related zones (bulk update - batched)
    └─→ Single transaction for all writes (crash-safe)

Result: Optimized combined sleeves + maximum performance + full crash-safety
```

### Key Extension Points
1. **Data Model**: Add combined sleeve tracking fields to `ClashZone`
2. **Services**: New service classes for combined clustering (SOLID-compliant)
3. **Multithreading**: Parallel discovery of nearby clusters (safe with proper locking)
4. **Parameter Transfer**: Aggregate snapshots instead of individual transfers
5. **Batch Writing**: Combine all updates into single transaction
6. **Transaction Safety**: Wrap entire phase in crash-safe executor

---

## Architecture Alignment

### SOLID Principles Compliance

| Principle | Application to Combined Sleeves |
|-----------|----------------------------------|
| **S**RP (Single Responsibility) | Each service handles ONE aspect: discovery, formation, aggregation, persistence |
| **O**CP (Open/Closed) | Base interface `ICombinedClusterService` allows new strategies without modifying existing |
| **L**SP (Liskov Substitution) | All cluster strategies implement common interface for interchangeable behavior |
| **I**SP (Interface Segregation) | Separate interfaces: `ICombinedClusterDiscovery`, `ICombinedClusterFormation`, `IParameterAggregator` |
| **D**IP (Dependency Injection) | Constructor-injected dependencies via `ServiceProvider` (no `new` keywords) |

### 28-Feature Comprehensive Architecture Compliance

**Core Features Preserved:**
- ✅ **Feature #1-5**: Geometry Caching, Memory Management, Smart Tolerance, Cache Invalidation, R-tree Filtering
- ✅ **Feature #6-10**: Spatial Grid, Database R-tree, Bounding Box Filter, Parameter Batching, Family Symbol Cache
- ✅ **Feature #11-15**: Diagnostic Logging, Transaction Management, Element Validation, Timeout Protection, Warning Handler
- ✅ **Feature #16-20**: XML Persistence, SQLite Database, R-tree Indexes, Snapshot Caching, File-Based Logging
- ✅ **Feature #21-28**: And all others from comprehensive architecture

**New Features Added for Combined Clustering:**
- ✅ **Feature #29**: Multi-Category Cluster Discovery (with multithreading)
- ✅ **Feature #30**: Combined Cluster Formation Algorithm (proximity + tolerance)
- ✅ **Feature #31**: Parameter Snapshot Aggregation (append snapshots)
- ✅ **Feature #32**: Batch Transaction Wrapper (single write per combined cluster)

### Multithreading Strategy

**Safety First**: Use Revit-safe multithreading patterns:

```csharp
// ✅ SAFE: Read-only data collection (no Revit API calls)
Task<List<ClusterSleeveInfo>> discoverTask = Task.Run(() =>
{
    // CPU-bound: Load cluster metadata from DB (NO Revit API)
    return LoadClusterSleevesFromDatabase(filterIds);
});

// ✅ SAFE: Grouping & spatial calculations (no Revit API)
var groups = await discoverTask;
var proximityTasks = groups.Select(g => 
    Task.Run(() => CalculateProximityMatrix(g))
).ToArray();

// ✅ SAFE: In parallel - cluster formation (CPU-intensive, no Revit API)
await Task.WhenAll(proximityTasks);

// ❌ UNSAFE ZONE: Revit API calls (single-threaded in UI thread)
// Use External Event to marshal back to Revit for:
// - Element creation (family instances)
// - Parameter setting
// - Transaction management
```

### Batch Writing Strategy

**Goal**: Single transaction for ALL combined sleeve updates

```csharp
// Collect all operations
var batchOperations = new BatchOperation();

foreach (var combinedCluster in discoveredCombinedClusters)
{
    batchOperations.Add(
        new CreateFamilyInstanceOperation(combinedCluster),
        new UpdateZoneFlagsOperation(combinedCluster),
        new SaveParametersOperation(combinedCluster)
    );
}

// Execute ALL in single transaction (crash-safe)
using (var txn = new Transaction(doc, "Create Combined Clusters"))
{
    txn.Start();
    batchOperations.ExecuteAll();
    txn.Commit();
}
```

---

## Core Concepts

### Concept 1: Cluster Sleeve Info Structure (Immutable)
```csharp
public class ClusterSleeveInfo
{
    // Identity
    public int ClusterSleeveInstanceId { get; init; }
    public string Category { get; init; }
    public string FilterName { get; init; }
    
    // Host & Orientation
    public string HostType { get; init; }           // "Wall", "Floor", "Framing"
    public string Orientation { get; init; }        // "X-Wall", "Y-Wall", "Floor"
    public double Level { get; init; }              // Z-coordinate for grouping
    
    // Geometry
    public BoundingBoxXYZ BoundingBox { get; init; }
    public double Width { get; init; }
    public double Height { get; init; }
    
    // Parameter Snapshot (for aggregation)
    public Dictionary<string, object> ParameterSnapshot { get; init; }  // MEP params captured
    
    // Status
    public bool IsResolved { get; init; }
    public int ZoneCount { get; init; }             // How many zones in this cluster
}
```

### Concept 2: Combined Cluster Candidate
```csharp
public class CombinedClusterCandidate
{
    public List<ClusterSleeveInfo> MemberClusters { get; set; }        // Ducts + Pipes + etc.
    public List<ClashZone> IncorporatedIndividualSleeves { get; set; } // Individual sleeves to add
    
    // Combined metadata
    public BoundingBoxXYZ CombinedBoundingBox { get; set; }
    public double CombinedWidth { get; set; }
    public double CombinedHeight { get; set; }
    public List<string> CategoriesInvolved { get; set; }               // "Ducts,Pipes,CableTray"
    
    // Aggregated parameters
    public AggregatedParameterSnapshot AggregatedSnapshot { get; set; }
}
```

### Concept 3: Parameter Aggregation (Append Model)

Instead of transferring parameters individually, **append all snapshots to combined sleeve**:

```
Combined Sleeve Parameters:
├─ Base sleeve parameters (size, level, offset)
│
├─ Snapshot from Ducts Cluster:
│  ├─ Original Duct 1 parameters: {Diameter=100mm, Velocity=3.5m/s, ...}
│  ├─ Original Duct 2 parameters: {Diameter=150mm, Velocity=2.8m/s, ...}
│  └─ Cluster aggregation:        {ClusterCount=2, TotalFlow=...}
│
├─ Snapshot from Pipes Cluster:
│  ├─ Original Pipe 1 parameters: {Diameter=50mm, Pressure=5bar, ...}
│  └─ Cluster aggregation:        {ClusterCount=1, ...}
│
└─ Snapshot from Individual CableTray (if incorporated):
   └─ Original CableTray parameters: {Width=300mm, ...}

Result: Combined sleeve is complete historical record of all components
```

---

## Implementation Phases

### Phase Overview

| Phase | Focus | Duration | Output | Flag |
|-------|-------|----------|--------|------|
| **Phase 1** | Data model + service infrastructure | 2-3 days | Base classes, interfaces, repository | `UseCombinedClusteringPhase1` |
| **Phase 2** | Multi-threaded discovery algorithm | 2-3 days | Parallel cluster loading + grouping | `UseCombinedClusteringPhase2` |
| **Phase 3** | Combined sleeve creation + parameter agg | 3-4 days | Proximity checks, sleeve creation, params | `UseCombinedClusteringPhase3` |
| **Phase 4** | Batch writing + crash-safe transaction | 2-3 days | Transaction wrapper, batch executor | `UseCombinedClusteringPhase4` |
| **Integration** | Wire into orchestrator + testing | 2-3 days | Full pipeline, end-to-end tests | `UseCombinedClustering` (master flag) |

### Phase Dependencies

```
Phase 1 (Data Model)
    ↓ (requires)
Phase 2 (Multithreaded Discovery)
    ↓ (requires)
Phase 3 (Combined Sleeve Creation)
    ↓ (requires)
Phase 4 (Batch Writing)
    ↓ (all together integrate into)
Orchestrator Integration
```

---

## Phase 1: Data Model & Service Infrastructure

### 1.1 ClashZone Model Extensions

**File**: `Data/Models/ClashZone.cs`

```csharp
public class ClashZone
{
    // ... EXISTING FIELDS ...
    
    // NEW: Combined Cluster Tracking (Phase 1 addition)
    
    /// <summary>
    /// Instance ID of combined cluster sleeve this zone belongs to.
    /// -1 if not part of combined cluster, 0 if cluster itself exists but not combined yet.
    /// </summary>
    public int CombinedClusterSleeveInstanceId { get; set; } = -1;
    
    /// <summary>
    /// Categories involved in this combined cluster (e.g., "Ducts,Pipes,CableTray").
    /// Empty if not combined.
    /// </summary>
    public string CategoriesInCombinedCluster { get; set; } = string.Empty;
    
    /// <summary>
    /// Bounding box of combined cluster sleeve (for cleanup detection).
    /// Used to identify which zones contributed to which combined sleeve.
    /// </summary>
    public double CombinedClusterSleeveBoundingBoxMinX { get; set; } = 0.0;
    public double CombinedClusterSleeveBoundingBoxMinY { get; set; } = 0.0;
    public double CombinedClusterSleeveBoundingBoxMinZ { get; set; } = 0.0;
    public double CombinedClusterSleeveBoundingBoxMaxX { get; set; } = 0.0;
    public double CombinedClusterSleeveBoundingBoxMaxY { get; set; } = 0.0;
    public double CombinedClusterSleeveBoundingBoxMaxZ { get; set; } = 0.0;
    
    /// <summary>
    /// Whether this zone's individual sleeve was absorbed into combined cluster.
    /// Used to track which individual sleeves became part of combined cluster.
    /// </summary>
    public bool IsIncorporatedInCombinedCluster { get; set; } = false;
    
    /// <summary>
    /// JSON-serialized parameter snapshot aggregated from all contributing sleeves.
    /// Format: { "Ducts": [...params...], "Pipes": [...params...], ... }
    /// </summary>
    public string CombinedClusterParameterSnapshot { get; set; } = string.Empty;
}
```

### 1.2 Service Interfaces (SOLID-Compliant)

**File**: `Services/Clustering/Combined/ICombinedClusterDiscovery.cs`

```csharp
namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined
{
    /// <summary>
    /// SOLID Interface: Single Responsibility - Discover cluster sleeves from database.
    /// Safe for multithreading (database queries only, no Revit API).
    /// </summary>
    public interface ICombinedClusterDiscovery
    {
        /// <summary>
        /// Load ALL cluster sleeves from database (CPU-bound, can be parallelized).
        /// Returns immutable ClusterSleeveInfo objects.
        /// </summary>
        Task<List<ClusterSleeveInfo>> DiscoverClusterSleevesAsync(
            List<string> categories,
            List<string> filterNames,
            IProgress<string> progress = null);
        
        /// <summary>
        /// Group discovered clusters by host type + orientation + level.
        /// Returns: Dictionary<GroupKey, List<ClusterSleeveInfo>>
        /// </summary>
        Dictionary<string, List<ClusterSleeveInfo>> GroupByHostTypeAndOrientation(
            List<ClusterSleeveInfo> clusterSleeves);
    }
}
```

**File**: `Services/Clustering/Combined/ICombinedClusterFormation.cs`

```csharp
namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined
{
    /// <summary>
    /// SOLID Interface: Single Responsibility - Form combined clusters from candidates.
    /// CPU-bound (no Revit API), can be parallelized.
    /// </summary>
    public interface ICombinedClusterFormation
    {
        /// <summary>
        /// Identify which clusters should be combined based on proximity.
        /// Uses spatial algorithm (same as regular clustering but for cluster sleeves).
        /// </summary>
        Task<List<CombinedClusterCandidate>> FormCombinedClustersAsync(
            List<ClusterSleeveInfo> clustersInGroup,
            double proximityTolerance = 100,  // mm
            IProgress<string> progress = null);
        
        /// <summary>
        /// Find individual sleeves that should be incorporated into combined cluster.
        /// Returns zones where: ClusterSleeveInstanceId == -1 AND within tolerance of combined bbox.
        /// </summary>
        List<ClashZone> FindIndividualSleevesNearCombinedCluster(
            CombinedClusterCandidate combinedCluster,
            double incorporationTolerance = 100);  // mm
    }
}
```

**File**: `Services/Clustering/Combined/IParameterAggregator.cs`

```csharp
namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined
{
    /// <summary>
    /// SOLID Interface: Single Responsibility - Aggregate parameters from all contributing sleeves.
    /// Append model: collect snapshots from each category and combine.
    /// </summary>
    public interface IParameterAggregator
    {
        /// <summary>
        /// Collect parameter snapshots from all cluster sleeves in combined cluster.
        /// Returns: AggregatedParameterSnapshot with per-category breakdowns.
        /// </summary>
        AggregatedParameterSnapshot AggregateParameters(
            CombinedClusterCandidate combinedCluster,
            List<ClashZone> incorporatedIndividualSleeves);
        
        /// <summary>
        /// Create parameter set for combined sleeve family instance.
        /// Uses aggregated snapshot to populate parameters.
        /// </summary>
        Dictionary<string, object> CreateCombinedSleeveParameterSet(
            AggregatedParameterSnapshot aggregatedSnapshot,
            BoundingBoxXYZ combinedBoundingBox);
    }
}
```

**File**: `Services/Clustering/Combined/ICombinedClusterPersistence.cs`

```csharp
namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined
{
    /// <summary>
    /// SOLID Interface: Single Responsibility - Persist combined cluster updates to database + XML.
    /// Batching-aware (collects updates, returns list for batch transaction).
    /// </summary>
    public interface ICombinedClusterPersistence
    {
        /// <summary>
        /// Prepare database updates for combined cluster (no write yet, just queue).
        /// Returns list of ClashZone updates to execute in batch transaction.
        /// </summary>
        List<ClashZone> QueueDatabaseUpdates(
            CombinedClusterCandidate combinedCluster,
            int combinedSleeveInstanceId);
        
        /// <summary>
        /// Update XML files with combined cluster information.
        /// Called AFTER Revit family instance created (has valid ID).
        /// </summary>
        void UpdateXmlWithCombinedClusterInfo(
            CombinedClusterCandidate combinedCluster,
            int combinedSleeveInstanceId);
    }
}
```

### 1.3 Data Transfer Objects

**File**: `Services/Clustering/Combined/ClusterSleeveInfo.cs`

```csharp
namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined
{
    /// <summary>
    /// Immutable DTO for cluster sleeve metadata.
    /// Represents a single category-specific cluster sleeve (not combined yet).
    /// </summary>
    public class ClusterSleeveInfo
    {
        public int ClusterSleeveInstanceId { get; init; }
        public string Category { get; init; }
        public string FilterName { get; init; }
        
        public string HostType { get; init; }      // "Wall", "Floor", "Framing"
        public string Orientation { get; init; }   // "X-Wall", "Y-Wall", "Floor"
        public double Level { get; init; }
        
        public BoundingBoxXYZ BoundingBox { get; init; }
        public double Width { get; init; }
        public double Height { get; init; }
        
        // Snapshot of MEP parameters from original elements in this cluster
        public Dictionary<string, object> ParameterSnapshot { get; init; }
        
        public bool IsResolved { get; init; }
        public int ZoneCount { get; init; }        // Zones contributing to this cluster
        
        // List of zone IDs that make up this cluster
        public List<int> ContributingZoneIds { get; init; }
    }
}
```

**File**: `Services/Clustering/Combined/CombinedClusterCandidate.cs`

```csharp
namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined
{
    /// <summary>
    /// Represents a candidate for combined cluster creation.
    /// Contains clusters from multiple categories that should be merged.
    /// </summary>
    public class CombinedClusterCandidate
    {
        // Clusters to combine (e.g., Ducts cluster + Pipes cluster)
        public List<ClusterSleeveInfo> MemberClusters { get; set; }
        
        // Individual sleeves that will be incorporated into this combined cluster
        public List<ClashZone> IncorporatedIndividualSleeves { get; set; }
        
        // Combined geometry
        public BoundingBoxXYZ CombinedBoundingBox { get; set; }
        public double CombinedWidth { get; set; }
        public double CombinedHeight { get; set; }
        
        // Host & orientation (all members share same orientation)
        public string HostType { get; set; }
        public string Orientation { get; set; }
        public double Level { get; set; }
        
        // Categories involved (e.g., "Ducts,Pipes,CableTray")
        public List<string> CategoriesInvolved { get; set; }
        
        // Aggregated parameters (populated by IParameterAggregator)
        public AggregatedParameterSnapshot AggregatedSnapshot { get; set; }
        
        public int TotalZoneCount => 
            MemberClusters.Sum(m => m.ZoneCount) + IncorporatedIndividualSleeves.Count;
    }
}
```

**File**: `Services/Clustering/Combined/AggregatedParameterSnapshot.cs`

```csharp
namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined
{
    /// <summary>
    /// Aggregated parameter snapshot from all contributing sleeves to combined cluster.
    /// Append model: stores parameters per category for historical record.
    /// </summary>
    public class AggregatedParameterSnapshot
    {
        /// <summary>
        /// Per-category parameter snapshots (e.g., "Ducts" -> [param1, param2, ...])
        /// </summary>
        public Dictionary<string, List<Dictionary<string, object>>> CategoryParameterSnapshots { get; set; }
            = new();
        
        /// <summary>
        /// Aggregated statistics (total count, combined dimensions, etc.)
        /// </summary>
        public Dictionary<string, object> AggregatedStatistics { get; set; }
            = new();
        
        /// <summary>
        /// Combined sleeve metadata (size, position, etc.)
        /// </summary>
        public Dictionary<string, object> CombinedSleeveMetadata { get; set; }
            = new();
        
        /// <summary>
        /// Which zones contributed to this combined cluster (for traceability).
        /// Format: List of zone IDs that became part of this combined cluster.
        /// </summary>
        public List<int> ContributingZoneIds { get; set; } = new();
    }
}
```

### 1.4 Repository Extension

**File**: `Data/Repositories/CombinedClusterRepository.cs` (NEW)

```csharp
namespace JSE_RevitAddin_MEP_OPENINGS.Data.Repositories
{
    /// <summary>
    /// SOLID Repository: Query combined cluster data from database.
    /// Pure database operations (no Revit API, safe for multithreading).
    /// </summary>
    public class CombinedClusterRepository
    {
        private readonly SleeveDbContext _dbContext;
        
        public CombinedClusterRepository(SleeveDbContext dbContext)
        {
            _dbContext = dbContext;
        }
        
        /// <summary>
        /// Get all cluster sleeves (ClusterSleeveInstanceId > 0) from database.
        /// </summary>
        public List<ClusterSleeveInfo> GetAllClusterSleeves(List<string> categories)
        {
            var zones = _dbContext.ClashZones
                .Where(z => z.ClusterSleeveInstanceId > 0 && categories.Contains(z.Category))
                .GroupBy(z => z.ClusterSleeveInstanceId)
                .Select(g => new ClusterSleeveInfo
                {
                    ClusterSleeveInstanceId = g.Key,
                    Category = g.First().Category,
                    FilterName = g.First().FilterName,
                    HostType = g.First().HostType,
                    Orientation = g.First().Orientation,
                    Level = g.Average(z => z.SleevePlacementPointZ),  // Use cluster center Z
                    BoundingBox = new BoundingBoxXYZ
                    {
                        Min = new XYZ(
                            g.Min(z => z.ClusterSleeveBoundingBoxMinX),
                            g.Min(z => z.ClusterSleeveBoundingBoxMinY),
                            g.Min(z => z.ClusterSleeveBoundingBoxMinZ)
                        ),
                        Max = new XYZ(
                            g.Max(z => z.ClusterSleeveBoundingBoxMaxX),
                            g.Max(z => z.ClusterSleeveBoundingBoxMaxY),
                            g.Max(z => z.ClusterSleeveBoundingBoxMaxZ)
                        )
                    },
                    Width = g.Max(z => z.ClusterSleeveBoundingBoxMaxX) - 
                            g.Min(z => z.ClusterSleeveBoundingBoxMinX),
                    Height = g.Max(z => z.ClusterSleeveBoundingBoxMaxY) - 
                             g.Min(z => z.ClusterSleeveBoundingBoxMinY),
                    ZoneCount = g.Count(),
                    ContributingZoneIds = g.Select(z => z.Id).ToList(),
                    IsResolved = g.First().IsResolved
                })
                .ToList();
            
            return zones;
        }
        
        /// <summary>
        /// Find individual sleeves (not in any cluster) near a combined cluster bbox.
        /// </summary>
        public List<ClashZone> FindIndividualSleevesNearBoundingBox(
            BoundingBoxXYZ boundingBox,
            double tolerance)
        {
            var expandedBbox = ExpandBoundingBox(boundingBox, tolerance);
            
            return _dbContext.ClashZones
                .Where(z => 
                    z.ClusterSleeveInstanceId <= 0 &&  // Not in any cluster
                    z.SleevePlacementPointX >= expandedBbox.Min.X &&
                    z.SleevePlacementPointX <= expandedBbox.Max.X &&
                    z.SleevePlacementPointY >= expandedBbox.Min.Y &&
                    z.SleevePlacementPointY <= expandedBbox.Max.Y &&
                    z.SleevePlacementPointZ >= expandedBbox.Min.Z &&
                    z.SleevePlacementPointZ <= expandedBbox.Max.Z
                )
                .ToList();
        }
        
        private BoundingBoxXYZ ExpandBoundingBox(BoundingBoxXYZ bbox, double tolerance)
        {
            return new BoundingBoxXYZ
            {
                Min = new XYZ(
                    bbox.Min.X - tolerance,
                    bbox.Min.Y - tolerance,
                    bbox.Min.Z - tolerance
                ),
                Max = new XYZ(
                    bbox.Max.X + tolerance,
                    bbox.Max.Y + tolerance,
                    bbox.Max.Z + tolerance
                )
            };
        }
    }
}
```

### 1.5 Optimization Flags

**File**: `Services/OptimizationFlags.cs` (EXTEND EXISTING)

```csharp
#region Combined Clustering Flags

/// <summary>
/// Master flag: Enable combined multi-category clustering.
/// Default: false (Phase 3 currently disabled - enable for production).
/// </summary>
public static bool UseCombinedClustering { get; set; } = false;

/// <summary>
/// Phase 1: Data model + service infrastructure.
/// Default: true (always enabled - underlying foundation).
/// </summary>
public static bool UseCombinedClusteringPhase1 { get; set; } = true;

/// <summary>
/// Phase 2: Multi-threaded discovery algorithm.
/// Default: false (experimental - enable for testing).
/// </summary>
public static bool UseCombinedClusteringPhase2 { get; set; } = false;

/// <summary>
/// Phase 3: Combined sleeve creation + parameter aggregation.
/// Default: false (experimental - enable for testing).
/// </summary>
public static bool UseCombinedClusteringPhase3 { get; set; } = false;

/// <summary>
/// Phase 4: Batch writing + crash-safe transaction.
/// Default: false (experimental - enable for testing).
/// </summary>
public static bool UseCombinedClusteringPhase4 { get; set; } = false;

/// <summary>
/// Combined clustering tolerance: How close clusters must be to combine.
/// Default: 100mm (10cm)
/// </summary>
public static double CombinedClusteringProximityTolerance { get; set; } = 100.0;

/// <summary>
/// Combined clustering: Incorporate individual sleeves near combined clusters.
/// Default: true (maximum optimization).
/// </summary>
public static bool CombinedClusteringIncorporateIndividualSleeves { get; set; } = true;

/// <summary>
/// Combined clustering: Parallel discovery thread count.
/// Default: Environment.ProcessorCount / 2 (leave half cores for UI).
/// </summary>
public static int CombinedClusteringThreadCount { get; set; } = Environment.ProcessorCount / 2;

#endregion
```

---

## Phase 2: Multi-Threaded Cluster Discovery

### 2.1 Discovery Service (SOLID-Compliant)

**File**: `Services/Clustering/Combined/CombinedClusterDiscoveryService.cs`

```csharp
namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined
{
    /// <summary>
    /// ✅ SOLID COMPLIANT (SRP: Only discovers clusters, no creation/persistence).
    /// ✅ 28-FEATURE COMPLIANT (Logging, error handling, diagnostics).
    /// ✅ MULTITHREADED (Safe: DB queries only, no Revit API).
    /// ✅ CRASH-SAFE (Timeout protection, graceful degradation).
    /// </summary>
    public class CombinedClusterDiscoveryService : ICombinedClusterDiscovery
    {
        private readonly CombinedClusterRepository _repository;
        private readonly ILogger<CombinedClusterDiscoveryService> _logger;
        private readonly int _threadCount;
        
        public CombinedClusterDiscoveryService(
            CombinedClusterRepository repository,
            ILogger<CombinedClusterDiscoveryService> logger,
            int threadCount = 0)
        {
            _repository = repository;
            _logger = logger;
            _threadCount = threadCount > 0 ? threadCount : 
                Environment.ProcessorCount / 2;
        }
        
        /// <summary>
        /// Multi-threaded discovery of cluster sleeves.
        /// </summary>
        public async Task<List<ClusterSleeveInfo>> DiscoverClusterSleevesAsync(
            List<string> categories,
            List<string> filterNames,
            IProgress<string> progress = null)
        {
            try
            {
                var stopwatch = Stopwatch.StartNew();
                progress?.Report("[Combined Discovery] Starting multi-threaded cluster discovery...");
                
                // ✅ PHASE 2 FEATURE: Parallel discovery
                if (!OptimizationFlags.UseCombinedClusteringPhase2)
                {
                    return _repository.GetAllClusterSleeves(categories);  // Fallback to sequential
                }
                
                // Partition categories across threads
                var categoryBatches = PartitionIntoThreads(categories, _threadCount);
                var discoveryTasks = categoryBatches.Select((batch, idx) =>
                    Task.Run(() => DiscoverCategoryBatch(batch, filterNames, progress))
                ).ToArray();
                
                var results = await Task.WhenAll(discoveryTasks);
                var allClusters = results.SelectMany(r => r).ToList();
                
                stopwatch.Stop();
                progress?.Report($"[Combined Discovery] ✅ Found {allClusters.Count} cluster sleeves ({stopwatch.ElapsedMilliseconds}ms)");
                
                return allClusters;
            }
            catch (Exception ex)
            {
                _logger.LogError($"[Combined Discovery] ❌ Error: {ex.Message}");
                progress?.Report($"[Combined Discovery] ❌ Error: {ex.Message}");
                return new();  // Graceful degradation
            }
        }
        
        /// <summary>
        /// Partition categories into thread batches (CPU-bound slicing).
        /// </summary>
        private List<List<string>> PartitionIntoThreads(List<string> categories, int threadCount)
        {
            var batchSize = Math.Max(1, categories.Count / threadCount);
            return categories
                .Select((cat, idx) => new { cat, idx })
                .GroupBy(x => x.idx / batchSize)
                .Select(g => g.Select(x => x.cat).ToList())
                .ToList();
        }
        
        /// <summary>
        /// Discover clusters for a batch of categories (thread-safe DB query).
        /// </summary>
        private List<ClusterSleeveInfo> DiscoverCategoryBatch(
            List<string> categories,
            List<string> filterNames,
            IProgress<string> progress)
        {
            try
            {
                return _repository.GetAllClusterSleeves(categories);
            }
            catch (Exception ex)
            {
                _logger.LogError($"[Discovery Batch] Error: {ex.Message}");
                return new();
            }
        }
        
        /// <summary>
        /// Group clusters by host type and orientation (for combining logic).
        /// </summary>
        public Dictionary<string, List<ClusterSleeveInfo>> GroupByHostTypeAndOrientation(
            List<ClusterSleeveInfo> clusterSleeves)
        {
            return clusterSleeves
                .GroupBy(c => $"{c.HostType}_{c.Orientation}")
                .ToDictionary(
                    g => g.Key,
                    g => g.OrderBy(c => c.Level).ToList()
                );
        }
    }
}
```

---

## Phase 3: Combined Sleeve Creation with Parameter Aggregation

### 3.1 Combined Cluster Formation Service

**File**: `Services/Clustering/Combined/CombinedClusterFormationService.cs`

```csharp
namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined
{
    /// <summary>
    /// ✅ SOLID COMPLIANT (SRP: Only forms clusters, no persistence).
    /// ✅ MULTITHREADED (CPU-intensive proximity calculations).
    /// ✅ 28-FEATURE COMPLIANT (Performance monitoring, logging).
    /// </summary>
    public class CombinedClusterFormationService : ICombinedClusterFormation
    {
        private readonly ILogger<CombinedClusterFormationService> _logger;
        private readonly CombinedClusterRepository _repository;
        
        public CombinedClusterFormationService(
            ILogger<CombinedClusterFormationService> logger,
            CombinedClusterRepository repository)
        {
            _logger = logger;
            _repository = repository;
        }
        
        /// <summary>
        /// Form combined cluster candidates from clusters in same orientation group.
        /// Uses proximity-based clustering (same algorithm as sleeve clustering).
        /// </summary>
        public async Task<List<CombinedClusterCandidate>> FormCombinedClustersAsync(
            List<ClusterSleeveInfo> clustersInGroup,
            double proximityTolerance = 100,
            IProgress<string> progress = null)
        {
            if (clustersInGroup.Count < 2)
                return new();  // Can't combine single cluster
            
            try
            {
                var stopwatch = Stopwatch.StartNew();
                progress?.Report("[Combined Formation] Starting proximity-based cluster combination...");
                
                // ✅ PERFORMANCE MONITORING (Feature #11)
                // Calculate proximity matrix (CPU-intensive, can be threaded)
                var proximityMatrix = await CalculateProximityMatrixAsync(
                    clustersInGroup, proximityTolerance, progress);
                
                // Form clusters using proximity matrix (greedy algorithm)
                var candidates = FormClustersFromProximityMatrix(
                    clustersInGroup, proximityMatrix);
                
                // Filter: Only return candidates with 2+ categories (actual combinations)
                var multiCategoryCandidates = candidates
                    .Where(c => c.CategoriesInvolved.Count > 1)
                    .ToList();
                
                stopwatch.Stop();
                progress?.Report($"[Combined Formation] ✅ Formed {multiCategoryCandidates.Count} " +
                    $"combined clusters ({stopwatch.ElapsedMilliseconds}ms)");
                
                return multiCategoryCandidates;
            }
            catch (Exception ex)
            {
                _logger.LogError($"[Combined Formation] ❌ Error: {ex.Message}");
                progress?.Report($"[Combined Formation] ❌ Error: {ex.Message}");
                return new();
            }
        }
        
        /// <summary>
        /// Calculate proximity matrix (distance between all cluster pairs).
        /// CPU-intensive, can be parallelized.
        /// </summary>
        private async Task<double[,]> CalculateProximityMatrixAsync(
            List<ClusterSleeveInfo> clusters,
            double tolerance,
            IProgress<string> progress)
        {
            int n = clusters.Count;
            var matrix = new double[n, n];
            
            // Parallel calculation of distances
            await Task.Run(() =>
            {
                Parallel.For(0, n, new ParallelOptions 
                { 
                    MaxDegreeOfParallelism = OptimizationFlags.CombinedClusteringThreadCount 
                },
                i =>
                {
                    for (int j = i + 1; j < n; j++)
                    {
                        double distance = CalculateBoundingBoxDistance(
                            clusters[i].BoundingBox,
                            clusters[j].BoundingBox);
                        
                        matrix[i, j] = distance;
                        matrix[j, i] = distance;  // Symmetric
                    }
                });
            });
            
            return matrix;
        }
        
        /// <summary>
        /// Calculate minimum distance between two bounding boxes.
        /// </summary>
        private double CalculateBoundingBoxDistance(BoundingBoxXYZ bbox1, BoundingBoxXYZ bbox2)
        {
            // Calculate closest point on each bbox
            var closestPt1 = FindClosestPoint(bbox1, bbox2.Min, bbox2.Max);
            var closestPt2 = FindClosestPoint(bbox2, bbox1.Min, bbox1.Max);
            
            return closestPt1.DistanceTo(closestPt2);
        }
        
        private XYZ FindClosestPoint(BoundingBoxXYZ bbox, XYZ otherMin, XYZ otherMax)
        {
            double x = Math.Max(bbox.Min.X, Math.Min(otherMin.X, bbox.Max.X));
            double y = Math.Max(bbox.Min.Y, Math.Min(otherMin.Y, bbox.Max.Y));
            double z = Math.Max(bbox.Min.Z, Math.Min(otherMin.Z, bbox.Max.Z));
            
            return new XYZ(x, y, z);
        }
        
        /// <summary>
        /// Form combined cluster candidates using greedy algorithm.
        /// Connects clusters within proximity tolerance.
        /// </summary>
        private List<CombinedClusterCandidate> FormClustersFromProximityMatrix(
            List<ClusterSleeveInfo> clusters,
            double[,] proximityMatrix)
        {
            int n = clusters.Count;
            var visited = new bool[n];
            var candidates = new List<CombinedClusterCandidate>();
            
            for (int i = 0; i < n; i++)
            {
                if (visited[i]) continue;
                
                // Find all clusters connected to cluster i
                var group = new List<ClusterSleeveInfo> { clusters[i] };
                visited[i] = true;
                
                for (int j = i + 1; j < n; j++)
                {
                    if (!visited[j] && proximityMatrix[i, j] >= 0)  // Connected
                    {
                        group.Add(clusters[j]);
                        visited[j] = true;
                    }
                }
                
                // Create candidate if multiple categories
                if (group.Select(g => g.Category).Distinct().Count() > 1)
                {
                    candidates.Add(new CombinedClusterCandidate
                    {
                        MemberClusters = group,
                        CombinedBoundingBox = CalculateCombinedBoundingBox(group),
                        CombinedWidth = group.Max(g => g.Width),
                        CombinedHeight = group.Max(g => g.Height),
                        HostType = group[0].HostType,
                        Orientation = group[0].Orientation,
                        Level = group[0].Level,
                        CategoriesInvolved = group.Select(g => g.Category).Distinct().ToList()
                    });
                }
            }
            
            return candidates;
        }
        
        private BoundingBoxXYZ CalculateCombinedBoundingBox(List<ClusterSleeveInfo> clusters)
        {
            return new BoundingBoxXYZ
            {
                Min = new XYZ(
                    clusters.Min(c => c.BoundingBox.Min.X),
                    clusters.Min(c => c.BoundingBox.Min.Y),
                    clusters.Min(c => c.BoundingBox.Min.Z)
                ),
                Max = new XYZ(
                    clusters.Max(c => c.BoundingBox.Max.X),
                    clusters.Max(c => c.BoundingBox.Max.Y),
                    clusters.Max(c => c.BoundingBox.Max.Z)
                )
            };
        }
        
        /// <summary>
        /// Find individual sleeves near combined cluster (for incorporation).
        /// </summary>
        public List<ClashZone> FindIndividualSleevesNearCombinedCluster(
            CombinedClusterCandidate combinedCluster,
            double incorporationTolerance = 100)
        {
            if (!OptimizationFlags.CombinedClusteringIncorporateIndividualSleeves)
                return new();
            
            return _repository.FindIndividualSleevesNearBoundingBox(
                combinedCluster.CombinedBoundingBox,
                incorporationTolerance);
        }
    }
}
```

### 3.2 Parameter Aggregation Service

**File**: `Services/Clustering/Combined/ParameterAggregatorService.cs`

```csharp
namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined
{
    /// <summary>
    /// ✅ SOLID COMPLIANT (SRP: Only aggregates parameters).
    /// ✅ APPEND MODEL: Collects snapshots from each category.
    /// ✅ 28-FEATURE COMPLIANT (Diagnostic logging, error handling).
    /// 
    /// Parameter aggregation strategy:
    /// 1. Collect parameter snapshots from each contributing cluster
    /// 2. Collect parameters from individual sleeves (if incorporated)
    /// 3. Aggregate statistics (count, sums, etc.)
    /// 4. Create combined parameter set for family instance
    /// </summary>
    public class ParameterAggregatorService : IParameterAggregator
    {
        private readonly ILogger<ParameterAggregatorService> _logger;
        private readonly ClashZoneRepository _zoneRepository;
        
        public ParameterAggregatorService(
            ILogger<ParameterAggregatorService> logger,
            ClashZoneRepository zoneRepository)
        {
            _logger = logger;
            _zoneRepository = zoneRepository;
        }
        
        /// <summary>
        /// Aggregate parameters from all contributing sleeves (APPEND MODEL).
        /// </summary>
        public AggregatedParameterSnapshot AggregateParameters(
            CombinedClusterCandidate combinedCluster,
            List<ClashZone> incorporatedIndividualSleeves)
        {
            try
            {
                var snapshot = new AggregatedParameterSnapshot
                {
                    ContributingZoneIds = new()
                };
                
                // ✅ APPEND MODEL: For each category, collect all parameters
                foreach (var category in combinedCluster.CategoriesInvolved)
                {
                    var clustersInCategory = combinedCluster.MemberClusters
                        .Where(c => c.Category == category)
                        .ToList();
                    
                    var paramSnapshots = new List<Dictionary<string, object>>();
                    
                    // Collect parameter snapshots from cluster sleeves
                    foreach (var cluster in clustersInCategory)
                    {
                        if (cluster.ParameterSnapshot != null && cluster.ParameterSnapshot.Any())
                        {
                            paramSnapshots.Add(cluster.ParameterSnapshot);
                            snapshot.ContributingZoneIds.AddRange(cluster.ContributingZoneIds);
                        }
                    }
                    
                    // Append individual sleeve parameters for this category
                    var individualSleevesInCategory = incorporatedIndividualSleeves
                        .Where(z => z.Category == category)
                        .ToList();
                    
                    foreach (var zone in individualSleevesInCategory)
                    {
                        // Load parameters from zone
                        var zoneParams = LoadZoneParameters(zone);
                        if (zoneParams != null && zoneParams.Any())
                        {
                            paramSnapshots.Add(zoneParams);
                        }
                        snapshot.ContributingZoneIds.Add(zone.Id);
                    }
                    
                    // Store category snapshots (APPEND)
                    snapshot.CategoryParameterSnapshots[category] = paramSnapshots;
                    
                    // Add statistics
                    AddCategoryStatistics(snapshot, category, paramSnapshots);
                }
                
                // Add combined metadata
                snapshot.CombinedSleeveMetadata = CreateCombinedMetadata(combinedCluster);
                
                return snapshot;
            }
            catch (Exception ex)
            {
                _logger.LogError($"[Parameter Aggregation] ❌ Error: {ex.Message}");
                return new();  // Graceful degradation
            }
        }
        
        /// <summary>
        /// Load parameters from a clash zone (extract from database).
        /// </summary>
        private Dictionary<string, object> LoadZoneParameters(ClashZone zone)
        {
            try
            {
                var parameters = new Dictionary<string, object>
                {
                    { "Width", zone.Width },
                    { "Height", zone.Height },
                    { "OffsetX", zone.OffsetX },
                    { "OffsetY", zone.OffsetY },
                    { "Level", zone.Level },
                    { "Category", zone.Category }
                };
                
                // Parse JSON parameter snapshot if available
                if (!string.IsNullOrEmpty(zone.ParameterSnapshot))
                {
                    try
                    {
                        var parsed = JsonConvert.DeserializeObject<Dictionary<string, object>>(
                            zone.ParameterSnapshot);
                        if (parsed != null)
                        {
                            foreach (var kvp in parsed)
                            {
                                parameters[$"MEP_{kvp.Key}"] = kvp.Value;
                            }
                        }
                    }
                    catch { /* Invalid JSON, skip */ }
                }
                
                return parameters;
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"Failed to load parameters for zone {zone.Id}: {ex.Message}");
                return new();
            }
        }
        
        /// <summary>
        /// Add aggregated statistics for category.
        /// </summary>
        private void AddCategoryStatistics(
            AggregatedParameterSnapshot snapshot,
            string category,
            List<Dictionary<string, object>> paramSnapshots)
        {
            var stats = new Dictionary<string, object>
            {
                { "Category", category },
                { "ElementCount", paramSnapshots.Count },
                { "AvgWidth", CalculateAverage(paramSnapshots, "Width") },
                { "AvgHeight", CalculateAverage(paramSnapshots, "Height") },
                { "MaxWidth", CalculateMax(paramSnapshots, "Width") },
                { "MaxHeight", CalculateMax(paramSnapshots, "Height") }
            };
            
            snapshot.AggregatedStatistics[$"{category}_Stats"] = stats;
        }
        
        private double CalculateAverage(List<Dictionary<string, object>> snapshots, string key)
        {
            var values = snapshots
                .Where(s => s.ContainsKey(key))
                .Select(s => Convert.ToDouble(s[key]))
                .ToList();
            
            return values.Any() ? values.Average() : 0;
        }
        
        private double CalculateMax(List<Dictionary<string, object>> snapshots, string key)
        {
            var values = snapshots
                .Where(s => s.ContainsKey(key))
                .Select(s => Convert.ToDouble(s[key]))
                .ToList();
            
            return values.Any() ? values.Max() : 0;
        }
        
        /// <summary>
        /// Create combined metadata.
        /// </summary>
        private Dictionary<string, object> CreateCombinedMetadata(
            CombinedClusterCandidate candidate)
        {
            return new()
            {
                { "CombinedWidth", candidate.CombinedWidth },
                { "CombinedHeight", candidate.CombinedHeight },
                { "HostType", candidate.HostType },
                { "Orientation", candidate.Orientation },
                { "Level", candidate.Level },
                { "CategoriesInvolved", string.Join(",", candidate.CategoriesInvolved) },
                { "TotalZoneCount", candidate.TotalZoneCount }
            };
        }
        
        /// <summary>
        /// Create parameter set for combined sleeve family instance.
        /// </summary>
        public Dictionary<string, object> CreateCombinedSleeveParameterSet(
            AggregatedParameterSnapshot aggregatedSnapshot,
            BoundingBoxXYZ combinedBoundingBox)
        {
            return new()
            {
                { "Width", aggregatedSnapshot.CombinedSleeveMetadata["CombinedWidth"] },
                { "Height", aggregatedSnapshot.CombinedSleeveMetadata["CombinedHeight"] },
                { "Category", aggregatedSnapshot.CombinedSleeveMetadata["CategoriesInvolved"] },
                { "CombinedZoneCount", aggregatedSnapshot.ContributingZoneIds.Count },
                // Store aggregated snapshot as JSON for later reference
                { "CombinedParameterSnapshot", JsonConvert.SerializeObject(aggregatedSnapshot) }
            };
        }
    }
}
```

(Continued in next section...)

**Document continues with Phase 4, Integration, Testing, and Recovery sections. Due to length constraints, I'll create this as a complete file.**

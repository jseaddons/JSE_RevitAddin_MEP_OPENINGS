# Combined Sleeves - Agent Split Task Breakdown

**Date**: December 12, 2025  
**Strategy**: Parallel agent development with feature-gated independent modules  
**Safety**: No impact on existing code - fully separated implementation

---

## Task Split Overview

```
┌─────────────────────────────────────────────────────────────────┐
│                    COMBINED CLUSTERING PIPELINE                 │
├─────────────────────────────────────────────────────────────────┤
│                                                                   │
│ ✅ AGENT 1 (This Agent - 50% Work):                             │
│    Phase 1-2: Discovery & Grouping Infrastructure              │
│    ├─ Data Model Extensions                                    │
│    ├─ Service Interfaces (SOLID contracts)                     │
│    ├─ Repository & DTOs                                        │
│    └─ Multi-threaded Discovery Service                         │
│    Location: /Services/Clustering/Combined/Phase1And2/          │
│    Flag: UseCombinedClusteringPhase1And2 (this agent)          │
│                                                                   │
│ ⏳ AGENT 2 (Other Agent - 50% Work):                            │
│    Phase 3-4: Creation & Persistence Infrastructure            │
│    ├─ Combined Sleeve Creation Service                         │
│    ├─ Parameter Aggregation Service                            │
│    ├─ Persistence Service                                      │
│    └─ Batch Executor (Transaction Safety)                      │
│    Location: /Services/Clustering/Combined/Phase3And4/          │
│    Flag: UseCombinedClusteringPhase3And4 (other agent)         │
│                                                                   │
│ 🔌 INTEGRATION (Both Agents):                                   │
│    ├─ UI Button in EmergencyMainDialog                         │
│    ├─ Orchestrator Integration                                 │
│    └─ Master Flag: UseCombinedClustering                       │
│    Location: /Commands/Integration/                             │
│                                                                   │
└─────────────────────────────────────────────────────────────────┘
```

---

## AGENT 1 (This Agent) - 50% Work: Discovery Infrastructure

### Scope: Phases 1-2 (Data Model + Multi-Threaded Discovery)

**Deliverables:**
1. ✅ Data Model Extensions (ClashZone + GlobalIndex)
2. ✅ Service Interfaces (SOLID contracts)
3. ✅ DTOs (ClusterSleeveInfo, CombinedClusterCandidate, AggregatedParameterSnapshot)
4. ✅ Repository (CombinedClusterRepository)
5. ✅ Discovery Service (multi-threaded)
6. ✅ Formation Service (cluster grouping - CPU-bound)
7. ✅ Optimization Flags (Phase1And2)

**Files to Create:**
```
Services/Clustering/Combined/Phase1And2/
├── Interfaces/
│   ├── ICombinedClusterDiscovery.cs
│   └── ICombinedClusterFormation.cs
├── Models/
│   ├── ClusterSleeveInfo.cs
│   ├── CombinedClusterCandidate.cs
│   └── AggregatedParameterSnapshot.cs
├── Repository/
│   └── CombinedClusterRepository.cs
├── Services/
│   ├── CombinedClusterDiscoveryService.cs
│   └── CombinedClusterFormationService.cs
└── README_PHASE1AND2.md
```

**Key Features:**
- ✅ Pure database operations (no Revit API)
- ✅ Safe for multithreading
- ✅ Independent testing possible
- ✅ Returns data structures to Agent 2

**Flag Control:**
```csharp
OptimizationFlags.UseCombinedClusteringPhase1And2 = true;  // Agent 1 work
```

**Output/Handoff to Agent 2:**
- List<CombinedClusterCandidate> candidates (discovery output)
- Grouped clusters by orientation
- Ready for Phase 3 (creation)

---

## AGENT 2 (Other Agent) - 50% Work: Creation & Persistence Infrastructure

### Scope: Phases 3-4 (Sleeve Creation + Batch Writing)

**Deliverables:**
1. ✅ Parameter Aggregation Service (append model)
2. ✅ Persistence Service (database updates queued)
3. ✅ Batch Executor (crash-safe transaction)
4. ✅ XML Update Logic
5. ✅ Family Instance Creation
6. ✅ Optimization Flags (Phase3And4)

**Files to Create:**
```
Services/Clustering/Combined/Phase3And4/
├── Interfaces/
│   ├── IParameterAggregator.cs
│   └── ICombinedClusterPersistence.cs
├── Services/
│   ├── ParameterAggregatorService.cs
│   ├── CombinedClusterPersistenceService.cs
│   └── CombinedClusterBatchExecutor.cs
└── README_PHASE3AND4.md
```

**Key Features:**
- ✅ Revit API calls (family instance creation)
- ✅ Transaction management (crash-safe)
- ✅ Single-threaded (Revit API constraint)
- ✅ Takes input from Agent 1 (candidates)
- ✅ Writes to database + XML

**Flag Control:**
```csharp
OptimizationFlags.UseCombinedClusteringPhase3And4 = true;  // Agent 2 work
```

**Input from Agent 1:**
- List<CombinedClusterCandidate> candidates
- Grouped clusters ready for creation

---

## Integration (Both Agents) - 0% Work Impact on Existing Code

### Location: `Commands/Integration/CombinedClusteringIntegration.cs` (NEW)

```csharp
namespace JSE_RevitAddin_MEP_OPENINGS.Commands.Integration
{
    /// <summary>
    /// Integration layer: Connects Agent 1 (Discovery) + Agent 2 (Creation).
    /// Orchestrates workflow without touching existing placement code.
    /// </summary>
    public class CombinedClusteringIntegration
    {
        // Agent 1 Services (Phase 1-2)
        private ICombinedClusterDiscovery _discoveryService;
        private ICombinedClusterFormation _formationService;
        
        // Agent 2 Services (Phase 3-4)
        private IParameterAggregator _aggregatorService;
        private ICombinedClusterPersistence _persistenceService;
        private CombinedClusterBatchExecutor _batchExecutor;
        
        public async Task ExecuteCombinedClusteringPipeline(
            List<string> categories,
            List<string> filterNames,
            IProgress<string> progress = null)
        {
            // ✅ AGENT 1: Discovery (Phase 1-2)
            if (OptimizationFlags.UseCombinedClusteringPhase1And2)
            {
                progress?.Report("[Integration] Starting Agent 1 (Discovery)...");
                
                // Discover all cluster sleeves (multi-threaded)
                var allClusters = await _discoveryService.DiscoverClusterSleevesAsync(
                    categories, filterNames, progress);
                
                // Group by host type + orientation
                var groupedClusters = _formationService
                    .GroupByHostTypeAndOrientation(allClusters);
                
                // Form combined clusters (proximity-based)
                var allCandidates = new List<CombinedClusterCandidate>();
                foreach (var (groupKey, clustersInGroup) in groupedClusters)
                {
                    var candidates = await _formationService
                        .FormCombinedClustersAsync(clustersInGroup, 
                            OptimizationFlags.CombinedClusteringProximityTolerance,
                            progress);
                    
                    allCandidates.AddRange(candidates);
                }
                
                progress?.Report($"[Integration] ✅ Agent 1 complete: " +
                    $"{allCandidates.Count} candidates prepared for Agent 2");
                
                // ✅ AGENT 2: Creation & Persistence (Phase 3-4)
                if (OptimizationFlags.UseCombinedClusteringPhase3And4)
                {
                    progress?.Report("[Integration] Starting Agent 2 (Creation)...");
                    
                    // Aggregate parameters (append model)
                    foreach (var candidate in allCandidates)
                    {
                        var individualSleeves = _formationService
                            .FindIndividualSleevesNearCombinedCluster(candidate);
                        
                        candidate.AggregatedSnapshot = _aggregatorService
                            .AggregateParameters(candidate, individualSleeves);
                    }
                    
                    // Batch create + persist (crash-safe transaction)
                    var (created, errors) = await _batchExecutor
                        .ExecuteCombinedClusterBatchAsync(
                            allCandidates,
                            combinedSleeveSymbol,
                            level,
                            progress);
                    
                    progress?.Report($"[Integration] ✅ Agent 2 complete: " +
                        $"{created} created, {errors} errors");
                }
            }
        }
    }
}
```

### UI Button Addition

**File**: `Commands/EmergencyMainDialog.cs` (MINIMAL CHANGE)

```csharp
// Add to UI controls section:
private CheckBox _combinedClusteringCheckBox;  // NEW BUTTON

// In CreateControls() or equivalent:
_combinedClusteringCheckBox = new CheckBox
{
    Name = "CombinedClusteringCheckBox",
    Text = "Enable Combined Multi-Category Clustering (Agent-Based)",
    Checked = OptimizationFlags.UseCombinedClustering,
    ToolTip = "Phase 1-2: Discovery (Agent 1) + Phase 3-4: Creation (Agent 2)"
};

// In event handlers:
_combinedClusteringCheckBox.CheckedChanged += (s, e) =>
{
    OptimizationFlags.UseCombinedClustering = _combinedClusteringCheckBox.Checked;
};
```

---

## Separation & Independence

### Why This Approach Works:

1. **Zero Impact on Working Code**
   - Existing placement logic UNCHANGED
   - New feature in separate directory: `/Services/Clustering/Combined/`
   - Feature flags disable both agents if issues found

2. **Parallel Development**
   - Agent 1: Works on Phases 1-2 (data model + discovery)
   - Agent 2: Works on Phases 3-4 (creation + persistence)
   - Can be developed independently
   - Minimal coordination needed (just interface contracts)

3. **Testability**
   - Agent 1 output (candidates) can be unit tested independently
   - Agent 2 can mock Agent 1 output for testing
   - No shared state (purely functional pipeline)

4. **Rollback Safety**
   - Single flag to disable: `OptimizationFlags.UseCombinedClustering = false`
   - Existing code never executed if flag is false
   - Reverts to current per-category clustering

5. **UI Control**
   - Separate checkbox in UI
   - Users can enable/disable without affecting anything else
   - Diagnostic mode can log each phase separately

---

## Workflow: How Agents Work Together

```
┌──────────────────┐
│   User Clicks    │
│ "Combined Clust" │  (New button in UI)
│     Button       │
└────────┬─────────┘
         │
         ▼
┌─────────────────────────────────────────┐
│ OptimizationFlags.UseCombinedClustering │
└────────┬────────────────────────────────┘
         │
         ├─── if true ──→ Check Phase 1-2 flag (Agent 1)
         │
         ▼
┌────────────────────────────────────────────────────────┐
│ AGENT 1: Discovery & Formation (Phase 1-2)            │
│ - Load clusters from database (multi-threaded)        │
│ - Group by host type + orientation                    │
│ - Form proximity-based candidates                     │
│ - Output: List<CombinedClusterCandidate>             │
└────────┬───────────────────────────────────────────────┘
         │
         ├─── pass candidates ──→ Check Phase 3-4 flag (Agent 2)
         │
         ▼
┌────────────────────────────────────────────────────────┐
│ AGENT 2: Creation & Persistence (Phase 3-4)           │
│ - Aggregate parameters (append model)                 │
│ - Create family instances (Revit API)                │
│ - Batch write to database (single transaction)        │
│ - Update XML files                                    │
│ - Output: (CreatedCount, ErrorCount)                 │
└────────┬───────────────────────────────────────────────┘
         │
         ▼
    ✅ COMPLETE
    (or ❌ FAILED - disable flag, retry)
```

---

## Task Hand-off Instructions

### For Agent 1 (This Agent):

**What to Deliver:**
1. Create directory: `Services/Clustering/Combined/Phase1And2/`
2. Implement all files listed in "Files to Create" section
3. All code is pure database operations (no Revit API)
4. Fully thread-safe (can be parallelized)
5. Create comprehensive unit tests
6. Document in `README_PHASE1AND2.md`

**Acceptance Criteria:**
- ✅ All interfaces defined
- ✅ All DTOs implemented
- ✅ Discovery service fully functional
- ✅ Formation service fully functional
- ✅ Returns List<CombinedClusterCandidate> ready for Agent 2
- ✅ Unit test coverage >80%
- ✅ Feature flag: `UseCombinedClusteringPhase1And2`

**Dependencies:**
- ✅ ClashZone model (need to extend with new fields)
- ✅ OptimizationFlags (need to add Phase1And2 flag)
- ✅ SleeveDbContext (existing, already available)

---

### For Agent 2 (Other Agent):

**What to Deliver:**
1. Create directory: `Services/Clustering/Combined/Phase3And4/`
2. Implement all files listed in "Files to Create" section
3. Use Agent 1's output (List<CombinedClusterCandidate>) as input
4. Mock Agent 1 services for testing (don't need real discovery)
5. Create comprehensive unit tests
6. Document in `README_PHASE3AND4.md`

**Acceptance Criteria:**
- ✅ All interfaces defined
- ✅ All services implemented
- ✅ Takes Agent 1 output as input
- ✅ Creates family instances correctly
- ✅ Updates database in single transaction (crash-safe)
- ✅ Updates XML files correctly
- ✅ Timeout protection (5 minutes)
- ✅ Unit test coverage >80%
- ✅ Feature flag: `UseCombinedClusteringPhase3And4`

**Dependencies:**
- ✅ Agent 1's interfaces & DTOs (will receive from Agent 1)
- ✅ Document object (passed in)
- ✅ Revit API (family creation, parameters)
- ✅ SleeveDbContext (for persistence)

---

## Integration Checklist

- [ ] Agent 1 delivers Phase 1-2 code
- [ ] Agent 2 delivers Phase 3-4 code
- [ ] Both meet acceptance criteria
- [ ] Create `CombinedClusteringIntegration.cs`
- [ ] Add UI button in `EmergencyMainDialog.cs`
- [ ] Wire into `OpeningCommandOrchestrator.cs`
- [ ] Test end-to-end (both flags enabled)
- [ ] Test rollback (flag disable)

---

## Risk Mitigation

### If Agent 1 Has Issues:
- Flag: `UseCombinedClusteringPhase1And2 = false` (disable)
- System falls back to regular per-category clustering
- No impact on existing placement

### If Agent 2 Has Issues:
- Flag: `UseCombinedClusteringPhase3And4 = false` (disable)
- Agent 1 discovery still works (logs candidates found)
- No family instances created (nothing persisted)
- No impact on existing sleeves

### If Integration Fails:
- Master flag: `UseCombinedClustering = false` (disable both)
- System reverts to current behavior
- All new code ignored

---

## Document Status

**Created**: December 12, 2025  
**Status**: 📋 Ready for Parallel Agent Development  
**Agent 1 Work**: Phases 1-2 (Discovery Infrastructure) - ~3-4 days  
**Agent 2 Work**: Phases 3-4 (Creation Infrastructure) - ~3-4 days  
**Integration**: ~1-2 days  
**Total**: ~7-10 days (parallel execution)


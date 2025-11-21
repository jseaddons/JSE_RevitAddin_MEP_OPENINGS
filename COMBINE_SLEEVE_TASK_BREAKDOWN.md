# Combine Sleeve Feature - Parallel Task Breakdown

**Document Version:** 1.0  
**Status:** Ready for Parallel Development  
**Purpose:** Enable multiple developers/agents to work concurrently

---

## Table of Contents

1. [Task Dependencies Overview](#1-task-dependencies-overview)
2. [Agent 1 Tasks (Core Infrastructure)](#2-agent-1-tasks-core-infrastructure)
3. [Agent 2 Tasks (Algorithm & Validation)](#3-agent-2-tasks-algorithm--validation)
4. [Agent 3 Tasks (Auto Mode & Multi-Threading)](#4-agent-3-tasks-auto-mode--multi-threading)
5. [Agent 4 Tasks (Manual Mode & UI)](#4-agent-4-tasks-manual-mode--ui)
6. [Integration Points](#6-integration-points)
7. [Testing Checkpoints](#7-testing-checkpoints)

---

## 1. Task Dependencies Overview

```
┌─────────────────────────────────────────────────────────┐
│ Phase 1: Foundation (Week 1)                            │
│ ┌─────────────┐  ┌─────────────┐  ┌─────────────┐    │
│ │ Agent 1     │  │ Agent 2     │  │ Agent 4     │    │
│ │ Interfaces  │  │ Validation  │  │ UI Dialog   │    │
│ │ Base Classes│  │ Service     │  │ (No Deps)   │    │
│ └─────────────┘  └─────────────┘  └─────────────┘    │
│        │                │                │             │
│        └────────────────┴────────────────┘             │
│                         │                               │
└─────────────────────────┼───────────────────────────────┘
                          │
┌─────────────────────────┼───────────────────────────────┐
│ Phase 2: Services (Week 2)                               │
│ ┌─────────────┐  ┌─────────────┐                       │
│ │ Agent 1     │  │ Agent 2     │                       │
│ │ Data Service│  │ Algorithm   │                       │
│ │ (Needs      │  │ Service     │                       │
│ │  Interfaces)│  │ (Needs      │                       │
│ │             │  │  Validation)│                       │
│ └─────────────┘  └─────────────┘                       │
│        │                │                                │
│        └────────────────┘                                │
│                         │                                │
└─────────────────────────┼────────────────────────────────┘
                          │
┌─────────────────────────┼────────────────────────────────┐
│ Phase 3: Implementation (Week 3)                        │
│ ┌─────────────┐  ┌─────────────┐                       │
│ │ Agent 3     │  │ Agent 4     │                       │
│ │ Auto Mode   │  │ Manual Mode │                       │
│ │ Multi-Thread│  │ (Needs      │                       │
│ │ (Needs      │  │  Services)  │                       │
│ │  Services)  │  │             │                       │
│ └─────────────┘  └─────────────┘                       │
│        │                │                                │
│        └────────────────┘                                │
│                         │                                │
└─────────────────────────┼────────────────────────────────┘
                          │
┌─────────────────────────┴────────────────────────────────┐
│ Phase 4: Integration (Week 4)                             │
│ All Agents: Integration Testing & Bug Fixes               │
└──────────────────────────────────────────────────────────┘
```

---

## 2. Agent 1 Tasks (Core Infrastructure)

### 2.1 Task 1.1: Create Interfaces (Day 1-2)

**Files to Create:**
- `Services/Combining/Interfaces/ICombineSleeveService.cs`
- `Services/Combining/Interfaces/ICombineValidationService.cs`
- `Services/Combining/Interfaces/ICombineAlgorithmService.cs`
- `Services/Combining/Interfaces/ICombinePlacementService.cs`
- `Services/Combining/Interfaces/ICombineDataService.cs`
- `Services/Combining/Interfaces/IAutoCombineService.cs`
- `Services/Combining/Interfaces/IManualCombineService.cs`

**Dependencies:** None (can start immediately)

**Deliverables:**
- All interface definitions with XML documentation
- Method signatures with parameter types
- Return types defined

**Acceptance Criteria:**
- ✅ All interfaces compile
- ✅ XML documentation complete
- ✅ Method signatures match requirements

**Code Template:**
```csharp
namespace JSE_RevitAddin_MEP_OPENINGS.Services.Combining
{
    /// <summary>
    /// Main service interface for combine sleeve operations
    /// </summary>
    public interface ICombineSleeveService
    {
        /// <summary>
        /// Execute auto mode: combine sleeves for selected category
        /// </summary>
        CombineResult ExecuteAuto(Document doc, string category, double tolerance);
        
        /// <summary>
        /// Execute manual mode: combine two selected sleeves
        /// </summary>
        bool ExecuteManual(Document doc, ElementId sleeve1Id, ElementId sleeve2Id);
    }
}
```

---

### 2.2 Task 1.2: Create Base Classes (Day 2-3)

**Files to Create:**
- `Services/Combining/Base/CrashSafeServiceBase.cs`
- `Services/Combining/Base/CombineExtensions.cs`
- `Services/Combining/Models/CombineResult.cs`
- `Services/Combining/Models/CombinedSleeve.cs`
- `Services/Combining/Models/WallGroup.cs`

**Dependencies:** Task 1.1 (needs interfaces)

**Deliverables:**
- Base class with crash-safe patterns
- Extension methods for null-safe access
- Model classes for results

**Acceptance Criteria:**
- ✅ Base class compiles
- ✅ Extension methods work
- ✅ Model classes serializable

**Code Template:**
```csharp
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
                $"[{DateTime.Now:HH:mm:ss}] {context}: {ex.Message}\n");
            return defaultValue;
        }
    }
}
```

---

### 2.3 Task 1.3: Create Data Service (Day 4-5)

**Files to Create:**
- `Services/Combining/Data/CombineDataService.cs`
- `Services/Combining/Data/ICombineDataService.cs` (already in Task 1.1)

**Dependencies:** 
- Task 1.1 (interface)
- Task 1.2 (base class)
- Existing: `IClusterDataService` (already exists)

**Deliverables:**
- Data service that wraps `ClusterDataService`
- Caching implementation
- Fast lookup methods

**Acceptance Criteria:**
- ✅ Loads sleeves from database
- ✅ Caching works correctly
- ✅ Lookup methods are O(1) after cache

**Integration Points:**
- Uses `IClusterDataService.LoadClashZonesFromRegularXml()`
- Caches results in `ConcurrentDictionary`

---

### 2.4 Task 1.4: Create Factory (Day 5)

**Files to Create:**
- `Services/Combining/CombineServiceFactory.cs`

**Dependencies:** All previous Agent 1 tasks + Agent 2's validation service

**Deliverables:**
- Factory method to create fully-wired service
- Dependency injection setup

**Acceptance Criteria:**
- ✅ Creates service with all dependencies
- ✅ All services properly injected

---

## 3. Agent 2 Tasks (Algorithm & Validation)

### 3.1 Task 2.1: Create Validation Service (Day 1-3)

**Files to Create:**
- `Services/Combining/Validation/CombineValidationService.cs`
- `Services/Combining/Validation/ICombineValidationService.cs` (already in Task 1.1)

**Dependencies:** 
- Task 1.1 (interface)
- Task 1.2 (base class, WallGroup enum)

**Deliverables:**
- Wall group detection logic
- Rotation validation (straight axis only)
- Compatibility checking methods

**Acceptance Criteria:**
- ✅ Detects Wall X, Wall Y, Floor correctly
- ✅ Validates straight axis (0°, 90°, 180°, 270°)
- ✅ Returns false for incompatible combinations

**Code Template:**
```csharp
public class CombineValidationService : CrashSafeServiceBase, ICombineValidationService
{
    private readonly ICombineDataService _dataService;
    
    public bool AreCompatibleWallGroups(ClashZone sleeve1, ClashZone sleeve2)
    {
        var group1 = _dataService.GetWallGroup(sleeve1);
        var group2 = _dataService.GetWallGroup(sleeve2);
        return group1 == group2;
    }
    
    public bool AreCompatibleRotations(ClashZone sleeve1, ClashZone sleeve2)
    {
        return _dataService.IsStraightAxisAligned(sleeve1) &&
               _dataService.IsStraightAxisAligned(sleeve2);
    }
}
```

**Unit Tests Required:**
- Test wall group detection
- Test rotation validation
- Test compatibility checks

---

### 3.2 Task 2.2: Create Spatial Index (Day 3-4)

**Files to Create:**
- `Services/Combining/Algorithm/SpatialIndex.cs`

**Dependencies:** None (independent utility)

**Deliverables:**
- Spatial grid implementation
- Fast proximity queries
- Thread-safe operations

**Acceptance Criteria:**
- ✅ Builds index in < 1 second for 1000 sleeves
- ✅ Finds nearby sleeves in O(log n) time
- ✅ Thread-safe for parallel access

**Code Template:**
```csharp
public class SpatialIndex
{
    private readonly ConcurrentDictionary<(int x, int y, int z), List<int>> _grid;
    private readonly double _cellSize;
    
    public void AddSleeve(int sleeveId, XYZ position) { ... }
    public List<int> FindNearby(int sleeveId, XYZ position, double radius) { ... }
}
```

---

### 3.3 Task 2.3: Create Algorithm Service (Day 4-5)

**Files to Create:**
- `Services/Combining/Algorithm/CombineAlgorithmService.cs`
- `Services/Combining/Algorithm/ICombineAlgorithmService.cs` (already in Task 1.1)

**Dependencies:**
- Task 1.1 (interface)
- Task 1.2 (base class)
- Task 2.1 (validation service)
- Task 2.2 (spatial index)
- Existing: `IClusterAlgorithmService` (already exists)

**Deliverables:**
- Algorithm service that wraps `ClusterAlgorithmService`
- Proximity finding logic
- Integration with spatial index

**Acceptance Criteria:**
- ✅ Finds nearby sleeves correctly
- ✅ Uses spatial index for performance
- ✅ Validates compatibility before returning

**Integration Points:**
- Uses `IClusterAlgorithmService` for proximity logic
- Uses `ICombineValidationService` for validation
- Uses `SpatialIndex` for fast lookups

---

## 4. Agent 3 Tasks (Auto Mode & Multi-Threading)

### 4.1 Task 3.1: Create Auto Service Base (Day 1-2)

**Files to Create:**
- `Services/Combining/Auto/AutoCombineService.cs`
- `Services/Combining/Auto/IAutoCombineService.cs` (already in Task 1.1)

**Dependencies:**
- Task 1.1 (interfaces)
- Task 1.2 (base class)
- Task 1.3 (data service)
- Task 2.3 (algorithm service)

**Deliverables:**
- Auto service structure
- Category processing logic
- Basic batch processing

**Acceptance Criteria:**
- ✅ Processes single category
- ✅ Groups by wall group
- ✅ Finds compatible sleeves

**Note:** Multi-threading comes in Task 3.2

---

### 4.2 Task 3.2: Add Multi-Threading (Day 3-4)

**Files to Modify:**
- `Services/Combining/Auto/AutoCombineService.cs`

**Dependencies:** Task 3.1

**Deliverables:**
- Parallel processing of wall groups
- Thread-safe collections
- Semaphore control
- Batch parallelization

**Acceptance Criteria:**
- ✅ Processes 1000 sleeves in < 30 seconds
- ✅ No race conditions
- ✅ Proper error handling in parallel context

**Code Pattern (Reuse from Clustering):**
```csharp
// Same pattern as ClusterAlgorithmService.FormClusters
Parallel.ForEach(wallGroups, new ParallelOptions 
{ 
    MaxDegreeOfParallelism = Environment.ProcessorCount 
}, processWallGroup);
```

---

### 4.3 Task 3.3: Create Result Aggregator (Day 4-5)

**Files to Create:**
- `Services/Combining/Auto/CombineResultAggregator.cs`

**Dependencies:** Task 1.2 (models)

**Deliverables:**
- Thread-safe result collection
- Progress tracking
- Error aggregation

**Acceptance Criteria:**
- ✅ Thread-safe operations
- ✅ Accurate progress reporting
- ✅ All errors captured

---

## 5. Agent 4 Tasks (Manual Mode & UI)

### 5.1 Task 4.1: Create UI Dialog (Day 1-3)

**Files to Create:**
- `Views/CombineSleeveDialog.xaml`
- `Views/CombineSleeveDialog.xaml.cs`
- `ViewModels/CombineSleeveViewModel.cs`

**Dependencies:** None (UI can be built independently)

**Deliverables:**
- WPF dialog with radio buttons
- Auto/Manual mode UI
- Category dropdown
- Pick element buttons

**Acceptance Criteria:**
- ✅ UI displays correctly
- ✅ Radio buttons work
- ✅ Controls enable/disable properly

**UI Structure:**
```xml
<Window>
    <RadioButton Content="Auto" IsChecked="{Binding IsAutoMode}"/>
    <RadioButton Content="Manual" IsChecked="{Binding IsManualMode}"/>
    
    <!-- Auto Mode Controls -->
    <ComboBox ItemsSource="{Binding Categories}" 
              SelectedItem="{Binding SelectedCategory}"
              IsEnabled="{Binding IsAutoMode}"/>
    
    <!-- Manual Mode Controls -->
    <Button Content="Pick First Sleeve" 
            Command="{Binding PickFirstCommand}"
            IsEnabled="{Binding IsManualMode}"/>
    <Button Content="Pick Second Sleeve" 
            Command="{Binding PickSecondCommand}"
            IsEnabled="{Binding IsManualMode}"/>
</Window>
```

---

### 5.2 Task 4.2: Create Manual Service (Day 3-5)

**Files to Create:**
- `Services/Combining/Manual/ManualCombineService.cs`
- `Services/Combining/Manual/IManualCombineService.cs` (already in Task 1.1)

**Dependencies:**
- Task 1.1 (interfaces)
- Task 1.2 (base class)
- Task 1.3 (data service)
- Task 2.1 (validation service)
- Task 2.3 (algorithm service)

**Deliverables:**
- Element picker integration
- Manual validation
- Two-sleeve combination logic

**Acceptance Criteria:**
- ✅ Picks elements from Revit
- ✅ Validates compatibility
- ✅ Combines two sleeves correctly

**Integration Points:**
- Uses `UIDocument.Selection.PickObject()` for element picking
- Uses `ICombineValidationService` for validation
- Uses `ICombinePlacementService` for placement

---

### 5.3 Task 4.3: Create Command (Day 5)

**Files to Create:**
- `Commands/CombineSleeveCommand.cs`

**Dependencies:**
- Task 4.1 (UI dialog)
- Task 4.2 (manual service)
- Task 1.4 (factory)

**Deliverables:**
- ICommand implementation
- Dialog integration
- Service execution

**Acceptance Criteria:**
- ✅ Command executes correctly
- ✅ Dialog opens and closes properly
- ✅ Results returned correctly

---

## 6. Integration Points

### 6.1 Interface Contracts

**Agent 1 provides:**
- All interfaces (ICombineSleeveService, etc.)
- Base classes (CrashSafeServiceBase)
- Models (CombineResult, WallGroup)

**Agent 2 provides:**
- ICombineValidationService implementation
- ICombineAlgorithmService implementation
- SpatialIndex utility

**Agent 3 provides:**
- IAutoCombineService implementation
- Multi-threading logic

**Agent 4 provides:**
- IManualCombineService implementation
- UI dialog

### 6.2 Shared Dependencies

**All agents need:**
- Existing clustering services (already exist)
- `IClusterDataService`
- `IClusterAlgorithmService`
- `IClusterPlacementService`

**No agent depends on another agent's implementation**, only on:
- Interfaces (Agent 1)
- Base classes (Agent 1)
- Models (Agent 1)

### 6.3 Integration Sequence

1. **Week 1 End:** All interfaces and base classes ready
2. **Week 2 End:** All services implemented
3. **Week 3 End:** Auto and Manual modes working
4. **Week 4:** Integration testing and bug fixes

---

## 7. Testing Checkpoints

### 7.1 Week 1 Checkpoint

**Agent 1:**
- ✅ All interfaces compile
- ✅ Base classes work
- ✅ Models defined

**Agent 2:**
- ✅ Validation service unit tests pass
- ✅ Spatial index performance tests pass

**Agent 4:**
- ✅ UI dialog displays correctly
- ✅ Controls work independently

### 7.2 Week 2 Checkpoint

**Agent 1:**
- ✅ Data service loads sleeves
- ✅ Caching works
- ✅ Factory creates services

**Agent 2:**
- ✅ Algorithm service finds nearby sleeves
- ✅ Integration tests pass

**All Agents:**
- ✅ Services integrate via interfaces
- ✅ No compilation errors

### 7.3 Week 3 Checkpoint

**Agent 3:**
- ✅ Auto mode processes categories
- ✅ Multi-threading works
- ✅ Performance targets met

**Agent 4:**
- ✅ Manual mode works
- ✅ Element picking works
- ✅ Validation errors display

**All Agents:**
- ✅ End-to-end tests pass
- ✅ No race conditions

### 7.4 Week 4 Checkpoint

**All Agents:**
- ✅ Integration tests pass
- ✅ Performance benchmarks met
- ✅ User acceptance tests pass
- ✅ No critical bugs

---

## 8. Communication Protocol

### 8.1 Interface Changes

**If Agent 1 needs to change an interface:**
1. Update interface immediately
2. Notify all other agents
3. Update this document
4. Other agents adapt their implementations

### 8.2 Blocking Issues

**If an agent is blocked:**
1. Document the blocker
2. Create a stub/mock implementation
3. Continue with other tasks
4. Resolve blocker in integration phase

### 8.3 Daily Sync Points

**Each agent should:**
- Commit code daily
- Update progress in this document
- Report blockers immediately
- Share test results

---

## 9. File Structure Summary

```
Services/Combining/
├── Interfaces/              [Agent 1 - Week 1]
│   ├── ICombineSleeveService.cs
│   ├── ICombineValidationService.cs
│   ├── ICombineAlgorithmService.cs
│   ├── ICombinePlacementService.cs
│   ├── ICombineDataService.cs
│   ├── IAutoCombineService.cs
│   └── IManualCombineService.cs
│
├── Base/                    [Agent 1 - Week 1]
│   ├── CrashSafeServiceBase.cs
│   └── CombineExtensions.cs
│
├── Models/                  [Agent 1 - Week 1]
│   ├── CombineResult.cs
│   ├── CombinedSleeve.cs
│   └── WallGroup.cs
│
├── Data/                    [Agent 1 - Week 2]
│   └── CombineDataService.cs
│
├── Validation/              [Agent 2 - Week 1]
│   └── CombineValidationService.cs
│
├── Algorithm/               [Agent 2 - Week 2]
│   ├── CombineAlgorithmService.cs
│   └── SpatialIndex.cs
│
├── Auto/                    [Agent 3 - Week 3]
│   ├── AutoCombineService.cs
│   └── CombineResultAggregator.cs
│
├── Manual/                  [Agent 4 - Week 3]
│   └── ManualCombineService.cs
│
├── Placement/               [Agent 1 - Week 2]
│   └── CombinePlacementService.cs
│
└── CombineServiceFactory.cs [Agent 1 - Week 2]

Commands/                    [Agent 4 - Week 3]
└── CombineSleeveCommand.cs

Views/                       [Agent 4 - Week 1]
├── CombineSleeveDialog.xaml
└── CombineSleeveDialog.xaml.cs

ViewModels/                  [Agent 4 - Week 1]
└── CombineSleeveViewModel.cs
```

---

## 10. Parallel Work Matrix

| Task | Agent | Week | Can Start | Blocks |
|------|-------|------|-----------|--------|
| Interfaces | 1 | 1 | Day 1 | None |
| Base Classes | 1 | 1 | Day 2 | Interfaces |
| UI Dialog | 4 | 1 | Day 1 | None |
| Validation Service | 2 | 1 | Day 1 | Interfaces, Base |
| Data Service | 1 | 2 | Day 4 | Interfaces, Base |
| Spatial Index | 2 | 1 | Day 3 | None |
| Algorithm Service | 2 | 2 | Day 4 | Validation, Spatial |
| Placement Service | 1 | 2 | Day 5 | Interfaces, Base |
| Factory | 1 | 2 | Day 5 | All Agent 1 tasks |
| Auto Service | 3 | 3 | Day 1 | Data, Algorithm |
| Multi-Threading | 3 | 3 | Day 3 | Auto Service |
| Manual Service | 4 | 3 | Day 3 | Validation, Algorithm |
| Command | 4 | 3 | Day 5 | Manual, UI, Factory |

---

**End of Task Breakdown**


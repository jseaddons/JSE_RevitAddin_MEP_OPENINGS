# ClashZoneService & FlagManager SOLID Refactoring - 3 Team Parallel Implementation Plan

**Date:** December 2025  
**Status:** Ready for Parallel Execution  
**Branch:** `restore-today`

---

## Executive Summary

This document outlines a **3-team parallel execution plan** for refactoring `ClashZoneService` and `FlagManager` to achieve full SOLID compliance while **preserving all salient features, optimizations, and fail-safe mechanisms**.

### Team Division

- **Team F:** FlagManager Refactoring (Interface + Core Services)
- **Team G:** ClashZoneService Refactoring (Interface + Split Services)
- **Team H:** Integration & Wiring (Orchestrator - You)

---

## Invariants To Preserve (MUST NOT REGRESS)

### ✅ Critical Features
- **Fail-Safe:** No unhandled exceptions abort operations; continue-on-error semantics remain
- **Batch Flag Updates:** `BatchUpdateFlags()` optimization preserved (4-6× faster)
- **Batch Parameter Updates:** `ParameterBatchingService` integration preserved
- **SectionBox Support:** `SectionBoxHelper` reuse for 3D view filtering
- **Database-First:** All operations prioritize database over XML
- **Session Protection:** Recently placed cluster sleeves protected from deletion
- **Performance Optimizations:** Batch collection, HashSet lookups, pre-loaded entries

### ✅ Existing Code Reuse
- **SectionBoxHelper:** `GetSectionBoxBounds()`, `FilterElementsBySectionBox()`
- **ClashZoneRepository:** `BatchUpdateFlags()`, `GetClashZonesByCategoryInSectionBox()`
- **ParameterBatchingService:** `DeferParameter()`, `Flush()`
- **GlobalIndexService:** `GetAllEntries()`, flag operations
- **SafeFileLogger:** All logging operations

### ✅ Data Integrity
- Transaction safety: All database operations wrapped in transactions
- Rollback on failure: Prevents partial data corruption
- Flag hierarchy: Cluster flags take precedence over individual flags
- Revit API verification: Always verify sleeves exist in Revit (authoritative source)

---

## Team F: FlagManager Refactoring

### F.1 Deliverables

#### F.1.1 Interfaces
**File:** `Services/Interfaces/Refactor/IFlagManager.cs`
```csharp
namespace JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor
{
    /// <summary>
    /// Core flag management operations for clash zones.
    /// SOLID: Single Responsibility - flag operations only.
    /// </summary>
    public interface IFlagManager
    {
        /// <summary>
        /// Resets flags for deleted sleeves (database-first, then XML fallback).
        /// Preserves batch update optimization.
        /// </summary>
        int ResetFlagsForDeletedSleeves(
            List<ClashZone> clashZones, 
            List<string> categories,
            string refreshLogName = null);
        
        /// <summary>
        /// Resets instance IDs for deleted sleeves (database-first).
        /// Preserves batch collection optimization.
        /// </summary>
        int ResetInstanceIdsForDeletedSleeves(
            List<string> categories,
            Dictionary<string, List<ClashZone>> clashZonesByCategory = null,
            string refreshLogName = null);
        
        /// <summary>
        /// Updates flags after sleeve placement (batch update).
        /// </summary>
        void UpdateFlagsAfterPlacement(
            List<(Guid clashZoneId, int sleeveInstanceId, bool isCluster)> placedSleeves);
    }
}
```

**File:** `Services/Interfaces/Refactor/IInstanceIdManager.cs`
```csharp
namespace JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor
{
    /// <summary>
    /// Instance ID management for sleeves (separated from flag management).
    /// SOLID: Single Responsibility - instance ID operations only.
    /// </summary>
    public interface IInstanceIdManager
    {
        /// <summary>
        /// Resets instance IDs for deleted sleeves.
        /// Uses batch collection optimization (collect all sleeves once).
        /// </summary>
        int ResetInstanceIdsForDeletedSleeves(
            List<string> categories,
            Dictionary<string, List<ClashZone>> clashZonesByCategory = null);
    }
}
```

**File:** `Services/Interfaces/Refactor/ISessionTracker.cs`
```csharp
namespace JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor
{
    /// <summary>
    /// Session tracking for recently placed cluster sleeves (removes static state).
    /// SOLID: Single Responsibility - session tracking only.
    /// </summary>
    public interface ISessionTracker
    {
        /// <summary>
        /// Register a cluster sleeve as recently placed (protects from deletion).
        /// </summary>
        void RegisterRecentlyPlacedClusterSleeve(int clusterSleeveId);
        
        /// <summary>
        /// Clear recently placed cluster sleeves list.
        /// </summary>
        void ClearRecentlyPlacedClusterSleeves();
        
        /// <summary>
        /// Check if a cluster sleeve was recently placed.
        /// </summary>
        bool IsRecentlyPlacedClusterSleeve(int clusterSleeveId);
    }
}
```

#### F.1.2 Implementations

**File:** `Services/FlagManagement/FlagManagerService.cs`
```csharp
namespace JSE_RevitAddin_MEP_OPENINGS.Services.FlagManagement
{
    /// <summary>
    /// SOLID-compliant flag management service.
    /// Uses dependency injection for all dependencies.
    /// </summary>
    public class FlagManagerService : IFlagManager
    {
        private readonly Document _document;
        private readonly IClashZoneRepository _repository;
        private readonly IGlobalIndexService _globalIndexService;
        private readonly ILogger _logger;
        private readonly ISessionTracker _sessionTracker;
        
        public FlagManagerService(
            Document document,
            IClashZoneRepository repository,
            IGlobalIndexService globalIndexService,
            ILogger logger,
            ISessionTracker sessionTracker)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
            _repository = repository ?? throw new ArgumentNullException(nameof(repository));
            _globalIndexService = globalIndexService ?? throw new ArgumentNullException(nameof(globalIndexService));
            _logger = logger ?? LoggerAdapter.Default;
            _sessionTracker = sessionTracker ?? throw new ArgumentNullException(nameof(sessionTracker));
        }
        
        public int ResetFlagsForDeletedSleeves(
            List<ClashZone> clashZones, 
            List<string> categories,
            string refreshLogName = null)
        {
            // ✅ PRESERVE: Batch collection optimization
            // ✅ PRESERVE: HashSet lookup optimization
            // ✅ PRESERVE: Pre-loaded entries optimization
            // ✅ PRESERVE: Flag hierarchy (cluster first)
            // ✅ PRESERVE: Revit API verification (authoritative source)
            // ✅ PRESERVE: Batch database update via BatchUpdateFlags()
            // ✅ REUSE: SectionBoxHelper for section box detection
            // Implementation details in F.2
        }
        
        // ... other methods
    }
}
```

**File:** `Services/FlagManagement/InstanceIdManagerService.cs`
```csharp
namespace JSE_RevitAddin_MEP_OPENINGS.Services.FlagManagement
{
    /// <summary>
    /// SOLID-compliant instance ID management service.
    /// </summary>
    public class InstanceIdManagerService : IInstanceIdManager
    {
        private readonly Document _document;
        private readonly IClashZoneRepository _repository;
        private readonly ILogger _logger;
        
        public InstanceIdManagerService(
            Document document,
            IClashZoneRepository repository,
            ILogger logger)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
            _repository = repository ?? throw new ArgumentNullException(nameof(repository));
            _logger = logger ?? LoggerAdapter.Default;
        }
        
        public int ResetInstanceIdsForDeletedSleeves(
            List<string> categories,
            Dictionary<string, List<ClashZone>> clashZonesByCategory = null)
        {
            // ✅ PRESERVE: Batch collection optimization (collect all sleeves once)
            // ✅ PRESERVE: Category filtering in memory (O(1) lookup)
            // ✅ PRESERVE: Batch database update
        }
    }
}
```

**File:** `Services/FlagManagement/SessionTrackerService.cs`
```csharp
namespace JSE_RevitAddin_MEP_OPENINGS.Services.FlagManagement
{
    /// <summary>
    /// SOLID-compliant session tracker (removes static state).
    /// </summary>
    public class SessionTrackerService : ISessionTracker
    {
        private readonly HashSet<int> _recentlyPlacedClusterSleeveIds;
        private readonly object _lock;
        private readonly ILogger _logger;
        
        public SessionTrackerService(ILogger logger = null)
        {
            _recentlyPlacedClusterSleeveIds = new HashSet<int>();
            _lock = new object();
            _logger = logger ?? LoggerAdapter.Default;
        }
        
        public void RegisterRecentlyPlacedClusterSleeve(int clusterSleeveId)
        {
            if (clusterSleeveId <= 0) return;
            
            lock (_lock)
            {
                _recentlyPlacedClusterSleeveIds.Add(clusterSleeveId);
            }
            
            _logger.Info($"[SESSION-TRACKER] Registered recently placed cluster sleeve ID={clusterSleeveId}");
        }
        
        // ... other methods
    }
}
```

#### F.1.3 Adapter (Coexistence)

**File:** `Services/FlagManagement/FlagManagerAdapter.cs`
```csharp
namespace JSE_RevitAddin_MEP_OPENINGS.Services.FlagManagement
{
    /// <summary>
    /// Adapter that wraps existing FlagManager for IFlagManager interface.
    /// Enables gradual migration without breaking existing code.
    /// </summary>
    public class FlagManagerAdapter : IFlagManager
    {
        private readonly FlagManager _legacyFlagManager;
        private readonly ILogger _logger;
        
        public FlagManagerAdapter(FlagManager legacyFlagManager, ILogger logger = null)
        {
            _legacyFlagManager = legacyFlagManager ?? throw new ArgumentNullException(nameof(legacyFlagManager));
            _logger = logger ?? LoggerAdapter.Default;
        }
        
        public int ResetFlagsForDeletedSleeves(
            List<ClashZone> clashZones, 
            List<string> categories,
            string refreshLogName = null)
        {
            // Delegate to legacy FlagManager
            return _legacyFlagManager.ResetFlagsForDeletedSleeves(clashZones, categories, refreshLogName);
        }
        
        // ... other methods delegate to legacy
    }
}
```

### F.2 Implementation Details

#### F.2.1 Preserve Batch Flag Update Optimization

**Current Code (FlagManager.cs, lines 2156-2177):**
```csharp
// ✅ PRESERVE: Batch database update
var dbUpdates = updatesWithData.Select(u => (
    ClashZoneId: u.Id,
    IsResolved: u.IsResolved,
    IsClusterResolved: u.IsClusterResolved,
    SleeveInstanceId: u.SleeveInstanceId,
    ClusterInstanceId: u.ClusterSleeveInstanceId,
    // ... other fields
)).ToList();

repository.BatchUpdateFlags(dbUpdates);  // ✅ BATCH UPDATE (4-6× faster)
```

**Refactored Code (FlagManagerService.cs):**
```csharp
// ✅ PRESERVE: Same batch update logic
var dbUpdates = updatesWithData.Select(u => (
    ClashZoneId: u.Id,
    IsResolved: u.IsResolved,
    IsClusterResolved: u.IsClusterResolved,
    SleeveInstanceId: u.SleeveInstanceId,
    ClusterInstanceId: u.ClusterSleeveInstanceId,
    // ... other fields
)).ToList();

_repository.BatchUpdateFlags(dbUpdates);  // ✅ SAME BATCH UPDATE
```

#### F.2.2 Preserve Batch Collection Optimization

**Current Code (FlagManager.cs, lines 108-136):**
```csharp
// ✅ PRESERVE: Collect ALL sleeves from Revit ONCE (not per category)
if (categories.Count > 1)
{
    var allSleeves = new FilteredElementCollector(_document)
        .OfClass(typeof(FamilyInstance))
        .Cast<FamilyInstance>()
        .Where(s => /* sleeve filter */)
        .ToList();
    
    // Build category index
    sleevesByCategory = new Dictionary<string, HashSet<int>>();
    // ... build index
}
```

**Refactored Code (InstanceIdManagerService.cs):**
```csharp
// ✅ PRESERVE: Same batch collection logic
if (categories.Count > 1)
{
    var allSleeves = new FilteredElementCollector(_document)
        .OfClass(typeof(FamilyInstance))
        .Cast<FamilyInstance>()
        .Where(s => /* sleeve filter */)
        .ToList();
    
    // ✅ SAME: Build category index
    sleevesByCategory = new Dictionary<string, HashSet<int>>();
    // ... build index
}
```

#### F.2.3 Preserve SectionBoxHelper Reuse

**Current Code (FlagManager.cs, lines 178-198):**
```csharp
// ✅ REUSE: SectionBoxHelper for section box detection
BoundingBoxXYZ sectionBox = null;
if (view3D != null && view3D.IsSectionBoxActive)
{
    sectionBox = SectionBoxHelper.GetSectionBoxBounds(view3D);
}

// ✅ REUSE: Section-box-aware query
if (sectionBox != null && OptimizationFlags.UseRTreeDatabaseIndex)
{
    dbZones = repository.GetClashZonesByCategoryInSectionBox(category, sectionBox);
}
```

**Refactored Code (FlagManagerService.cs):**
```csharp
// ✅ REUSE: Same SectionBoxHelper usage
BoundingBoxXYZ sectionBox = null;
if (view3D != null && view3D.IsSectionBoxActive)
{
    sectionBox = SectionBoxHelper.GetSectionBoxBounds(view3D);  // ✅ REUSE
}

// ✅ REUSE: Same section-box-aware query
if (sectionBox != null && _config.Detection.UseRTreeIndex)
{
    dbZones = _repository.GetClashZonesByCategoryInSectionBox(category, sectionBox);  // ✅ REUSE
}
```

### F.3 Constraints

- ✅ **Must preserve** all batch update optimizations
- ✅ **Must reuse** `SectionBoxHelper` for section box operations
- ✅ **Must preserve** flag hierarchy (cluster flags first)
- ✅ **Must preserve** Revit API verification (authoritative source)
- ✅ **Must preserve** fail-safe error handling
- ✅ **Must not** break existing callers (use adapter pattern)

### F.4 Acceptance Criteria

- ✅ `IFlagManager` interface created
- ✅ `IInstanceIdManager` interface created
- ✅ `ISessionTracker` interface created
- ✅ All implementations use dependency injection
- ✅ All batch optimizations preserved
- ✅ SectionBoxHelper reused
- ✅ Adapter created for coexistence
- ✅ Unit tests can mock all dependencies

---

## Team G: ClashZoneService Refactoring

### G.1 Deliverables

#### G.1.1 Interfaces

**File:** `Services/Interfaces/Refactor/IClashZoneService.cs`
```csharp
namespace JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor
{
    /// <summary>
    /// Main clash zone service interface (orchestrator).
    /// SOLID: Depends on abstractions only.
    /// </summary>
    public interface IClashZoneService
    {
        /// <summary>
        /// Cleanup invalid clash zones (removes null elements and duplicates).
        /// </summary>
        int CleanupInvalidClashZones(Document document);
        
        /// <summary>
        /// Filter clash zones by current selection parameters.
        /// </summary>
        List<ClashZone> FilterClashZonesByCurrentSelection(
            List<string> selectedReferenceFiles,
            Dictionary<string, double> currentClearanceSettings,
            string currentPrefix,
            Document document);
    }
}
```

**File:** `Services/Interfaces/Refactor/IClashZoneCleanupService.cs`
```csharp
namespace JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor
{
    /// <summary>
    /// Clash zone cleanup operations (removed from ClashZoneService).
    /// SOLID: Single Responsibility - cleanup only.
    /// </summary>
    public interface IClashZoneCleanupService
    {
        /// <summary>
        /// Remove invalid clash zones (null elements) and duplicates.
        /// </summary>
        int CleanupInvalidClashZones(
            List<ClashZone> clashZones,
            Document document);
    }
}
```

**File:** `Services/Interfaces/Refactor/IClashZoneFilterService.cs`
```csharp
namespace JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor
{
    /// <summary>
    /// Clash zone filtering operations (removed from ClashZoneService).
    /// SOLID: Single Responsibility - filtering only.
    /// </summary>
    public interface IClashZoneFilterService
    {
        /// <summary>
        /// Filter clash zones by current selection parameters.
        /// Uses SectionBoxHelper for section box filtering.
        /// </summary>
        List<ClashZone> FilterClashZonesByCurrentSelection(
            List<ClashZone> clashZones,
            List<string> selectedReferenceFiles,
            Dictionary<string, double> currentClearanceSettings,
            string currentPrefix,
            Document document);
    }
}
```

**File:** `Services/Interfaces/Refactor/IClashZoneValidationService.cs`
```csharp
namespace JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor
{
    /// <summary>
    /// Clash zone validation operations (removed from ClashZoneService).
    /// SOLID: Single Responsibility - validation only.
    /// </summary>
    public interface IClashZoneValidationService
    {
        /// <summary>
        /// Validate clash zone elements exist and still intersect.
        /// </summary>
        bool ValidateClashZone(ClashZone clashZone, Document document);
    }
}
```

#### G.1.2 Implementations

**File:** `Services/ClashZoneManagement/ClashZoneService.cs`
```csharp
namespace JSE_RevitAddin_MEP_OPENINGS.Services.ClashZoneManagement
{
    /// <summary>
    /// SOLID-compliant clash zone service (orchestrator).
    /// Delegates to focused services.
    /// </summary>
    public class ClashZoneService : IClashZoneService
    {
        private readonly IClashZoneCleanupService _cleanupService;
        private readonly IClashZoneFilterService _filterService;
        private readonly IClashZoneValidationService _validationService;
        private readonly ClashZoneStorage? _storage;
        private readonly ILogger _logger;
        
        public ClashZoneService(
            IClashZoneCleanupService cleanupService,
            IClashZoneFilterService filterService,
            IClashZoneValidationService validationService,
            ClashZoneStorage? storage,
            ILogger logger)
        {
            _cleanupService = cleanupService ?? throw new ArgumentNullException(nameof(cleanupService));
            _filterService = filterService ?? throw new ArgumentNullException(nameof(filterService));
            _validationService = validationService ?? throw new ArgumentNullException(nameof(validationService));
            _storage = storage;
            _logger = logger ?? LoggerAdapter.Default;
        }
        
        public int CleanupInvalidClashZones(Document document)
        {
            if (_storage?.ClashZones == null)
            {
                _logger.Info("No clash zones to clean up");
                return 0;
            }
            
            // ✅ DELEGATE: Use focused cleanup service
            var removed = _cleanupService.CleanupInvalidClashZones(_storage.ClashZones, document);
            _storage.ClashZones.RemoveAll(cz => /* removed logic */);
            
            return removed;
        }
        
        // ... other methods delegate to focused services
    }
}
```

**File:** `Services/ClashZoneManagement/ClashZoneCleanupService.cs`
```csharp
namespace JSE_RevitAddin_MEP_OPENINGS.Services.ClashZoneManagement
{
    /// <summary>
    /// SOLID-compliant clash zone cleanup service.
    /// </summary>
    public class ClashZoneCleanupService : IClashZoneCleanupService
    {
        private readonly ILogger _logger;
        
        public ClashZoneCleanupService(ILogger logger = null)
        {
            _logger = logger ?? LoggerAdapter.Default;
        }
        
        public int CleanupInvalidClashZones(
            List<ClashZone> clashZones,
            Document document)
        {
            // ✅ PRESERVE: Original cleanup logic from ClashZoneService.CleanupInvalidClashZones()
            // Step 1: Remove invalid clash zones (null elements)
            // Step 2: Remove duplicates (keep first occurrence)
            // ✅ PRESERVE: Fail-safe error handling
        }
    }
}
```

**File:** `Services/ClashZoneManagement/ClashZoneFilterService.cs`
```csharp
namespace JSE_RevitAddin_MEP_OPENINGS.Services.ClashZoneManagement
{
    /// <summary>
    /// SOLID-compliant clash zone filter service.
    /// Uses SectionBoxHelper for section box filtering.
    /// </summary>
    public class ClashZoneFilterService : IClashZoneFilterService
    {
        private readonly ILogger _logger;
        
        public ClashZoneFilterService(ILogger logger = null)
        {
            _logger = logger ?? LoggerAdapter.Default;
        }
        
        public List<ClashZone> FilterClashZonesByCurrentSelection(
            List<ClashZone> clashZones,
            List<string> selectedReferenceFiles,
            Dictionary<string, double> currentClearanceSettings,
            string currentPrefix,
            Document document)
        {
            // ✅ PRESERVE: Original filtering logic
            // ✅ REUSE: SectionBoxHelper.GetSectionBoxBounds() for section box filtering
            // ✅ REUSE: SectionBoxHelper.IsClashZoneVisibleInCurrentSectionBox() if exists
            // ✅ PRESERVE: Fail-safe error handling
        }
    }
}
```

#### G.1.3 Adapter (Coexistence)

**File:** `Services/ClashZoneManagement/ClashZoneServiceAdapter.cs`
```csharp
namespace JSE_RevitAddin_MEP_OPENINGS.Services.ClashZoneManagement
{
    /// <summary>
    /// Adapter that wraps existing ClashZoneService for IClashZoneService interface.
    /// Enables gradual migration without breaking existing code.
    /// </summary>
    public class ClashZoneServiceAdapter : IClashZoneService
    {
        private readonly Services.ClashZoneService _legacyService;
        private readonly ILogger _logger;
        
        public ClashZoneServiceAdapter(
            Services.ClashZoneService legacyService,
            ILogger logger = null)
        {
            _legacyService = legacyService ?? throw new ArgumentNullException(nameof(legacyService));
            _logger = logger ?? LoggerAdapter.Default;
        }
        
        public int CleanupInvalidClashZones(Document document)
        {
            // Delegate to legacy ClashZoneService
            return _legacyService.CleanupInvalidClashZones(document);
        }
        
        // ... other methods delegate to legacy
    }
}
```

### G.2 Implementation Details

#### G.2.1 Preserve SectionBoxHelper Reuse

**Current Code (ClashZoneService.cs, lines 1317-1336):**
```csharp
private bool IsClashZoneVisibleInCurrentSectionBox(ClashZone clashZone, Document document)
{
    if (!(document.ActiveView is View3D view3D) || !view3D.IsSectionBoxActive)
        return true;
    
    // ✅ REUSE: SectionBoxHelper
    var sectionBox = SectionBoxHelper.GetSectionBoxBounds(view3D);
    if (sectionBox == null)
        return true;
    
    // Check if clash zone intersection point is within section box
    // ...
}
```

**Refactored Code (ClashZoneFilterService.cs):**
```csharp
private bool IsClashZoneVisibleInCurrentSectionBox(ClashZone clashZone, Document document)
{
    if (!(document.ActiveView is View3D view3D) || !view3D.IsSectionBoxActive)
        return true;
    
    // ✅ REUSE: Same SectionBoxHelper
    var sectionBox = SectionBoxHelper.GetSectionBoxBounds(view3D);  // ✅ REUSE
    if (sectionBox == null)
        return true;
    
    // ✅ PRESERVE: Same intersection point check
    // ...
}
```

### G.3 Constraints

- ✅ **Must preserve** all filtering logic
- ✅ **Must reuse** `SectionBoxHelper` for section box operations
- ✅ **Must preserve** fail-safe error handling
- ✅ **Must not** break existing callers (use adapter pattern)
- ✅ **Must preserve** cleanup logic (invalid + duplicates)

### G.4 Acceptance Criteria

- ✅ `IClashZoneService` interface created
- ✅ `IClashZoneCleanupService` interface created
- ✅ `IClashZoneFilterService` interface created
- ✅ `IClashZoneValidationService` interface created
- ✅ All implementations use dependency injection
- ✅ SectionBoxHelper reused
- ✅ Adapter created for coexistence
- ✅ Unit tests can mock all dependencies

---

## Team H: Integration & Wiring (You)

### H.1 Deliverables

#### H.1.1 Factory Classes

**File:** `Services/FlagManagement/FlagManagerFactory.cs`
```csharp
namespace JSE_RevitAddin_MEP_OPENINGS.Services.FlagManagement
{
    /// <summary>
    /// Factory for creating fully-wired flag management services.
    /// Handles dependency injection and coexistence.
    /// </summary>
    public static class FlagManagerFactory
    {
        /// <summary>
        /// Create refactored flag management services (SOLID-compliant).
        /// </summary>
        public static (IFlagManager flagManager, IInstanceIdManager instanceIdManager, ISessionTracker sessionTracker) CreateRefactored(
            Document document,
            IClashZoneRepository repository = null,
            IGlobalIndexService globalIndexService = null,
            ILogger logger = null,
            IOptimizationConfig config = null)
        {
            // ✅ WIRING: Create dependencies
            var dbContext = new SleeveDbContext();
            repository = repository ?? new ClashZoneRepository(dbContext);
            globalIndexService = globalIndexService ?? new GlobalIndexService();
            logger = logger ?? LoggerAdapter.Default;
            config = config ?? OptimizationConfigFactory.Default;
            
            // ✅ WIRING: Create session tracker
            var sessionTracker = new SessionTrackerService(logger);
            
            // ✅ WIRING: Create instance ID manager
            var instanceIdManager = new InstanceIdManagerService(document, repository, logger);
            
            // ✅ WIRING: Create flag manager
            var flagManager = new FlagManagerService(
                document,
                repository,
                globalIndexService,
                logger,
                sessionTracker);
            
            return (flagManager, instanceIdManager, sessionTracker);
        }
        
        /// <summary>
        /// Create adapter for legacy FlagManager (coexistence).
        /// </summary>
        public static IFlagManager CreateAdapter(
            FlagManager legacyFlagManager,
            ILogger logger = null)
        {
            return new FlagManagerAdapter(legacyFlagManager, logger ?? LoggerAdapter.Default);
        }
    }
}
```

**File:** `Services/ClashZoneManagement/ClashZoneServiceFactory.cs`
```csharp
namespace JSE_RevitAddin_MEP_OPENINGS.Services.ClashZoneManagement
{
    /// <summary>
    /// Factory for creating fully-wired clash zone services.
    /// Handles dependency injection and coexistence.
    /// </summary>
    public static class ClashZoneServiceFactory
    {
        /// <summary>
        /// Create refactored clash zone service (SOLID-compliant).
        /// </summary>
        public static IClashZoneService CreateRefactored(
            ClashZoneStorage? storage = null,
            ILogger logger = null)
        {
            logger = logger ?? LoggerAdapter.Default;
            
            // ✅ WIRING: Create focused services
            var cleanupService = new ClashZoneCleanupService(logger);
            var filterService = new ClashZoneFilterService(logger);
            var validationService = new ClashZoneValidationService(logger);
            
            // ✅ WIRING: Create main service
            var service = new ClashZoneService(
                cleanupService,
                filterService,
                validationService,
                storage,
                logger);
            
            return service;
        }
        
        /// <summary>
        /// Create adapter for legacy ClashZoneService (coexistence).
        /// </summary>
        public static IClashZoneService CreateAdapter(
            Services.ClashZoneService legacyService,
            ILogger logger = null)
        {
            return new ClashZoneServiceAdapter(legacyService, logger ?? LoggerAdapter.Default);
        }
    }
}
```

#### H.1.2 Integration Points

**File:** `Services/Integration/ClashZoneFlagManagerIntegration.cs`
```csharp
namespace JSE_RevitAddin_MEP_OPENINGS.Services.Integration
{
    /// <summary>
    /// Integration service that wires ClashZoneService and FlagManager together.
    /// Handles migration from legacy to refactored services.
    /// </summary>
    public class ClashZoneFlagManagerIntegration
    {
        private readonly IClashZoneService _clashZoneService;
        private readonly IFlagManager _flagManager;
        private readonly IInstanceIdManager _instanceIdManager;
        private readonly ISessionTracker _sessionTracker;
        private readonly ILogger _logger;
        
        public ClashZoneFlagManagerIntegration(
            IClashZoneService clashZoneService,
            IFlagManager flagManager,
            IInstanceIdManager instanceIdManager,
            ISessionTracker sessionTracker,
            ILogger logger = null)
        {
            _clashZoneService = clashZoneService ?? throw new ArgumentNullException(nameof(clashZoneService));
            _flagManager = flagManager ?? throw new ArgumentNullException(nameof(flagManager));
            _instanceIdManager = instanceIdManager ?? throw new ArgumentNullException(nameof(instanceIdManager));
            _sessionTracker = sessionTracker ?? throw new ArgumentNullException(nameof(sessionTracker));
            _logger = logger ?? LoggerAdapter.Default;
        }
        
        /// <summary>
        /// Execute refresh with integrated services.
        /// </summary>
        public void ExecuteRefresh(
            Document document,
            List<string> categories,
            Dictionary<string, List<ClashZone>> clashZonesByCategory)
        {
            // ✅ WIRING: Use refactored services
            // 1. Cleanup invalid clash zones
            var cleanupCount = _clashZoneService.CleanupInvalidClashZones(document);
            _logger.Info($"Cleaned up {cleanupCount} invalid clash zones");
            
            // 2. Reset flags for deleted sleeves
            foreach (var category in categories)
            {
                var clashZones = clashZonesByCategory.GetValueOrDefault(category, new List<ClashZone>());
                var resetCount = _flagManager.ResetFlagsForDeletedSleeves(clashZones, new List<string> { category });
                _logger.Info($"Reset flags for {resetCount} deleted sleeves in category '{category}'");
            }
            
            // 3. Reset instance IDs
            var instanceResetCount = _instanceIdManager.ResetInstanceIdsForDeletedSleeves(categories, clashZonesByCategory);
            _logger.Info($"Reset instance IDs for {instanceResetCount} deleted sleeves");
        }
    }
}
```

#### H.1.3 Migration Strategy

**File:** `Services/Integration/MigrationStrategy.cs`
```csharp
namespace JSE_RevitAddin_MEP_OPENINGS.Services.Integration
{
    /// <summary>
    /// Migration strategy for gradually moving from legacy to refactored services.
    /// </summary>
    public static class MigrationStrategy
    {
        /// <summary>
        /// Create services based on feature flag (gradual migration).
        /// </summary>
        public static (IClashZoneService clashZoneService, IFlagManager flagManager) CreateServices(
            Document document,
            ClashZoneStorage? storage = null,
            FlagManager? legacyFlagManager = null,
            Services.ClashZoneService? legacyClashZoneService = null,
            bool useRefactored = false)
        {
            if (useRefactored)
            {
                // ✅ WIRING: Use refactored services
                var clashZoneService = ClashZoneServiceFactory.CreateRefactored(storage);
                var (flagManager, instanceIdManager, sessionTracker) = FlagManagerFactory.CreateRefactored(document);
                return (clashZoneService, flagManager);
            }
            else
            {
                // ✅ WIRING: Use adapters for legacy services (coexistence)
                IClashZoneService clashZoneService = null;
                if (legacyClashZoneService != null)
                {
                    clashZoneService = ClashZoneServiceFactory.CreateAdapter(legacyClashZoneService);
                }
                
                IFlagManager flagManager = null;
                if (legacyFlagManager != null)
                {
                    flagManager = FlagManagerFactory.CreateAdapter(legacyFlagManager);
                }
                
                return (clashZoneService, flagManager);
            }
        }
    }
}
```

### H.2 Wiring Strategy

#### H.2.1 Update Existing Callers

**File:** `refresh refactor/refresh_service_refactored.cs`
```csharp
// ✅ WIRING: Update to use refactored services
private IClashZoneService _clashZoneService;
private IFlagManager _flagManager;

// In constructor or initialization:
var (clashZoneService, flagManager) = MigrationStrategy.CreateServices(
    document,
    storage: context.ExistingClashZones,
    useRefactored: OptimizationFlags.UseRefactoredClashZoneFlagServices);

_clashZoneService = clashZoneService;
_flagManager = flagManager;
```

**File:** `Services/Clustering/RefactoredClusterService.cs`
```csharp
// ✅ WIRING: Update to use IFlagManager interface
private readonly IFlagManager? _flagManager;

public RefactoredClusterService(
    // ... existing parameters ...
    IFlagManager? flagManager = null)  // ✅ Changed from FlagManager to IFlagManager
{
    // ... existing code ...
    _flagManager = flagManager;
}
```

### H.3 Constraints

- ✅ **Must preserve** all existing functionality
- ✅ **Must enable** gradual migration (feature flag)
- ✅ **Must not** break existing callers
- ✅ **Must wire** all dependencies correctly
- ✅ **Must handle** coexistence (legacy + refactored)

### H.4 Acceptance Criteria

- ✅ Factory classes created
- ✅ Integration service created
- ✅ Migration strategy implemented
- ✅ Existing callers updated (with feature flag)
- ✅ All dependencies wired correctly
- ✅ Coexistence pattern working
- ✅ No breaking changes

---

## Implementation Sequence

### Phase 1: Interface Definition (Week 1)
- **Team F:** Create `IFlagManager`, `IInstanceIdManager`, `ISessionTracker` interfaces
- **Team G:** Create `IClashZoneService`, `IClashZoneCleanupService`, `IClashZoneFilterService`, `IClashZoneValidationService` interfaces
- **Team H:** Review interfaces, provide feedback

### Phase 2: Implementation (Week 2)
- **Team F:** Implement `FlagManagerService`, `InstanceIdManagerService`, `SessionTrackerService`, `FlagManagerAdapter`
- **Team G:** Implement `ClashZoneService`, `ClashZoneCleanupService`, `ClashZoneFilterService`, `ClashZoneValidationService`, `ClashZoneServiceAdapter`
- **Team H:** Create factory classes, integration service, migration strategy

### Phase 3: Wiring & Integration (Week 3)
- **Team H:** Wire services together, update existing callers, test integration
- **Team F & G:** Support Team H with any needed adjustments

### Phase 4: Testing & Migration (Week 4)
- **All Teams:** Unit tests, integration tests, performance benchmarks
- **Team H:** Enable feature flag, monitor for issues, gradual rollout

---

## Safety Checklist

### ✅ Fail-Safe Mechanisms
- [ ] All operations wrapped in try-catch
- [ ] Continue-on-error semantics preserved
- [ ] Transaction rollback on failure
- [ ] Logging for all errors

### ✅ Optimizations Preserved
- [ ] Batch flag updates (`BatchUpdateFlags()`)
- [ ] Batch collection (collect all sleeves once)
- [ ] HashSet lookups (O(1) instead of O(n))
- [ ] Pre-loaded entries (calculate once, use many times)

### ✅ Code Reuse
- [ ] `SectionBoxHelper` reused for section box operations
- [ ] `ClashZoneRepository.BatchUpdateFlags()` reused
- [ ] `GlobalIndexService.GetAllEntries()` reused
- [ ] `SafeFileLogger` reused for logging

### ✅ Data Integrity
- [ ] Flag hierarchy preserved (cluster flags first)
- [ ] Revit API verification preserved (authoritative source)
- [ ] Database-first operations preserved
- [ ] Transaction safety preserved

---

## Risk Mitigation

1. **Feature Flag:** Use `OptimizationFlags.UseRefactoredClashZoneFlagServices` to enable/disable
2. **Adapter Pattern:** Legacy services wrapped in adapters for coexistence
3. **Gradual Migration:** Start with new code, migrate existing code gradually
4. **Comprehensive Testing:** Unit tests, integration tests, performance benchmarks
5. **Rollback Plan:** Can disable feature flag to revert to legacy services

---

## Success Metrics

- ✅ All interfaces created and implemented
- ✅ All optimizations preserved (performance within 2% of baseline)
- ✅ All fail-safe mechanisms preserved
- ✅ All code reuse implemented
- ✅ Zero breaking changes
- ✅ All unit tests passing
- ✅ Integration tests passing
- ✅ Performance benchmarks met

---

**Document Status:** ✅ Ready for Parallel Execution  
**Next Steps:** Teams F, G, and H begin Phase 1 (Interface Definition)


# FilterManagementService SOLID Refactoring - 2 Team Parallel Implementation Plan

**Date:** December 2025  
**Status:** Ready for Parallel Execution  
**Branch:** `restore-today`

---

## Executive Summary

This document outlines a **2-team parallel execution plan** for refactoring `FilterManagementService` to achieve full SOLID compliance while **preserving all salient features, optimizations, and fail-safe mechanisms**.

### Team Division

- **Team I:** Filter Persistence & Data Operations (Database + XML)
- **Team J:** Filter Business Logic & UI Orchestration (CRUD Operations)

---

## Invariants To Preserve (MUST NOT REGRESS)

### ✅ Critical Features
- **Database-First:** All operations prioritize database over XML
- **UI State Persistence:** SelectedHostCategories and OpeningSettings preserved
- **Category Validation:** No fallback to stale filter object data
- **Backward Compatibility:** XML fallback for legacy systems
- **Fail-Safe:** No unhandled exceptions abort operations; continue-on-error semantics remain
- **Filter Registration:** Filters must be registered in database before UI operations

### ✅ Existing Code Reuse
- **FilterRepository:** `EnsureFilter()`, `UpdateFilterName()`, `DeleteFilter()`, `SaveFilterUIState()`
- **ProjectPathService:** `GetFiltersDirectory()`, `EnsureFiltersDirectory()`
- **FilterUiStateProvider:** `GetSelectedMepCategoryNames`, `GetSelectedHostCategories`
- **MepCategoryConstants:** `Normalize()` for category standardization
- **SafeFileLogger:** All logging operations

### ✅ Data Integrity
- Transaction safety: All database operations wrapped in transactions
- Rollback on failure: Prevents partial data corruption
- Category validation: Must get from UI state, not stale filter object
- Filter uniqueness: No duplicate filter names per category

---

## Current Architecture Analysis

### Current Responsibilities (Violates SRP)
`FilterManagementService` currently handles:
1. **Filter CRUD Operations** (Create, Copy, Rename, Delete)
2. **Filter Persistence** (Database + XML save/load)
3. **UI State Management** (CreateFilterFromCurrentUIState, etc.)
4. **Filter Validation** (IsFilterSaved, category validation)
5. **UI Interaction** (GetFilterNameFromUser, ShowError, ListBox operations)
6. **File System Operations** (XML file I/O)

### SOLID Violations
- **Single Responsibility Principle (SRP):** Service has 6+ responsibilities
- **Dependency Inversion Principle (DIP):** Direct dependencies on `FilterRepository`, `SleeveDbContext`, file system
- **Open/Closed Principle (OCP):** Hard to extend without modifying existing code
- **Interface Segregation Principle (ISP):** No interfaces, clients must depend on entire service

---

## Team I: Filter Persistence & Data Operations

### I.1 Deliverables

#### I.1.1 Interfaces

**File:** `Services/Interfaces/Refactor/IFilterRepository.cs`
```csharp
namespace JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor
{
    /// <summary>
    /// Team I: Database operations for filters.
    /// SOLID: Single Responsibility - database operations only.
    /// </summary>
    public interface IFilterRepository
    {
        /// <summary>
        /// Ensures filter exists in database (creates if not exists).
        /// Returns FilterId if successful, -1 if failed.
        /// </summary>
        int EnsureFilter(string filterName, string category);
        
        /// <summary>
        /// Updates filter name in database.
        /// </summary>
        void UpdateFilterName(string oldName, string category, string newName);
        
        /// <summary>
        /// Deletes filter from database.
        /// </summary>
        void DeleteFilter(string filterName, string category);
        
        /// <summary>
        /// Gets all filters from database.
        /// </summary>
        List<FilterInfo> GetAllFilters();
        
        /// <summary>
        /// Gets filter ID by name and category.
        /// </summary>
        int GetFilterId(string filterName, string category);
        
        /// <summary>
        /// Saves filter UI state to database.
        /// </summary>
        void SaveFilterUIState(string filterName, string category, List<string> selectedHostCategories, OpeningSettings settings);
        
        /// <summary>
        /// Loads filter UI state from database.
        /// </summary>
        (List<string> selectedHostCategories, OpeningSettings settings) LoadFilterUIState(string filterName, string category);
    }
    
    /// <summary>
    /// Filter information from database.
    /// </summary>
    public class FilterInfo
    {
        public int FilterId { get; set; }
        public string FilterName { get; set; }
        public string Category { get; set; }
    }
}
```

**File:** `Services/Interfaces/Refactor/IFilterPersistenceService.cs`
```csharp
namespace JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor
{
    /// <summary>
    /// Team I: Filter persistence operations (database + XML).
    /// SOLID: Single Responsibility - persistence operations only.
    /// </summary>
    public interface IFilterPersistenceService
    {
        /// <summary>
        /// Saves filter to database (primary storage).
        /// Returns FilterId if successful, -1 if failed.
        /// </summary>
        int SaveFilterToDatabase(string filterName, string category, OpeningFilter filter);
        
        /// <summary>
        /// Loads filter from database.
        /// Returns null if not found.
        /// </summary>
        OpeningFilter LoadFilterFromDatabase(string filterName, string category);
        
        /// <summary>
        /// Saves filter to XML file (backward compatibility).
        /// </summary>
        void SaveFilterToXmlFile(OpeningFilter filter, string filePath);
        
        /// <summary>
        /// Loads filter from XML file (backward compatibility).
        /// Returns null if file not found.
        /// </summary>
        OpeningFilter LoadFilterFromXmlFile(string filePath);
        
        /// <summary>
        /// Checks if filter is saved (exists in database or XML).
        /// </summary>
        bool IsFilterSaved(string filterName, string category = null);
        
        /// <summary>
        /// Gets all saved filter names from database and XML.
        /// </summary>
        List<string> GetAllSavedFilterNames();
    }
}
```

#### I.1.2 Implementations

**File:** `Services/FilterManagement/FilterRepositoryAdapter.cs`
```csharp
namespace JSE_RevitAddin_MEP_OPENINGS.Services.FilterManagement
{
    /// <summary>
    /// Team I: Adapter for FilterRepository to implement IFilterRepository.
    /// SOLID: Adapter pattern for dependency inversion.
    /// </summary>
    public class FilterRepositoryAdapter : IFilterRepository
    {
        private readonly Document _document;
        private readonly ILogger _logger;
        
        public FilterRepositoryAdapter(Document document, ILogger logger = null)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
            _logger = logger ?? LoggerAdapter.Default;
        }
        
        public int EnsureFilter(string filterName, string category)
        {
            int filterId = -1;
            UseFilterRepository(repo => 
            {
                filterId = repo.EnsureFilter(filterName, category);
            });
            return filterId;
        }
        
        // ... other methods delegate to FilterRepository
        
        private void UseFilterRepository(Action<FilterRepository> action)
        {
            if (action == null) return;
            
            try
            {
                using (var context = new SleeveDbContext(_document, msg => _logger.Info($"[FilterRepositoryAdapter] {msg}")))
                {
                    var repository = new FilterRepository(context, msg => _logger.Info($"[FilterRepositoryAdapter] {msg}"));
                    action(repository);
                }
            }
            catch (Exception ex)
            {
                _logger.Error($"FilterRepository operation failed: {ex.Message}", ex, "FilterRepositoryAdapter");
                throw;
            }
        }
    }
}
```

**File:** `Services/FilterManagement/FilterPersistenceService.cs`
```csharp
namespace JSE_RevitAddin_MEP_OPENINGS.Services.FilterManagement
{
    /// <summary>
    /// Team I: SOLID-compliant filter persistence service.
    /// Handles database-first persistence with XML fallback.
    /// </summary>
    public class FilterPersistenceService : IFilterPersistenceService
    {
        private readonly IFilterRepository _repository;
        private readonly Document _document;
        private readonly ILogger _logger;
        
        public FilterPersistenceService(
            IFilterRepository repository,
            Document document,
            ILogger logger = null)
        {
            _repository = repository ?? throw new ArgumentNullException(nameof(repository));
            _document = document ?? throw new ArgumentNullException(nameof(document));
            _logger = logger ?? LoggerAdapter.Default;
        }
        
        public int SaveFilterToDatabase(string filterName, string category, OpeningFilter filter)
        {
            // ✅ PRESERVE: Database-first approach
            // ✅ PRESERVE: UI state persistence
            // Implementation details in I.2
        }
        
        public OpeningFilter LoadFilterFromDatabase(string filterName, string category)
        {
            // ✅ PRESERVE: Database-first loading
            // Implementation details in I.2
        }
        
        public void SaveFilterToXmlFile(OpeningFilter filter, string filePath)
        {
            // ✅ PRESERVE: XML serialization logic
            // ✅ REUSE: Existing SaveFilterToXmlFile implementation
        }
        
        public OpeningFilter LoadFilterFromXmlFile(string filePath)
        {
            // ✅ PRESERVE: XML deserialization logic
            // ✅ REUSE: Existing LoadFilterFromXmlFile implementation
        }
        
        public bool IsFilterSaved(string filterName, string category = null)
        {
            // ✅ PRESERVE: Database-first check, XML fallback
            // Implementation details in I.2
        }
        
        public List<string> GetAllSavedFilterNames()
        {
            // ✅ PRESERVE: Load from database first, then XML
            // Implementation details in I.2
        }
    }
}
```

### I.2 Implementation Details

#### I.2.1 Database Operations
- **Migrate:** `RegisterFilterInDatabase()`, `UpdateFilterNameInDatabase()`, `DeleteFilterFromDatabase()`
- **Preserve:** Database-first approach, transaction safety
- **Reuse:** `FilterRepository` via adapter pattern

#### I.2.2 XML Operations
- **Migrate:** `SaveFilterToXmlFile()`, `LoadFilterFromXmlFile()`
- **Preserve:** Backward compatibility, XML serialization/deserialization
- **Reuse:** Existing XML serialization logic

#### I.2.3 Filter Validation
- **Migrate:** `IsFilterSaved()` logic
- **Preserve:** Database-first check, XML fallback
- **Reuse:** `ProjectPathService` for directory operations

---

## Team J: Filter Business Logic & UI Orchestration

### J.1 Deliverables

#### J.1.1 Interfaces

**File:** `Services/Interfaces/Refactor/IFilterCrudService.cs`
```csharp
namespace JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor
{
    /// <summary>
    /// Team J: Filter CRUD operations.
    /// SOLID: Single Responsibility - CRUD operations only.
    /// </summary>
    public interface IFilterCrudService
    {
        /// <summary>
        /// Creates a new filter from current UI state.
        /// </summary>
        OpeningFilter CreateFilter(string filterName);
        
        /// <summary>
        /// Copies an existing filter.
        /// </summary>
        OpeningFilter CopyFilter(OpeningFilter sourceFilter, string newName);
        
        /// <summary>
        /// Renames a filter.
        /// </summary>
        void RenameFilter(string oldName, string newName, string category);
        
        /// <summary>
        /// Deletes a filter.
        /// </summary>
        void DeleteFilter(string filterName, string category);
        
        /// <summary>
        /// Validates filter before save.
        /// </summary>
        (bool isValid, string errorMessage) ValidateFilter(OpeningFilter filter);
    }
}
```

**File:** `Services/Interfaces/Refactor/IFilterUiOrchestrator.cs`
```csharp
namespace JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor
{
    /// <summary>
    /// Team J: Filter UI orchestration (coordinates CRUD + Persistence).
    /// SOLID: Orchestrator pattern - coordinates focused services.
    /// </summary>
    public interface IFilterUiOrchestrator
    {
        /// <summary>
        /// Creates a new filter and saves it (database + XML).
        /// </summary>
        void CreateNewFilter(ListBox filterListBox);
        
        /// <summary>
        /// Copies a filter and saves it.
        /// </summary>
        void CopyFilter(ListBox filterListBox);
        
        /// <summary>
        /// Renames a filter and updates persistence.
        /// </summary>
        void RenameFilter(ListBox filterListBox);
        
        /// <summary>
        /// Deletes a filter from persistence and UI.
        /// </summary>
        void DeleteFilter(ListBox filterListBox);
        
        /// <summary>
        /// Saves filter to persistence (database + XML).
        /// </summary>
        void SaveFilter(ListBox filterListBox);
        
        /// <summary>
        /// Auto-saves filter (no user prompt).
        /// </summary>
        void SaveFilterAuto(ListBox filterListBox);
        
        /// <summary>
        /// Loads all saved filters into UI.
        /// </summary>
        void LoadAllSavedFilters(ListBox filterListBox);
        
        /// <summary>
        /// Seeds default filters into UI.
        /// </summary>
        void SeedDefaultFilters(ListBox filterListBox, IEnumerable<string> defaultNames);
    }
}
```

#### J.1.2 Implementations

**File:** `Services/FilterManagement/FilterCrudService.cs`
```csharp
namespace JSE_RevitAddin_MEP_OPENINGS.Services.FilterManagement
{
    /// <summary>
    /// Team J: SOLID-compliant filter CRUD service.
    /// Handles filter creation, copying, renaming, deletion, and validation.
    /// </summary>
    public class FilterCrudService : IFilterCrudService
    {
        private readonly ILogger _logger;
        
        public FilterCrudService(ILogger logger = null)
        {
            _logger = logger ?? LoggerAdapter.Default;
        }
        
        public OpeningFilter CreateFilter(string filterName)
        {
            // ✅ PRESERVE: CreateFilterFromCurrentUIState logic
            // ✅ PRESERVE: Category validation from UI state
            // ✅ REUSE: FilterUiStateProvider for UI state access
            // Implementation details in J.2
        }
        
        public OpeningFilter CopyFilter(OpeningFilter sourceFilter, string newName)
        {
            // ✅ PRESERVE: Deep copy logic
            // ✅ PRESERVE: Category handling
            // Implementation details in J.2
        }
        
        public void RenameFilter(string oldName, string newName, string category)
        {
            // ✅ PRESERVE: Rename logic
            // ✅ PRESERVE: Database update
            // Implementation details in J.2
        }
        
        public void DeleteFilter(string filterName, string category)
        {
            // ✅ PRESERVE: Delete logic
            // ✅ PRESERVE: Database deletion
            // Implementation details in J.2
        }
        
        public (bool isValid, string errorMessage) ValidateFilter(OpeningFilter filter)
        {
            // ✅ PRESERVE: Category validation
            // ✅ PRESERVE: Name validation
            // Implementation details in J.2
        }
    }
}
```

**File:** `Services/FilterManagement/FilterUiOrchestrator.cs`
```csharp
namespace JSE_RevitAddin_MEP_OPENINGS.Services.FilterManagement
{
    /// <summary>
    /// Team J: SOLID-compliant filter UI orchestrator.
    /// Coordinates CRUD operations with persistence operations.
    /// </summary>
    public class FilterUiOrchestrator : IFilterUiOrchestrator
    {
        private readonly IFilterCrudService _crudService;
        private readonly IFilterPersistenceService _persistenceService;
        private readonly IUserInteractionService _userInteraction;
        private readonly ILogger _logger;
        private readonly Action<string> _updateStatus;
        
        public FilterUiOrchestrator(
            IFilterCrudService crudService,
            IFilterPersistenceService persistenceService,
            IUserInteractionService userInteraction,
            Action<string> updateStatus,
            ILogger logger = null)
        {
            _crudService = crudService ?? throw new ArgumentNullException(nameof(crudService));
            _persistenceService = persistenceService ?? throw new ArgumentNullException(nameof(persistenceService));
            _userInteraction = userInteraction ?? throw new ArgumentNullException(nameof(userInteraction));
            _updateStatus = updateStatus ?? throw new ArgumentNullException(nameof(updateStatus));
            _logger = logger ?? LoggerAdapter.Default;
        }
        
        public void CreateNewFilter(ListBox filterListBox)
        {
            // ✅ ORCHESTRATION: Get name from user → Create filter → Validate → Save → Update UI
            // ✅ PRESERVE: Database-first registration
            // ✅ PRESERVE: UI state persistence
            // Implementation details in J.2
        }
        
        public void CopyFilter(ListBox filterListBox)
        {
            // ✅ ORCHESTRATION: Get source → Copy → Get name → Validate → Save → Update UI
            // Implementation details in J.2
        }
        
        public void RenameFilter(ListBox filterListBox)
        {
            // ✅ ORCHESTRATION: Get old/new names → Rename → Update persistence → Update UI
            // Implementation details in J.2
        }
        
        public void DeleteFilter(ListBox filterListBox)
        {
            // ✅ ORCHESTRATION: Get filter → Confirm → Delete from persistence → Update UI
            // Implementation details in J.2
        }
        
        public void SaveFilter(ListBox filterListBox)
        {
            // ✅ ORCHESTRATION: Get filter → Validate → Save to persistence → Update UI
            // Implementation details in J.2
        }
        
        public void SaveFilterAuto(ListBox filterListBox)
        {
            // ✅ ORCHESTRATION: Get filter → Save to persistence (no user prompt)
            // Implementation details in J.2
        }
        
        public void LoadAllSavedFilters(ListBox filterListBox)
        {
            // ✅ ORCHESTRATION: Load from persistence → Add to UI
            // ✅ PRESERVE: Database-first, XML fallback
            // Implementation details in J.2
        }
        
        public void SeedDefaultFilters(ListBox filterListBox, IEnumerable<string> defaultNames)
        {
            // ✅ ORCHESTRATION: Create default filters → Add to UI
            // Implementation details in J.2
        }
    }
}
```

**File:** `Services/Interfaces/Refactor/IUserInteractionService.cs`
```csharp
namespace JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor
{
    /// <summary>
    /// Team J: User interaction abstraction (for testability).
    /// SOLID: Interface Segregation - focused interface for user interactions.
    /// </summary>
    public interface IUserInteractionService
    {
        /// <summary>
        /// Gets filter name from user via input dialog.
        /// </summary>
        string GetFilterNameFromUser(string title, string prompt, string defaultValue);
        
        /// <summary>
        /// Shows error message to user.
        /// </summary>
        void ShowError(string message);
        
        /// <summary>
        /// Shows confirmation dialog to user.
        /// </summary>
        bool ShowConfirmation(string message, string title);
    }
}
```

**File:** `Services/FilterManagement/UserInteractionService.cs`
```csharp
namespace JSE_RevitAddin_MEP_OPENINGS.Services.FilterManagement
{
    /// <summary>
    /// Team J: User interaction service implementation.
    /// </summary>
    public class UserInteractionService : IUserInteractionService
    {
        public string GetFilterNameFromUser(string title, string prompt, string defaultValue)
        {
            // ✅ REUSE: Existing GetFilterNameFromUser implementation
            return Microsoft.VisualBasic.Interaction.InputBox(prompt, title, defaultValue);
        }
        
        public void ShowError(string message)
        {
            MessageBox.Show(message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        
        public bool ShowConfirmation(string message, string title)
        {
            return MessageBox.Show(message, title, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes;
        }
    }
}
```

### J.2 Implementation Details

#### J.2.1 Filter Creation
- **Migrate:** `CreateNewFilter()` → `FilterUiOrchestrator.CreateNewFilter()`
- **Preserve:** Category validation from UI state, database-first registration, UI state persistence
- **Reuse:** `FilterUiStateProvider`, `MepCategoryConstants.Normalize()`

#### J.2.2 Filter Copying
- **Migrate:** `CopyFilter()` → `FilterUiOrchestrator.CopyFilter()`
- **Preserve:** Deep copy logic, category handling
- **Reuse:** Existing copy logic

#### J.2.3 Filter Renaming
- **Migrate:** `RenameFilter()` → `FilterUiOrchestrator.RenameFilter()`
- **Preserve:** Database update, UI update
- **Reuse:** `FilterRepository.UpdateFilterName()`

#### J.2.4 Filter Deletion
- **Migrate:** `DeleteFilter()` → `FilterUiOrchestrator.DeleteFilter()`
- **Preserve:** Confirmation dialog, database deletion, UI update
- **Reuse:** `FilterRepository.DeleteFilter()`

#### J.2.5 Filter Saving
- **Migrate:** `SaveFilter()`, `SaveFilterAuto()` → `FilterUiOrchestrator.SaveFilter()`, `SaveFilterAuto()`
- **Preserve:** Database-first save, XML fallback, UI state persistence
- **Reuse:** `FilterPersistenceService`

#### J.2.6 Filter Loading
- **Migrate:** `LoadAllSavedFilters()` → `FilterUiOrchestrator.LoadAllSavedFilters()`
- **Preserve:** Database-first load, XML fallback, deduplication
- **Reuse:** `FilterPersistenceService`

---

## Integration & Wiring (You)

### Wiring Responsibilities

1. **Create Factory:** `FilterManagementServiceFactory` to wire all services
2. **Create Adapter:** `FilterManagementServiceAdapter` for backward compatibility
3. **Update Callers:** Replace direct `FilterManagementService` usage with new interfaces
4. **Migration Strategy:** Feature flag for gradual migration

### Factory Implementation

**File:** `Services/FilterManagement/FilterManagementServiceFactory.cs`
```csharp
namespace JSE_RevitAddin_MEP_OPENINGS.Services.FilterManagement
{
    /// <summary>
    /// Factory for creating fully-wired filter management services.
    /// </summary>
    public static class FilterManagementServiceFactory
    {
        public static IFilterUiOrchestrator CreateRefactored(
            Document document,
            Action<string> updateStatus,
            ILogger logger = null)
        {
            logger = logger ?? LoggerAdapter.Default;
            
            // ✅ WIRING: Create Team I services (Persistence)
            var filterRepository = new FilterRepositoryAdapter(document, logger);
            var persistenceService = new FilterPersistenceService(filterRepository, document, logger);
            
            // ✅ WIRING: Create Team J services (CRUD + Orchestration)
            var crudService = new FilterCrudService(logger);
            var userInteraction = new UserInteractionService();
            var orchestrator = new FilterUiOrchestrator(
                crudService,
                persistenceService,
                userInteraction,
                updateStatus,
                logger);
            
            return orchestrator;
        }
        
        public static FilterManagementService CreateLegacy(
            Document document,
            Action<string> log,
            Action<string> updateStatus)
        {
            return new FilterManagementService(document, log, updateStatus);
        }
    }
}
```

### Adapter Implementation

**File:** `Services/FilterManagement/FilterManagementServiceAdapter.cs`
```csharp
namespace JSE_RevitAddin_MEP_OPENINGS.Services.FilterManagement
{
    /// <summary>
    /// Adapter for backward compatibility during migration.
    /// Delegates to refactored services when flag is enabled.
    /// </summary>
    public class FilterManagementServiceAdapter : FilterManagementService
    {
        private readonly IFilterUiOrchestrator _orchestrator;
        private readonly bool _useRefactored;
        
        public FilterManagementServiceAdapter(
            Document document,
            Action<string> log,
            Action<string> updateStatus,
            bool useRefactored = false)
            : base(document, log, updateStatus)
        {
            _useRefactored = useRefactored;
            
            if (_useRefactored)
            {
                _orchestrator = FilterManagementServiceFactory.CreateRefactored(document, updateStatus);
            }
        }
        
        public override void CreateNewFilter(ListBox filterListBox)
        {
            if (_useRefactored && _orchestrator != null)
            {
                _orchestrator.CreateNewFilter(filterListBox);
            }
            else
            {
                base.CreateNewFilter(filterListBox);
            }
        }
        
        // ... other methods delegate similarly
    }
}
```

---

## Migration Strategy

### Phase 1: Parallel Development
- **Team I:** Implement persistence services (database + XML)
- **Team J:** Implement CRUD services and orchestrator
- **You:** Create factory and adapter

### Phase 2: Integration Testing
- Wire services together
- Test all CRUD operations
- Verify database-first approach
- Verify XML fallback

### Phase 3: Gradual Migration
- Add feature flag: `OptimizationFlags.UseRefactoredFilterManagementService`
- Update callers to use adapter
- Test in production with flag disabled
- Enable flag when stable

### Phase 4: Legacy Removal
- Remove legacy `FilterManagementService` after full migration
- Move to backup folder

---

## Testing Checklist

### Team I (Persistence)
- [ ] `SaveFilterToDatabase()` saves correctly
- [ ] `LoadFilterFromDatabase()` loads correctly
- [ ] `SaveFilterToXmlFile()` saves correctly
- [ ] `LoadFilterFromXmlFile()` loads correctly
- [ ] `IsFilterSaved()` checks database first, then XML
- [ ] `GetAllSavedFilterNames()` returns deduplicated list

### Team J (CRUD + Orchestration)
- [ ] `CreateNewFilter()` creates and saves filter
- [ ] `CopyFilter()` copies and saves filter
- [ ] `RenameFilter()` renames and updates persistence
- [ ] `DeleteFilter()` deletes from persistence and UI
- [ ] `SaveFilter()` saves to persistence
- [ ] `SaveFilterAuto()` auto-saves without prompt
- [ ] `LoadAllSavedFilters()` loads all filters
- [ ] `SeedDefaultFilters()` seeds default filters

### Integration
- [ ] Factory wires all services correctly
- [ ] Adapter delegates to refactored services when flag enabled
- [ ] All existing callers work with adapter
- [ ] No regression in functionality
- [ ] Performance maintained or improved

---

## File Structure

```
Services/
├── FilterManagement/
│   ├── FilterRepositoryAdapter.cs          (Team I)
│   ├── FilterPersistenceService.cs         (Team I)
│   ├── FilterCrudService.cs                (Team J)
│   ├── FilterUiOrchestrator.cs             (Team J)
│   ├── UserInteractionService.cs          (Team J)
│   ├── FilterManagementServiceFactory.cs   (You)
│   └── FilterManagementServiceAdapter.cs   (You)
├── Interfaces/
│   └── Refactor/
│       ├── IFilterRepository.cs            (Team I)
│       ├── IFilterPersistenceService.cs    (Team I)
│       ├── IFilterCrudService.cs          (Team J)
│       ├── IFilterUiOrchestrator.cs       (Team J)
│       └── IUserInteractionService.cs     (Team J)
└── FilterManagementService.cs              (Legacy - move to backup after migration)
```

---

## Success Criteria

1. ✅ **SOLID Compliance:** All services follow SOLID principles
2. ✅ **Testability:** All services are fully testable via dependency injection
3. ✅ **No Regression:** All existing functionality preserved
4. ✅ **Performance:** No performance degradation
5. ✅ **Maintainability:** Code is easier to understand and modify
6. ✅ **Extensibility:** Easy to add new filter operations

---

## Notes

- **Preserve All Optimizations:** Database-first approach, batch operations, etc.
- **Reuse Existing Code:** FilterRepository, ProjectPathService, etc.
- **Gradual Migration:** Use feature flag for safe rollout
- **Backward Compatibility:** Adapter pattern ensures no breaking changes

---

**End of Plan**


# Refresh Service Feature Flag Implementation

## Overview

A feature flag system has been implemented to allow switching between the **legacy RefreshService** and the **refactored RefreshServiceRefactored** implementations. This enables gradual migration and testing of the new optimized refresh service.

## Architecture

### Components

1. **`RefreshServiceFactory`** (`Services/RefreshServiceFactory.cs`)
   - Factory class that creates the appropriate refresh service based on feature flag
   - Implements `IRefreshService` interface for both implementations
   - Provides wrapper classes for compatibility

2. **`IRefreshService` Interface**
   - Common interface for both legacy and refactored services
   - Methods: `SetUIReferences()`, `ExecuteRefresh()`, `LoadExistingClashZoneData()`

3. **Feature Flag** (`SettingsModel.UseRefactoredRefreshService`)
   - Boolean flag in `SettingsModel` class
   - Default: `false` (uses legacy service)
   - Can be toggled via application settings

## Usage

### Current Implementation

The `EmergencyMainDialog.cs` now uses the factory:

```csharp
// ✅ FEATURE FLAG: Use factory to create appropriate refresh service
var refreshService = Services.RefreshServiceFactory.Create(
    document, 
    _uiDocument, 
    _appProfileService);
refreshService.SetUIReferences(_statusLabel, _progressBar, _refreshButton);
refreshService.LoadExistingClashZoneData();
refreshService.ExecuteRefresh(selectedFilterItems, selectedMepCategories, 
    selectedReferenceFiles, selectedHostFiles, clearanceSettings);
```

### Enabling Refactored Service

To enable the refactored refresh service:

1. **Via Code** (for testing):
   ```csharp
   var settings = _appProfileService.GetCurrentSettings();
   settings.UseRefactoredRefreshService = true;
   _appProfileService.SaveCurrentSettings(settings);
   ```

2. **Via Settings UI** (when implemented):
   - Add a checkbox/toggle in the settings dialog
   - Bind to `SettingsModel.UseRefactoredRefreshService`

3. **Via Configuration File**:
   - Edit the user's settings XML file
   - Set `UseRefactoredRefreshService` to `true`

## Service Differences

### Legacy RefreshService
- Original implementation (~5,700 LOC)
- Monolithic structure
- All logic in one class
- Requires `LoadExistingClashZoneData()` call before `ExecuteRefresh()`

### Refactored RefreshServiceRefactored
- New optimized implementation (~200 LOC orchestrator)
- Modular architecture with helper services:
  - `RefreshContext` - Shared state management
  - `XmlCacheManager` - XML loading/caching
  - `ValidationService` - Hash-based validation
  - `IntersectionProcessor` - Lazy intersection detection
  - `PerformanceMonitor` - Performance tracking
- Loads existing clash zones internally
- `LoadExistingClashZoneData()` is a no-op (for compatibility)

## Benefits

1. **Gradual Migration**: Test refactored service without affecting production
2. **Easy Rollback**: Switch back to legacy service if issues occur
3. **A/B Testing**: Compare performance between implementations
4. **Zero Breaking Changes**: Both services implement same interface

## Logging

The factory logs which service is being used:

```
[RefreshServiceFactory] Using LEGACY RefreshService
// or
[RefreshServiceFactory] Using REFACTORED RefreshService
```

## Future Work

1. **Settings UI**: Add toggle in application settings dialog
2. **Performance Comparison**: Add metrics to compare both implementations
3. **Migration Path**: Once refactored service is stable, make it default
4. **Deprecation**: Remove legacy service after full migration

## Notes

- Both services maintain the same external interface
- The refactored service is still in `refresh refactor` folder (can be moved later)
- Feature flag defaults to `false` for safety (legacy service)
- All helper classes for refactored service are in `Services.Refresh` namespace


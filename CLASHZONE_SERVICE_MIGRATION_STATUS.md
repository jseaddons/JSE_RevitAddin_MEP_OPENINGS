# ClashZoneService Migration Status

## Naming Differences

### Legacy ClashZoneService
- **Namespace:** `JSE_RevitAddin_MEP_OPENINGS.Services`
- **Full Name:** `Services.ClashZoneService`
- **Location:** `Services/ClashZoneService.cs`
- **Size:** ~4,600 lines (monolithic)
- **Status:** ⚠️ Still in use for detection operations

### Refactored ClashZoneService
- **Namespace:** `JSE_RevitAddin_MEP_OPENINGS.Services.ClashZoneManagement`
- **Full Name:** `Services.ClashZoneManagement.ClashZoneService`
- **Location:** `Services/ClashZoneManagement/ClashZoneService.cs`
- **Size:** ~100 lines (focused, SOLID-compliant)
- **Status:** ✅ In use for cleanup and filtering operations

## Current Migration Status

### ✅ Migrated Operations
1. **Cleanup** (`CleanupInvalidClashZones`) - Uses refactored service
2. **Filtering** (`FilterClashZonesByCurrentSelection`) - Uses refactored service

### ⚠️ Still Using Legacy
1. **Detection** (`DetectNewClashZones`) - Still uses legacy service in `intersection_processor.cs`

## Migration Plan

### Step 1: Add Detection Method to Interface ✅
- Added `DetectNewClashZones` to `IClashZoneService` interface

### Step 2: Implement Detection in Refactored Service
- Option A: Delegate to legacy service (temporary bridge)
- Option B: Create new `ClashZoneDetectionService` (proper SOLID approach)

### Step 3: Update intersection_processor.cs
- Use refactored service from context instead of creating new legacy service

### Step 4: Remove Legacy Service
- Once fully migrated, move legacy service to backup folder

## Code Locations

### Legacy Service Usage
- `refresh refactor/intersection_processor.cs:254` - Direct instantiation
- `Services/ClashZoneService.cs` - Legacy implementation

### Refactored Service Usage
- `refresh refactor/refresh_service_refactored.cs:180` - Created via MigrationStrategy
- `Services/ClashZoneManagement/ClashZoneService.cs` - Refactored implementation

## Feature Flag
- `OptimizationFlags.UseRefactoredClashZoneFlagServices = true` ✅


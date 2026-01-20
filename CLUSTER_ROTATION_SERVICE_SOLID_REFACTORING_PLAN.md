# ClusterRotationService SOLID Refactoring Plan

## Current Status: ~60% SOLID Compliant

### ✅ Completed
1. **DIP Compliance**: Created interfaces for static dependencies
   - `IRcsTransformer` - abstraction for WallRcsTransformer
   - `RcsTransformerAdapter` - adapter for static class
   - `IRevitUnitConversionService` - already exists
   - `ILogger` - already exists

2. **SRP Compliance**: Extracted focused services
   - `ClusterRotationCacheService` - caching logic only
   - `ClusterPlacementPointService` - placement point calculation only

### 🔄 In Progress / Pending

#### 1. Extract ClusterRotationAngleService (SRP)
**Location**: `Services/Clustering/Rotation/Services/ClusterRotationAngleService.cs`

**Responsibilities**:
- Determine rotation angle for clusters
- Handle element-type-specific logic (pipes, ducts, walls, floors)
- Use strategy pattern for element types (OCP compliance)

**Dependencies to Inject**:
- `IClusterRotationCacheService` - for ClashZone lookups
- `ILogger` - for logging (optional, can use static for now)

**Methods to Extract**:
- `DetermineRotationAngle()` - lines 62-340 in current file

**Strategy Pattern Implementation**:
```csharp
public interface IElementRotationStrategy
{
    double CalculateRotationAngle(ClashZone clashZone, List<dynamic> cluster);
    bool AppliesTo(ClashZone clashZone);
}

public class PipeRotationStrategy : IElementRotationStrategy { }
public class DuctRotationStrategy : IElementRotationStrategy { }
public class WallRotationStrategy : IElementRotationStrategy { }
public class FloorRotationStrategy : IElementRotationStrategy { }
```

#### 2. Extract ClusterBoundingBoxService (SRP)
**Location**: `Services/Clustering/Rotation/Services/ClusterBoundingBoxService.cs`

**Responsibilities**:
- Calculate rotated bounding boxes
- Handle corner-based calculations
- Handle RCS transformations for walls/framing
- Handle WCS calculations for floors

**Dependencies to Inject**:
- `IClusterRotationCacheService` - for ClashZone lookups
- `IRcsTransformer` - for RCS transformations
- `IRevitUnitConversionService` - for unit conversions
- `IClusterPlacementPointService` - for placement point calculation

**Methods to Extract**:
- `CalculateRotatedBoundingBox()` - lines 346-865 in current file
- Corner-based calculation logic
- RCS transformation logic
- WCS calculation logic

#### 3. Refactor ClusterRotationService (Orchestrator)
**Location**: `Services/Clustering/Rotation/ClusterRotationService.cs`

**New Responsibilities** (SRP Compliant):
- Orchestrate rotation angle determination
- Orchestrate bounding box calculation
- Manage rotation data storage/retrieval
- Coordinate between services

**Dependencies to Inject**:
- `IClusterRotationAngleService`
- `IClusterBoundingBoxService`
- `IClusterPlacementPointService`
- `IClusterRotationCacheService`
- `IRcsTransformer`
- `IRevitUnitConversionService`
- `ILogger` (optional)

**Methods to Keep**:
- `GetRotationData()` - data retrieval
- `StoreRotationData()` - data storage
- `ClearRotationData()` - data clearing

**Methods to Delegate**:
- `DetermineRotationAngle()` → `_rotationAngleService.DetermineRotationAngle()`
- `CalculateRotatedBoundingBox()` → `_boundingBoxService.CalculateRotatedBoundingBox()`

#### 4. Update Interface (ISP Compliance)
**Location**: `Services/Clustering/Rotation/IClusterRotationService.cs`

**Consider Splitting**:
- `IRotationAngleService` - angle determination only
- `IRotatedBoundingBoxService` - bounding box calculation only
- `IRotationDataStorage` - data storage/retrieval only

**OR Keep Single Interface** (if backward compatibility is critical):
- Keep current interface but implement via composition

## Implementation Steps

### Phase 1: Complete Service Extraction ✅ (Partially Done)
- [x] Create interfaces
- [x] Extract CacheService
- [x] Extract PlacementPointService
- [ ] Extract RotationAngleService
- [ ] Extract BoundingBoxService

### Phase 2: Refactor Main Service
- [ ] Update ClusterRotationService constructor to inject dependencies
- [ ] Replace method implementations with service delegations
- [ ] Remove extracted code from main service
- [ ] Update cache management to use CacheService

### Phase 3: Strategy Pattern (OCP)
- [ ] Create IElementRotationStrategy interface
- [ ] Implement strategies for each element type
- [ ] Update RotationAngleService to use strategies

### Phase 4: Testing & Validation
- [ ] Ensure backward compatibility
- [ ] Test all rotation scenarios
- [ ] Verify cache behavior
- [ ] Performance testing

## SOLID Compliance After Refactoring

### ✅ Single Responsibility Principle (SRP)
- **ClusterRotationService**: Orchestration only
- **ClusterRotationAngleService**: Rotation angle determination only
- **ClusterBoundingBoxService**: Bounding box calculation only
- **ClusterPlacementPointService**: Placement point calculation only
- **ClusterRotationCacheService**: Caching only

### ✅ Open/Closed Principle (OCP)
- Strategy pattern for element types
- New element types can be added without modifying existing code

### ✅ Liskov Substitution Principle (LSP)
- All services implement interfaces
- Can be substituted with alternative implementations

### ✅ Interface Segregation Principle (ISP)
- Interfaces are focused and minimal
- Clients only depend on methods they use

### ✅ Dependency Inversion Principle (DIP)
- All dependencies are injected via interfaces
- No static class dependencies
- Testable and mockable

## Estimated SOLID Score After Refactoring: **95%+**

## Notes
- Maintain backward compatibility with existing `IClusterRotationService` interface
- Use adapter pattern for static classes (WallRcsTransformer, etc.)
- Consider gradual migration if needed
- All services should be unit testable


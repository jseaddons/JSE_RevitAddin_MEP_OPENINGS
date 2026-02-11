# Sleeve Placement Methodology - Refactored Architecture

**Document Version:** 2.1 (December 2025)  
**Status:** Current Architecture - Database-Driven, Refactored Services Only  
**Last Updated:** Added damper intersection point calculation using insertion point (LocationPoint) for proper centering

---

## Table of Contents

1. [Overview](#1-overview)
2. [Architecture Principles](#2-architecture-principles)
3. [Refresh Service Architecture](#3-refresh-service-architecture)
4. [Three Paths System](#4-three-paths-system)
5. [Database-Driven Persistence](#5-database-driven-persistence)
6. [Service Components](#6-service-components)
7. [Sleeve Placement Flow](#7-sleeve-placement-flow)
8. [Clustering Flow](#8-clustering-flow)
9. [Combine Sleeve Feature](#9-combine-sleeve-feature)
10. [Comprehensive Flag Management System](#10-comprehensive-flag-management-system)
    - 10.6 [Cross-Filter Consistency Principles](#106-cross-filter-consistency-principles)

---

## 9. Combine Sleeve Feature

**Status:** Planning Phase  
**Related Document:** `COMBINE_SLEEVE_PLAN.md`

### 9.1 Overview

The **Combine Sleeve** feature allows combining cluster sleeves with individual sleeves from different categories, with strict wall group restrictions (Wall X, Wall Y, Floor).

**Key Features:**
- ✅ Auto Mode: Select category, system automatically combines compatible sleeves
- ✅ Manual Mode: User manually selects 2 sleeves to combine
- ✅ Wall Group Restrictions: Wall X only with Wall X, Wall Y only with Wall Y, Floor only with Floor
- ✅ Straight Axis Only: Phase 1 focuses on straight axis-aligned sleeves (0°, 90°, 180°, 270°)

**For detailed implementation plan, see:** `COMBINE_SLEEVE_PLAN.md`

---

## 1. Overview

This document describes the **current refactored architecture** for sleeve placement in MEP openings. All legacy services have been removed, and the system now uses:

- ✅ **Database-First Architecture** - SQLite database as primary data store
- ✅ **Refactored Refresh Service** - Clean, maintainable refresh orchestration
- ✅ **Path Strategy Pattern** - Three distinct execution paths based on file combo state
- ✅ **Repository Pattern** - Database operations abstracted through repositories
- ✅ **XML as Fallback Only** - XML used only when database has no data

### 1.1 Key Services

| Service | Location | Purpose |
|--------|----------|---------|
| **RefreshServiceRefactored** | `refresh refactor/refresh_service_refactored.cs` | Main refresh orchestrator |
| **RefreshPathDeterminer** | `refresh refactor/refresh_path_strategy.cs` | Path selection logic (PATH 1/2/3) |
| **IntersectionProcessor** | `refresh refactor/intersection_processor.cs` | MEP vs Structural intersection detection |
| **ValidationService** | `refresh refactor/validation_service.cs` | 3-point validation for existing zones |
| **ClashZoneRepository** | `Data/Repositories/ClashZoneRepository.cs` | Database operations for clash zones |
| **FilterRepository** | `Data/Repositories/FilterRepository.cs` | Database operations for filters |
| **UniversalSleevePlacerService** | `Services/UniversalSleevePlacerService.cs` | Individual sleeve placement (Current) |
| **NewSleevePlacerService** | `Services/NewSleevePlacerService.cs` | Individual sleeve placement (Future - OOP architecture, SOLID principles) |
| **UniversalClusterService** | `Services/UniversalClusterService.cs` | Cluster sleeve formation and placement (Legacy - 8,765 lines) |
| **RefactoredClusterService** | `Services/Clustering/RefactoredClusterService.cs` | ✅ Modern cluster orchestrator (598 lines) - Phase 1-10 services |
| **ClusterServiceFactory** | `Services/Clustering/ClusterServiceFactory.cs` | Factory for creating fully-wired clustering services |

---

## 2. Architecture Principles

### 2.1 Database-First

**Primary Data Store:** SQLite database (`SleeveDbContext`)

**Database Tables:**
- `Filters` - Filter metadata and UI state (SelectedHostCategories, OpeningSettings)
- `FileCombos` - File combination tracking with `IsFilterComboNew` flag
- `ClashZones` - All clash zone data (PRIMARY source)
- `SleeveSnapshots` - Sleeve placement snapshots
- `ClusterSleeves` - Cluster sleeve calculation results (enables PATH 1 replay without recalculation)

**XML Role:** Fallback only (if database has no data)

### 2.2 Service Separation

- **Refresh Service** - Orchestrates detection and validation
- **Placement Service** - Handles individual sleeve placement
- **Cluster Service** - Handles cluster sleeve formation
- **Repository Services** - Database operations only

### 2.3 Path Strategy Pattern

Three distinct execution paths determined by:
1. `IsFilterComboNew` flag in `FileCombos` table
2. `enableThreePointValidation` setting (Adopt to Modified Document)

---

## 3. Refresh Service Architecture

### 3.1 RefreshServiceRefactored

**Location:** `refresh refactor/refresh_service_refactored.cs`

**Purpose:** Main orchestrator for refresh operations

**Key Methods:**
- `ExecuteRefresh()` - Main entry point
- `ExecuteRefreshInternal()` - Internal orchestration logic

**Execution Flow:**
```
1. Validate UI selections
2. Load XML cache (once, eliminates redundant loads)
3. Load existing clash zones from database (PRIMARY) or XML (fallback)
4. Determine path strategy (PATH 1/2/3)
5. Sync flags (if required by path)
6. Validate clash zones (if required by path)
7. Process zones (zone splitting for PATH 3)
8. Run intersection detection (if required by path)
9. Save clash zones to database
10. Reset flags for deleted sleeves (if required by path)
```

### 3.2 RefreshPathDeterminer

**Location:** `refresh refactor/refresh_path_strategy.cs`

**Purpose:** Determines which execution path to use

**Key Methods:**
- `DeterminePath()` - Selects PATH 1, 2, or 3 based on context
- `DeterminePlacementPath()` - Determines placement path (Replay vs Sizing)
- `UpdateIsFilterComboNewFlagBasedOnFileCombos()` - Updates flag based on database state

**Path Selection Logic:**
```csharp
IF enableThreePointValidation = false
  → PATH 1 (Replay Mode)

ELSE IF enableThreePointValidation = true
  IF filter has existing clash zones
    → PATH 3 (Full Detection with Validation)
  ELSE
    → PATH 2 (Fresh Placement Mode)
```

### 3.3 Path Strategy Interface

**Interface:** `IRefreshPathStrategy`

**Implementations:**
- `Path1Strategy` - Replay mode
- `Path2Strategy` - Fresh placement mode
- `Path3Strategy` - Full detection with validation

**Strategy Properties:**
- `AllowStructuralUpdates` - Whether to allow structural geometry updates
- `EnableThreePointValidation` - Whether to validate existing zones
- `ShouldResetFlags` - Whether to reset flags for deleted sleeves
- `ShouldSyncFlags` - Whether to sync flags from Global XML
- `ShouldCheckGuids` - Whether to check GUIDs

---

## 4. Three Paths System

### 4.1 PATH 1: REPLAY MODE

**Trigger:**
- `enableThreePointValidation = false` (Adopt to Modified Document = OFF)
- OR `IsFilterComboNew = 0` (all file combos already processed)

**Characteristics:**
- ✅ **No intersection detection** - Uses existing zones from database
- ✅ **No validation** - Assumes all zones are valid
- ✅ **Flag reset only** - Resets flags for deleted sleeves
- ✅ **No structural updates** - Only flag/ID changes
- ✅ **Fastest path** - Minimal processing
- ✅ **Conditional clustering** - Checks `IsClusterResolved` flags, skips if all resolved

**Database Operations:**
- **READ**: Load clash zones from `ClashZones` table
- **UPDATE**: Update flags in `ClashZones` table for deleted sleeves
- **NO INSERT**: No new clash zones created

**Clustering:**
- **Check ClusterSleeves table**: Query if cluster data exists for ComboId
- **If cluster data exists**: Load and place using stored data (skip calculation)
- **If no cluster data**: Run clustering calculation → Save to ClusterSleeves table

**Use Case:** Replay existing zones, use pre-calculated cluster data if available, otherwise calculate

### 4.2 PATH 2: FRESH PLACEMENT MODE

**Trigger:**
- `IsFilterComboNew = 1` (new file combos in database)
- AND `enableThreePointValidation = false` (Adopt to Modified Document = OFF)
- **Note:** Adopt setting is IRRELEVANT for PATH 2 - does NOT affect behavior

**Characteristics:**
- ✅ **Run intersection detection** - Find all MEP vs Structural intersections
- ✅ **No validation** - Assumes all zones are valid (fresh placement)
- ✅ **Save to database** - All clash zones saved to `ClashZones` table
- ✅ **Structural updates allowed** - Can update intersection geometry
- ✅ **Fast path** - No validation overhead
- ✅ **Always cluster** - Fresh detection always requires clustering

**Database Operations:**
- **INSERT**: Create new FileCombo with `IsFilterComboNew=1`
- **INSERT**: Save clash zones to `ClashZones` table (all zones are new - fresh detection)
- **INSERT**: Save sleeve snapshots to `SleeveSnapshots` table (all snapshots are new)
- **NO RESET**: Flag stays = 1 until after cluster sleeve placement

**Clustering:**
- **Always run calculation** - Fresh detection means new sleeves always need clustering
- **Save to ClusterSleeves** - Store calculation results for future PATH 1 replay

**Use Case:** First time adding filter, fresh detection and placement

### 4.3 PATH 3: FULL DETECTION WITH VALIDATION

**Trigger:**
- `enableThreePointValidation = true` (Adopt to Modified Document = ON)
- **Note:** Adopt setting MUST be ON to trigger PATH 3

**Characteristics:**
- ✅ **Run intersection detection** - Find all MEP vs Structural intersections
- ✅ **3-Point Validation ENABLED** - Validate existing zones:
  - MEP Element exists
  - Structural Element exists
  - Elements still intersect
- ✅ **Zone splitting** - Process three categories:
  - **Validated zones** → PATH 1 logic (check sleeve presence, place if missing)
  - **Invalidated zones** → **MERGE REQUIRED** (update intersection points)
  - **New zones** → PATH 2 logic (save to database)
- ✅ **Only invalidated zones merge** - Validated zones use PATH 1, new zones use PATH 2
- ✅ **Most thorough path** - Validates all existing zones

**Database Operations:**
- **READ**: Load existing clash zones from `ClashZones` table for validation
- **INSERT**: Create new FileCombo with `IsFilterComboNew=1`
- **INSERT/UPDATE**: Save validated/invalidated/new clash zones to `ClashZones` table
- **UPDATE**: Update existing zones if intersection points changed
- **INSERT/UPDATE**: Save sleeve snapshots to `SleeveSnapshots` table
- **NO RESET**: Flag stays = 1 until after cluster sleeve placement
- **Invalid Zone Removal:** If a zone is found to be invalid (elements no longer intersect or are deleted), it is **REMOVED** from the database.

**Clustering (Zone-Type Based):**
- **Validated zones**: Check `ClusterSleeves` table (same as PATH 1)
  - If cluster data exists: Load and place using stored data
  - If no cluster data: Run clustering calculation → Save to ClusterSleeves
- **Invalidated zones**: Always run clustering calculation (geometry changed, needs re-clustering) → Save to ClusterSleeves
- **New zones**: Always run clustering calculation (fresh detection, new sleeves) → Save to ClusterSleeves

**Use Case:** Re-validate existing zones, handle moved/deleted elements, merge invalidated zones

### 4.4 Path Selection Summary

| Path | Trigger | Detection | Validation | Merge | Database Save | Clustering | Use Case |
|------|---------|-----------|------------|-------|--------------|------------|----------|
| **PATH 1** | Adopt OFF OR IsFilterComboNew=0 | ❌ Skip | ❌ Skip | ❌ No | ❌ No | Conditional (check flags) | Replay existing zones |
| **PATH 2** | IsFilterComboNew=1 AND Adopt OFF | ✅ Yes | ❌ Skip | ❌ No | ✅ Yes | Always run | Fresh placement |
| **PATH 3** | Adopt ON | ✅ Yes | ✅ Yes | ✅ Yes (invalidated only) | ✅ Yes | Zone-type based | Full detection with validation |

**Clustering Details:**
- **PATH 1**: Check `ClusterSleeves` table → Use stored cluster data if exists (skip calculation), else calculate
- **PATH 2**: Always run clustering calculation → Save results to `ClusterSleeves` table
- **PATH 3**: 
  - Validated zones → Check `ClusterSleeves` table (same as PATH 1)
  - Invalidated zones → Always run calculation → Save to `ClusterSleeves`
  - New zones → Always run calculation → Save to `ClusterSleeves`

**Important:** `IsClusterResolved` flag is NOT used to skip clustering. It's only used to verify if cluster sleeve element still exists in Revit (for flag reset if deleted).

---

## 5. Database-Driven Persistence

### 5.1 Database Schema

**Filters Table:**
```sql
CREATE TABLE Filters (
    FilterId INTEGER PRIMARY KEY,
    FilterName TEXT NOT NULL,
    Category TEXT NOT NULL,
    SelectedHostCategories TEXT,  -- JSON array
    OpeningSettings TEXT,            -- JSON object
    CreatedAt DATETIME,
    UpdatedAt DATETIME
)

CREATE TABLE ClusterSleeves (
    ClusterSleeveId INTEGER PRIMARY KEY,
    ClusterInstanceId INTEGER NOT NULL UNIQUE,  -- Revit Element ID
    ComboId INTEGER NOT NULL,
    FilterId INTEGER NOT NULL,
    Category TEXT NOT NULL,
    -- Cluster bounding box (calculated during clustering)
    BoundingBoxMinX REAL NOT NULL,
    BoundingBoxMinY REAL NOT NULL,
    BoundingBoxMinZ REAL NOT NULL,
    BoundingBoxMaxX REAL NOT NULL,
    BoundingBoxMaxY REAL NOT NULL,
    BoundingBoxMaxZ REAL NOT NULL,
    -- Cluster dimensions (width, height, depth)
    ClusterWidth REAL NOT NULL,
    ClusterHeight REAL NOT NULL,
    ClusterDepth REAL NOT NULL,
    -- Cluster rotation (if rotated bounding box was used)
    RotationAngleDeg REAL DEFAULT 0.0,
    IsRotated INTEGER NOT NULL DEFAULT 0,
    -- Cluster placement point (center of cluster)
    PlacementX REAL NOT NULL,
    PlacementY REAL NOT NULL,
    PlacementZ REAL NOT NULL,
    -- Host type and orientation (for grouping)
    HostType TEXT,
    HostOrientation TEXT,
    -- List of ClashZoneIds that are part of this cluster (JSON array of GUIDs)
    ClashZoneIdsJson TEXT NOT NULL,
    -- Metadata
    CreatedAt DATETIME NOT NULL,
    UpdatedAt DATETIME NOT NULL,
    FOREIGN KEY(ComboId) REFERENCES FileCombos(ComboId) ON DELETE CASCADE,
    FOREIGN KEY(FilterId) REFERENCES Filters(FilterId) ON DELETE CASCADE
)
```

**FileCombos Table:**
```sql
CREATE TABLE FileCombos (
    ComboId INTEGER PRIMARY KEY,
    FilterId INTEGER NOT NULL,
    LinkedFileKey TEXT NOT NULL,
    HostFileKey TEXT NOT NULL,
    IsFilterComboNew INTEGER NOT NULL DEFAULT 1,  -- 0 = processed, 1 = new
    ProcessedAt DATETIME,
    CreatedAt DATETIME,
    UpdatedAt DATETIME,
    FOREIGN KEY (FilterId) REFERENCES Filters(FilterId)
)
```

**ClashZones Table:**
```sql
CREATE TABLE ClashZones (
    ClashZoneId INTEGER PRIMARY KEY,
    ClashZoneGuid TEXT UNIQUE NOT NULL,
    FilterId INTEGER NOT NULL,
    ComboId INTEGER NOT NULL,
    -- ... all clash zone properties ...
    IsResolvedFlag INTEGER DEFAULT 0,
    IsClusterResolvedFlag INTEGER DEFAULT 0,
    SleeveInstanceId INTEGER DEFAULT -1,
    ClusterSleeveInstanceId INTEGER DEFAULT -1,
    FOREIGN KEY (FilterId) REFERENCES Filters(FilterId),
    FOREIGN KEY (ComboId) REFERENCES FileCombos(ComboId)
)
```

### 5.2 Repository Pattern

**ClashZoneRepository:**
- `InsertOrUpdateClashZones()` - Save clash zones to database
- `GetClashZonesByFilter()` - Load clash zones from database
- `UpdateSleeveBoundingBoxes()` - Update bounding box coordinates
- `UpdateFlags()` - Update flags in database

**FilterRepository:**
- `GetFilterId()` - Get filter ID from database
- `EnsureFilter()` - Create filter if it doesn't exist
- `SaveFilterUIState()` - Save UI state (SelectedHostCategories, OpeningSettings)
- `LoadFilterUIState()` - Load UI state from database

**ClusterSleeveRepository:**
- `SaveClusterSleeve()` - Save cluster calculation results to ClusterSleeves table
- `LoadClusterSleevesForCombo()` - Load cluster data for a ComboId (used by PATH 1)
- `HasClusterDataForCombo()` - Check if cluster data exists (used by PATH 1 to skip recalculation)
- `DeleteClusterSleeve()` - Delete cluster data when cluster sleeve is deleted

### 5.3 Transaction Management

All database operations are wrapped in SQLite transactions:
- **Atomic commits** - Ensures data consistency
- **Rollback on error** - Prevents partial data

---

## 6. Service Components

### 6.0 OOP Damper and Insulation Architecture (SOLID Principles)

**✅ NEW: Object-Oriented Refactoring for Damper Detection and Insulation Awareness**

The system now uses a fully OOP architecture for damper connector detection and insulation-aware sizing, following SOLID principles.

#### 6.0.1 Damper Detection Services

**Location:** `Services/DamperDetection/`

**Services:**
- **`IDamperTypeDetector`** - Interface for detecting damper type from family/type name
- **`DamperTypeDetector`** - Implementation that classifies dampers (MSFD, MSD, MD, Motorized, Standard)
- **`IDamperConnectorDetector`** - Interface for detecting MEP connector presence and direction
- **`DamperConnectorDetector`** - Implementation that detects connector side using world/local coordinates
- **`DamperConnectorService`** - Orchestrator that combines type detection and connector detection

**Key Methods:**
- `DetectDamperType(familyTypeName)` - Returns damper type string (MSFD, MSD, MD, Motorized, Standard)
- `RequiresMepSideClearance(damperType)` - Determines if damper requires asymmetric clearance
- `IsStandardDamper(damperType)` - Checks if damper uses symmetric clearance
- `HasMepConnector(damper)` - Checks if damper has MEP connector
- `DetectConnectorSide(damper, useWorldCoordinates, out connector)` - Detects connector direction (Left, Right, Top, Bottom)

**Usage in ClashZoneService:**
```csharp
var damperConnectorService = new DamperConnectorService();
var connectorInfo = damperConnectorService.DetectConnectorInfo(mepElement, mepCategory);

// Populate ClashZone properties:
clashZone.IsMSFDDamper = connectorInfo.HasMepConnector && !string.IsNullOrEmpty(connectorInfo.ConnectorSide);
clashZone.IsStandardDamper = connectorInfo.IsStandardDamper;
clashZone.DamperConnectorSide = connectorInfo.ConnectorSide;
clashZone.MepElementSizeData.DamperType = connectorInfo.DamperType;
```

**SOLID Principles Applied:**
- **Single Responsibility:** Each service has one clear purpose (type detection vs connector detection)
- **Open/Closed:** Can extend for new damper types without modifying existing code
- **Liskov Substitution:** Interfaces allow different implementations
- **Interface Segregation:** Focused interfaces (IDamperTypeDetector, IDamperConnectorDetector)
- **Dependency Inversion:** High-level code depends on abstractions (interfaces)

#### 6.0.1.1 Damper Intersection Point Calculation

**Location:** `Services/MepIntersectionService.cs` - `FindDamperIntersectionsInternal()`

**Purpose:** Calculate the intersection point for dampers at the center of the damper body (excluding connectors), projected onto the wall plane.

**Critical Fix:** For dampers, the intersection point must be at the **center of the damper's width and height** (geometric center of damper body), not the intersection bounding box center.

**Problem:**
- Using intersection bounding box center can be off-center if the damper is not perfectly centered on the wall
- Using damper bounding box center may include connector geometry, shifting the center away from the damper body center
- This causes incorrect offset calculations and lopsided sleeve placement

**Solution:**
1. **Use Damper's Insertion Point** - Use `FamilyInstance.Location` (LocationPoint) or `GetTransform().Origin`
   - This is the geometric center of the damper body, excluding connectors
   - Connectors are separate entities and do not affect the insertion point
2. **Project onto Wall Plane** - Project the damper's insertion point onto the wall face along the wall normal
3. **Use as Intersection Point** - This projected point becomes the intersection point stored in ClashZone

**Implementation:**
```csharp
// For dampers, use damper's insertion point (LocationPoint) as center
XYZ damperCenter;
if (damperElement is FamilyInstance familyInstance)
{
    // Use insertion point - this is the geometric center of the damper body, excluding connectors
    var locationPoint = familyInstance.Location as LocationPoint;
    if (locationPoint != null)
    {
        damperCenter = locationPoint.Point;
        // Transform if damper is in a linked document
        if (damperLinkTransform != null)
        {
            damperCenter = damperLinkTransform.OfPoint(damperCenter);
        }
    }
    else
    {
        // Fallback to transform origin (geometric center)
        var transform = familyInstance.GetTransform();
        damperCenter = transform.Origin;
        if (damperLinkTransform != null)
        {
            damperCenter = damperLinkTransform.OfPoint(damperCenter);
        }
    }
}

// Project damper center onto wall plane
if (structuralElement is Wall wall)
{
    var wallNormal = wall.Orientation;
    var wallLocation = wall.Location as LocationCurve;
    XYZ wallFaceOrigin = wallLocation?.Curve?.GetEndPoint(0) ?? damperCenter;
    
    // Project damper center onto wall plane
    double distance = (damperCenter - wallFaceOrigin).DotProduct(wallNormal);
    intersectionPoint = damperCenter - wallNormal.Multiply(distance);
}
```

**Why This Matters:**
- **Correct Centering:** Intersection point is at the true center of damper body (width × height)
- **Proper Offset Calculation:** When offset is applied (25mm toward connector side), it starts from the correct center
- **Accurate Clearance Distribution:** Results in 100mm clearance on connector side, 50mm on other side

**Data Flow:**
1. **Refresh Phase:** `MepIntersectionService.FindDamperIntersectionsInternal()` calculates intersection point using damper's insertion point
2. **Database Save:** Intersection point saved to `ClashZones.IntersectionPointX/Y/Z`
3. **Placement Phase:** `UniversalSleevePlacerService` uses intersection point as base placement point, then applies offset from `DamperPlacementStrategy`

**Key Point:** The damper's insertion point (LocationPoint) is the geometric center of the damper body, excluding connectors. Connectors are separate Revit entities and do not affect the insertion point calculation.

#### 6.0.1.2 Damper MEP Connector Side Detection and Clearance Assignment

**Location:** `Services/DamperDetection/DamperConnectorDetector.cs` and `Services/Strategies/DamperPlacementStrategy.cs`

**Purpose:** Detect which side of a damper has an MEP connector and assign asymmetric clearance accordingly (100mm on MEP connector side, 50mm on other side).

**Overview:**
- **Detection Method:** Rotation-aware approach using damper's transform to map connector's family-space direction to world coordinates
- **Detection Result:** World coordinate direction (`"+X"`, `"-X"`, `"+Y"`, `"-Y"`, `"+Z"`, `"-Z"`)
- **Clearance Assignment:** Maps world coordinate direction to sleeve clearance sides (Left, Right, Top, Bottom) based on wall orientation
- **Placement Offset:** Calculates offset vector to achieve correct clearance distribution (25mm toward connector side)

##### Detection Logic (Rotation-Aware Approach)

**Service:** `DamperConnectorDetector.DetectConnectorSideWorld()`

**Key Principle:** Uses the damper's transform to map the connector's family-space `BasisX` direction to world coordinates. This correctly handles damper rotation (e.g., family +X → world -Y when rotated 270°).

**Why Rotation-Aware is Needed:**
- Dampers can be rotated or flipped within the linked file
- The connector's `BasisX` is in the connector's local coordinate system (aligned with damper family)
- We need to transform it to world coordinates to determine which world direction the connector faces
- Example: If a damper is rotated 270°, a connector that points in the family's +X direction becomes -Y in world space

**Algorithm:**
1. **Get Damper Center:**
   - Use transform origin: `damper.GetTransform().Origin`
   - This is the geometric center of the damper body, excluding connectors
   - More accurate than bounding box center (which may include connector geometry)

2. **Select MEP Connector:**
   - **Single Connector:** Use it directly
   - **Multiple Connectors:** Select the one that is:
     - Furthest from damper center (by distance)
     - AND whose BasisX direction (facing direction) aligns with its position relative to center
     - Score: `distance * (1.0 + alignment)` where alignment is dot product of normalized position and BasisX

3. **Get Connector's Local Direction:**
   ```csharp
   XYZ connectorBasisXLocal = connector.CoordinateSystem.BasisX; // Family space
   ```

4. **Transform to World Coordinates:**
   ```csharp
   Transform damperTransform = damper.GetTransform();
   XYZ connectorBasisXWorld = damperTransform.OfVector(connectorBasisXLocal);
   ```
   - The damper's transform represents how the family is oriented in world space
   - `OfVector()` transforms the local direction vector to world coordinates
   - This correctly accounts for rotation (e.g., 270° rotation maps family +X to world -Y)

5. **Determine Dominant World Axis:**
   - Calculate absolute values: `absX`, `absY`, `absZ` from `connectorBasisXWorld`
   - Find which world axis has the largest component (threshold: 0.7)
   - Use the sign of that component to determine direction:
     - If `absZ >= max(absX, absY) && absZ > 0.7` → `+Z` or `-Z` (vertical)
     - Else if `absY >= absX && absY > 0.7` → `+Y` or `-Y` (horizontal)
     - Else if `absX > 0.7` → `+X` or `-X` (horizontal)

6. **Position-Based Fallback (if ambiguous):**
   - If no clear direction from transformed BasisX (all components < 0.7)
   - Use connector position relative to damper center:
     ```csharp
     XYZ connectorToCenter = connectorOrigin - damperCenter;
     ```
   - Determine dominant axis from position vector (same logic as step 5)

7. **Return World Coordinate Direction:**
   - Returns: `"+X"`, `"-X"`, `"+Y"`, `"-Y"`, `"+Z"`, `"-Z"`
   - Example: If connector's transformed BasisX is `(0, -0.9, 0)` → returns `"-Y"`

**Why Rotation-Aware Works:**
- ✅ **Handles Rotation:** Correctly maps family-space directions to world coordinates
- ✅ **Handles Flips:** Damper flips are accounted for in the transform
- ✅ **Accurate for Rotated Dampers:** Works correctly when damper is rotated (e.g., 270° rotation)
- ✅ **Preserves Verticality:** Vertical connectors (+Z/-Z) remain vertical after rotation
- ✅ **Fallback Safety:** Position-based fallback handles ambiguous cases

##### Comprehensive Logging

**Log File:** `damper_connector_debug.log`

**Logged Information:**
```
[HH:mm:ss.fff] [DetectConnectorSideWorld] 
  Damper ID={id}, 
  Document='{title}', 
  Family='{family}', 
  Type='{type}', 
  TotalConnectors={count}, 
  MaxDistance={distance}ft ({distance}mm), 
  DamperCenter=({x}, {y}, {z}) [Transform Origin], 
  ConnectorOrigin=({x}, {y}, {z}), 
  ConnectorBasisXLocal=({x}, {y}, {z}) [family space], 
  ConnectorBasisXWorld=({x}, {y}, {z}) [world space after transform], 
  WorldAbs: X={absX}, Y={absY}, Z={absZ}, 
  WallOrientation='{orientation}' ({wallAwareInfo}), 
  DetectedDirection='{direction}' (based on ROTATION: connector direction transformed to world coordinates)
```

**Purpose:**
- Verify detection correctness
- Debug incorrect detections
- Trace connector selection logic
- Understand rotation transformation (family space → world space)
- Verify wall orientation is passed correctly

##### Clearance Assignment Logic

**Service:** `DamperPlacementStrategy.GetDamperPlacementAdjustment()`

**Input:** 
- `ClashZone.DamperConnectorSide` (world coordinate direction: `"+X"`, `"-X"`, `"+Y"`, `"-Y"`, `"+Z"`, `"-Z"`)
- `ClashZone.HostOrientation` (wall orientation: `"X"` or `"Y"`)
- `ClashZone.StructuralElementType` (host type: `"Wall"`, `"Floor"`, etc.)

**Clearance Values:**
- **MEP Side Clearance:** 100mm (from `DuctAccessoryMepNormal` setting)
- **Other Side Clearance:** 50mm (from `DuctAccessoryOtherNormal` setting)

**Mapping Logic (Wall-Aware):**

**For X-Walls (width along X-axis):**
- `"+X"` → Right side → `ClearanceRight = 100mm`, `ClearanceLeft = 50mm`
- `"-X"` → Left side → `ClearanceLeft = 100mm`, `ClearanceRight = 50mm`
- `"+Y"` → Right side (perpendicular to wall) → `ClearanceRight = 100mm`, `ClearanceLeft = 50mm`
- `"-Y"` → Left side (perpendicular to wall) → `ClearanceLeft = 100mm`, `ClearanceRight = 50mm`

**For Y-Walls (width along Y-axis):**
- `"+Y"` → Right side → `ClearanceRight = 100mm`, `ClearanceLeft = 50mm`
- `"-Y"` → Left side → `ClearanceLeft = 100mm`, `ClearanceRight = 50mm`
- `"+X"` → Right side (perpendicular to wall) → `ClearanceRight = 100mm`, `ClearanceLeft = 50mm`
- `"-X"` → Left side (perpendicular to wall) → `ClearanceLeft = 100mm`, `ClearanceRight = 50mm`

**For Vertical (Z-axis):**
- `"+Z"` → Top → `ClearanceTop = 100mm`, `ClearanceBottom = 50mm`
- `"-Z"` → Bottom → `ClearanceBottom = 100mm`, `ClearanceTop = 50mm`

**Storage in ClashZone:**
```csharp
clashZone.ClearanceLeft = left;      // 100mm or 50mm
clashZone.ClearanceRight = right;    // 100mm or 50mm
clashZone.ClearanceTop = top;        // 100mm or 50mm
clashZone.ClearanceBottom = bottom; // 100mm or 50mm
```

##### Placement Offset Calculation

**Purpose:** Move sleeve toward connector side to achieve correct clearance distribution.

**Problem:** If sleeve is centered on damper, clearance is equally distributed (75mm on each side).

**Solution:** Offset sleeve by `(MEP_clearance - Other_clearance) / 2` toward connector side.

**Formula:**
```
OffsetAmount = (100mm - 50mm) / 2 = 25mm toward connector direction
```

**After Offset:**
- Connector side: 75mm + 25mm = **100mm** ✓
- Other side: 75mm - 25mm = **50mm** ✓

**Offset Vector Calculation:**

**For X-Walls:**
- `"+X"` connector → Offset in `+X` direction: `(offsetAmount, 0, 0)`
- `"-X"` connector → Offset in `-X` direction: `(-offsetAmount, 0, 0)`
- `"+Y"` connector → Offset in `+X` direction (maps to right side): `(offsetAmount, 0, 0)`
- `"-Y"` connector → Offset in `-X` direction (maps to left side): `(-offsetAmount, 0, 0)`

**For Y-Walls:**
- `"+Y"` connector → Offset in `+Y` direction: `(0, offsetAmount, 0)`
- `"-Y"` connector → Offset in `-Y` direction: `(0, -offsetAmount, 0)`
- `"+X"` connector → Offset in `+Y` direction (maps to right side): `(0, offsetAmount, 0)`
- `"-X"` connector → Offset in `-Y` direction (maps to left side): `(0, -offsetAmount, 0)`

**For Vertical:**
- `"+Z"` connector → Offset in `+Z` direction: `(0, 0, offsetAmount)`
- `"-Z"` connector → Offset in `-Z` direction: `(0, 0, -offsetAmount)`

**Critical Rule:** Offset is **ONLY along wall axis (width direction)**, NOT perpendicular to wall.

##### Final Sleeve Dimensions

**Formula:**
```
FinalWidth = BaseWidth + InsulationContribution + ClearanceLeft + ClearanceRight
FinalHeight = BaseHeight + InsulationContribution + ClearanceTop + ClearanceBottom
```

**Example (MSFD Damper with connector on -Y side, Y-wall):**
- Base: 500mm × 500mm
- Insulation: 0mm (dampers don't have insulation)
- Clearance: Left=100mm, Right=50mm, Top=50mm, Bottom=50mm
- **Final:** 650mm × 600mm
- **Offset:** 25mm in `-Y` direction (along wall width axis)

##### Placement Flow Integration

**1. Refresh Phase:**
- `ClashZoneService_Legacy.CreateClashZone()` calls `DamperConnectorService.DetectConnectorInfo()`
- `DamperConnectorDetector.DetectConnectorSideWorld()` detects connector side
- Stores `HasMepConnector` and `DamperConnectorSide` in `ClashZone`
- Saves to database: `ClashZones.HasMepConnector`, `ClashZones.DamperConnectorSide`

**2. Placement Phase:**
- `UniversalSleevePlacerService` calls `DamperPlacementStrategy.GetDamperPlacementAdjustment()`
- Strategy reads `ClashZone.DamperConnectorSide` (world coordinate direction)
- Maps world direction to clearance sides based on wall orientation
- Calculates offset vector and final dimensions
- Stores individual clearance values in `ClashZone` (for sleeve parameter setting)
- Returns `(offsetVector, finalWidth, finalHeight)`

**3. Sleeve Placement:**
- `UniversalSleevePlacerService` applies offset to placement point:
  ```csharp
  adjustedPlacementPoint = placementPointChosen + placementOffset;
  ```
- Sets sleeve dimensions: `Width = finalWidth`, `Height = finalHeight`
- Sets individual clearance parameters on sleeve:
  ```csharp
  ClearanceLeft = clashZone.ClearanceLeft;
  ClearanceRight = clashZone.ClearanceRight;
  ClearanceTop = clashZone.ClearanceTop;
  ClearanceBottom = clashZone.ClearanceBottom;
  ```

##### Logging for Placement

**Log File:** `damper_placement_trace.log`

**Logged Information:**
- Connector detection results (`HasMepConnector`, `DamperConnectorSide`)
- Clearance values from conditions (MEP=100mm, Other=50mm)
- Clearance assignment (which side gets MEP clearance)
- Offset calculation (offset amount and vector)
- Final dimensions calculation
- Wall orientation and mapping logic

**Example Log Entry:**
```
[HH:mm:ss.fff] [STRATEGY-CONNECTOR-DETECTION] Zone {id}: 
  HasMepConnector=True, 
  DamperConnectorSide='-Y' (World Coordinate Direction), 
  HostOrientation='Y', 
  StructuralElementType='Wall'

[HH:mm:ss.fff] [STRATEGY-CLEARANCE-ASSIGNMENT] Zone {id}: 
  ConnectorDirection='-Y', 
  HostOrientation='Y', 
  MEPClearance=100.0mm on Left, 
  OtherClearance=50.0mm on opposite side

[HH:mm:ss.fff] [STRATEGY-OFFSET-DETAIL] Zone {id}: 
  OffsetAmount=25.0mm, 
  ConnectorDirection='-Y', 
  OffsetVector=(0.0, -25.0, 0.0)mm

[HH:mm:ss.fff] [STRATEGY-MSFD-FINAL] Zone {id}: 
  Base(500.0mm) + Clearance(150.0mm width, 100.0mm height) = 
  Final(650.0mm x 600.0mm), Offset=25.0mm toward -Y direction
```

##### Benefits of Rotation-Aware Detection

✅ **Handles Rotation:** Correctly maps family-space connector directions to world coordinates using damper's transform  
✅ **Accurate for Rotated Dampers:** Works correctly when damper is rotated (e.g., 270° rotation maps family +X to world -Y)  
✅ **Preserves Verticality:** Vertical connectors (+Z/-Z) remain vertical after rotation, don't become horizontal  
✅ **Handles Flips:** Damper flips are accounted for in the transform  
✅ **Fallback Safety:** Position-based fallback handles ambiguous cases when transformed direction is unclear  
✅ **Comprehensive Logging:** Full traceability for debugging (shows both local and world-space directions)  
✅ **Wall-Aware Mapping:** Correctly maps world directions to clearance sides based on wall orientation  
✅ **Database Persistence:** Detection results saved to database for consistency  

#### 6.0.2 Insulation Detection Services

**Location:** `Services/InsulationDetection/`

**Services:**
- **`IInsulationDetector`** - Interface for detecting insulation status and thickness
- **`InsulationDetector`** - Implementation that checks insulation parameters and MepElementSize data

**Key Methods:**
- `IsInsulated(element)` - Returns true if element has insulation
- `GetInsulationThickness(element)` - Returns insulation thickness in Revit internal units (feet)
- `GetInsulationInfo(element, mepElementSize)` - Returns tuple (isInsulated, thickness) with priority to MepElementSize data

**Usage in ClashZoneService:**
```csharp
var insulationDetector = new InsulationDetector();
var (isInsulated, thickness) = insulationDetector.GetInsulationInfo(mepElement, mepElementSize);

// Populate ClashZone properties:
clashZone.IsInsulated = isInsulated;
clashZone.InsulationThickness = thickness;
```

**Database Persistence:**
- `IsInsulated` and `InsulationThickness` are saved to `ClashZones` table during refresh
- Used by placement services for consistent sizing calculations

#### 6.0.3 Insulation-Aware Sizing Service

**Location:** `Services/Sizing/`

**Services:**
- **`IInsulationAwareSizingService`** - Interface for calculating final sleeve dimensions with insulation
- **`InsulationAwareSizingService`** - Implementation that applies insulation thickness to sizing calculations

**Key Methods:**
- `CalculateFinalDimensions(rawWidth, rawHeight, rawDiameter, isInsulated, insulationThickness, clearance)` - Core calculation
- `CalculateFinalDimensionsFromClashZone(rawWidth, rawHeight, rawDiameter, clashZone, clearance)` - Convenience method using ClashZone properties

**Formula Applied:**
```
FinalSize = RawSize + (2 × InsulationThickness) + (2 × Clearance)
```

**Usage in Placement Services:**
```csharp
var sizingService = new InsulationAwareSizingService();
var (finalWidth, finalHeight, finalDiameter) = sizingService.CalculateFinalDimensionsFromClashZone(
    rawWidth, rawHeight, rawDiameter, clashZone, clearance);
```

**Integration Points:**
- **UniversalSleevePlacerService** - Uses sizing service for pipes, ducts, and fallback cases
- **NewSleevePlacerService** - Uses sizing service for all MEP categories (future replacement for UniversalSleevePlacerService)
- **DamperPlacementStrategy** - Uses sizing service for symmetric cases, calculates insulation contribution for asymmetric cases
- **CableTrayPlacementStrategy** - Uses sizing service for consistent insulation handling
- **ParallelSleevePlacementPlanner** - Uses sizing service for planning calculations

**SOLID Principles Applied:**
- **Single Responsibility:** Only responsible for dimension calculations with insulation
- **Open/Closed:** Can be extended for category-specific logic without modification
- **Dependency Inversion:** Placement services depend on `IInsulationAwareSizingService` interface

#### 6.0.4 ClashZone Model Extensions

**Location:** `Models/ClashZone.cs`

**New Properties:**
- `IsInsulated` (bool) - Whether MEP element is insulated (detected via IInsulationDetector)
- `InsulationThickness` (double) - Insulation thickness in Revit internal units (feet)
- `IsMSFDDamper` (bool) - Whether MEP connector was detected and side determined
- `IsStandardDamper` (bool) - Whether this is a standard damper family
- `DamperConnectorSide` (string) - Connector side direction (Left, Right, Top, Bottom)

**Data Flow:**
1. **Refresh Phase:** `ClashZoneService_Legacy` uses OOP services to populate insulation and damper properties
2. **Database Save:** Properties saved to `ClashZones` table
3. **Placement Phase:** Placement services read properties from ClashZone and use OOP sizing service

#### 6.0.5 Benefits of OOP Architecture

✅ **Consistency:** All placement services use the same OOP methods for insulation and damper detection  
✅ **Maintainability:** Changes to detection logic isolated to specific services  
✅ **Testability:** Interfaces allow easy mocking for unit tests  
✅ **Extensibility:** New damper types or insulation detection methods can be added without modifying existing code  
✅ **SOLID Compliance:** Follows all five SOLID principles  
✅ **Database-First:** Insulation and damper data saved to database during refresh, used during placement

### 6.1 Clustering Services (Phase 1-10)

**✅ NEW: Refactored Clustering Architecture**

The clustering system is now organized into 10 distinct service phases, each with clear responsibilities:

**Phase 1: Geometry Services** (Static utility classes)
- `DistanceCalculator` - 2D/3D minimum distance calculations
- `RotationMatrixCalculator` - Rotation matrix creation and application
- `CoordinateTransformer` - Coordinate space transformations
- **Location:** `Services/Clustering/Geometry/`

**Phase 2: Proximity Services**
- `IProximityChecker` - Interface for proximity checking strategies
- `BoundingBoxProximityChecker` - Bounding box overlap distance
- `EdgeToEdgeProximityChecker` - Edge-to-edge distance for round pipes/ducts
- `RotatedProximityChecker` - Rotated coordinate system proximity
- `ProximityCheckerFactory` - Factory for selecting appropriate checker
- **Location:** `Services/Clustering/Proximity/`

**Phase 3: BoundingBox Services**
- `IBoundingBoxCalculator` - Interface for bounding box calculations
- `AxisAlignedBoundingBoxCalculator` - Simple union of axis-aligned boxes
- `RotatedBoundingBoxCalculator` - Corner-based rotated bounding boxes
- `CornerBasedBoundingBoxCalculator` - Corner transformation helpers
- **Location:** `Services/Clustering/BoundingBox/`

**Phase 4: Strategy Services**
- `IClusteringStrategy` - Interface for clustering strategies
- `ClusteringStrategyBase` - Base class with common functionality
- `FloorCircularClusteringStrategy` - Floor circular sleeves (2D X-Y, edge-to-edge)
- `FloorRectangularClusteringStrategy` - Floor rectangular sleeves (2D X-Y, bbox overlap)
- `WallAxisAlignedStrategy` - Wall axis-aligned sleeves (2D orientation-Z plane)
- `WallRotatedClusteringStrategy` - Wall rotated sleeves (rotation angle validation)
- `ClusteringStrategyFactory` - Factory for selecting appropriate strategy
- **Location:** `Services/Clustering/Strategy/`

**Phase 5: Placement Services**
- `IClusterPlacementService` - Interface for cluster sleeve placement
- `ClusterPlacementService` - Implementation (creation, sizing, metadata, family loading)
- **Location:** `Services/Clustering/Placement/`

**Phase 6: Rotation Services**
- `IClusterRotationService` - Interface for rotation calculations
- `ClusterRotationService` - Rotation angle determination, rotated bounding boxes
- **Location:** `Services/Clustering/Rotation/`

**Phase 7: Cleanup Services**
- `IClusterCleanupService` - Interface for cleanup operations
- `ClusterCleanupService` - Delete individual sleeves within clusters, reset flags
- **Location:** `Services/Clustering/Cleanup/`

**Phase 8: Algorithm Services**
- `IClusterAlgorithmService` - Interface for clustering algorithms
- `ClusterAlgorithmService` - Spatial grid, flood-fill expansion, proximity checking
- **Location:** `Services/Clustering/Algorithm/`

**Phase 9: Data Services**
- `IClusterDataService` - Interface for data loading and caching
- `ClusterDataService` - Load clash zones from database (PRIMARY) or XML (fallback)
- **Location:** `Services/Clustering/Data/`

**Phase 10: Timeout Services**
- `IClusterTimeoutService` - Interface for timeout protection
- `ClusterTimeoutService` - Timeout monitoring, progress tracking
- **Location:** `Services/Clustering/Timeout/`

**Factory:**
- `ClusterServiceFactory` - Factory for creating fully-wired clustering services
- `CreateRefactored()` - Creates RefactoredClusterService with all Phase 1-10 services
- **Location:** `Services/Clustering/ClusterServiceFactory.cs`

**Orchestrator:**
- `RefactoredClusterService` - Clean orchestrator for all Phase 1-10 services (598 lines)
- **Location:** `Services/Clustering/RefactoredClusterService.cs`

### 6.1 IntersectionProcessor

**Location:** `refresh refactor/intersection_processor.cs`

**Purpose:** Detects MEP vs Structural intersections

**Modes:**
- `Replace` - Replace existing zones (PATH 2/3)
- `Replay` - Replay existing zones (PATH 1)
- `FullDetection` - Full detection with validation (PATH 3)

**Output:** List of `ClashZone` objects

### 6.2 ValidationService

**Location:** `refresh refactor/validation_service.cs`

**Purpose:** 3-point validation for existing clash zones

**Validation Points:**
1. **MEP Element Exists** - Verify `MepElementId` exists in document
2. **Structural Element Exists** - Verify `StructuralElementId` exists in document
3. **Elements Still Intersect** - Verify `CalculateIntersectionPoint()` returns valid point

**Output:**
- `ValidZones` - Zones that pass all 3 points
- `InvalidZones` - Zones that fail any point

### 6.3 ParameterCaptureService

**Location:** `refresh refactor/parameter_capture_service.cs`

**Purpose:** Captures MEP element parameters during detection

**Captured Parameters:**
- Element dimensions (width, height, diameter)
- System type
- Service type
- Other placement-relevant parameters

### 6.4 XmlCacheManager

**Location:** `refresh refactor/xml_cache_manager.cs`

**Purpose:** Loads XML data once (eliminates redundant loads)

**Cached Data:**
- Filter XML files
- Global XML files

**Usage:** XML used as fallback if database has no data

### 6.5 PerformanceMonitor

**Location:** `refresh refactor/performance_monitor.cs`

**Purpose:** Tracks performance metrics for each operation

**Metrics:**
- Operation duration
- Item count processed
- Memory usage

---

## 7. Sleeve Placement Flow

### 7.1 Individual Sleeve Placement

**Service:** `UniversalSleevePlacerService`

**Flow:**
```
1. Load clash zones from database (PRIMARY) or XML (fallback)
2. Sync flags from database (PRIMARY) or Global XML (fallback)
3. For each clash zone:
   - Check if sleeve exists (by MEP+Host+Point)
   - If exists → SKIP
   - If NOT exists → Place sleeve
4. Update flags in database (PRIMARY) and Global XML (fallback)
5. Save sleeve instance IDs
```

**Placement Paths:**
- **PATH 1 (Replay):** Load zones from DB → Sync flags → Check sleeve exists → Place if missing → Update flags → Cluster
- **PATH 2 (Sizing):** Load conditions from DB → Calculate size (with OOP sizing service) → Place sleeve → Update flags → Cluster
- **PATH 3 Invalidated:** Load conditions from DB → Run detection for moved elements → Delete affected sleeves → Reset flags for deleted → Place new sleeves (with OOP sizing service) → Update flags for placed → Recalculate clusters → Handle cluster removal/addition

**NewSleevePlacerService (Future):**
- ✅ **OOP Architecture:** Uses `IInsulationAwareSizingService` for consistent sizing calculations
- ✅ **SOLID Principles:** Dependency injection, single responsibility, open/closed principle
- ✅ **Insulation-Aware:** Automatically accounts for insulation thickness from ClashZone properties
- ✅ **Smart Replay:** Can use saved data for faster placement when conditions haven't changed
- ✅ **Ready for Future Use:** Fully integrated with OOP damper and insulation detection services

### 7.2 Coordinate Update and Bounding Box Persistence

**Service:** `SleeveCoordinateService`

**Purpose:** After individual sleeves are placed, update their coordinates and bounding boxes in the database so clustering can find them.

**Critical Flow:**
```
1. After sleeve placement completes:
   - Sleeves are placed in Revit with SleeveInstanceId set in memory
   - But SleeveInstanceId is NOT yet saved to database (still -1 in DB)

2. UpdateSleeveCoordinatesInXml() is called:
   - Loads ALL clash zones for category from database (not just SleeveInstanceId > 0)
   - Reason: SleeveInstanceId hasn't been saved yet, so we need to match by position
   
3. UpdateSleeveCoordinates() matches sleeves to zones:
   - For each zone: Try direct ID match first (if SleeveInstanceId > 0 in memory)
   - If no match: Try position matching (compare SleevePlacementPoint with sleeve bounding box)
   - When match found: Update SleeveInstanceId in memory
   - Retrieve bounding box from Revit sleeve element
   - Update bounding box coordinates in memory

4. Save to database:
   - Save SleeveInstanceId to ClashZones table (via UpdateSleeveInstanceId)
   - Save bounding box coordinates to ClashZones table (via UpdateSleeveBoundingBoxes)
   - Both operations use ClashZoneGuid for matching
   
5. Clustering can now proceed:
   - Loads zones from database with SleeveInstanceId > 0
   - Finds valid bounding boxes for proximity calculation
   - Forms clusters based on actual sleeve positions
```

**Key Implementation Details:**

- **Database-First Loading:** Loads ALL zones for category (not filtered by `SleeveInstanceId > 0`) because IDs aren't saved yet
- **Position Matching:** Uses `SleevePlacementPoint` to match sleeves when `SleeveInstanceId` is not yet in database
- **Dual Save:** Saves both `SleeveInstanceId` and bounding boxes to database after matching
- **Why This Matters:** Without this, clustering finds 0 sleeves with valid bounding boxes and cannot form clusters

**Repository Methods:**
- `UpdateSleeveInstanceId(Guid clashZoneGuid, int sleeveInstanceId)` - Updates SleeveInstanceId by GUID
- `UpdateSleeveBoundingBoxes(Guid clashZoneGuid, double minX, minY, minZ, maxX, maxY, maxZ)` - Updates bounding box coordinates

### 7.3 Flag Management (Atomic Session Logic)

**Service:** `FlagManager` & `ClashZoneRepository`

**Critical Improvement: Atomic Session Flags**
Instead of a two-step process (Reset All -> Set In-Scope), we now use a **Single Atomic UPDATE** operation during refresh. This eliminates timing issues and ensures perfect synchronization with the current section box and filters.

**The Atomic Logic:**
```sql
UPDATE ClashZones
SET ReadyForPlacementFlag = CASE 
        WHEN (zone IN scope) THEN 1 
        ELSE 0 END,
    IsCurrentClashFlag = CASE 
        WHEN (zone IN scope) THEN 1 
        ELSE 0 END
WHERE (zone matches filter + category)
```

**Definition of "IN SCOPE":**
A zone is considered "In Scope" (Flag = 1) if it meets ALL conditions:
1. Matches current **Filter**
2. Matches current **Category**
3. Is **Unresolved** (`IsResolved=0` AND `IsClusterResolved=0` AND `IsCombinedResolved=0`)
4. Is within the current **Section Box**

**Zones "OUT OF SCOPE" (Flag = 0):**
Any zone that matches the filter/category but fails any other condition (e.g., is resolved OR is outside section box) automatically gets Flag = 0.

**Flag Management by Path:**
- **Refresh Phase:** `SetReadyForPlacementBatchOptimized` runs the atomic UPDATE.
- **Placement Phase:** `VerifyExistingSleevesAndResetFlags` runs at start:
  - Checks if resolved sleeves still exist in Revit.
  - If deleted: Sets `IsResolved=0` AND `IsCurrentClashFlag=1` (forces it back into scope).
- **After Placement:** `BulkResetReadyForPlacementFlags` runs to mark zones as consumed (Flag = 0).

**Why This Works:**
- **No Stale Data:** Moving the section box automatically clears flags for zones now outside it.
- **No Timing Issues:** Set/Clear happens in one SQL statement.
- **Deleted Sleeve Recovery:** `VerifyExisting` explicitly re-enables flags for deleted sleeves so they are placed again.

#### 7.3.1 Combined Sleeve Flag Strategy (CRITICAL)

**Objective:**
Ensure that placing a combined sleeve correctly marks the *constituent* individual clash zones as resolved, preventing double placement, without creating new database rows or losing the connection to the original detection data.

**Core Principle:**
**"Update Existing, Do Not Create New"** - Combined sleeves are virtual entities that "consume" existing clash zones. The database operation must always be an `UPDATE` on the existing `ClashZoneGuid` rows, never an `INSERT`.

**Detailed Workflow:**

1.  **Identification**:
    *   The system uses `ClashZoneGuid` to deterministically identify the original individual clash zones that will form the combined sleeve.
    *   These IDs are collected during the combined sleeve candidates formation phase.

2.  **State Updates (The "Consumption" Logic)**:
    *   When a combined sleeve is successfully placed, the `CombinedClusterPersistenceService` must update the *existing* rows for all constituent zones with the following state:
        *   `IsCombinedResolved = 1` (TRUE) - **PRIMARY FLAG**: This zone is now resolved by a combined sleeve.
        *   `IsResolved = 0` (FALSE) - Reset individual resolution (it's no longer individually resolved).
        *   `IsClusterResolved = 0` (FALSE) - Reset cluster resolution (it's no longer cluster resolved).
        *   `SleeveInstanceId = -1` - Reset individual sleeve ID.
        *   `ClusterSleeveInstanceId = -1` - Reset cluster sleeve ID.
        *   `CombinedClusterSleeveInstanceId = [NewCombinedSleeveId]` - Link to the new combined sleeve.

3.  **Persistence Mechanism**:
    *   **Service**: `CombinedClusterPersistenceService.PersistCombinedCluster`
    *   **Operation**: Calls `ClashZoneRepository.UpdateCombinedResolutionFlags(List<Guid> zoneGuids, int combinedSleeveId)`
    *   **SQL Logic**:
        ```sql
        UPDATE ClashZones
        SET IsCombinedResolved = 1,
            IsResolvedFlag = 0,
            IsClusterResolvedFlag = 0,
            SleeveInstanceId = -1,
            ClusterInstanceId = -1,
            CombinedClusterSleeveInstanceId = @CombinedSleeveId,
            UpdatedAt = CURRENT_TIMESTAMP
        WHERE ClashZoneGuid IN (@Guids)
        ```

4.  **Re-Placement Prevention**:
    *   **Orchestrator Check**: `OpeningCommandOrchestrator` and `CombinedSleeveManager` must filter out zones where `IsCombinedResolved = 1`.
    *   **Query**: `SELECT * FROM ClashZones WHERE IsCombinedResolved = 0 ...`
    *   This ensures that once a zone is part of a combined sleeve, it is invisible to the individual and cluster placement logic.

5.  **Deletion & Release Handling**:
    *   When a combined sleeve is deleted from Revit (detected during Refresh or Flag Sync):
    *   **Service**: `ClashZoneRepository.VerifyExistingSleevesAndResetFlags`
    *   **Action**:
        *   identify missing combined sleeves.
        *   "Release" the constituent zones back to the pool.
    *   **State Update**:
        *   `IsCombinedResolved = 0`
        *   `CombinedClusterSleeveInstanceId = -1`
        *   (Crucially, `IsResolved` and `IsClusterResolved` remain `0`, making the zone eligible for individual/cluster placement again).

---

## 9. FileCombo Flag System

### 9.1 IsFilterComboNew Flag

**Location:** `FileCombos` table, `IsFilterComboNew` column

**Values:**
- `0` = File combo already processed (used)
- `1` = New file combo needs processing (fresh)

### 9.2 Flag Lifecycle

1. **Creation:** New FileCombo created with `IsFilterComboNew = 1` (during clash zone save in PATH 2/3)
2. **During Placement:** Flag stays = 1, allowing PATH 2 (Sizing/Detection) to run for sizing calculations
3. **Reset:** After cluster sleeves are placed, flag reset to `0` via `ClashZoneRepository` (AFTER all placement complete)
4. **Next Refresh:** Flag = 0 triggers PATH 1 (Replay Mode) - skips detection, uses existing zones

**Key Point:** Flag reset happens **after cluster sleeve placement**, not after clash zones saved. This allows PATH 2 (Sizing) to run during placement, then next refresh uses PATH 1 (skip detection).

### 9.3 Flag Decision Logic

**Refresh Path Selection:**
- `IsFilterComboNew = 0` → PATH 1 (Replay Mode)
- `IsFilterComboNew = 1` AND `enableThreePointValidation = false` → PATH 2 (Fresh Placement Mode)
- `IsFilterComboNew = 1` AND `enableThreePointValidation = true` → PATH 3 (Full Detection with Validation)

---

## Damper Connector Detection and Sleeve Placement (Critical Logic)

**Overview:**
- Each damper (duct accessory) is analyzed to determine its connector side using world coordinates: `+X`, `-X`, `+Y`, `-Y`, `+Z`, `-Z`.
- The detection process handles damper flipping, rotation, and arbitrary placement in linked files, ensuring the connector direction is always mapped correctly to the host wall orientation.
- This was developed through extensive debugging of MSFD dampers with connectors on all sides (Left/Right/Top/Bottom) and required multiple iterations to handle the BasisX inward-pointing behavior.

### Connector Detection (DamperConnectorDetector.cs)

**Step 1: Find Damper Center**
```csharp
// Get damper center from bounding box (preferred) or transform origin (fallback)
XYZ damperCenter;
var bbox = damper.get_BoundingBox(null);
if (bbox != null)
{
    damperCenter = (bbox.Min + bbox.Max) / 2.0;
}
else
{
    damperCenter = damper.GetTransform().Origin;
}
```
- This is the geometric center of the damper body, EXCLUDING connectors.
- Critical for correct offset and clearance calculations.
- Bounding box center is preferred because it represents the damper body more accurately than the transform origin (which can be at the connector).

**Step 2: Select the Best Connector**
```csharp
// Find the MEP connector - prefer the one furthest from center
Connector best = null;
double maxDistance = 0;

foreach (Connector c in cm.Connectors)
{
    double distance = c.Origin.DistanceTo(damperCenter);
    if (distance > maxDistance)
    {
        maxDistance = distance;
        best = c;
    }
}
```
- If the damper has multiple connectors, select the one furthest from the damper center.
- This ensures the main airflow connector is chosen, not accessory connectors (e.g., control wiring).
- Distance-based selection is robust against connector naming inconsistencies across different damper families.

**Step 3: Get Connector's BasisX (THE CRITICAL INSIGHT)**
```csharp
// ✅ KEY INSIGHT: The connector's BasisX in Revit points INWARD (toward the damper body)
// NOT outward. We need to NEGATE it to get the side the connector is on.
// 
// The connector.CoordinateSystem is already in WORLD coordinates for placed instances.
// We don't need to transform it - Revit gives us the world-space direction directly.
// But we DO need to negate it because BasisX points inward, not outward.

XYZ connectorBasisX = -best.CoordinateSystem.BasisX; // Already in world coordinates, negated to get outward direction
```
- **STRUGGLE #1**: Initially assumed BasisX pointed outward (toward the connected duct). This caused 180° errors.
- **DISCOVERY**: BasisX points INWARD (into the damper body), representing the direction air flows FROM the connector INTO the damper.
- **FIX**: Negate BasisX to get the outward direction (the side the connector is on).
- **STRUGGLE #2**: Initially thought we needed to transform from family to world coordinates using `damperTransform.OfVector()`.
- **DISCOVERY**: For placed instances, `connector.CoordinateSystem` is ALREADY in world coordinates. Revit handles the transformation automatically.
- **FIX**: Use `connector.CoordinateSystem.BasisX` directly (after negation), no manual transformation needed.

**Step 4: Determine Dominant World Axis**
```csharp
// Calculate absolute values
double absX = Math.Abs(connectorBasisX.X);
double absY = Math.Abs(connectorBasisX.Y);
double absZ = Math.Abs(connectorBasisX.Z);

// Check Z first (vertical)
if (absZ >= absX && absZ >= absY && absZ > 0.5)
{
    detectedDirection = connectorBasisX.Z > 0 ? "+Z" : "-Z";
}
// For Y-walls, prioritize position along Y (fallback to BasisX if near zero)
else if (wallOrientation == "Y")
{
    if (Math.Abs(positionVector.Y) >= axisPosThreshold)
        detectedDirection = positionVector.Y > 0 ? "-Y" : "+Y"; // FLIPPED: connector offset +Y means damper is on -Y side
    else
        detectedDirection = connectorBasisX.Y > 0 ? "+Y" : "-Y";
}
// For X-walls, prioritize position along X (fallback to BasisX if near zero)
else if (wallOrientation == "X")
{
    if (Math.Abs(positionVector.X) >= axisPosThreshold)
        detectedDirection = positionVector.X > 0 ? "-X" : "+X"; // FLIPPED: connector offset +X means damper is on -X side
    else
        detectedDirection = connectorBasisX.X > 0 ? "+X" : "-X";
}
```
- **STRUGGLE #3**: BasisX alone sometimes gave ambiguous results (e.g., both X and Y components ~0.7 for 45° rotations).
- **DISCOVERY**: For walls, the connector's position relative to the damper center along the wall width axis is more reliable than BasisX.
- **FIX**: Prioritize position vector along wall width axis (X for X-walls, Y for Y-walls), fallback to BasisX sign if position is near zero.
- **CRITICAL**: Position vector is FLIPPED: if connector is at +Y from center, the damper body is on the -Y side (and vice versa).
- Threshold for "dominant" axis: 0.5 (meaning ≥50% of magnitude in that direction).
- Threshold for "significant position": 0.02 ft (~6mm) to filter out floating-point noise.

**Step 5: Logging for Verification**
```csharp
string logMessage = $"[{DateTime.Now:HH:mm:ss.fff}] [DetectConnectorSideWorld] " +
    $"Damper ID={damper?.Id?.IntegerValue ?? -1}, " +
    $"Family='{damper?.Symbol?.Family?.Name ?? "Unknown"}', " +
    $"Type='{damper?.Symbol?.Name ?? "Unknown"}', " +
    $"FacingFlipped={isFacingFlipped}, HandFlipped={isHandFlipped}, " +
    $"TotalConnectors={connectorCount}, " +
    $"DamperCenter=({damperCenter.X:F4}, {damperCenter.Y:F4}, {damperCenter.Z:F4}), " +
    $"ConnectorOrigin=({connectorOrigin.X:F4}, {connectorOrigin.Y:F4}, {connectorOrigin.Z:F4}), " +
    $"ConnectorBasisX=({connectorBasisX.X:F4}, {connectorBasisX.Y:F4}, {connectorBasisX.Z:F4}) [WORLD COORDS - connector facing direction], " +
    $"BasisXAbs: X={absX:F4}, Y={absY:F4}, Z={absZ:F4}, " +
    $"PositionVector=({positionVector.X:F4}, {positionVector.Y:F4}, {positionVector.Z:F4}), " +
    $"PosAbs: X={posAbsX:F4}, Y={posAbsY:F4}, Z={posAbsZ:F4}, " +
    $"WallOrientation='{wallOrientation ?? "null"}', " +
    $"DetectedDirection='{detectedDirection}'\n";

SafeFileLogger.SafeAppendText("damper_connector_debug.log", logMessage);
```
- Comprehensive logging to `damper_connector_debug.log` for debugging and verification.
- Includes flip state, connector count, positions, BasisX components, position vector, and final detected direction.
- **CRITICAL for debugging**: Without this logging, it was impossible to diagnose BasisX inward-pointing behavior and position vector flipping.

### Sleeve Sizing and Placement (DamperPlacementStrategy.cs)

**Step 6: Handle Vertical Connectors on Walls (THE Z-SWAP BUG)**
```csharp
// ✅ WALL-ONLY FIX: If connector is vertical (+Z/-Z), the damper's family Width/Height are rotated
// relative to the wall axes. Swap the base damperWidth/damperHeight before adding clearances.
if (isWallHost && (connectorDir == "+Z" || connectorDir == "-Z"))
{
    double originalWidth = damperWidth;
    double originalHeight = damperHeight;
    damperWidth = originalHeight;
    damperHeight = originalWidth;
    DebugLogger.Info($"[DamperStrategy] WALL Z-CONNECTOR: Swapped base dimensions for vertical connector");
}
```
- **STRUGGLE #4**: For dampers with Top/Bottom (+Z/-Z) connectors on walls, sleeves were incorrectly sized (width and height swapped).
- **DISCOVERY**: Damper families define "Width" and "Height" in their local coordinate system. When a damper is rotated so its connector points vertically (+Z/-Z), the family's Width/Height no longer align with the wall's horizontal/vertical axes.
- **FIX**: For vertical connectors on walls ONLY, swap the base damperWidth and damperHeight BEFORE adding clearances.
- **CRITICAL**: This swap is ONLY for walls. For floors and framing, the swap is NOT needed (and would cause bugs).
- **WHY**: Walls have a specific horizontal/vertical orientation expectation. Floors/framing do not.

**Step 7: Map Connector Direction to Clearance Sides**
```csharp
// ✅ NEW: Map world coordinate directions to clearance sides (wall-aware)
switch (connectorDir)
{
    case "+X":
        if (isXWall)
        {
            right = mepSideClearance; // X-wall: width is along X-axis, so +X = right side
        }
        else if (isYWall)
        {
            right = mepSideClearance; // Y-wall: +X direction means right side of width
        }
        break;
    
    case "+Z":
        top = mepSideClearance; // Vertical: always affects height
        break;
    
    // ... similar for -X, +Y, -Y, -Z
}
```
- World coordinate direction (`+X`, `-X`, etc.) is mapped to clearance sides (left/right/top/bottom).
- For X/Y directions: mapping depends on wall orientation (which axis is width).
- For Z direction: always affects height (top/bottom), regardless of wall orientation.
- MEP side gets `mepSideClearance` (100mm default), other side gets `otherSideClearance` (50mm default).

**Step 8: Calculate Offset for Asymmetric Clearance**
```csharp
// ✅ OFFSET ALONG WALL AXIS: Calculate offset to achieve correct clearance distribution
// Methodology: 
// 1. Sleeve is sized: Base (500mm) + MEP clearance (100mm) + Other clearance (50mm) = 650mm total
// 2. When centered on damper: clearance is 75mm on each side
// 3. To achieve 100mm on connector side and 50mm on other side, move by difference/2
//    Offset = (mepClearance - otherClearance) / 2 = (100 - 50) / 2 = 25mm toward connector
// 4. After move: Connector side = 75 + 25 = 100mm ✓, Other side = 75 - 25 = 50mm ✓
double offsetAmount = (mepSideClearance - otherSideClearance) / 2.0;

switch (connectorDir)
{
    case "+X":
        if (isXWall)
            offsetVector = new XYZ(offsetAmount, 0, 0); // X-wall: offset in +X direction
        else if (isYWall)
            offsetVector = new XYZ(0, offsetAmount, 0); // Y-wall: offset in +Y direction
        break;
    
    case "+Z":
        offsetVector = new XYZ(0, 0, offsetAmount); // Vertical: offset in +Z direction
        break;
    
    // ... similar for -X, +Y, -Y, -Z
}
```
- **STRUGGLE #5**: Initially tried to offset by full difference (50mm), which moved the sleeve too far.
- **DISCOVERY**: Sleeve is already sized with total clearance (100 + 50 = 150mm). Centering it gives 75mm on each side.
- **FIX**: Offset by HALF the difference (25mm) to redistribute clearance from symmetric (75/75) to asymmetric (100/50).
- **CRITICAL**: Offset direction MUST match connector direction and wall orientation. For example, +X connector on X-wall offsets in +X; +X connector on Y-wall offsets in +Y (because Y-wall width is along Y-axis).

### Fail-Proof Mechanisms and Optimizations

**Optimization 1: Bounding Box Center vs Transform Origin**
```csharp
XYZ damperCenter;
var bbox = damper.get_BoundingBox(null);
if (bbox != null)
{
    damperCenter = (bbox.Min + bbox.Max) / 2.0;  // PREFERRED
}
else
{
    damperCenter = damperTransform.Origin;        // FALLBACK
}
```
- **WHY**: Bounding box center represents the damper body geometry center, excluding connectors.
- **FAIL-PROOF**: If bounding box is null (rare), falls back to transform origin.
- **OPTIMIZATION**: Using bbox center improves accuracy for dampers with offset insertion points.

**Optimization 2: Distance-Based Connector Selection**
```csharp
// Find the MEP connector - prefer the one furthest from center
Connector best = null;
double maxDistance = 0;

foreach (Connector c in cm.Connectors)
{
    double distance = c.Origin.DistanceTo(damperCenter);
    if (distance > maxDistance)
    {
        maxDistance = distance;
        best = c;
    }
}
```
- **WHY**: Dampers may have multiple connectors (airflow + control wiring).
- **FAIL-PROOF**: Furthest connector is always the main airflow connector.
- **OPTIMIZATION**: Avoids hardcoded connector naming or indexing assumptions.

**Fail-Proof 3: Multiple Fallback Strategies for Direction Detection**
```csharp
// Priority hierarchy:
// 1. Z-axis (vertical) - highest priority, clearest signal
// 2. Wall orientation + position vector - wall-aware, robust
// 3. Dominant BasisX axis - general fallback
// 4. "+X" default - ultimate fallback

if (absZ >= absX && absZ >= absY && absZ > 0.5)
{
    detectedDirection = connectorBasisX.Z > 0 ? "+Z" : "-Z";  // PRIORITY 1
}
else if (wallOrientation == "Y" && Math.Abs(positionVector.Y) >= axisPosThreshold)
{
    detectedDirection = positionVector.Y > 0 ? "-Y" : "+Y";   // PRIORITY 2
}
else if (wallOrientation == "Y")
{
    detectedDirection = connectorBasisX.Y > 0 ? "+Y" : "-Y";  // PRIORITY 3
}
else
{
    detectedDirection = "+X";  // PRIORITY 4 (ultimate fallback)
}
```
- **FAIL-PROOF**: Multiple detection strategies ensure we ALWAYS get a valid direction.
- **OPTIMIZATION**: Priority order maximizes accuracy (vertical first, then wall-aware position, then BasisX, then default).

**Fail-Proof 4: Threshold-Based Ambiguity Filtering**
```csharp
const double axisPosThreshold = 0.02; // ~6mm in feet - filters floating-point noise
const double dominantAxisThreshold = 0.5; // 50% magnitude minimum

if (absZ >= absX && absZ >= absY && absZ > 0.5)  // ← 0.5 threshold ensures clear dominance
{
    detectedDirection = connectorBasisX.Z > 0 ? "+Z" : "-Z";
}
```
- **WHY**: Floating-point arithmetic can produce noise (e.g., 0.0001 instead of 0.0).
- **FAIL-PROOF**: Thresholds filter out ambiguous cases and numerical noise.
- **OPTIMIZATION**: 0.02 ft (~6mm) for position, 0.5 (50%) for axis dominance are empirically validated.

**Fail-Proof 5: Null-Safety Throughout**
```csharp
var cm = damper.MEPModel?.ConnectorManager;
if (cm == null)
{
    SafeFileLogger.SafeAppendText("damper_connector_debug.log", 
        $"[{DateTime.Now:HH:mm:ss.fff}] [DetectConnectorSideWorld] Damper ID={damper?.Id?.IntegerValue ?? -1}: No ConnectorManager, returning '+X' (fallback)\n");
    return "+X"; // Default fallback
}
```
- **FAIL-PROOF**: Null-checks on every potential null reference (damper, MEPModel, ConnectorManager, etc.).
- **OPTIMIZATION**: Early returns prevent expensive operations on invalid data.
- **LOGGING**: Every fallback path is logged for debugging.

**Fail-Proof 6: Comprehensive Logging for Debugging**
```csharp
string logMessage = $"[{DateTime.Now:HH:mm:ss.fff}] [DetectConnectorSideWorld] " +
    $"Damper ID={damper?.Id?.IntegerValue ?? -1}, " +
    $"Family='{damper?.Symbol?.Family?.Name ?? "Unknown"}', " +
    $"FacingFlipped={isFacingFlipped}, HandFlipped={isHandFlipped}, " +
    $"ConnectorBasisX=({connectorBasisX.X:F4}, {connectorBasisX.Y:F4}, {connectorBasisX.Z:F4}), " +
    $"PositionVector=({positionVector.X:F4}, {positionVector.Y:F4}, {positionVector.Z:F4}), " +
    $"DetectedDirection='{detectedDirection}'\n";

SafeFileLogger.SafeAppendText("damper_connector_debug.log", logMessage);
```
- **FAIL-PROOF**: Every detection is logged with full context (ID, family, flip state, vectors, result).
- **OPTIMIZATION**: Timestamped logs enable rapid debugging of production issues.
- **CRITICAL**: Without this logging, diagnosing the BasisX inward-pointing issue would have been impossible.

**Fail-Proof 7: Build Stamp for Version Tracking**
```csharp
try
{
    var asm = typeof(DamperConnectorDetector).Assembly;
    string asmLoc = asm.Location;
    DateTime asmWrite = File.Exists(asmLoc) ? File.GetLastWriteTime(asmLoc) : DateTime.MinValue;
    string asmVer = asm.GetName().Version?.ToString() ?? "unknown";
    SafeFileLogger.SafeAppendText("damper_connector_debug.log",
        $"[{DateTime.Now:HH:mm:ss.fff}] [BUILD-STAMP] AssemblyLastWrite={asmWrite:yyyy-MM-dd HH:mm:ss}, Version={asmVer}\n");
}
catch { }
```
- **FAIL-PROOF**: Try-catch ensures logging never crashes the detection logic.
- **OPTIMIZATION**: Build timestamp in logs confirms which version is running in production.
- **CRITICAL**: Enables verification that latest fix is deployed when user reports bugs.

**Optimization 8: Wall-Only Z-Swap Guard**
```csharp
if (isWallHost && (connectorDir == "+Z" || connectorDir == "-Z"))
{
    // Swap only for walls
    damperWidth = originalHeight;
    damperHeight = originalWidth;
}
```
- **FAIL-PROOF**: `isWallHost` check prevents applying the swap to floors/framing where it would cause bugs.
- **OPTIMIZATION**: Targeted fix minimizes code complexity and testing surface.

**Optimization 9: Wall Orientation Context Passing**
```csharp
public string DetectConnectorSide(FamilyInstance damper, bool useWorldCoordinates, 
                                  out Connector connector, string wallOrientation = null)
```
- **OPTIMIZATION**: Passing wall orientation ("X" or "Y") enables wall-aware position vector prioritization.
- **FAIL-PROOF**: Optional parameter (default null) ensures backward compatibility.

**Optimization 10: Offset Calculation Math Validation**
```csharp
// Methodology: 
// 1. Sleeve is sized: Base (500mm) + MEP clearance (100mm) + Other clearance (50mm) = 650mm total
// 2. When centered on damper: clearance is 75mm on each side
// 3. To achieve 100mm on connector side and 50mm on other side, move by difference/2
//    Offset = (mepClearance - otherClearance) / 2 = (100 - 50) / 2 = 25mm toward connector
// 4. After move: Connector side = 75 + 25 = 100mm ✓, Other side = 75 - 25 = 50mm ✓
double offsetAmount = (mepSideClearance - otherSideClearance) / 2.0;
```
- **FAIL-PROOF**: Mathematical validation in comments ensures correctness and prevents future regressions.
- **OPTIMIZATION**: Half-difference offset minimizes sleeve movement while achieving asymmetric clearance.

### Why This Works (Summary of Lessons Learned)

1. **BasisX Inward-Pointing**: Revit's connector BasisX points INTO the damper, not out. Must negate to get the side.
2. **Already World Coordinates**: For placed instances, `connector.CoordinateSystem` is already in world space. No manual transform needed.
3. **Position Vector Flipped**: Connector at +Y from center means damper is on -Y side (and vice versa).
4. **Z-Swap for Walls**: Vertical connectors on walls require swapping Width/Height before adding clearances.
5. **Wall-Aware Mapping**: Offset and clearance mapping must account for wall orientation (X-wall vs Y-wall).
6. **Half-Difference Offset**: To redistribute symmetric clearance to asymmetric, offset by half the difference, not the full difference.
7. **Multiple Fallbacks**: Priority hierarchy ensures we ALWAYS get a valid direction, even for edge cases.
8. **Threshold Filtering**: Ambiguity and floating-point noise are eliminated with empirically validated thresholds.
9. **Null-Safety Everywhere**: Every potential null reference is checked, with logged fallbacks.
10. **Comprehensive Logging**: Every detection is logged with full context for rapid debugging.

**Result**: Robust, fail-proof connector detection and sleeve placement for dampers with connectors on any side (Left/Right/Top/Bottom), on any wall orientation (X/Y), with correct clearances and positioning. The system handles edge cases, invalid data, and production debugging through multiple optimization layers and fail-safe mechanisms.

---

## 10. Comprehensive Flag Management System

**Status:** Current Architecture - Database-Only (SQLITE)  
**Source of Truth:** All flags are managed exclusively in the SQLite database (`ClashZones` and `FileCombos` tables). Legacy Global/Filter XML files are NO LONGER used for flag management.

### 10.1 Core Flag Definitions

| Flag | Table | Purpose |
| :--- | :--- | :--- |
| **`ReadyForPlacementFlag`** | `ClashZones` | **Primary Processor Flag:** The ONLY flag used by placement services to identify zones for processing. Set to `1` only if all session and physical constraints are met. |
| **`IsCurrentClashFlag`** | `ClashZones` | **Initialization Helper:** Its ONLY role is to facilitate the setting of `ReadyForPlacementFlag` by identifying zones that match the active Context Session (Filters and File Combos). |
| **`IsResolvedFlag`** | `ClashZones` | **Individual Resolution:** `1` if an individual sleeve has been placed for this clash. |
| **`IsClusterResolvedFlag`** | `ClashZones` | **Cluster Resolution:** `1` if this clash is covered by a cluster sleeve. |
| **`IsCombinedResolved`** | `ClashZones` | **Manual/Refactored Combined Resolution:** `1` if part of a manual or refactored combined sleeve. |
| **`IsFilterComboNew`** | `FileCombos` | **Optimization Path:** `1` for fresh detection runs, `0` for subsequent "Replay" (Path 1) runs. |

### 10.2 Flag Hierarchy & Logic Flow

The system uses a strict hierarchy to determine if a zone is "eligible" for placement:

1.  **Context Session Activation (Refresh):**
    - `SetReadyForPlacementBatchOptimized` performs a two-step **Atomic Session Update**:
        - **Step A (Reset):** Set `ReadyForPlacementFlag = 0` for all zones in the current filter/category context.
        - **Step B (Activate):** Set `IsCurrentClashFlag = 1` for ALL zones matching the active Context Session (Filters and File Combos).
        - **Step C (Identify Ready Zones):** For zones with `IsCurrentClashFlag = 1`, set `ReadyForPlacementFlag = 1` ONLY if the zone is:
            - Within the Section Box.
            - **UNRESOLVED** (`IsResolved=0` AND `IsCluster=0` AND `IsCombined=0`).

2.  **Path-Based Resolution (Path 1 vs Path 3):**
    - **Path 1 (Replay Mode):** Uses existing zones. Flags are set based on current Revit existence and resolution status.
    - **Path 3 (Adopt/Force Detection):** 
        - **Validity Check:** Every existing zone is checked for validity (elements must exist and intersect).
        - **Invalid Zone Removal:** If a zone is invalid, it is **DELETED** from the database.
        - **Valid Zone Handling:** Valid zones are treated similarly to Path 1 (flags set for placement if unresolved).

2.  **Autonomous Placement (BulkPlacementService):**
    - Queries EXCLUSIVELY for `ReadyForPlacementFlag = 1`. 
    - The `IsCurrentClashFlag` is ignored during this phase as its role was completed during the Refresh/Activation phase.
    - Once placed, sets `IsResolvedFlag = 1` (or cluster/combined flag) and resets consumed flags.

3.  **Completion Reset:**
    - After successful placement, `ResetProcessedFlags` updates the `FileCombos` table, setting `IsFilterComboNew = 0`.
    - This ensures the next run transitions from Path 2/3 (Detection) to Path 1 (Fast Replay).

### 10.3 Automatic Recovery Mechanisms

- **Sleeve Deletion:** If a user deletes a sleeve in Revit, the `VerifyExistingSleevesAndResetFlags` service detects the missing ElementId and resets the corresponding database flags (`IsResolved=0`, etc.) and re-enables `IsCurrentClashFlag=1` so it can be re-placed.
- **Section Box Move:** Because the refresh uses an Atomic Update, zones moving out of the section box automatically have their `ReadyForPlacementFlag` set to `0`.

### 10.4 Atomic Session Logic (SQL)

```sql
UPDATE ClashZones
SET ReadyForPlacementFlag = CASE 
        WHEN (zone IN SectionBox AND MatchesFilters AND NOT IsResolved) THEN 1 
        ELSE 0 END,
    IsCurrentClashFlag = CASE 
        WHEN (zone IN SectionBox AND MatchesFilters) THEN 1 
        ELSE 0 END
WHERE (FilterId = @FilterId)
```

### 10.5 Performance Optimization through Flag Management

The primary performance benefit of this architecture is **Redundant Operation Elimination**:

- **Targeted Processing:** By using `ReadyForPlacementFlag` as the sole trigger for the `BulkPlacementService`, the system ignores thousands of "Resolved" or "Out-of-Scope" zones, focusing Revit API resources only on what needs to be placed **now**.

### 10.6 Cross-Filter Consistency Principles

The flag system operates with **Filter Independence** to ensure global data integrity across different user UI selections:

- **Universal Resolution:** If a clash zone (identified by its deterministic MEP+Host+Point GUID) is resolved in one filter (e.g., "Plumbing"), it is automatically considered resolved in any other filter (e.g., "Electrical") that might detect the same intersection.
- **Cross-Filter Skip:** When a new filter is run, the system checks the database for existing resolution flags for those intersections. If a sleeve already exists (processed by a previous filter), the new filter will skip placement for that zone.
- **Global Flag Reset:** If a sleeve is deleted in Revit, the flag reset mechanism (`VerifyExistingSleevesAndResetFlags`) resets the flags for that intersection globally. This ensures that any filter covering that category will re-detect the zone as "Unresolved" and re-enable it for placement in the next run.
- **Reliability:** This prevents double-placement of sleeves and ensures that the model remains the single source of truth, synchronized with the database flags regardless of which filter combination is currently active.


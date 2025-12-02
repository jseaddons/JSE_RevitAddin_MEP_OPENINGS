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

### 7.3 Flag Management

**Service:** `FlagManager`

**Database-First Operations:**
- **READ**: Load flags from `ClashZones` table (PRIMARY)
- **UPDATE**: Update flags in `ClashZones` table (PRIMARY)
- **FALLBACK**: Use Global XML if database has no data

**Flags Managed:**
- `IsResolved` - Individual sleeve placed (checks if element exists in Revit)
- `IsClusterResolved` - Cluster sleeve placed (checks if element exists in Revit)
- `SleeveInstanceId` - Individual sleeve ElementId
- `ClusterSleeveInstanceId` - Cluster sleeve ElementId

**Important:** Flags are NOT used to skip clustering calculation. They only verify if sleeve elements still exist in Revit (for flag reset if deleted).

**Flag Management by Path:**
- **PATH 1:** Checks flags before placement (if sleeve exists → skip), updates flags after placement
- **PATH 2:** No flag check before placement, updates flags after placement (`IsResolved = true`, `SleeveInstanceId`)
- **PATH 3 Invalidated:** Resets flags for deleted sleeves (`IsResolved = true` to prevent re-placement), updates flags for placed sleeves (`IsResolved = true`, `SleeveInstanceId`)

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

**Connector Detection:**
- Each damper (duct accessory) is analyzed to determine its connector side using world coordinates: `+X`, `-X`, `+Y`, `-Y`, `+Z`, `-Z`.
- The connector side is detected by evaluating the damper's position and orientation relative to the host wall or floor, using the `DamperConnectorDetector` and position vector logic.
- The detected connector direction is stored in the `ClashZone` as `DamperConnectorSide` and used throughout placement and sizing.

**Top/Bottom Connection (+Z/-Z) and Width/Height Swap:**
- For most damper families, the parameters `Damper Width` and `Damper Height` are defined relative to the damper's product orientation, not the host wall axes.
- When the connector is on the left or right (`+X`, `-X`, `+Y`, `-Y`), the damper's width aligns with the wall's horizontal axis, and height aligns with the vertical axis.
- **Critical Case:** When the connector is on the top or bottom (`+Z`, `-Z`), the damper's width and height are rotated 90° relative to the wall axes. This means:
    - The family parameter `Width` now aligns with the wall's vertical axis.
    - The family parameter `Height` now aligns with the wall's horizontal axis.
- **Bug Fix:** To ensure correct sleeve placement, the system swaps the damper's base width and height before adding clearances when the connector is vertical (+Z/-Z) and the host is a wall.
    - This swap is performed in `DamperPlacementStrategy.GetDamperPlacementAdjustment`.
    - Logging is added to `damper_placement_trace.log` for traceability.

**Why the Swap Is Needed:**
- Without swapping, sleeves placed for top/bottom connectors would have their dimensions transposed, resulting in incorrect opening sizes in the wall.
- The swap ensures that:
    - For left/right connectors, sleeve width = damper width, sleeve height = damper height.
    - For top/bottom connectors, sleeve width = damper height, sleeve height = damper width (after swap).
- This logic guarantees that all damper sleeves are placed with correct orientation and sizing, regardless of connector direction.

---
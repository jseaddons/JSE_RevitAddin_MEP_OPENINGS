# Sleeve Placement Methodology - Refactored Architecture

**Document Version:** 2.0 (December 2025)  
**Status:** Current Architecture - Database-Driven, Refactored Services Only

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
| **UniversalSleevePlacerService** | `Services/UniversalSleevePlacerService.cs` | Individual sleeve placement |
| **UniversalClusterService** | `Services/UniversalClusterService.cs` | Cluster sleeve formation and placement |

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
- **PATH 2 (Sizing):** Load conditions from DB → Calculate size → Place sleeve → Update flags → Cluster
- **PATH 3 Invalidated:** Load conditions from DB → Run detection for moved elements → Delete affected sleeves → Reset flags for deleted → Place new sleeves → Update flags for placed → Recalculate clusters → Handle cluster removal/addition

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

## 8. Clustering Flow

### 8.1 Clustering Decision Logic

**Service:** `UniversalClusterService`

**Clustering is path-dependent** - Different paths have different clustering requirements:

#### PATH 1: REPLAY MODE - Conditional Clustering

**Decision:** Check `IsClusterResolved` flags from database

**Logic:**
```
1. Query database: COUNT(*) FROM ClashZones 
   WHERE ComboId = @ComboId 
     AND SleeveInstanceId > 0 
     AND IsClusterResolvedFlag = 0

2. If count = 0 (all zones resolved):
   → SKIP clustering (all sleeves already clustered)

3. If count > 0 (any unresolved zones):
   → RUN clustering (only for unresolved zones)
```

**Use Case:** Replay mode - only cluster if there are new or unresolved sleeves

#### PATH 2: FRESH PLACEMENT MODE - Always Cluster

**Decision:** Always run clustering

**Logic:**
```
→ Always run clustering (fresh detection, new sleeves placed)
→ No flag check needed
```

**Use Case:** Fresh detection - new sleeves always need clustering

#### PATH 3: FULL DETECTION WITH VALIDATION - Zone-Type Based

**Decision:** Depends on zone category

**Placement Routing:**
- **Validated zones**: Route to PATH 1 placement logic (check sleeve exists, place if missing)
- **Invalidated zones**: Route to PATH 3 invalidated placement logic (distinct flow)
- **New zones**: Route to PATH 2 placement logic (calculate size, place)

**Clustering Logic:**
```
Validated Zones (all 3 points pass):
  → Use PATH 1 placement logic (check sleeve exists, place if missing)
  → After placement: Always run clustering recalculation
  → Reason: Nearby zones may have changed (invalidated/new), affecting cluster formation
  → Always calculate fresh (don't use stored data from ClusterSleeves table)
  → Save cluster data to ClusterSleeves table

Invalidated Zones (any point fails):
  → Always run clustering recalculation
  → Check if cluster needed for affected zones
  → If cluster not needed: Remove cluster sleeve
  → If cluster needed: Add cluster sleeve
  → Save/update cluster data to ClusterSleeves table

New Zones (not in existing):
  → Use PATH 2 placement logic (calculate size, place)
  → After placement: Always run clustering calculation (same as PATH 2)
  → Always calculate fresh
  → Save new cluster data to ClusterSleeves table
```

**Use Case:** Validation mode - validated zones use PATH 1 logic, invalidated zones use distinct PATH 3 placement, new zones use PATH 2 logic

### 8.2 Clustering Decision Summary

| Path | Zone Type | Clustering Decision | Reason |
|------|-----------|-------------------|--------|
| **PATH 1** | All zones | Check ClusterSleeves table → Use stored data if exists, else calculate | Replay mode, use pre-calculated cluster data |
| **PATH 2** | All zones | Always run calculation → Save to ClusterSleeves | Fresh detection, new sleeves |
| **PATH 3** | Validated | Check ClusterSleeves table → Use stored data if exists, else calculate | Same as PATH 1 |
| **PATH 3** | Invalidated | Always run calculation → Save to ClusterSleeves | Geometry changed, needs re-clustering |
| **PATH 3** | New | Always run calculation → Save to ClusterSleeves | Fresh detection, new sleeves |

### 8.3 Cluster Formation Process

**Service:** `UniversalClusterService` + `ClusterSleeveRepository`

**Flow (when clustering calculation is required - PATH 2/3 or PATH 1 when no cluster data exists):**
```
1. Load clash zones from database (PRIMARY) or XML (fallback)
   - Filter: SleeveInstanceId > 0 (placed sleeves)
   - Filter: IsClusterResolved = false (not yet clustered)
2. Group sleeves by:
   - Host type (wall, floor, framing)
   - System type
   - Orientation
3. Calculate proximity using bounding boxes
4. Form clusters based on join distance
5. Calculate cluster bounding boxes (rotated if needed)
6. Calculate cluster data:
   - Bounding box coordinates (minX, minY, minZ, maxX, maxY, maxZ)
   - Dimensions (width, height, depth)
   - Rotation angle (if rotated)
   - Placement point (center of cluster)
   - List of ClashZoneIds in cluster
7. Save cluster data to ClusterSleeves table (via ClusterSleeveRepository)
8. Place cluster sleeves using calculated data
9. Delete individual sleeves within clusters
10. Update flags in database (PRIMARY) and Global XML (fallback)
    - Set IsClusterResolved = true
    - Set ClusterSleeveInstanceId
    - Clear SleeveInstanceId
11. Reset IsFilterComboNew = 0 (after cluster placement complete)
```

**Flow (when cluster data exists - PATH 1 replay):**
```
1. Check ClusterSleeves table: HasClusterDataForCombo(comboId, category)
2. If cluster data exists:
   - Load cluster data from ClusterSleeves table (via ClusterSleeveRepository)
   - Load: bounding box, dimensions, rotation, placement point, ClashZoneIds
   - Place cluster sleeves using stored data (SKIP calculation)
   - Update flags in database
   - Reset IsFilterComboNew = 0 (after cluster placement complete)
3. If no cluster data:
   - Run full clustering calculation (see flow above)
   - Reset IsFilterComboNew = 0 (after cluster placement complete)
```

### 8.4 Rotated Bounding Boxes

**Service:** `ClusterBoundingBoxServices`

**Purpose:** Calculate bounding boxes in rotated coordinate system

**Logic:**
- Determine dominant rotation angle from individual sleeves
- If angle is near axis-aligned (0°, 90°, 180°, 270°) → Use axis-aligned bounding box
- Otherwise → Calculate bounding box in rotated coordinate system

**Benefits:**
- Cluster sleeves accurately follow outlines of individual sleeves
- Handles non-orthogonal angles correctly

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
- `IsFilterComboNew = 1` AND `enableThreePointValidation = false` → PATH 2 (Fresh Placement)
- `enableThreePointValidation = true` → PATH 3 (Full Detection with Validation)

**Placement Path Selection:**
- `IsFilterComboNew = 0` → Check if conditions changed:
  - **If conditions changed** → PATH 2 (Sizing) - Recalculate sizes and clusters
  - **If conditions NOT changed** → PATH 1 (Replay) - Uses saved data
- `IsFilterComboNew = 1` → PATH 2 (Sizing/Detection) - Runs detection
- **PATH 3 zones route directly** (bypass flag check):
  - Validated zones → PATH 1 placement
  - Invalidated zones → PATH 3 invalidated placement
  - New zones → PATH 2 placement

**Condition Change Detection:**
- Compare current `OpeningSettings` from UI with saved `OpeningSettings` in database
- If different → Conditions changed → Route to PATH 2 placement
- If same → Conditions unchanged → Use PATH 1 placement

---

## 10. PATH 3 Invalidated Zones Placement Flow

### 10.1 Distinct Placement Logic for Invalidated Zones

**Service:** `UniversalSleevePlacerService` (with PATH 3 invalidated mode)

**Flow:**
```
1. Load conditions from database (Filters table, OpeningSettings column)
2. Run intersection detection for moved MEP/host elements
3. Delete placed sleeves affected by invalid clash zone
4. Reset flags for deleted sleeves → Set IsResolved = true (so they won't be placed again)
5. Calculate sleeve size (read clearance settings from DB, apply to MEP dimensions)
6. Place new sleeves with calculated size (NO flag check before placement)
7. Update flags for placed sleeves → Set IsResolved = true, SleeveInstanceId
8. Always run clustering recalculation
9. Check if cluster needed for affected zones:
   - If cluster not needed → Remove cluster sleeve
   - If cluster needed → Add cluster sleeve
10. Update cluster flags accordingly
```

**Key Differences from PATH 1 and PATH 2:**
- **No flag check before placement** (like PATH 2)
- **Deletes affected sleeves first** (unique to PATH 3)
- **Resets flags for deleted sleeves** (unique to PATH 3)
- **Updates flags for placed sleeves** (like PATH 1 and PATH 2)
- **Handles cluster removal/addition** based on need (unique to PATH 3)

**Flag Management:**
- Resets flags for deleted sleeves: `IsResolved = true` (prevents re-placement)
- Updates flags for placed sleeves: `IsResolved = true`, `SleeveInstanceId`
- Updates cluster flags: `IsClusterResolved = true`, `ClusterSleeveInstanceId` (or clears if cluster removed)

---

## 11. UI State Persistence

### 11.1 Filter UI State

**Stored in:** `Filters` table

**Fields:**
- `SelectedHostCategories` - JSON array of selected host categories (Walls, Floors, Structural Framing, etc.)
- `OpeningSettings` - JSON object containing opening configuration

### 11.2 Save Operations

**Service:** `FilterRepository.SaveFilterUIState()`

**Called From:**
- `FilterManagementService.SaveFilter()`
- `FilterManagementService.CreateFilter()`
- `FilterManagementService.CopyFilter()`
- `FilterManagementService.SaveFilterAuto()`

### 11.3 Load Operations

**Service:** `FilterRepository.LoadFilterUIState()`

**Called From:**
- `FilterManagementService.CreateFilterFromCurrentUIState()`
- `FilterManagementService.LoadFilterFromXmlFile()`

---

## 12. Migration Notes

### 12.1 Legacy Services Removed

- ❌ `Services/RefreshService.cs` - **EXCLUDED from compilation**
- ❌ All references to legacy refresh service removed

### 12.2 Current Services Only

- ✅ `RefreshServiceRefactored` - **PRIMARY SERVICE**
- ✅ `RefreshPathDeterminer` - **PATH SELECTION LOGIC**
- ✅ `IntersectionProcessor` - **INTERSECTION DETECTION**
- ✅ `ValidationService` - **3-POINT VALIDATION**
- ✅ `ClashZoneRepository` - **DATABASE OPERATIONS**
- ✅ `FilterRepository` - **FILTER OPERATIONS**

### 12.3 XML Role

- **PRIMARY:** Database (SQLite)
- **FALLBACK:** XML (only if database has no data)
- **WRITE:** XML writes disabled when `DeploymentConfiguration.DisableXmlCreation = true`

---

## 13. Error Handling and Resilience

### 13.1 Error Handling Strategy

The refactored architecture implements comprehensive error handling at every critical operation point to ensure system resilience and data integrity.

#### 13.1.1 Error Categories

**Critical Errors (Operation Cancellation):**
- Database connection failures
- Transaction commit failures
- Invalid data structure errors
- Missing required configuration

**Non-Critical Errors (Continue with Logging):**
- Individual sleeve placement failures
- Flag update failures for single zones
- Cluster calculation errors (fallback to skip)
- Element deletion failures (log warning, continue)

#### 13.1.2 Error Handling Points

**Database Operations:**
```
CreateFilter → CheckCreateFilter → RollbackCreateFilter (on failure) → Cancel
SaveUIState → CheckSaveUIState → RollbackSaveUIState (on failure) → Cancel
SaveDB2 → CheckSaveDB2 → RollbackSaveDB2 (on failure) → Cancel
Commit2 → CheckCommit2 → RollbackCommit2 (on failure) → Cancel
```

**Data Loading:**
```
LoadFilter → CheckLoadFilter → LogLoadFilterError (on failure) → Cancel
LoadExisting1 → CheckLoadExisting1 → LogLoadExistingError (on failure) → Cancel
LoadConditions2 → CheckLoadConditions2 → LogLoadConditionsError2 (on failure) → Cancel
```

**Sleeve Placement:**
```
PlaceSleeve1 → CheckPlaceSleeve1 → LogPlaceError1 (on failure) → Continue with Next Zone
PlaceSleeve2 → CheckPlaceSleeve2 → LogPlaceError2 (on failure) → Continue with Next Zone
PlaceSleeve3 → CheckPlaceSleeve3 → LogPlaceError3 (on failure) → Continue with Next Zone
```

**Flag Updates:**
```
UpdateFlags1 → CheckUpdateFlags1 → LogFlagError1 (on failure) → Continue with Next Zone
UpdateFlags2 → CheckUpdateFlags2 → LogFlagError2 (on failure) → Continue with Next Zone
UpdateFlags3 → CheckUpdateFlags3 → LogFlagError3 (on failure) → Continue with Next Zone
```

**Clustering Operations:**
```
TryCalcCluster1 → ClusterCalcError1 (on failure) → LogError28 → ResetFlag
TryDelIndiv1 → DelError1 (on failure) → LogError33 → Continue Anyway
```

### 13.2 Transaction Management

**All database operations are wrapped in transactions:**

1. **Transaction Start:** Before any database write operation
2. **Error Detection:** After each critical operation
3. **Rollback on Failure:** If any critical operation fails, rollback entire transaction
4. **Commit on Success:** Only commit if all operations succeed

**Example Flow:**
```csharp
using (var transaction = connection.BeginTransaction())
{
    try
    {
        // Create filter
        if (!CreateFilter(transaction)) throw new Exception("Filter creation failed");
        
        // Save UI state
        if (!SaveUIState(transaction)) throw new Exception("UI state save failed");
        
        // Commit only if all succeed
        transaction.Commit();
    }
    catch (Exception ex)
    {
        transaction.Rollback();
        LogError(ex);
        Cancel();
    }
}
```

### 13.3 Error Logging

**All errors are logged with context:**
- Error type and message
- Operation that failed
- Affected data (filter, file combo, clash zone)
- Timestamp
- Stack trace (for debugging)

**Log Levels:**
- **Error:** Critical failures requiring user attention
- **Warning:** Non-critical failures that don't stop execution
- **Info:** Normal operation flow tracking

### 13.4 Fallback Mechanisms

**Database Connection Failure:**
- Show user-friendly error message
- Suggest checking SQLite file location
- Cancel operation (no partial data)

**Data Load Failure:**
- Fallback to XML if database has no data
- If XML also fails, cancel operation
- Log both attempts for debugging

**Sleeve Placement Failure:**
- Log error for specific zone
- Continue with next zone
- Don't fail entire operation
- Report failures in summary

**Cluster Calculation Failure:**
- Log error
- Skip clustering for affected category
- Continue with other categories
- Reset flag to allow retry

### 13.5 User Feedback

**Error Messages:**
- Clear, actionable error messages
- No technical jargon for end users
- Suggestions for resolution when possible

**Progress Reporting:**
- Show progress during long operations
- Report partial success (e.g., "5 of 10 sleeves placed")
- Summary of errors at completion

---

## 14. Key Principles

1. **Database-First** - All operations prioritize database over XML
2. **Path Strategy** - Three distinct paths based on file combo state
3. **Repository Pattern** - Database operations abstracted through repositories
4. **Transaction Safety** - All database operations wrapped in transactions
5. **Fallback Support** - XML used only when database has no data
6. **Clear Separation** - Each service has distinct responsibilities
7. **Performance** - Single XML load, database-first reads, optimized queries
8. **Error Resilience** - Comprehensive error handling with graceful degradation
9. **Data Integrity** - Transaction rollback prevents partial data corruption
10. **User Experience** - Clear error messages and progress reporting

---

**End of Document**


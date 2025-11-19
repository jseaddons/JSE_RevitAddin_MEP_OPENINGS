# Chapter X: Sleeve Placement Methodology

## Table of Contents

### Part I: Introduction and Overview
1. [Introduction](#1-introduction)
2. [Latest Updates](#2-latest-updates)
3. [Key Concepts](#3-key-concepts)

### Part II: Architecture and Data Flow
4. [Complete Workflow Architecture](#4-complete-workflow-architecture)
   - 4.1 [Phase 1: Process Clash Zones](#41-phase-1-process-clash-zones-button-click)
   - 4.2 [Phase 1.5: 3-Point Validation](#42-phase-15-3-point-validation)
   - 4.3 [Phase 2: Individual Sleeve Placement](#43-phase-2-individual-sleeve-placement)
   - 4.4 [Phase 3: Cluster Sleeve Placement](#44-phase-3-cluster-sleeve-placement)
5. [Data Architecture](#5-data-architecture)
   - 5.1 [Global XML Structure](#51-global-xml-structure)
   - 5.2 [Filter XML Structure](#52-filter-xml-structure)
   - 5.3 [Why Both Are Needed](#53-why-both-are-needed)
6. [Coordinate System Architecture](#6-coordinate-system-architecture)
7. [Flag Management System](#7-flag-management-system)

### Part III: Implementation Details
8. [Class Responsibilities](#8-class-responsibilities)
9. [Flag System Details](#9-flag-system-details)
   - 9.1 [Flag Definitions](#91-flag-definitions)
   - 9.2 [Flag Hierarchy](#92-flag-hierarchy)
   - 9.3 [Flag Logic Flow](#93-flag-logic-flow)
   - 9.4 [Flag Initialization](#94-flag-initialization)
   - 9.5 [Flag Reset Mechanisms](#95-flag-reset-mechanisms)
10. [Clustering Implementation](#10-clustering-implementation)
    - 10.1 [Cluster Formation](#101-cluster-formation)
    - 10.2 [Cluster Placement Process](#102-cluster-placement-process)
    - 10.3 [Bounding Box Calculations](#103-bounding-box-calculations)
    - 10.4 [Rotated Non-Axis-Aligned Clustering](#104-rotated-non-axis-aligned-clustering)
11. [Sleeve Placement Implementation](#11-sleeve-placement-implementation)
    - 11.1 [Pre-calculation Strategy](#111-pre-calculation-strategy)
    - 11.2 [Placement Point Calculation](#112-placement-point-calculation)
    - 11.3 [Host Element Handling](#113-host-element-handling)
    - 11.4 [Clearance Calculations](#114-clearance-calculations)

### Part IV: Performance and Optimization
12. [Multi-Threading Optimizations](#12-multi-threading-optimizations)
    - 12.1 [Clustering Parallelization](#121-clustering-parallelization)
    - 12.2 [Pre-filtering Parallelization](#122-pre-filtering-parallelization)
    - 12.3 [Revit API Limitations](#123-revit-api-limitations)
13. [Performance Optimizations](#13-performance-optimizations)
    - 13.1 [MEP+Host+Point Matching](#131-mephostpoint-matching)
    - 13.2 [Elimination of Duplicate Detection](#132-elimination-of-duplicate-detection)
    - 13.3 [XML Data Usage](#133-xml-data-usage)

### Part V: Critical Features and Fixes
14. [3-Point Validation System](#14-3-point-validation-system)
15. [Filter XML Append-Only Principle](#15-filter-xml-append-only-principle)
16. [Fresh Run Detection Architecture](#16-fresh-run-detection-architecture)
17. [GUID Management System](#17-guid-management-system)
    - 17.1 [Deterministic GUID Generation](#171-deterministic-guid-generation)
    - 17.2 [GUID Storage and Recovery](#172-guid-storage-and-recovery)
18. [Critical Bug Fixes](#18-critical-bug-fixes)
    - 18.1 [Flag Initialization Bug](#181-flag-initialization-bug)
    - 18.2 [Category Mismatch Fix](#182-category-mismatch-fix)
    - 18.3 [Vertical MEP Orientation Fix](#183-vertical-mep-orientation-fix)
    - 18.4 [Duct-Damper Optimization](#184-duct-damper-optimization)

### Part VI: User Interface and Configuration
19. [User Interface Updates](#19-user-interface-updates)
20. [Configuration Management](#20-configuration-management)
    - 20.1 [Clearance Settings](#201-clearance-settings)
    - 20.2 [Pipe Opening Type Rules](#202-pipe-opening-type-rules)
    - 20.3 [Configuration Resolution](#203-configuration-resolution)

### Part VII: Appendices
21. [Code Examples](#21-code-examples)
22. [Implementation Changes Log](#22-implementation-changes-log)
23. [Debugging Guide](#23-debugging-guide)
24. [Migration Strategy](#24-migration-strategy)

---

## 1. Introduction

This chapter documents the complete methodology for sleeve placement in MEP openings, including individual sleeve placement and clustering flow. The system handles clash detection, sleeve placement, and cluster optimization for MEP elements intersecting structural elements in Revit.

### 1.1 Purpose
The sleeve placement system automates the creation of openings in structural elements (walls, floors, structural framing) for MEP elements (ducts, pipes, cable trays). It handles both individual sleeve placement and intelligent clustering of proximate sleeves.

### 1.2 Scope
This documentation covers:
- Clash detection and data preparation
- Individual sleeve placement
- Cluster sleeve formation and placement
- Flag management and state tracking
- Performance optimizations
- Data architecture and persistence

---

## 2. Latest Updates

### 2.1 Global XML Flag Management (Single Source of Truth)
- ✅ Flags (`IsResolved`, `IsClusterResolved`, `SleeveInstanceId`, `ClusterSleeveInstanceId`) managed exclusively in Global XML
- ❌ Filter XML no longer stores flags - eliminated redundant flag writes
- ✅ Improved cross-filter consistency and data integrity

### 2.2 UI Changes
- ✅ "Refresh" button → "Process Clash Zones" (moved to header panel)
- ✅ "OK" button → "Place Sleeves" (in header panel)
- ✅ Better button organization and clearer user intent

### 2.3 Multi-Threading Optimizations
- ✅ Clustering formation parallelized (XML-only operations)
- ✅ Sleeve placement pre-filtering parallelized (XML-only validation)
- ❌ Revit API calls remain sequential (Revit requirement)
- ✅ Significant performance improvements on multi-core CPUs (i5/i7/i9)

### 2.4 Deterministic GUID Generation (November 6, 2025)
- ✅ **Deterministic GUID Generation**: GUIDs are now generated deterministically from MEP+Host+Point hash using MD5
- ✅ **Stable Identifiers**: Same intersection always gets the same GUID across detection runs
- ✅ **Recovery Fix**: Fixed tolerance mismatch (0.01ft → 0.1ft) to match `FindByMepHostAndPoint` default tolerance
- ✅ **Sleeve Metadata Parameters**: Added `LinkedFile` and `HostFile` text parameters to sleeve families
- ✅ **GUID Storage**: GUIDs stored on sleeves via `ClashZone_GUID` parameter for recovery and tracking

**Implementation Details:**
1. **GUID Generation** (`GuidManager.GenerateDeterministicGuid`):
   - Uses MD5 hash of: `MEP_ID|HOST_ID|X|Y|Z` (coordinates rounded to 0.1ft tolerance)
   - Same 3-point combo always produces same GUID
   - Falls back to random GUID if inputs are invalid

2. **Clash Zone Creation** (`ClashZoneService.CreateClashZone`):
   - Sets `clashZone.Id` deterministically using `GuidManager.GenerateDeterministicGuid`
   - Ensures consistency across detection runs even if Global XML is deleted/recreated

3. **Sleeve GUID Setting** (`UniversalSleevePlacerService.SetClashZoneGuidOnSleeveStable`):
   - Sets `ClashZone_GUID` parameter on sleeves using `clashZone.Id` (deterministic)
   - No Global XML lookup needed - deterministic GUID ensures consistency

4. **Recovery Mechanism** (`FlagManager.RecoverSleeveFlagsFromRevit`):
   - Uses deterministic GUID generation for new entries
   - Ensures recovery creates entries with same GUID that `CreateClashZone` would generate
   - Fixed tolerance from 0.01ft to 0.1ft to match matching logic

**Benefits:**
- ✅ **Stability**: Same intersection always gets same GUID across detection runs
- ✅ **Recovery**: Works correctly even when Global XML is deleted/recreated
- ✅ **Consistency**: Recovery and detection create matching GUIDs for same intersections
- ✅ **Industry Standard**: Follows best practices for deterministic clash identification
- ✅ **Tolerance Matching**: Recovery tolerance matches `FindByMepHostAndPoint` default (0.1ft)

**Why Keep GUIDs:**
- XML serialization compatibility
- Standard format for future database integration
- Cross-system interoperability (IFC, COBie, etc.)
- Human-readable format for debugging
- Large namespace (128-bit) reduces collision risk

### 2.5 MEP+Host+Point Matching (Primary Matching Method)
- ✅ **Matching uses MEP+Host+Point** as primary method (not GUID parameter on sleeves)
- ✅ **GUID parameter recommended** for recovery and tracking (`ClashZone_GUID` parameter)
- ✅ **Cross-filter matching works automatically** using MEP+Host+Point

**Benefits:**
- ✅ **Performance**: O(1) Dictionary lookup vs O(n) sleeve iteration
- ✅ **Reliability**: Matches by actual intersection data (MEP+Host+Point) instead of GUID
- ✅ **Fallback**: Works even if GUID parameter is missing on sleeves
- ✅ **Cross-Filter**: Works automatically when switching filters
- ✅ **GUID Value**: GUID parameter helps with recovery and provides stable identifier

### 2.6 Placement Path Services (November 2025)
- ✅ **Path 1 – Replay Placement**: `SleevePlacementReplayService` places sleeves strictly from the persisted snapshot (intersection point, dimensions, offsets, cluster payload). After placement it calls `SaveClashZones(..., allowStructuralUpdates: false)` so only flag/ID changes flow back to Global XML.
- ✅ **Path 2 – Clearance/Type Recalc**: `SleeveSizingService` recomputes sleeve width/height/diameter and opening type when UI overrides change. It updates the snapshot via `SaveClashZones(..., false)` and immediately reuses Path 1 to execute placement with the refreshed data.
- ✅ **Path 3 – Detection Rebuild**: `SleeveDetectionService` owns “Adopt to modified document” runs and global configuration changes. It reruns full clash detection and persists results with `SaveClashZones(..., true)`, which is the only path allowed to overwrite intersection geometry.
- ✅ **Persistence Impact**: Replay (Path 1) defines the canonical placement snapshot; any service that mutates geometry/size must hand control back to Path 1 for post-placement synchronization. Global XML remains the flag ledger while Filter XML captures the latest placement payload.

---

## 3. Key Concepts

### 3.1 Clash Zone
A clash zone represents an intersection between a MEP element and a structural element. It contains all the data needed for sleeve placement, including coordinates, dimensions, host information, and placement flags.

### 3.2 Individual Sleeve
A single sleeve placed at the intersection point of one MEP element with a structural element. Individual sleeves are placed first before clustering.

### 3.3 Cluster Sleeve
A combined sleeve that replaces multiple individual sleeves when they are proximate to each other. Cluster sleeves are created after individual sleeves are placed.

### 3.4 Flag System
Flags track the state of sleeve placement:
- `IsResolved` - Individual sleeve has been placed
- `IsClusterResolved` - Cluster sleeve has been placed
- `MarkedForClusteringSleeveProcess` - Indicates whether a sleeve should be clustered

### 3.5 Global XML vs Filter XML
- **Global XML**: Single source of truth for flags and cross-filter state tracking
- **Filter XML**: Stores placement data (coordinates, dimensions, host information)

---

## 4. Complete Workflow Architecture

### 4.1 Phase 1: Process Clash Zones Button Click

**Purpose:** Detect clashes and save coordinate data for both individual and cluster processing

**UI Button:** "Process Clash Zones" (formerly "Refresh" button)

**Services Involved:**
- `RefreshServiceRefactored` - Main orchestrator (✅ REFACTORED SERVICE - Database-driven)
- `IntersectionProcessor` - Detects MEP vs Structural clashes (with Replace/Replay/FullDetection modes)
- `ClashZoneService` - Creates and manages ClashZone objects
- `FlagManager` - Syncs flags from Database (PRIMARY) or Global XML (fallback) to in-memory objects
- `ClashZoneRepository` - Database operations for clash zones (PRIMARY source)
- `FilterRepository` - Database operations for filters and UI state

**What Happens:**
1. **Detect Intersections:** Find MEP elements intersecting structural elements
2. **Create ClashZones:** One ClashZone per (MEP + Structural) pair
3. **Save Dual Coordinates:**
   - `IntersectionPoint` = Linked document coordinates (for individual sleeve placement)
   - `SleevePlacementPointActiveDocument` = Active document coordinates (for proximity calculation)
4. **Initialize Flags:** All flags start as `false` or `null`

**Output:** XML files with ClashZone data containing both coordinate systems

### 4.2 Phase 1.5: 3-Point Validation

**Purpose:** Validate existing clash zones and remove stale entries when elements are deleted or no longer intersect

**Services Involved:**
- `RefreshServiceRefactored` - Performs validation during refresh (✅ REFACTORED SERVICE)
- `ValidationService` - 3-point validation logic
- `ClashZoneRepository` - Database operations (PRIMARY source)
- `GlobalIndexService` - Clears invalid entries from Global XML (fallback only)

**What Happens:**
1. **Load Existing ClashZones:** From Filter XML files created in previous runs
2. **3-Point Validation (Per ClashZone):**
   - ✅ **Point 1: MEP Element Exists** → Verify `MepElementId` exists in document
   - ✅ **Point 2: Structural Element Exists** → Verify `StructuralElementId` exists in document
   - ✅ **Point 3: Elements Still Intersect** → Verify `CalculateIntersectionPoint()` returns valid point
3. **Handle Valid ClashZones (All 3 Points Pass):**
   - Keep clash zone and preserve GUID
   - Update `IntersectionPoint` if elements moved (>1mm difference)
   - Update `SleevePlacementPointActiveDocument` with new coordinates
   - Preserve existing flags (`IsResolved`, `IsClusterResolved`)
4. **Handle Invalid ClashZones (Any Point Fails):**
   - Remove clash zone from Filter XML
   - Clear entry from Global XML (`CategoryGlobalIndex`)
   - Log removal reason (MEP deleted / Structural deleted / No longer intersect)

**Key Principle:**
- **Clash Zone Validity ≠ Sleeve Placement State**
  - Clash zone validity depends on: MEP Element ID + Structural Element ID + Intersection Point
  - Sleeve placement state depends on: Flags (`IsResolved`, `IsClusterResolved`) and sleeve existence
  - A clash zone is valid if elements still intersect, regardless of whether sleeves exist
  - Flags track sleeve placement independently of clash zone validity

**Benefits:**
- ✅ Prevents stale clash zones from accumulating in XML
- ✅ Updates intersection points when elements move but still intersect
- ✅ Automatically cleans up when elements are deleted
- ✅ Maintains data integrity across multiple refresh cycles

**Output:** Validated clash zones with updated coordinates, stale entries removed

### 4.3 Phase 2: Individual Sleeve Placement

**Purpose:** Place individual sleeves using linked coordinates

**Services Involved:**
- `UniversalSleevePlacerService` - Places individual sleeves
- `OpeningCommandOrchestrator` - Coordinates the process

**What Happens:**
1. **Load ClashZones:** From XML files created during refresh
2. **Flag Check:** 
   - If `IsClustered = true` → Check if cluster sleeve exists at cluster placement point
     - If cluster sleeve exists → SKIP
     - If cluster sleeve NOT found → Reset both `IsClusterResolved = false` and `IsResolved = false`, proceed
   - If `IsResolved = true` → Check if individual sleeve exists at placement point
     - If sleeve exists → SKIP
     - If sleeve NOT found → Reset `IsResolved = false` and proceed
3. **Place Sleeves:** Using `IntersectionPoint` (linked coordinates)
4. **Update Flags:** Set `IsResolved = true`, save `SleeveInstanceId`
5. **Update Coordinates:** Save actual placed coordinates to `SleevePlacementPoint`

**Output:** Individual sleeves placed, flags updated, coordinates saved

### 4.4 Phase 3: Cluster Sleeve Placement

**Purpose:** Group proximate sleeves and place cluster sleeves

**UI Button:** "Place Sleeves" (cluster processing happens after individual placement)

**Services Involved:**
- `UniversalClusterService` - Main cluster orchestrator (performs clustering directly)
- `ClusterBoundingBoxServices` - Calculates cluster bounding boxes from actual sleeves
- `ClusterConfigurationManager` - Manages join distance settings
- `FlagManager` - Updates flags in Global XML after cluster placement

**What Happens:**
1. **Load ClashZones:** From XML files with `IsResolved = true` and `IsClusterResolved = false`
2. **Cluster Flag Check FIRST:** Filter sleeves with `IsClusterResolved = true` (skip already clustered)
3. **Group Sleeves:** 
   - `UniversalClusterService.FormClusters()` groups sleeves by host type, system type, and orientation
   - Uses `HostOrientation` from XML for walls/framing, `MepElementOrientationDirection` for floors
4. **Calculate Clusters:** Using bounding box overlap from XML
   - Reads bounding box coordinates from XML (saved during individual sleeve placement)
   - Uses `BoundingBoxesOverlapFromXml()` to check overlap with tolerance
   - No Revit API calls needed - all data from XML
5. **Place Cluster Sleeves:** 
   - Convert XML cluster to actual Revit `FamilyInstance` objects
   - `ClusterBoundingBoxServices.GetClusterBoundingBox()` calculates cluster dimensions using Revit API
   - Returns: `(width, height, depth, mid)` for cluster sleeve placement
6. **Set Parameters:** Apply correct width/height/depth based on host type and orientation
7. **Delete Individual Sleeves:** Remove individual sleeves within clusters
8. **Update Flags:** Set `IsClusterResolved = true`, save `ClusterSleeveInstanceId`, clear `SleeveInstanceId`

**Output:** Cluster sleeves placed, individual sleeves deleted, flags updated

---

## 4.5. Refresh Service Architecture - Three Paths System

### 4.5.1 Overview

The refresh process uses a **refactored service** (`RefreshServiceRefactored`) that implements a **three-path execution system** based on the `IsFilterComboNew` flag in the `FileCombos` database table. The legacy `RefreshService.cs` has been **completely removed** from the codebase.

### 4.5.2 Service Architecture

**Refactored Service:**
- **Location**: `refresh refactor/refresh_service_refactored.cs`
- **Factory**: `RefreshServiceFactory.Create()` - Always returns `RefreshServiceRefactoredWrapper`
- **Access**: All refresh operations go through the refactored service
- **Database-First**: All operations prioritize database over XML

**Legacy Service (Removed):**
- ❌ `Services/RefreshService.cs` - **EXCLUDED from compilation**
- ❌ All references removed
- ❌ No longer accessible in codebase

### 4.5.3 FileCombo Flag Decision System

The `IsFilterComboNew` flag in the `FileCombos` table is the **primary decision mechanism** for path selection:

**Database Query:**
```sql
SELECT IsFilterComboNew 
FROM FileCombos 
WHERE FilterId = @FilterId
```

**Flag Values:**
- `IsFilterComboNew = 0`: File combo already processed → **PATH 1**
- `IsFilterComboNew = 1`: New file combo needs processing → **PATH 2** (if Adopt OFF) or **PATH 3** (if Adopt ON)

**Flag Reset:**
- After successful sleeve placement: `IsFilterComboNew` reset to `0` via `OpeningCommandOrchestrator.ResetFilterComboFlagAfterPlacement()`
- Next refresh will use PATH 1 (Replay Mode)

### 4.5.4 Three Paths Detailed Flow

#### PATH 1: REPLAY MODE

**Trigger Conditions:**
- `IsFilterComboNew = 0` (from FileCombos table in DB)
- File combos already processed in database

**Execution Flow:**
1. Load existing clash zones from Database (PRIMARY source)
2. Check if sleeves exist (validation)
3. If sleeve NOT present → Place sleeve using saved data from DB
4. ResetFlagsForDeletedSleeves - Update flags in DB for deleted sleeves
5. Skip MergeAndSave - No database save of new zones
6. Go directly to Place Sleeves - Only for deleted sleeves

**Database Operations:**
- **READ**: `SELECT * FROM ClashZones WHERE FilterId = @FilterId`
- **UPDATE**: `UPDATE ClashZones SET IsResolvedFlag = 0 WHERE SleeveInstanceId = @SleeveId AND Sleeve NOT EXISTS`
- **NO INSERT**: No new clash zones created

**Key Characteristics:**
- ✅ **No merge required** - Uses existing zones as-is
- ✅ **No sync required** - Only flag updates
- ✅ **Fastest path** - Minimal processing
- ✅ **Adopt setting affects behavior** - If ON, may need validation checks

#### PATH 2: FRESH PLACEMENT MODE

**Trigger Conditions:**
- `IsFilterComboNew = 1` (from FileCombos table in DB - new file combos)
- **Adopt to Modified Document setting is IRRELEVANT** - Does NOT affect PATH 2 behavior

**Execution Flow:**
1. Run intersection detection (find all MEP vs Structural intersections)
2. **SKIP 3-Point Validation** - No validation required (fresh placement)
3. Create clash zones from new intersections
4. MergeAndSave to Database - Save all clash zones to DB
5. Database save sequence:
   - Get/Create Filter (DB)
   - Get/Create FileCombo (DB, IsFilterComboNew=1)
   - Insert/Update ClashZones (DB)
   - Insert/Update SleeveSnapshots (DB)
   - Commit Transaction

**Database Operations:**
- **INSERT**: `INSERT INTO FileCombos (FilterId, LinkedFileKey, HostFileKey, IsFilterComboNew=1)`
- **INSERT/UPDATE**: `INSERT INTO ClashZones (...) ON CONFLICT UPDATE ...`
- **INSERT/UPDATE**: `INSERT INTO SleeveSnapshots (...) ON CONFLICT UPDATE ...`

**Key Characteristics:**
- ✅ **No merge required** - Fresh detection only, no existing zones to merge
- ✅ **No validation** - Assumes all zones are valid (fresh placement)
- ✅ **Fast path** - No validation overhead
- ✅ **Adopt setting irrelevant** - Does not affect behavior

#### PATH 3: FULL DETECTION WITH VALIDATION

**Trigger Conditions:**
- `enableThreePointValidation = true` (Adopt to Modified Document = ON)
- **Adopt setting MUST be ON** to trigger PATH 3

**Execution Flow:**
1. Run intersection detection (find all MEP vs Structural intersections)
2. **3-Point Validation ENABLED** - Validate existing zones from DB:
   - MEP Element exists
   - Structural Element exists
   - Elements still intersect
3. **Process three zone categories**:
   - **Validated zones** (all 3 points pass):
     - Go to PATH 1 logic
     - Check sleeve presence
     - Place if missing using saved data from DB
     - **NO merge required**
   - **Invalidated zones** (any point fails):
     - **MERGE REQUIRED**
     - Update intersection points
     - Handle moved elements
     - Merge with new zones
   - **New zones** (not in existing zones):
     - Go to PATH 2 logic
     - Save to DB (duplicate data)
     - **NO merge required**
4. MergeAndSave to Database - Save validated/invalidated/new zones to DB
5. Database save sequence (same as PATH 2)

**Database Operations:**
- **READ**: `SELECT * FROM ClashZones WHERE FilterId = @FilterId` (for validation)
- **INSERT**: `INSERT INTO FileCombos (FilterId, LinkedFileKey, HostFileKey, IsFilterComboNew=1)`
- **INSERT/UPDATE**: `INSERT INTO ClashZones (...) ON CONFLICT UPDATE ...` (for all zone types)
- **UPDATE**: `UPDATE ClashZones SET IntersectionX/Y/Z = @NewCoords WHERE ClashZoneId = @Id` (for invalidated zones)
- **INSERT/UPDATE**: `INSERT INTO SleeveSnapshots (...) ON CONFLICT UPDATE ...`

**Key Characteristics:**
- ✅ **Merge required ONLY for invalidated zones**
- ✅ **Validated zones use PATH 1** - No merge, just check sleeve presence
- ✅ **New zones use PATH 2** - No merge, direct save
- ✅ **Most thorough path** - Validates all existing zones
- ✅ **Adopt setting MUST be ON** - Required to trigger PATH 3

### 4.5.5 Database-Driven Operations

**All operations prioritize database over XML:**

**Read Operations:**
- ✅ Clash zones loaded from `ClashZones` table (PRIMARY)
- ✅ FileCombo flags read from `FileCombos` table (PRIMARY)
- ✅ Filter metadata read from `Filters` table (PRIMARY)
- ✅ XML used only as fallback if database has no data

**Write Operations:**
- ✅ All clash zones saved to `ClashZones` table (PRIMARY)
- ✅ All flags updated in `ClashZones` table (PRIMARY)
- ✅ FileCombo flags updated in `FileCombos` table (PRIMARY)
- ✅ XML writes disabled when `DeploymentConfiguration.DisableXmlCreation = true`

**Transaction Management:**
- ✅ All database operations wrapped in SQLite transactions
- ✅ Atomic commits ensure data consistency
- ✅ Rollback on any error prevents partial data

### 4.5.6 Path Selection Summary

| Path | Trigger | Validation | Merge | Database Save | Use Case |
|------|---------|------------|-------|--------------|----------|
| **PATH 1** | IsFilterComboNew = 0 (DB) | Sleeve existence check only | ❌ No | ❌ No | Replay existing zones |
| **PATH 2** | IsFilterComboNew = 1 (DB) AND enableThreePointValidation = false | ❌ Skip | ❌ No | ✅ Yes | Fresh placement, no validation |
| **PATH 3** | enableThreePointValidation = true | ✅ Yes (3-point) | ✅ Yes (invalidated only) | ✅ Yes | Full detection with validation |

**Key Distinctions:**
- **PATH 1 vs PATH 2**: PATH 1 skips detection, PATH 2 runs detection
- **PATH 2 vs PATH 3**: PATH 2 has no validation, PATH 3 has validation
- **PATH 3 Merge**: Only invalidated zones merge, validated zones use PATH 1, new zones use PATH 2

---

## 5. Data Architecture

### 5.1 Global XML Structure

**Purpose:** Filter-independent flag tracking and skip decision

**File Naming:** `{category}_global.xml` (e.g., `Pipes_global.xml`)

**Hierarchical Schema Structure (Tree Organization):**

```
CategoryGlobalIndex (Category)
  └─ Filter (FilterName)
      └─ FileCombo (LinkedFile + HostFile + IsProcessed)
          └─ Entry (ClashZone data: GUID, flags, MEP+Host+Point)
```

**Tree Structure Benefits:**
- ✅ **Better Organization**: Clash zones grouped by Filter → FileCombo → ClashZones
- ✅ **Clear Hierarchy**: Easy to navigate and understand data relationships
- ✅ **Efficient Lookup**: Can search within specific filter/file combo context
- ✅ **Data Isolation**: Each filter's data is clearly separated

**Contains:**
- ClashZone GUID (`Id`) - Internal identifier
- `IsResolved` (individual sleeve flag)
- `IsClusterResolved` (cluster sleeve flag)
- `SleeveInstanceId` (individual sleeve ElementId)
- `ClusterSleeveInstanceId` (cluster sleeve ElementId)
- ✅ **MEP+Host+Point Data** (for O(1) matching by intersection point)
  - `MepElementId` (MEP element ID)
  - `StructuralElementId` (Structural element ID)
  - `IntersectionPointX/Y/Z` (Intersection point coordinates)
- `FilterName` - ⚠️ DEPRECATED: Filter name stored in FilterGroup.Name (kept for backward compatibility)

**What it answers:**
- ✅ **"Is this clash zone already resolved?"**
- ✅ **"Should I skip placement for this clash zone?"**
- ✅ **"Does this intersection point already have a sleeve?"** (cross-filter matching)

**Matching Strategy:**
- ✅ **O(1) Dictionary lookup** by `(MEP+Host+Point)` key
- ✅ **No GUID parameter needed** on sleeves (matching uses MEP+Host+Point instead)
- ✅ **Cross-filter matching** works without needing GUID parameter on sleeves

**What it does NOT contain:**
- ❌ Placement coordinates (only intersection point for matching)
- ❌ Family names
- ❌ Sizing parameters
- ❌ Host orientation
- ❌ Bounding boxes
- ❌ Other placement-specific data

### 5.2 Filter XML Structure

**Purpose:** Complete placement data storage (filter-specific)

**File Naming:** `{filter}_{category}.xml` (e.g., `Ventilation_ducts.xml`)

**Contains:**
- Full ClashZone object with ALL placement data:
  - `IntersectionPoint` coordinates (where to place)
  - `SleevePlacementPointActiveDocument` coordinates
  - `SleeveFamilyName` (which family to use)
  - `MepElementWidth/Height` (sizing parameters)
  - `HostOrientation` (placement direction)
  - `StructuralElementIdValue` (host element ID)
  - Bounding box coordinates (for clustering)
  - MEP/Structural element IDs
  - All other ClashZone properties

**What it answers:**
- ✅ **"Where should I place the sleeve?"**
- ✅ **"What family should I use?"**
- ✅ **"What size should it be?"**
- ✅ **"What host information do I need?"**

**Append-Only Principle:**
- ✅ New clash zones are ADDED to existing Filter XML
- ❌ Existing clash zones are NEVER removed or overwritten on rerun
- ⚠️ EXCEPTION: Only remove zones if "Adopt Document" is enabled AND 3-point validation FAILED

**Naming & Timing Guardrails:**
- ✅ **Normalize the base filter name before saving.** Always pass `Plumbing` (base) into `ClashZonePersistenceService.SaveClashZones`. The helper adds the category suffix internally (`Plumbing_pipes`). Feeding `Plumbing_pipes` as the base produces `Plumbing_pipes_pipes`, leaving the real branch untouched.
- ✅ **Persist before clearing Revit API objects.** `ClearRevitApiObjects()` must run only *after* every persistence call (individual placement, coordinate regeneration, clustering). Clearing first zeroes out placement points and bounding boxes, so the next run falls back to host centroids even though placement succeeded.
- ✅ **Clone post-placement state.** The clones you give to the persistence service must contain the post-placement sleeve IDs, placement points, and bounding boxes. If you clone before those fields are populated, the XML files stay empty even though sleeves exist in Revit.

### 5.3 Why Both Are Needed

**Cannot use only Global XML because:**
- Global XML has NO placement data (coordinates, family, size, etc.)
- Global XML only tells you IF to place, not HOW or WHERE

**Cannot use only Filter XML because:**
- Filter XML has no filter-independent flag tracking (**flags removed from Filter XML**)
- Without Global XML, you can't check if a clash zone was already resolved in another filter
- Would require expensive Revit API calls to check if sleeves exist
- Flags are synced FROM Global XML TO in-memory objects during refresh

**Placement Flow:**
1. Load clash zones from Filter XML → Get ALL placement data
2. Sync flags FROM Global XML TO in-memory objects via `FlagManager.SyncFlagsFromGlobal()`
3. Check Global XML by MEP+Host+Point (O(1) Dictionary lookup) → "Should I skip?"
4. Use Filter XML data → Get coordinates, family, size, host info
5. Place sleeve using Filter XML data
6. Update Global XML ONLY → Set flags, save sleeve IDs

---

## 6. Coordinate System Architecture

### Dual Coordinate System

**Refresh Phase:**
- `IntersectionPoint` = Linked document coordinates (for individual placement)
- `SleevePlacementPointActiveDocument` = Active document coordinates (for clustering)

**Individual Placement:**
- Uses `IntersectionPoint` (linked coordinates)
- Updates `SleevePlacementPoint` with actual placed coordinates

**Cluster Processing:**
- Uses `SleevePlacementPointActiveDocument` (active coordinates from refresh)
- Calculates proximity using active coordinate system

---

## 7. Flag Management System

### 7.1 Global XML as Single Source of Truth

**Key Principle:** **One Source of Truth** - Global XML exclusively manages all flag data.

**What Changed:**
- ✅ **Global XML** (`{category}_global.xml`) is now the **ONLY** source for flag management
- ❌ **Filter XML** (`{filter}_{category}.xml`) **NO LONGER** stores flags
- ✅ Flags are synced FROM Global XML TO in-memory ClashZone objects during refresh
- ✅ All flag updates are written TO Global XML ONLY

**Flags Managed in Global XML:**
- `IsResolved` - Individual sleeve placement status
- `IsClusterResolved` - Cluster sleeve placement status  
- `SleeveInstanceId` - Individual sleeve Revit ElementId
- `ClusterSleeveInstanceId` - Cluster sleeve Revit ElementId
- `MepElementId` + `StructuralElementId` + `IntersectionPoint` - 3-point matching data

**Flags NOT Stored in Filter XML:**
- `IsResolved` - ❌ Removed (read from Global XML)
- `IsClusterResolved` - ❌ Removed (read from Global XML)
- `SleeveInstanceId` - ❌ Removed (read from Global XML)
- `ClusterSleeveInstanceId` - ❌ Removed (read from Global XML)

**Metadata Still Stored in Filter XML:**
- `MarkedForClusteringSleeveProcess` - ✅ Kept (metadata, not a flag)
- `IsCurrentClash` - ✅ Kept (metadata, not a flag)
- Sleeve dimensions (Width, Height, Diameter) - ✅ Kept
- Sleeve placement points - ✅ Kept
- Bounding box coordinates - ✅ Kept

**Benefits:**
- ✅ **Single Source of Truth** - No conflicting flag values between files
- ✅ **Cross-Filter Consistency** - Flags work across all filters automatically
- ✅ **Simplified Maintenance** - Update flags in one place only
- ✅ **Performance** - No redundant flag writes to Filter XML
- ✅ **Data Integrity** - Eliminates flag synchronization issues

### 7.2 Flag Hierarchy

1. **`MarkedForClusteringSleeveProcess`** - CLEAR FLAG for clustering decisions
   - `true` = Sleeve is proximate to other sleeves and should be clustered
   - `false` = Sleeve should remain individual (not proximate)
   - `null` = Not yet processed for clustering

2. **`IsClusterResolved`** - Cluster sleeve exists
   - `true` = Cluster sleeve placed
   - `false` = No cluster sleeve

3. **`IsResolved`** - Individual sleeve exists
   - `true` = Individual sleeve placed
   - `false` = No individual sleeve

### 7.3 Flag Logic

**Individual Sleeve Placement:**
```
if (MarkedForClusteringSleeveProcess == true) → SKIP (sleeve should be clustered, not individual)
if (IsResolved == true) → Check if sleeve exists at placement point
  - If sleeve exists → SKIP
  - If sleeve NOT found → Reset IsResolved = false, proceed with placement
else → PROCESS (place individual sleeve)
```

**Cluster Processing:**
```
if (MarkedForClusteringSleeveProcess == true) → PROCESS (sleeve is proximate, should be clustered)
if (MarkedForClusteringSleeveProcess == false) → SKIP (sleeve is not proximate, keep individual)
if (MarkedForClusteringSleeveProcess == null) → PROCESS (not yet processed, check proximity)
```

### 7.4 Edge Case: Individual Sleeves in Cluster Zones

**Problem:**
When cluster sleeves are placed, individual sleeves within the cluster are deleted. However, there may be individual sleeves that fall within the cluster zone but weren't part of the original cluster formation. When these "edge case" sleeves are deleted, we need to set `IsClusterResolved=true` for their entries.

**Solution:**
During cluster placement, when deleting individual sleeves:
1. Check for other individual sleeves that fall within the cluster bounding box
2. These sleeves weren't part of the original cluster formation but fall in the cluster zone
3. When deleting these edge case sleeves, set `IsClusterResolved=true` for their entries

**Implementation:**
- **Location:** `Services/UniversalClusterService.cs` - `PlaceClusterSleeve` method
- **Process:**
  1. After deleting cluster formation sleeves, check for other individual sleeves within cluster bounding box
  2. For each edge case sleeve found:
     - Delete the sleeve
     - Find its ClashZone entry (by SleeveInstanceId or GUID)
     - **DATABASE FIRST:** Update database flags using `ClashZoneRepository.BatchUpdateFlags()` or `UpdateFlags()`
       - Set `IsClusterResolvedFlag=true`, `ClusterSleeveInstanceId=clusterSleeveId`, `SleeveInstanceId=-1`
     - **XML SECOND:** Then update Global XML via `FlagManager.UpdateFlagsForPlacement()` or `GlobalIndexService.UpsertFlagsWithIdsAndClashZoneData()`
       - Set `IsClusterResolved=true`, `ClusterSleeveInstanceId=clusterSleeveId`, `SleeveInstanceId=-1`
  3. **Priority:** Database is always updated first as the primary source of truth (XML will be discarded in future)

**Expected Behavior:**
1. **Normal cluster deletion:** Sleeves that formed the cluster are deleted and marked as `IsClusterResolved=true` (existing behavior)
   - Database updated first, then Global XML
2. **Edge case deletion:** Individual sleeves that fall within cluster zone but weren't part of cluster formation are also deleted and marked as `IsClusterResolved=true` (new behavior)
   - Database updated first, then Global XML

**Key Principle:**
- Database is the primary source of truth for flags
- All flag updates must update database FIRST, then sync to Global XML
- This ensures consistency and prepares for future XML removal

---

## 17. GUID Management System

### 17.1 Deterministic GUID Generation

**Purpose:** Ensure stable, consistent GUIDs for clash zones across multiple detection runs, even when Global XML is deleted/recreated.

**Implementation:**
- **Algorithm**: MD5 hash of stable identifiers
- **Input Format**: `MEP_ID|HOST_ID|X|Y|Z` (coordinates rounded to 0.1ft tolerance)
- **Output**: 128-bit GUID (16 bytes)

**Code Location:**
- `Services/GuidManager.cs` - `GenerateDeterministicGuid()` method
- `Services/ClashZoneService.cs` - `CreateClashZone()` method (sets deterministic GUID)

**Key Features:**
1. **Deterministic**: Same intersection always produces same GUID
2. **Tolerance**: Uses 0.1ft (~30mm) tolerance for coordinate rounding
3. **Fallback**: Returns random GUID if inputs are invalid
4. **Standards-Compliant**: Follows industry best practices for clash identification

**Example:**
```csharp
// Same intersection always gets same GUID
Guid guid1 = GenerateDeterministicGuid(700223, 432131, 27.819, 51.634, -0.770);
Guid guid2 = GenerateDeterministicGuid(700223, 432131, 27.819, 51.634, -0.770);
// guid1 == guid2 (deterministic)
```

**Benefits:**
- ✅ **Stability**: GUID remains constant across detection runs
- ✅ **Recovery**: Works correctly even when Global XML is deleted
- ✅ **Consistency**: Recovery and detection create matching GUIDs
- ✅ **Standard Format**: GUIDs are XML-compatible and database-ready

### 17.2 GUID Storage and Recovery

**Storage Locations:**
1. **Global XML**: `CategoryGlobalIndexEntry.Id` (string GUID)
2. **Filter XML**: `ClashZone.Id` (GUID)
3. **Sleeve Parameter**: `ClashZone_GUID` text parameter on sleeve family instances

**Recovery Mechanism:**
- **Method**: `FlagManager.RecoverSleeveFlagsFromRevit()`
- **Process**:
  1. Scan all sleeves in Revit for given category
  2. Extract MEP ID, Host ID, and intersection point from sleeve
  3. Generate deterministic GUID using same algorithm as `CreateClashZone`
  4. Match to existing Global XML entries by MEP+Host+Point (not GUID)
  5. Update flags and sleeve IDs for matching entries
  6. Create new entries with deterministic GUIDs for unmatched sleeves

**⚠️ IMPORTANT: Recovery is Only for Edge Cases (November 6, 2025)**
- **Primary Path**: Intersection detection (`ClashZoneService.CreateClashZone`) now immediately updates Global XML when sleeves are found
  - Uses optimized spatial index lookup (O(1) complexity)
  - Pre-collects sleeves by category at START of detection (calculate once, use many times)
  - Updates Global XML immediately: `IsResolved=true`, `SleeveInstanceId` set
  - **Result**: Flags are set correctly during detection, eliminating need for recovery in normal flow
- **Recovery Purpose**: Only needed for edge cases:
  1. **Corrupted XML**: Global XML was manually edited or corrupted
  2. **Manual Sleeve Creation**: Sleeves were created outside the normal placement flow
  3. **XML Deletion**: Global XML was deleted but sleeves still exist in Revit
  4. **Data Migration**: Migrating from old system without proper flag initialization
- **Performance**: Recovery uses same optimization pattern as detection:
  - Pre-collects sleeves by category (filtered by `MEP_Category` parameter)
  - Builds HashSet for O(1) lookup
  - Minimizes Revit API calls
- **Architecture**: Recovery is now redundant for normal operation but kept for edge case handling

**Critical Fix (November 6, 2025):**
- **Tolerance Mismatch**: Fixed recovery tolerance from 0.01ft to 0.1ft
- **Reason**: Must match `FindByMepHostAndPoint` default tolerance
- **Impact**: Recovery now correctly matches entries to sleeves

**Sleeve Metadata Parameters:**
- **LinkedFile**: Text parameter storing MEP element's linked file name
- **HostFile**: Text parameter storing host element's file name
- **Purpose**: Enables recovery to determine file combo without expensive Revit API calls
- **Set During**: `UniversalSleevePlacerService.SetSleeveMetadata()`

**Matching Strategy:**
- **Primary**: MEP+Host+Point (used for all matching operations)
- **GUID Role**: Provides stable identifier, not used for matching
- **Recovery**: Uses deterministic GUID generation to ensure consistency

**Benefits:**
- ✅ **Fast Recovery**: Uses sleeve parameters instead of expensive API calls
- ✅ **Deterministic**: Same sleeve always gets same GUID in recovery
- ✅ **Consistent**: Recovery and detection create matching entries
- ✅ **Robust**: Works even when Global XML is missing

---

## 10. Clustering Implementation

### 10.4 Rotated Non-Axis-Aligned Clustering

**Date Implemented:** November 2025  
**Status:** ✅ Complete - Fixed for rotated floor sleeves

#### Overview

This section documents the methodology for clustering sleeves that are rotated at non-axis-aligned angles (e.g., -45°, 135°, 225°). The algorithm uses a **corner-based watertight approach** that works for all scenarios: single sleeves, stacked sleeves, inline sleeves, diagonal arrangements, and grid patterns.

#### Problem Statement

**Original Issue:**
- Cluster sleeves with rotated individual sleeves (e.g., -45° and 135°) were incorrectly sized
- Averaging rotation angles (e.g., -45° and 135° averaging to 225°) produced incorrect results
- Bounding box calculations based on sleeve centers or simple unions failed for complex arrangements
- Cluster dimensions were incorrect (e.g., expected 550mm × 400mm but got 530mm × 730mm)

**Root Cause:**
- Previous approach used sleeve centers and offsets, which fails for diagonal/grid arrangements
- Rotation angle averaging was incorrect for sleeves on the same axis but 180° apart
- Bounding box calculation did not account for all corner points after transformation

#### Solution: Corner-Based Watertight Algorithm

**Key Principle:**
> **"Dump Once, Use Many Times"** - Pre-calculate and store sleeve corner coordinates and rotation matrix components in the database during individual sleeve placement, then reuse them during clustering.

#### Algorithm Steps

**Phase 1: Individual Sleeve Placement (Pre-calculation)**

During individual sleeve placement (`UniversalSleevePlacerService`), the following data is calculated and stored in the database:

1. **Sleeve Center (Active Document Coordinates)**
   - Uses `SleevePlacementPointActiveDocumentX/Y/Z` (where sleeve is actually placed)
   - Saved to database via `UpdateSleevePlacement()` method

2. **Rotation Matrix Components**
   - Pre-calculates `cos(rotationAngle)` and `sin(rotationAngle)`
   - Stored in database columns: `MepRotationCos` and `MepRotationSin`
   - Avoids redundant trigonometric calculations during clustering

3. **Four Corner Coordinates (World Space)**
   - Calculates all 4 corners of each sleeve in world coordinates
   - Corner order: 1=Bottom-left, 2=Bottom-right, 3=Top-left, 4=Top-right
   - Stored in database columns: `SleeveCorner1X/Y/Z` through `SleeveCorner4X/Y/Z`

**Corner Calculation Process:**
```csharp
// Step 1: Calculate corners in local coordinate system (before rotation)
double halfW = sleeveWidth / 2.0;
double halfH = sleeveHeight / 2.0;
var localCorners = new[]
{
    new XYZ(-halfW, -halfH, 0),  // Corner 1: Bottom-left
    new XYZ(halfW, -halfH, 0),   // Corner 2: Bottom-right
    new XYZ(-halfW, halfH, 0),   // Corner 3: Top-left
    new XYZ(halfW, halfH, 0)     // Corner 4: Top-right
};

// Step 2: Rotate corners by sleeve rotation angle to get world-space corners
double cosSleeve = Math.Cos(rotationAngleRad);
double sinSleeve = Math.Sin(rotationAngleRad);
for (int j = 0; j < 4; j++)
{
    double localX = localCorners[j].X;
    double localY = localCorners[j].Y;
    
    // Rotate corner by sleeve rotation matrix
    double worldX = localX * cosSleeve - localY * sinSleeve;
    double worldY = localX * sinSleeve + localY * cosSleeve;
    
    // Translate to sleeve center (placement point)
    worldCorners[j] = new XYZ(
        sleeveCenter.X + worldX,
        sleeveCenter.Y + worldY,
        sleeveCenter.Z
    );
}

// Step 3: Save world-space corners to database
repository.UpdateSleeveCorners(
    zone.Id,
    worldCorners[0].X, worldCorners[0].Y, worldCorners[0].Z,  // Corner 1
    worldCorners[1].X, worldCorners[1].Y, worldCorners[1].Z,  // Corner 2
    worldCorners[2].X, worldCorners[2].Y, worldCorners[2].Z,  // Corner 3
    worldCorners[3].X, worldCorners[3].Y, worldCorners[3].Z   // Corner 4
);
```

**Phase 2: Cluster Bounding Box Calculation**

During clustering (`UniversalClusterService.GetClusterBoundingBoxWithRotatedCoordinates`), the algorithm:

1. **Determines Cluster's Intended Rotated Axis**
   - Uses the **first sleeve's actual rotation angle** (from Revit `LocationPoint.Rotation`)
   - Falls back to `ClashZone.MepElementRotationAngle` if `LocationPoint` is unavailable
   - ⚠️ **Critical:** This is NOT an average of sleeve angles - it's the intended axis direction for the cluster

2. **Loads Pre-calculated Data**
   - Retrieves sleeve centers from `SleevePlacementPointActiveDocumentX/Y/Z`
   - Retrieves pre-calculated corners from `SleeveCorner1X/Y/Z` through `SleeveCorner4X/Y/Z`
   - Retrieves rotation matrix components from `MepRotationCos` and `MepRotationSin`
   - Falls back to recalculation if pre-calculated data is missing

3. **Transforms All Corners to Cluster's Rotated Coordinate System**
   - For each sleeve, uses pre-calculated world-space corners (or recalculates if missing)
   - Transforms each corner into the cluster's intended rotated axis coordinate system
   - This aligns all corners to a common rotated frame for min/max calculation

4. **Finds Min/Max Extents**
   - Finds `minX`, `maxX`, `minY`, `maxY` across all transformed corners
   - Calculates cluster width = `maxX - minX`
   - Calculates cluster height = `maxY - minY`

5. **Calculates Cluster Center**
   - Cluster center in rotated coordinate system: `((minX + maxX) / 2, (minY + maxY) / 2)`
   - Transforms center back to world coordinates using inverse rotation matrix (transpose)

**Transformation Process:**
```csharp
// Step 1: Choose reference point (first sleeve center as origin)
XYZ origin = firstSleeveCenter;

// Step 2: Pre-calculate cluster rotation matrix components
// rotationAngle is the cluster's INTENDED ROTATED AXIS, not average of sleeve angles
double cosCluster = Math.Cos(rotationAngle);
double sinCluster = Math.Sin(rotationAngle);

// Step 3: For each sleeve, transform all 4 corners
foreach (var sleeve in cluster)
{
    // Use pre-calculated corners from database (or recalculate if missing)
    XYZ[] worldCorners = GetPreCalculatedCorners(sleeve) ?? RecalculateCorners(sleeve);
    
    // Transform each corner to cluster's rotated coordinate system
    for (int j = 0; j < 4; j++)
    {
        // Translate relative to origin
        double relX = worldCorners[j].X - origin.X;
        double relY = worldCorners[j].Y - origin.Y;
        
        // Rotate to cluster's intended axis coordinate system
        double clusterX = relX * cosCluster - relY * sinCluster;
        double clusterY = relX * sinCluster + relY * cosCluster;
        
        allTransformedCorners.Add(new XYZ(clusterX, clusterY, worldCorners[j].Z));
    }
}

// Step 4: Find min/max extents across all transformed corners
double minX = allTransformedCorners.Min(c => c.X);
double minY = allTransformedCorners.Min(c => c.Y);
double maxX = allTransformedCorners.Max(c => c.X);
double maxY = allTransformedCorners.Max(c => c.Y);

// Step 5: Calculate cluster dimensions
double width = maxX - minX;
double height = maxY - minY;

// Step 6: Calculate cluster center in rotated coordinate system
double centerX_rotated = (minX + maxX) / 2.0;
double centerY_rotated = (minY + maxY) / 2.0;

// Step 7: Transform center back to world coordinates (inverse rotation = transpose)
double centerX_world = origin.X + centerX_rotated * cosCluster + centerY_rotated * sinCluster;
double centerY_world = origin.Y - centerX_rotated * sinCluster + centerY_rotated * cosCluster;
XYZ clusterCenter = new XYZ(centerX_world, centerY_world, origin.Z);
```

#### Database Schema

**New Columns Added to `ClashZones` Table:**

| Column Name | Type | Description |
|------------|------|-------------|
| `SleevePlacementActiveX` | REAL | Active document X coordinate of sleeve center |
| `SleevePlacementActiveY` | REAL | Active document Y coordinate of sleeve center |
| `SleevePlacementActiveZ` | REAL | Active document Z coordinate of sleeve center |
| `SleeveCorner1X` | REAL | World X coordinate of corner 1 (Bottom-left) |
| `SleeveCorner1Y` | REAL | World Y coordinate of corner 1 (Bottom-left) |
| `SleeveCorner1Z` | REAL | World Z coordinate of corner 1 (Bottom-left) |
| `SleeveCorner2X` | REAL | World X coordinate of corner 2 (Bottom-right) |
| `SleeveCorner2Y` | REAL | World Y coordinate of corner 2 (Bottom-right) |
| `SleeveCorner2Z` | REAL | World Z coordinate of corner 2 (Bottom-right) |
| `SleeveCorner3X` | REAL | World X coordinate of corner 3 (Top-left) |
| `SleeveCorner3Y` | REAL | World Y coordinate of corner 3 (Top-left) |
| `SleeveCorner3Z` | REAL | World Z coordinate of corner 3 (Top-left) |
| `SleeveCorner4X` | REAL | World X coordinate of corner 4 (Top-right) |
| `SleeveCorner4Y` | REAL | World Y coordinate of corner 4 (Top-right) |
| `SleeveCorner4Z` | REAL | World Z coordinate of corner 4 (Top-right) |
| `MepRotationCos` | REAL | Pre-calculated cos(rotationAngle) |
| `MepRotationSin` | REAL | Pre-calculated sin(rotationAngle) |

#### Key Implementation Details

**1. Cluster Rotation Angle Determination**

The cluster's intended rotated axis is determined by:
1. **First Priority:** First sleeve's actual Revit `LocationPoint.Rotation` (if available)
2. **Second Priority:** `ClashZone.MepElementRotationAngle` from database
3. **Third Priority:** `DetermineDominantRotationAngle()` (only if angle is still 0 or unavailable)

⚠️ **Important:** The cluster rotation angle is NOT an average of sleeve angles. It represents the intended axis direction for the cluster coordinate system.

**2. Cluster Sleeve Placement**

- Cluster sleeve is placed **axis-aligned** (0° rotation in Revit)
- Cluster dimensions already account for rotation (calculated in rotated coordinate system)
- Cluster sleeve does NOT need to be rotated - its width/height reflect the rotated bounding box

**3. Performance Optimization**

- **Pre-calculation:** Corner coordinates and rotation matrix components are calculated once during individual sleeve placement
- **Database Storage:** Pre-calculated values stored in database for fast retrieval during clustering
- **Fallback:** If pre-calculated data is missing, corners are recalculated on-the-fly

#### Supported Scenarios

This algorithm correctly handles:

✅ **Single Sleeve:** One rotated sleeve  
✅ **Stacked Sleeves:** Multiple sleeves stacked vertically  
✅ **Inline Sleeves:** Multiple sleeves arranged horizontally  
✅ **Diagonal Arrangements:** Sleeves at various angles  
✅ **Grid Patterns:** Complex 2D grid arrangements  
✅ **Mixed Rotations:** Sleeves with different rotation angles in same cluster

#### Code Locations

**Individual Sleeve Placement (Pre-calculation):**
- `Services/UniversalSleevePlacerService.cs` - Lines 1797-1979
- `Data/Repositories/ClashZoneRepository.cs` - `UpdateSleevePlacement()` and `UpdateSleeveCorners()` methods

**Cluster Bounding Box Calculation:**
- `Services/UniversalClusterService.cs` - `GetClusterBoundingBoxWithRotatedCoordinates()` method (Lines 2719-3100)

**Database Schema:**
- `Data/SleeveDbContext.cs` - `ClashZones` table schema (Lines 366-394)
- `Data/Entities/SleeveClashZone.cs` - Entity model properties
- `Models/ClashZone.cs` - Model properties

#### Benefits

✅ **Watertight Algorithm:** Works for all scenarios (single, stacked, inline, diagonal, grid)  
✅ **Performance:** Pre-calculated data avoids redundant computations  
✅ **Accuracy:** Corner-based approach ensures correct bounding box for any arrangement  
✅ **Maintainability:** Clear separation between pre-calculation and clustering phases  
✅ **Robustness:** Fallback logic handles missing pre-calculated data gracefully

---

## 18. Critical Bug Fixes

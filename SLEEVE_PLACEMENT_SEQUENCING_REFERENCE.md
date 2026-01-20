# Sleeve Placement Sequencing Reference

## Overview

This document details the complete sequencing of sleeve placement operations, including individual sleeve placement, clustering, and coordinate updates. The dominant sequencing is critical for proper bounding box capture and cleanup.

---

## Complete Sequencing Flow

### Phase 1: Individual Sleeve Placement

1. **Place Individual Sleeves**
   - Location: `Services/OpeningCommandOrchestrator.cs` line 512
   - Method: `ExecuteUniversalSleevePlacement()`
   - Action: Places all individual sleeves via `UniversalSleevePlacementCommand`

2. **Pause (300ms)**
   - Location: `Services/OpeningCommandOrchestrator.cs` line 526
   - Purpose: Allow document state to stabilize after placement

3. **Regenerate Document**
   - Location: `Services/OpeningCommandOrchestrator.cs` line 532
   - Method: `_document.Regenerate()`
   - Purpose: Ensure bounding boxes are available for individual sleeves

4. **Pause (300ms)**
   - Location: `Services/OpeningCommandOrchestrator.cs` line 534
   - Purpose: Wait for regeneration to complete

5. **Read Individual Sleeve Data from Revit and Save to XML**
   - Location: `Services/OpeningCommandOrchestrator.cs` lines 540-541
   - Class: `SleeveCoordinateService`
   - Method: `UpdateSleeveCoordinatesInXml(xmlFilePath)`
   - Purpose: Get bounding box data from Revit and persist to XML

### Phase 2: Clustering

6. **Trigger Clustering Command**
   - Location: `Services/OpeningCommandOrchestrator.cs` line 231
   - Method: `ExecuteCommandSequence()` calls `ExecuteClusteringForCategory()`
   - Entry: `Services/OpeningCommandOrchestrator.cs` line 237

7. **Clustering: Place Cluster Sleeves + Stage 1 Cleanup**
   - Location: `Services/OpeningCommandOrchestrator.cs` line 275
   - Class: `UniversalClusterService`
   - Method: `ClusterSleeves()`
   - Actions:
     - Find individual sleeves within proximity (default 200mm)
     - Place cluster sleeve to encompass them
     - **Delete individual sleeves that formed the cluster** (Stage 1 cleanup)
     - Set metadata on cluster sleeve (MEP_Category, Filter Name)

8. **Pause (200ms)**
   - Location: `Services/OpeningCommandOrchestrator.cs` line 294
   - Purpose: Allow document state to stabilize after clustering

9. **Regenerate for Cluster Sleeves**
   - Location: `Services/OpeningCommandOrchestrator.cs` line 291
   - Method: `_document.Regenerate()`
   - Purpose: Ensure bounding boxes are available for cluster sleeves

10. **Read Cluster Sleeve Data from Revit and Save to XML**
    - Location: `Services/OpeningCommandOrchestrator.cs` line 306
    - Class: `SleeveCoordinateService`
    - Method: `UpdateSleeveCoordinatesInXml(xmlFilePath)`
    - Purpose: Get cluster sleeve bounding box data from Revit and persist to XML

### Phase 3: Stage 2 Cleanup

11. **Stage 2 Cleanup: Delete Individual Sleeves Within Cluster Bounding Boxes**
    - Location: `Services/OpeningCommandOrchestrator.cs` line 325
    - Class: `UniversalClusterService`
    - Method: `CleanupSleevesWithinClustersAfterXmlSave()`
    - Purpose: Delete individual sleeves that fall within cluster sleeve bounding boxes but were not part of the original cluster formation

---

## Code Flow Diagram

```
SleevePlacementExternalEvent
  │
  ├─> Execute()
  │     │
  │     └─> OpeningCommandOrchestrator.ExecuteMultipleFilters()
  │           │
  │           └─> ExecuteCommandSequence()
  │                 │
  │                 ├─> ExecuteUniversalSleevePlacement()
  │                 │     │
  │                 │     ├─> UniversalSleevePlacementCommand.Execute()
  │                 │     │     │
  │                 │     │     └─> UniversalSleevePlacerService.PlaceAllSleevesInTransaction()
  │                 │     │           │
  │                 │     │           ├─> Place sleeves in Revit
  │                 │     │           └─> SaveUpdatedXmlFiles() [Saves SleeveInstanceId]
  │                 │     │
  │                 │     ├─> Sleep 300ms
  │                 │     ├─> _document.Regenerate()
  │                 │     ├─> Sleep 300ms
  │                 │     └─> SleeveCoordinateService.UpdateSleeveCoordinatesInXml()
  │                 │           [Reads individual sleeve bounding boxes from Revit]
  │                 │
  │                 └─> ExecuteClusteringForCategory()
  │                       │
  │                       ├─> UniversalClusterService.ClusterSleeves()
  │                       │     │
  │                       │     ├─> Collect individual sleeves from Revit
  │                       │     ├─> Calculate proximity (default 200mm)
  │                       │     ├─> Place cluster sleeves
  │                       │     └─> DELETE individual sleeves (Stage 1 cleanup)
  │                       │
  │                       ├─> Sleep 200ms
  │                       ├─> _document.Regenerate()
  │                       ├─> SleeveCoordinateService.UpdateSleeveCoordinatesInXml()
  │                       │     [Reads cluster sleeve bounding boxes from Revit]
  │                       │
  │                       └─> UniversalClusterService.CleanupSleevesWithinClustersAfterXmlSave()
  │                             │
  │ Soul └─> LoadClashZoneCacheForCleanup() [Reloads XML with cluster bboxes]
  │                             └─> Delete individual sleeves within cluster bboxes (Stage 2)
```

---

## Key Classes and Methods

### 1. SleevePlacementExternalEvent
**File:** `Services/SleevePlacementExternalEvent.cs`

**Purpose:** External event handler that provides Revit API context

**Key Method:**
- `Execute(UIApplication app)` - Entry point for sleeve placement

---

### 2. OpeningCommandOrchestrator
**File:** `Services/OpeningCommandOrchestrator.cs`

**Purpose:** Orchestrates the entire sleeve placement flow

**Key Methods:**

#### `ExecuteMultipleFilters(List<OpeningFilter> filters)`
- Line: 51
- Orchestrates execution of placement and clustering for all filters

#### `ExecuteCommandSequence()`
- Line: 215
- Defines the order of operations:
  - ExecuteUniversalSleevePlacement (line 226)
  - ExecuteClusteringForCategory (line 231)

#### `ExecuteUniversalSleevePlacement(OpeningFilter filter)`
- Line: 458
- **Places individual sleeves**
- Lines 512-541: Runs placement, then saves individual sleeve coordinates
- Critical steps:
  - Line 512: Execute placement command
  - Line 526: Sleep 300ms
  - Line 532: Regenerate document
  - Line 534: Sleep 300ms
  - Line 541: Update coordinates in XML

#### `ExecuteClusteringForCategory(OpeningFilter filter)`
- Line: 237
- **Handles clustering and Stage 2 cleanup**
- Critical steps:
  - Line 275: Cluster sleeves (includes Stage 1 cleanup)
  - Line 291: Regenerate document
  - Line 294: Sleep 200ms
  - Line 306: Update cluster coordinates in XML
  - Line 325: Stage 2 cleanup

---

### 3. UniversalSleevePlacerService
**File:** `Services/UniversalSleevePlacerService.cs`

**Purpose:** Handles the actual placement of individual sleeves

**Key Method:**

#### `PlaceAllSleevesInTransaction(List<ClashZone> clashZones)`
- Line: 211
- Places individual sleeves in Revit
- Line 872: Calls `SaveUpdatedXmlFiles()` to save SleeveInstanceId values

---

### 4. UniversalSleevePlacementCommand
**File:** `Commands/UniversalSleevePlacementCommand.cs`

**Purpose:** Command wrapper that executes placement in a transaction

**Key Method:**
- `Execute(ExternalCommandData)` - Manages transaction for placement

---

### 5. UniversalClusterService
**File:** `Services/UniversalClusterService.cs`

**Purpose:** Handles clustering logic and cleanup

**Key Methods:**

#### `ClusterSleeves(Document doc, string category, ...)`
- Line: 62
- **Stage 1: Places cluster sleeves and deletes forming individual sleeves**
- Collects individual sleeves from Revit
- Calculates proximity using XML data
- Places cluster sleeves
- Deletes individual sleeves that formed clusters

#### `CleanupSleevesWithinClustersAfterXmlSave(Document doc, List<FamilyInstance> placedClusters)`
- Line: 1902
- **Stage 2: Deletes individual sleeves within cluster bounding boxes**
- Uses XML cached data for bounding boxes (no expensive Revit API calls)
- Compares individual sleeve positions to cluster bounding boxes
- Deletes sleeves that fall within cluster bounds

---

### 6. SleeveCoordinateService
**File:** `Services/SleeveCoordinateService.cs`

**Purpose:** Updates sleeve coordinates and bounding boxes in XML

**Key Method:**

#### `UpdateSleeveCoordinatesInXml(string xmlFilePath)`
- Line: 82
- Reads all sleeves from Revit (individual and cluster)
- Gets bounding boxes for each sleeve
- Matches sleeves to ClashZones in XML
- Updates ClashZone with bounding box coordinates
- Saves updated XML

**Details:**
- Line 87-91: Collects all sleeves from Revit
- Line 93: Creates SleeveCoordinateUpdater
- Line 96: Loads clash zones from XML file
- Line 99: Updates coordinates via updater
- Line 102: Saves updated XML

---

## Data Model: ClashZone

**File:** `Models/ClashZone.cs`

**Key Properties:**

### Individual Sleeve Data
- `SleeveInstanceId` (int): Revit element ID of individual sleeve (-1 if deleted/clustered)
- `SleeveBoundingBoxMinX/Y/Z` (double): Min bounding box coordinates
- `SleeveBoundingBoxMaxX/Y/Z` (double): Max bounding box coordinates
- `SleevePlacementPointX/Y/Z` (double): Placement point coordinates

### Cluster Sleeve Data
- `ClusterSleeveInstanceId` (int): Revit element ID of cluster sleeve
- `ClusterSleeveBoundingBoxMinX/Y/Z` (double): Cluster min bounding box
- `ClusterSleeveBoundingBoxMaxX/Y/Z` (double): Cluster max bounding box

### Flags
- `IsResolved` (bool): True if individual sleeve is placed
- `IsClusterResolved` (bool): True if cluster sleeve is placed

---

## XML File Structure

**Location:** `C:\Users\<username>\AppData\Roaming\JSE_MEP_Openings\Projects\Default\Filters\`

**File Naming:** `{filterName}_{category}.xml`

**Example:** `Ventilation_duct_accessories.xml`

**Key XML Elements:**
```xml
<ClashZone>
  <Id>{guid}</Id>
  
  <!-- Individual sleeve data -->
  <SleeveInstanceId>912452</SleeveInstanceId>
  <SleeveBoundingBoxMinX>0.742000</SleeveBoundingBoxMinX>
  <SleeveBoundingBoxMinY>22.128000</SleeveBoundingBoxMinY>
  <SleeveBoundingBoxMinZ>7.787000</SleeveBoundingBoxMinZ>
  <SleeveBoundingBoxMaxX>1.342000</SleeveBoundingBoxMaxX>
  <SleeveBoundingBoxMaxY>22.728000</SleeveBoundingBoxMaxY>
  <SleeveBoundingBoxMaxZ>8.387000</SleeveBoundingBoxMaxZ>
  
  <!-- Cluster sleeve data -->
  <ClusterSleeveInstanceId>-1</ClusterSleeveInstanceId>
  <ClusterSleeveBoundingBoxMinX>0.000000</ClusterSleeveBoundingBoxMinX>
  <ClusterSleeveBoundingBoxMinY>0.000000</ClusterSleeveBoundingBoxMinY>
  <ClusterSleeveBoundingBoxMinZ>0.000000</ClusterSleeveBoundingBoxMinZ>
  <ClusterSleeveBoundingBoxMaxX>0.000000</ClusterSleeveBoundingBoxMaxX>
  <ClusterSleeveBoundingBoxMaxY>0.000000</ClusterSleeveBoundingBoxMaxY>
  <ClusterSleeveBoundingBoxMaxZ>0.000000</ClusterSleeveBoundingBoxMaxZ>
  
  <!-- Flags -->
  <IsResolved>true</IsResolved>
  <IsClusterResolved>false</IsClusterResolved>
</ClashZone>
```

---

## Critical Timing Requirements

### Why Regeneration is Needed

1. **After Individual Sleeve Placement:**
   - Sleeves are created in Revit but bounding boxes may not be immediately available
   - `Regenerate()` ensures geometry is updated
   - Wait 300ms to allow regeneration to complete

2. **After Cluster Sleeve Placement:**
   - Cluster sleeves are created with dimensions that encompass individual sleeves
   - Bounding boxes need regeneration to reflect actual cluster size
   - Wait 200ms to allow regeneration to complete

### Why Individual Coordinates Must Be Saved Before Clustering

- Individual sleeves are **deleted during clustering** (Stage 1 cleanup)
- If coordinates are not saved before clustering, they will be lost forever
- Stage 2 cleanup needs individual sleeve coordinates to determine which sleeves are within cluster bounds

---

## Stage 1 vs Stage 2 Cleanup

### Stage 1 Cleanup (During Clustering)
**Location:** `Services/UniversalClusterService.cs`, `ClusterSleeves()` method
- Deletes individual sleeves that **actively form** a cluster
- Happens immediately after cluster sleeve placement
- Uses proximity calculation (200mm default)
- **Individual sleeves deleted here have SleeveInstanceId set to -1**

### Stage 2 Cleanup (After XML Save)
**Location:** `Services/UniversalClusterService.cs`, `CleanupSleevesWithinClustersAfterXmlSave()` method
- Deletes individual sleeves that fall **within cluster bounding box** but were not part of original cluster
- Happens after cluster coordinates are saved to XML
- Uses XML cached bounding box data (no expensive Revit API calls)
- Compares individual sleeve positions to cluster bounds

---

## Performance Optimizations

### 1. XML-Based Bounding Box Caching
- Bounding boxes are read from Revit once and cached in XML
- Stage 2 cleanup uses XML data instead of expensive Revit API calls
- Avoids querying `element.get_BoundingBox()` for every sleeve/cluster pair

### 2. Dynamic XML File Path
- No hardcoded file names or paths
- Constructed from filter name and category: `{filterName}_{category}.xml`
- Filter names and categories are passed from UI

### 3. Selective XML Loading
- Only the relevant XML file for the current category is loaded
- Avoids loading all 22 XML files during clustering

---

## Error Handling

### Document Modification Errors
- `_document.Regenerate()` is wrapped in try-catch (line 296-299)
- If regeneration fails, coordinate updates proceed anyway

### Missing XML Files
- SleeveCoordinateService validates file existence before processing
- Logs warnings but continues execution

### Missing Sleeves in Revit
- During coordinate update, if a sleeve ID from XML is not found in Revit, it's logged and skipped
- Allows process to continue with valid sleeves

---

## Debugging Logs

### Key Log Files

1. **orchestrator_debug.log**
   - Location: `Log\orchestrator_debug.log`
   - Contains sequencing events and timing information

2. **coordinate_update.log**
   - Location: `Log\coordinate_update.log`
   - Contains bounding box read/write operations
   - Shows which sleeves are found and updated

3. **placement_debug.log**
   - Location: `Log\placement_debug.log`
   - Contains individual sleeve placement details

---

## Summary

The sequencing is designed to:
1. Place individual sleeves and capture their bounding boxes before they are deleted
2. Perform clustering with Stage 1 cleanup (delete forming sleeves)
3. Capture cluster sleeve bounding boxes after regeneration
4. Perform Stage 2 cleanup using cached XML data (no expensive Revit API calls)

This two-stage cleanup ensures all sleeves within cluster bounds are deleted while maintaining optimal performance through XML caching.




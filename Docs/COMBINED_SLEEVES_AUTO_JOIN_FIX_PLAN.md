# Implementation Plan - Fix Auto Join for Combined Sleeves

## Problem Statement
The "Auto Join" feature fails to correctly group and place combined sleeves because:
1.  **Selection:** It uses imprecise "Point-in-Box" logic and misses `ClusterSleeves` data.
2.  **Proximity:** It uses simple AABB checks instead of the robust spatial indexing available in `CrossCategoryProximityService`.
3.  **Placement:** It relies on incomplete code that lacks `JoinGeometry` and robust flag management.

---

## Section 1: Selection (Discovery & Data Loading)
**Goal:** Accurately retrieve all relevant sleeves (Individual & Cluster) located within the 3D Section Box.

### 1.1 Data Model Update ✅ COMPLETE
#### [MODIFY] [ClusterSleeveInfo.cs](file:///C:/JSE_CSharp_Projects/JSE_MEPOPENING_23/Services/Clustering/Combined/Phase1And2/Models/ClusterSleeveInfo.cs)
| Task | Detail |
|---|---|
| **Add Properties** | ✅ Added `Corner1X`, `Corner1Y`, `Corner1Z` ... `Corner4Z` (12 `double` properties). |
| **Update `FromClashZone`** | ✅ Mapped directly from `ClashZone.SleeveCorner1X`, etc. |

### 1.2 Discovery Service Enhancement (Revit-First Filtering)
#### [MODIFY] [CombinedClusterDiscoveryService.cs](file:///C:/JSE_CSharp_Projects/JSE_MEPOPENING_23/Services/Clustering/Combined/Phase1And2/Services/CombinedClusterDiscoveryService.cs)

**Flow (Exactly Like ParameterService):**
1. **Get Section Box Bounds** → `SectionBoxHelper.GetSectionBoxBounds(view3D)` returns `BoundingBoxXYZ`.
2. **Revit-Level Filter** → Create `BoundingBoxIntersectsFilter(outline)` and apply to `FilteredElementCollector` for sleeve family instances.
3. **Collect Element IDs** → Get `List<int>` of sleeve `ElementId.IntegerValue` that pass the filter.
4. **Query DB by IDs** → Load only those specific sleeves from `ClashZones` and `ClusterSleeves` tables using the collected Revit IDs.
5. **Category Filter** → Further filter in-memory by user-selected categories (`MepElementCategory` / `Category`).

| Step | Code Reuse / Implementation |
|---|---|
| **Step 1: Get Section Box** | **REUSE:** `SectionBoxHelper.GetSectionBoxBounds(view3D)` in [SectionBoxHelper.cs](file:///C:/JSE_CSharp_Projects/JSE_MEPOPENING_23/Helpers/SectionBoxHelper.cs). |
| **Step 2: Create Filter** | **REUSE:** `new Outline(bounds.Min, bounds.Max)` + `new BoundingBoxIntersectsFilter(outline)`. Pattern at [ParameterTransferService.cs:L3056-3058](file:///C:/JSE_CSharp_Projects/JSE_MEPOPENING_23/Services/ParameterTransferService.cs#L3056). |
| **Step 3: Collect Sleeve IDs** | `FilteredElementCollector(doc).OfClass(typeof(FamilyInstance)).WherePasses(sectionBoxFilter)` → Loop, get `elem.Id.IntegerValue` for sleeves matching our family names. |
| **Step 4a: Load Individuals (by IDs + UI Categories)** | SQL: `SELECT * FROM ClashZones WHERE SleeveInstanceId IN (...) AND MepElementCategory IN ('Ducts', 'Pipes', ...)` |
| **Step 4b: Load Clusters (by IDs + UI Categories)** | SQL: `SELECT * FROM ClusterSleeves WHERE ClusterInstanceId IN (...) AND Category IN ('Ducts', 'Pipes', ...)` |

> [!NOTE]
> The UI-selected categories (`userSelectedCategories`) are passed from the ViewModel to the Discovery Service, and applied **at the DB query level** - NOT in-memory filter. This avoids loading unwanted sleeves.

---

## Section 2: Proximity Check (Formation)
**Goal:** Group sleeves using precise geometric proximity - REUSE existing cluster proximity code.

### 2.1 Existing Proximity Logic (Working Code to Reuse)
The following code **already works** for cluster sleeves and should be extended for Combined Sleeves:

| Component | Location | Description |
|---|---|---|
| **`UnifiedSleeve.IsWithinProximity`** | [UnifiedSleeve.cs:L117](file:///C:/JSE_CSharp_Projects/JSE_MEPOPENING_23/Services/Combined/Models/UnifiedSleeve.cs#L117) | Returns `true` if `GetDistanceTo(other) <= threshold`. |
| **`UnifiedSleeve.GetDistanceTo`** | [UnifiedSleeve.cs:L107](file:///C:/JSE_CSharp_Projects/JSE_MEPOPENING_23/Services/Combined/Models/UnifiedSleeve.cs#L107) | Calculates center-to-center distance. |
| **`CrossCategoryProximityService.DetectProximityGroups`** | [CrossCategoryProximityService.cs:L31](file:///C:/JSE_CSharp_Projects/JSE_MEPOPENING_23/Services/Combined/CrossCategoryProximityService.cs#L31) | Uses `SimplifiedSpatialIndex` + Union-Find to group sleeves. |
| **`SimplifiedSpatialIndex`** | [Spatial/SimplifiedSpatialIndex.cs](file:///C:/JSE_CSharp_Projects/JSE_MEPOPENING_23/Services/Combined/Spatial/SimplifiedSpatialIndex.cs) | O(N log N) spatial indexing for fast proximity queries. |

### 2.2 Extension for Combined Sleeves
#### [MODIFY] [CombinedClusterFormationService.cs](file:///C:/JSE_CSharp_Projects/JSE_MEPOPENING_23/Services/Clustering/Combined/Phase1And2/Services/CombinedClusterFormationService.cs)
| Step | Code Reuse / Implementation |
|---|---|
| **Convert to UnifiedSleeve** | Convert `ClusterSleeveInfo` → `UnifiedSleeve` using `UnifiedSleeve.FromClusterSleeve()` (already exists). |
| **Inject Proximity Service** | Inject `ICrossCategoryProximityService` via DI. |
| **Call Existing Method** | `var groups = _proximityService.DetectProximityGroups(unifiedSleeves, toleranceThreshold);` |
| **Map Results** | Convert `List<ProximityGroup>` → `List<CombinedClusterCandidate>`. |

**Why Reuse `CrossCategoryProximityService`?**
- It uses **Spatial Indexing** (O(N log N)) via `SimplifiedSpatialIndex`.
- It uses **Union-Find** for grouping.
- It already filters for **Cross-Category** pairs.

### 2.3 How Cluster-to-Individual Proximity Works
```
┌─────────────────┐         ┌─────────────────┐
│   ClashZone     │         │ ClusterSleeves  │
│  (Individual)   │         │   (Cluster)     │
└────────┬────────┘         └────────┬────────┘
         │                           │
         ▼                           ▼
┌─────────────────────────────────────────────┐
│           UnifiedSleeve (Abstraction)       │
│  - Id, Type, Category, BoundingBox          │
│  - PlacementPoint, Corners                  │
│  - GetCenter() → XYZ                        │
│  - GetDistanceTo(other) → double            │
│  - IsWithinProximity(other, tol) → bool     │
└─────────────────────────────────────────────┘
         │
         ▼
┌─────────────────────────────────────────────┐
│  CrossCategoryProximityService              │
│  DetectProximityGroups(List<UnifiedSleeve>) │
│  → Returns List<ProximityGroup>             │
└─────────────────────────────────────────────┘
```
**Answer:** Since both Individual (`ClashZone`) and Cluster (`ClusterSleeves`) are converted to `UnifiedSleeve`, the proximity check `IsWithinProximity(other, threshold)` works **identically** regardless of whether the pair is:
- Individual ↔ Individual
- Individual ↔ Cluster
- Cluster ↔ Cluster

### 2.4 Required Update: Populate ClusterSleeve Corners ✅ COMPLETE
#### [MODIFY] [UnifiedSleeve.cs](file:///C:/JSE_CSharp_Projects/JSE_MEPOPENING_23/Services/Combined/Models/UnifiedSleeve.cs#L187)
| Issue | Current Code | Fix |
|---|---|---|
| **Missing Corners** | ~~`FromClusterSleeve` creates `corners = new List<XYZ>()` (empty).~~ | ✅ Added corner properties to `ClusterSleeveData`, updated `FromClusterSleeve()` to populate corners. |

---

## Section 3: Placement & Persistence (Execution)
**Goal:** Place the new Combined Sleeve in Revit, join it to the host, and update all database flags.

### 3.1 Placement Service Enhancement
#### [MODIFY] [CombinedSleevePlacementService.cs](file:///C:/JSE_CSharp_Projects/JSE_MEPOPENING_23/Services/Combined/CombinedSleevePlacementService.cs)
| Step | Code Reuse / Implementation |
|---|---|
| **Revit Family Placement** | **REUSE:** Pattern from `CombinedSleeveViewModel.CreateCombinedSleeve(...)` at [CombinedSleeveViewModel.cs:L500+](file:///C:/JSE_CSharp_Projects/JSE_MEPOPENING_23/UI/CombinedSleeveViewModel.cs#L500). Copy the `doc.Create.NewFamilyInstance(...)` logic. |
| **Auto-Join Geometry** | **NEW:** Add `JoinGeometryUtils.JoinGeometry(doc, host, combinedSleeveInstance)` after placement. Wrap in `try/catch` to log warnings. |
| **Database Insert** | **REUSE:** `ICombinedSleeveRepository.SaveCombinedSleeve(CombinedSleeve)` at [CombinedSleeveRepository.cs:L30](file:///C:/JSE_CSharp_Projects/JSE_MEPOPENING_23/Data/Repositories/CombinedSleeveRepository.cs#L30). This already inserts into `CombinedSleeves` and `CombinedSleeveConstituents`. |
| **Flag Management** | **REUSE:** `ICombinedSleeveRepository.MarkConstituentsAsResolved(List<SleeveConstituent>)` at [CombinedSleeveRepository.cs:L416](file:///C:/JSE_CSharp_Projects/JSE_MEPOPENING_23/Data/Repositories/CombinedSleeveRepository.cs#L416). This sets `IsCombinedResolved = 1` on `ClashZones` (Individual) and `ClusterSleeves` (Cluster). |
| **Batch Mode** | Wrap the loop in a single `using (var transaction = new Transaction(...))` for Revit, and a single `BeginTransaction()` for SQLite. |

---

## Summary of Code Reuse

| Section | Reused Component | File | Method/Function |
|---|---|---|---|
| **1. Selection** | Section Box Filtering | `EfficientIntersectionService.cs` | `IsPointInBoundingBox` (L379), `IsElementInSectionBoxAndIntersectsMep` (L520) |
| **1. Selection** | Load Individuals | `ClashZoneRepository.cs` | `GetClashZonesByCategory` (L315) |
| **1. Selection** | Load Clusters | `ClashZoneRepository.cs` | `GetClusterSleevesByInstanceIds` (L7065) |
| **2. Proximity** | Spatial Indexing & Grouping | `CrossCategoryProximityService.cs` | `DetectProximityGroups` (L31), `SimplifiedSpatialIndex` |
| **3. Placement** | Revit Creation Pattern | `CombinedSleeveViewModel.cs` | `CreateCombinedSleeve` (L500+) |
| **3. Placement** | Database Insert | `CombinedSleeveRepository.cs` | `SaveCombinedSleeve` (L30) |
| **3. Placement** | Flag Reset | `CombinedSleeveRepository.cs` | `MarkConstituentsAsResolved` (L416) |

---

## Section 4: Wiring & Integration
**Goal:** Connect all services using existing SOLID architecture and DI patterns.

### 4.1 Architecture Diagram
```
┌───────────────────────────────────────────────────────────────────────────────┐
│                         CombinedSleeveViewModel                                │
│  (UI Layer - Orchestrates the workflow, receives user-selected categories)    │
└───────────────────┬───────────────────────────────────────────────────────────┘
                    │
    ┌───────────────┼───────────────────────────┐
    │               │                           │
    ▼               ▼                           ▼
┌───────────────┐  ┌───────────────────────┐  ┌─────────────────────────────────┐
│ Discovery     │  │ Formation Service     │  │ Placement Service               │
│ Service       │  │                       │  │                                 │
│ (Section 1)   │  │ (Section 2)           │  │ (Section 3)                     │
└───────┬───────┘  └───────────┬───────────┘  └────────────────┬────────────────┘
        │                      │                               │
        ▼                      ▼                               ▼
┌──────────────────────────────────────────────────────────────────────────────┐
│                         IClashZoneRepository                                  │
│                         ICombinedSleeveRepository                             │
│                         ICrossCategoryProximityService                        │
│                         ISleeveCornerCalculationService                       │
└──────────────────────────────────────────────────────────────────────────────┘
```

### 4.2 Existing DI Constructor Patterns (Already SOLID)
| Service | Location | Constructor Signature |
|---|---|---|
| **CombinedSleevePlacementService** | [L24-34](file:///C:/JSE_CSharp_Projects/JSE_MEPOPENING_23/Services/Combined/CombinedSleevePlacementService.cs#L24) | `(Document, ICombinedSleeveRepository, ICrossCategoryProximityService, ISleeveCornerCalculationService)` |
| **CombinedSleeveViewModel** | [L70+](file:///C:/JSE_CSharp_Projects/JSE_MEPOPENING_23/UI/CombinedSleeveViewModel.cs#L70) | Receives `_discoveryService`, `_formationService` (add `_placementService`) |

### 4.3 Required Wiring Changes
#### [MODIFY] [CombinedSleeveViewModel.cs](file:///C:/JSE_CSharp_Projects/JSE_MEPOPENING_23/UI/CombinedSleeveViewModel.cs)
| Task | Detail |
|---|---|
| **Inject PlacementService** | Add `private readonly CombinedSleevePlacementService _placementService;` to fields and constructor. |
| **Replace Private Method** | In `CreateCombinedSleeves()` (L172), replace call to local `CreateCombinedSleeve(candidate)` with `_placementService.PlaceCombinedSleeves(individuals, clusters, comboId, filterId)`. |
| **Pass UI Categories** | Pass `userSelectedCategories` to `_discoveryService.Discover(...)` (already done at L204-208). |

### 4.4 ViewModel Integration Flow
```csharp
// In CombinedSleeveViewModel.CreateCombinedSleeves()
_requestHandler.SetAction(async (uiapp) =>
{
    // 1. Get UI-selected categories (EXISTING CODE at L204-208)
    var categories = GetSelectedCategories(); 
    
    // 2. Get Section Box (EXISTING CODE at L222-227)
    var sectionBox = GetActiveSectionBox(doc);
    
    // 3. SECTION 1: Discover sleeves (NEW: pass categories to filter at DB level)
    var (individuals, clusters) = _discoveryService.DiscoverWithCategories(
        categories, sectionBox);
    
    // 4. SECTION 2 + 3: Place combined sleeves (REUSE existing service)
    //    CombinedSleevePlacementService handles:
    //    - ConvertToUnifiedSleeves (Section 2)
    //    - DetectProximityGroups (Section 2)
    //    - PlaceCombinedSleevesInRevit (Section 3)
    //    - SaveToDatabase + MarkConstituentsAsResolved (Section 3)
    var placedSleeves = _placementService.PlaceCombinedSleeves(
        individuals, clusters, comboId, filterId, proximityThreshold);
    
    StatusMessage = $"Placed {placedSleeves.Count} combined sleeves.";
});
```

### 4.5 Service Instantiation (Entry Point)
#### [MODIFY] Service Factory / Entry Point
Services are instantiated where `CombinedSleeveViewModel` is created (likely in `OpeningCommandOrchestrator.cs` or a factory class).

```csharp
// Example instantiation pattern (pseudo-code)
var dbContext = new SleeveDbContext(dbPath);
var clashZoneRepo = new ClashZoneRepository(dbContext, logger);
var combinedSleeveRepo = new CombinedSleeveRepository(dbContext, logger);
var cornerService = new SleeveCornerCalculationService();
var proximityService = new CrossCategoryProximityService(logger);

var placementService = new CombinedSleevePlacementService(
    doc, combinedSleeveRepo, proximityService, cornerService);

var discoveryService = new CombinedClusterDiscoveryService(
    doc, clashZoneRepo); // Inject repository for DB queries

var viewModel = new CombinedSleeveViewModel(
    doc, discoveryService, formationService, placementService, ...);
```

---

## Verification Plan
1.  **Selection:** Log discovered sleeve counts (Individual vs. Cluster) and verify only sleeves inside section box are picked.
2.  **Proximity:** Test with (a) Overlapping, (b) Close (< Tol), (c) Far (> Tol) sleeves.
3.  **Placement:** Verify `JoinGeometry` call logs success. Check Revit model for wall cut.
4.  **Flags:** Query DB: `SELECT COUNT(*) FROM ClashZones WHERE IsCombinedResolved = 1`.


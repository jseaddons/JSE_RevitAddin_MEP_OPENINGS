# Refresh Flow Diagram - GUID & Flag Management

## Overview
This document illustrates the complete refresh flow, showing how RefreshService, ClashZoneService, and GlobalIndexService interact to manage clash zone GUIDs and sleeve placement flags.

---

## 🔄 REFRESH FLOW - Complete Sequence

```
┌─────────────────────────────────────────────────────────────────────────────┐
│                         REFRESH CLICK TRIGGERED                               │
│                    (EmergencyMainDialog → RefreshService)                     │
└─────────────────────────────────────────────────────────────────────────────┘
                                      │
                                      ▼
┌─────────────────────────────────────────────────────────────────────────────┐
│ STEP 1: RefreshService.ExecuteRefreshInternal()                            │
│                                                                              │
│  ┌──────────────────────────────────────────────────────────────────────┐  │
│  │ 1.1 Get Resolved GUIDs from Global XML                                │  │
│  │    GlobalIndexService.GetResolvedGuids(category)                     │  │
│  │    → Returns: HashSet<Guid> resolved, HashSet<Guid> clusterResolved  │  │
│  │    → Purpose: Filter out already-resolved clash zones                │  │
│  │    → NO Revit API calls, just reads flags                            │  │
│  └──────────────────────────────────────────────────────────────────────┘  │
│                                      │                                       │
│                                      ▼                                       │
│  ┌──────────────────────────────────────────────────────────────────────┐  │
│  │ 1.2 Load Existing Clash Zones from Filter XML                        │  │
│  │    LoadExistingClashZonesFromFilterXml()                             │  │
│  │    → Deserializes {filter}_{category}.xml                             │  │
│  │    → Creates ClashZone objects (MepElementId may be null initially) │  │
│  └──────────────────────────────────────────────────────────────────────┘  │
│                                      │                                       │
│                                      ▼                                       │
│  ┌──────────────────────────────────────────────────────────────────────┐  │
│  │ 1.3 Sync Flags from Global XML to Filter XML                         │  │
│  │    For each ClashZone:                                               │  │
│  │      - Load GlobalIndexService.LoadOrCreate(category)                │  │
│  │      - Find entry by ClashZone.Id (GUID)                             │  │
│  │      - Update: IsResolved, IsClusterResolved, SleeveInstanceId       │  │
│  │      - Update: ClusterSleeveInstanceId                               │  │
│  │    → Ensures Filter XML has latest flag states                      │  │
│  └──────────────────────────────────────────────────────────────────────┘  │
│                                      │                                       │
│                                      ▼                                       │
│  ┌──────────────────────────────────────────────────────────────────────┐  │
│  │ 1.4 3-POINT VALIDATION (RefreshService)                              │  │
│  │    For each existing ClashZone:                                      │  │
│  │                                                                       │  │
│  │    POINT 1: MEP Element Exists?                                      │  │
│  │      GetElementFromDocumentOrLinked(MepElementId)                    │  │
│  │      ❌ NULL → Invalid (MEP deleted)                                │  │
│  │      ✅ EXISTS → Continue                                           │  │
│  │                                                                       │  │
│  │    POINT 2: Structural Element Exists?                               │  │
│  │      GetElementFromDocumentOrLinked(StructuralElementId)            │  │
│  │      ❌ NULL → Invalid (Host deleted)                               │  │
│  │      ✅ EXISTS → Continue                                           │  │
│  │                                                                       │  │
│  │    POINT 3: Elements Still Intersect?                                │  │
│  │      CalculateIntersectionPoint(mepElement, structuralElement)      │  │
│  │      ❌ NULL → Invalid (No longer intersect)                      │  │
│  │      ✅ EXISTS → Valid (Update intersection point if moved)        │  │
│  │                                                                       │  │
│  │    IF VALIDATION FAILS (Invalid):                                    │  │
│  │      → Add to invalidClashZones list                                │  │
│  │      → Remove from Global XML (RefreshService does this):           │  │
│  │         • Load GlobalIndexService.LoadOrCreate(category)            │  │
│  │         • Remove entry: globalIndex.Entries.RemoveAll(GUID match)   │  │
│  │         • Save: GlobalIndexService.Save(globalIndex)                │  │
│  │      → Remove from existingClashZones list                          │  │
│  │                                                                       │  │
│  │    IF VALIDATION PASSES (Valid):                                     │  │
│  │      → Add to validClashZones list                                   │  │
│  │      → Update IntersectionPoint if moved (>1mm)                     │  │
│  │      → Update SleevePlacementPointActiveDocument                    │  │
│  └──────────────────────────────────────────────────────────────────────┘  │
│                                      │                                       │
│                                      ▼                                       │
│  ┌──────────────────────────────────────────────────────────────────────┐  │
│  │ 1.5 Replace with Validated Clash Zones Only                          │  │
│  │    existingClashZones.ClashZones = validClashZones                   │  │
│  │    → Invalid ones removed, only valid clash zones remain             │  │
│  └──────────────────────────────────────────────────────────────────────┘  │
│                                      │                                       │
│                                      ▼                                       │
│  ┌──────────────────────────────────────────────────────────────────────┐  │
│  │ 1.6 Initialize ClashZoneService                                      │  │
│  │    _clashZoneService = new ClashZoneService(existingClashZones, ...)│  │
│  │    → Stores reference to existingClashZones as _clashZoneStorage     │  │
│  │    → This is used by DetectNewClashZones for duplicate checking     │  │
│  └──────────────────────────────────────────────────────────────────────┘  │
│                                      │                                       │
│                                      ▼                                       │
│  ┌──────────────────────────────────────────────────────────────────────┐  │
│  │ 1.7 Detect New Clash Zones (ClashZoneService.DetectNewClashZones)    │  │
│  │    For each current intersection:                                    │  │
│  │                                                                       │  │
│  │    STEP 1: Check if Already Exists                                   │  │
│  │      FindExistingClashZone(MepElementId, StructuralElementId)       │  │
│  │      → Checks _clashZoneStorage.ClashZones                          │  │
│  │      → Compares by IntegerValue (not ElementId objects)             │  │
│  │      ✅ FOUND → Skip (already exists)                               │  │
│  │      ❌ NOT FOUND → Continue to create new                          │  │
│  │                                                                       │  │
│  │    STEP 2: Create New ClashZone (if not found)                       │  │
│  │      new ClashZone {                                                 │  │
│  │        Id = Guid.NewGuid(),          // ← NEW GUID GENERATED        │  │
│  │        MepElementId = mepElement.Id,                                 │  │
│  │        StructuralElementId = structuralElement.Id,                  │  │
│  │        IsResolved = false,           // ← Default: unresolved        │  │
│  │        IsClusterResolved = false     // ← Default: unresolved        │  │
│  │      }                                                               │  │
│  │      → Add to newClashZones list                                     │  │
│  │                                                                       │  │
│  │    STEP 3: Ensure Global XML Entry Exists                           │  │
│  │      GlobalIndexService.EnsureEntries(category, [new GUIDs])        │  │
│  │      → Creates entry with IsResolved=false, IsClusterResolved=false  │  │
│  │      → SleeveInstanceId=0, ClusterSleeveInstanceId=0                 │  │
│  └──────────────────────────────────────────────────────────────────────┘  │
│                                      │                                       │
│                                      ▼                                       │
│  ┌──────────────────────────────────────────────────────────────────────┐  │
│  │ 1.8 Reset Flags for Deleted Sleeves                                  │  │
│  │    ClashZoneService.ResetResolvedFlagForDeletedSleeves()             │  │
│  │                                                                       │  │
│  │    For each ClashZone with IsResolved=true OR IsClusterResolved=true:│  │
│  │                                                                       │  │
│  │    STEP 1: Check Global XML First                                    │  │
│  │      GlobalIndexService.LoadOrCreate(category)                       │  │
│  │      → If Global XML says resolved → Trust it, skip Revit check     │  │
│  │                                                                       │  │
│  │    STEP 2: Check Cluster Sleeve (if IsClusterResolved=true)          │  │
│  │      document.GetElement(ClusterSleeveInstanceId)                   │  │
│  │      ✅ EXISTS → Keep flags true                                     │  │
│  │      ❌ NOT FOUND → Reset ALL flags:                                │  │
│  │         • IsClusterResolved = false                                  │  │
│  │         • IsResolved = false                                        │  │
│  │         • ClusterSleeveInstanceId = -1                              │  │
│  │         • SleeveInstanceId = -1                                     │  │
│  │                                                                       │  │
│  │    STEP 3: Check Individual Sleeve (if IsClusterResolved=false)     │  │
│  │      document.GetElement(SleeveInstanceId)                           │  │
│  │      ✅ EXISTS → Keep IsResolved=true                                │  │
│  │      ❌ NOT FOUND → Reset:                                          │  │
│  │         • IsResolved = false                                        │  │
│  │         • SleeveInstanceId = -1                                     │  │
│  │                                                                       │  │
│  │    STEP 4: Save Reset Flags to Both XMLs                             │  │
│  │      SaveResetFlagsToBothXmls()                                      │  │
│  │      → Updates Filter XML (_clashZoneStorage already modified)       │  │
│  │      → Updates Global XML via GlobalIndexService.UpsertFlagsWithIds()│  │
│  └──────────────────────────────────────────────────────────────────────┘  │
│                                      │                                       │
│                                      ▼                                       │
│  ┌──────────────────────────────────────────────────────────────────────┐  │
│  │ 1.9 Save All Clash Zones to Filter XML                               │  │
│  │    Save existingClashZones + newClashZones to {filter}_{category}.xml│  │
│  │    → Persists all clash zones (valid + new) to disk                  │  │
│  └──────────────────────────────────────────────────────────────────────┘  │
└─────────────────────────────────────────────────────────────────────────────┘
```

---

## 📊 GUID MANAGEMENT FLOW

```
┌─────────────────────────────────────────────────────────────────────────────┐
│                    GUID LIFECYCLE & VALIDATION                               │
└─────────────────────────────────────────────────────────────────────────────┘

┌─────────────────────────────────────────────────────────────────────────────┐
│ GENERATION                                                                   │
│ ──────────────────────────────────────────────────────────────────────────  │
│                                                                              │
│  When: New ClashZone created during DetectNewClashZones                     │
│  Where: ClashZoneService.DetectNewClashZones()                              │
│  Code: new ClashZone { Id = Guid.NewGuid() }                                │
│                                                                              │
│  ✅ GUID is UNIQUE and PERSISTENT                                           │
│  ✅ GUID does NOT change for the same clash zone                            │
│  ✅ GUID is used as key in Global XML (CategoryGlobalIndexEntry.Id)         │
└─────────────────────────────────────────────────────────────────────────────┘
                                      │
                                      ▼
┌─────────────────────────────────────────────────────────────────────────────┐
│ VALIDATION                                                                   │
│ ──────────────────────────────────────────────────────────────────────────  │
│                                                                              │
│  When: During Refresh (3-Point Validation)                                  │
│  Where: RefreshService.ExecuteRefreshInternal()                              │
│                                                                              │
│  Validation Criteria (All 3 Must Pass):                                     │
│  ┌─────────────────────────────────────────────────────────────────────┐   │
│  │ ✅ Point 1: MEP Element Exists                                      │   │
│  │    GetElementFromDocumentOrLinked(MepElementId) != null            │   │
│  └─────────────────────────────────────────────────────────────────────┘   │
│  ┌─────────────────────────────────────────────────────────────────────┐   │
│  │ ✅ Point 2: Structural Element Exists                                │   │
│  │    GetElementFromDocumentOrLinked(StructuralElementId) != null      │   │
│  └─────────────────────────────────────────────────────────────────────┘   │
│  ┌─────────────────────────────────────────────────────────────────────┐   │
│  │ ✅ Point 3: Elements Still Intersect                                 │   │
│  │    CalculateIntersectionPoint() != null                              │   │
│  └─────────────────────────────────────────────────────────────────────┘   │
│                                                                              │
│  IF VALIDATION FAILS:                                                        │
│    → GUID entry removed from Global XML                                     │
│    → ClashZone removed from Filter XML                                      │
│    → GUID becomes INVALID (no longer exists)                                │
│                                                                              │
│  IF VALIDATION PASSES:                                                       │
│    → GUID entry kept in Global XML                                          │
│    → ClashZone kept in Filter XML                                           │
│    → GUID remains VALID                                                     │
└─────────────────────────────────────────────────────────────────────────────┘
```

---

## 🏷️ FLAG MANAGEMENT FLOW

```
┌─────────────────────────────────────────────────────────────────────────────┐
│                      FLAG STATES & TRANSITIONS                               │
└─────────────────────────────────────────────────────────────────────────────┘

┌─────────────────────────────────────────────────────────────────────────────┐
│ INITIAL STATE (New ClashZone)                                                │
│ ──────────────────────────────────────────────────────────────────────────  │
│  IsResolved = false                                                          │
│  IsClusterResolved = false                                                   │
│  SleeveInstanceId = -1 (or 0)                                               │
│  ClusterSleeveInstanceId = -1 (or 0)                                         │
│                                                                              │
│  Global XML Entry: Created by EnsureEntries()                              │
│    Entry {                                                                   │
│      Id = clashZone.Id (GUID),                                              │
│      IsResolved = false,                                                    │
│      IsClusterResolved = false,                                              │
│      SleeveInstanceId = 0,                                                  │
│      ClusterSleeveInstanceId = 0                                             │
│    }                                                                         │
└─────────────────────────────────────────────────────────────────────────────┘
                                      │
                                      ▼
┌─────────────────────────────────────────────────────────────────────────────┐
│ INDIVIDUAL SLEEVE PLACED                                                    │
│ ──────────────────────────────────────────────────────────────────────────  │
│  Where: UniversalSleevePlacerService.PlaceSleeve()                          │
│                                                                              │
│  Filter XML Update:                                                         │
│    clashZone.IsResolved = true                                              │
│    clashZone.SleeveInstanceId = sleeveId.IntegerValue                       │
│                                                                              │
│  Global XML Update:                                                         │
│    GlobalIndexService.UpsertFlagsWithIds(category, [                        │
│      (clashZone.Id, true, false, sleeveId, 0)                               │
│    ])                                                                        │
│                                                                              │
│  Result State:                                                              │
│    IsResolved = true                                                        │
│    IsClusterResolved = false                                                 │
│    SleeveInstanceId = 12345 (valid ID)                                      │
│    ClusterSleeveInstanceId = 0                                              │
└─────────────────────────────────────────────────────────────────────────────┘
                                      │
                                      ▼
┌─────────────────────────────────────────────────────────────────────────────┐
│ CLUSTER SLEEVE PLACED (Individual Deleted)                                 │
│ ──────────────────────────────────────────────────────────────────────────  │
│  Where: UniversalClusterService.PlaceClusterSleeve()                        │
│                                                                              │
│  Filter XML Update:                                                         │
│    clashZone.IsClusterResolved = true                                        │
│    clashZone.IsResolved = true              // ← STAYS TRUE                 │
│    clashZone.SleeveInstanceId = -1          // ← Individual deleted        │
│    clashZone.ClusterSleeveInstanceId = clusterId.IntegerValue               │
│                                                                              │
│  Global XML Update:                                                         │
│    GlobalIndexService.UpsertFlagsWithIds(category, [                        │
│      (clashZone.Id, true, true, -1, clusterId)  // ← Both flags true        │
│    ])                                                                        │
│                                                                              │
│  Result State:                                                              │
│    IsResolved = true                        // ← Individual was placed      │
│    IsClusterResolved = true                  // ← Then deleted, cluster    │
│    SleeveInstanceId = -1                     // ← Individual deleted        │
│    ClusterSleeveInstanceId = 67890 (valid ID)                               │
└─────────────────────────────────────────────────────────────────────────────┘
                                      │
                                      ▼
┌─────────────────────────────────────────────────────────────────────────────┐
│ REFRESH: Flag Reset Check                                                   │
│ ──────────────────────────────────────────────────────────────────────────  │
│  Where: ClashZoneService.ResetResolvedFlagForDeletedSleeves()              │
│                                                                              │
│  Flow:                                                                       │
│  ┌─────────────────────────────────────────────────────────────────────┐   │
│  │ STEP 1: Check Global XML                                            │   │
│  │    If Global XML says resolved → Trust it, skip Revit check        │   │
│  └─────────────────────────────────────────────────────────────────────┘   │
│  ┌─────────────────────────────────────────────────────────────────────┐   │
│  │ STEP 2: Check Cluster Sleeve (if IsClusterResolved=true)            │   │
│  │    document.GetElement(ClusterSleeveInstanceId)                     │   │
│  │    ✅ EXISTS → Keep all flags true                                  │   │
│  │    ❌ DELETED → Reset ALL flags to false                            │   │
│  └─────────────────────────────────────────────────────────────────────┘   │
│  ┌─────────────────────────────────────────────────────────────────────┐   │
│  │ STEP 3: Check Individual Sleeve (if IsClusterResolved=false)       │   │
│  │    document.GetElement(SleeveInstanceId)                            │   │
│  │    ✅ EXISTS → Keep IsResolved=true                                 │   │
│  │    ❌ DELETED → Reset IsResolved=false                              │   │
│  └─────────────────────────────────────────────────────────────────────┘   │
│                                                                              │
│  After Reset: Save to Both XMLs                                              │
│    → Filter XML: _clashZoneStorage already modified                         │
│    → Global XML: GlobalIndexService.UpsertFlagsWithIds()                    │
└─────────────────────────────────────────────────────────────────────────────┘
```

---

## 🔄 SERVICE INTERACTIONS

```
┌─────────────────────────────────────────────────────────────────────────────┐
│                    REFRESHSERVICE                                            │
│ ──────────────────────────────────────────────────────────────────────────  │
│  Responsibilities:                                                          │
│  • Orchestrates refresh flow                                                 │
│  • Performs 3-Point Validation                                              │
│  • Removes invalid clash zones from Global XML                              │
│  • Updates intersection points                                               │
│  • Calls ClashZoneService for new detection                                 │
│                                                                              │
│  Direct Global XML Operations:                                              │
│  • LoadOrCreate() - Gets resolved GUIDs                                     │
│  • Remove entries - When validation fails                                    │
│  • Save() - After removing invalid entries                                  │
└─────────────────────────────────────────────────────────────────────────────┘
                                      │
                                      │ calls
                                      ▼
┌─────────────────────────────────────────────────────────────────────────────┐
│                    CLASHZONESERVICE                                          │
│ ──────────────────────────────────────────────────────────────────────────  │
│  Responsibilities:                                                          │
│  • Detects new clash zones (DetectNewClashZones)                            │
│  • Checks for duplicates (FindExistingClashZone)                            │
│  • Resets flags for deleted sleeves                                         │
│  • Manages _clashZoneStorage (Filter XML data)                              │
│                                                                              │
│  Direct Global XML Operations:                                              │
│  • EnsureEntries() - Creates entries for new clash zones                    │
│  • LoadOrCreate() - Checks Global XML during reset                           │
│  • UpsertFlagsWithIds() - Saves reset flags                                 │
└─────────────────────────────────────────────────────────────────────────────┘
                                      │
                                      │ uses
                                      ▼
┌─────────────────────────────────────────────────────────────────────────────┐
│                    GLOBALINDEXSERVICE                                        │
│ ──────────────────────────────────────────────────────────────────────────  │
│  Responsibilities:                                                          │
│  • Load/Save {category}_global.xml                                          │
│  • GetResolvedGuids() - Returns resolved GUID sets                           │
│  • UpsertFlagsWithIds() - Updates flags and sleeve IDs                      │
│  • EnsureEntries() - Creates entries for new GUIDs                           │
│  • Save() - Persists to disk                                                │
│                                                                              │
│  NO Revit API Calls:                                                        │
│  • Pure data store - just reads/writes XML                                  │
│  • Does NOT validate elements exist                                          │
│  • Does NOT reset flags                                                     │
└─────────────────────────────────────────────────────────────────────────────┘
```

---

## 📋 KEY DATA STRUCTURES

### Global XML (`{category}_global.xml`)
```xml
<CategoryGlobalIndex Category="Pipes">
  <Entry 
    Id="5db13c59-5938-4e66-95aa-ec1078eed78d"    <!-- ClashZone GUID -->
    IsResolved="true"
    IsClusterResolved="true"
    SleeveInstanceId="-1"                         <!-- Individual deleted -->
    ClusterSleeveInstanceId="950122"             <!-- Cluster sleeve ID -->
  />
</CategoryGlobalIndex>
```

### Filter XML (`{filter}_{category}.xml`)
```xml
<OpeningFilter>
  <ClashZoneStorage>
    <ClashZone>
      <Id>5db13c59-5938-4e66-95aa-ec1078eed78d</Id>
      <MepElementIdValue>12345</MepElementIdValue>
      <StructuralElementIdValue>67890</StructuralElementIdValue>
      <IsResolved>true</IsResolved>
      <IsClusterResolved>true</IsClusterResolved>
      <SleeveInstanceId>-1</SleeveInstanceId>
      <ClusterSleeveInstanceId>950122</ClusterSleeveInstanceId>
      <!-- ... other properties ... -->
    </ClashZone>
  </ClashZoneStorage>
</OpeningFilter>
```

---

## ⚠️ CRITICAL RULES

1. **GUID Uniqueness**: Each ClashZone has a unique GUID that never changes
2. **GUID Validation**: Only validated by 3-Point Validation (MEP/host exists + intersect)
3. **Flag Hierarchy**: Cluster flags (`IsClusterResolved`) take precedence over individual flags (`IsResolved`)
4. **Global XML Authority**: Global XML is the source of truth for flags across filters
5. **Filter XML History**: Filter XML keeps all clash zones (even resolved) for OK button logic
6. **No Redundant Validation**: GlobalIndexService does NOT check Revit - only stores data
7. **Reset Triggers**: Flags reset only when sleeves are actually deleted (via Revit API check)

---

## 🔍 DEBUGGING CHECKLIST

When duplicates appear, check:
- ✅ Are ElementIds properly initialized after XML load? (Use IntegerValue comparison)
- ✅ Is FindExistingClashZone checking the correct storage?
- ✅ Are invalid clash zones removed from both Filter XML and Global XML?
- ✅ Is ClashZoneService initialized with validated clash zones only?
- ✅ Are flags synced from Global XML before duplicate checking?


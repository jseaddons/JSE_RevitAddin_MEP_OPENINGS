# Refresh Flow - GUID and Flag Management

## Overview

This document illustrates the complete refresh flow for GUID management and sleeve flag management across three key services:
- **RefreshService**: Orchestrates refresh, validates clash zones, manages XML
- **ClashZoneService**: Manages clash zone detection, matching, and flag reset
- **GlobalIndexService**: Manages filter-independent flag storage

---

## Complete Refresh Flow (CORRECT ORDER)

```
┌─────────────────────────────────────────────────────────────────────┐
│                    REFRESH CLICK (RefreshService)                    │
│                  ExecuteRefreshInternal()                            │
└─────────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────────┐
│ STEP 1: Load Existing Clash Zones (RefreshService)                  │
│                                                                       │
│ LoadExistingClashZonesFromFilterXml()                                │
│   ├─ Load {filter}_{category}.xml                                   │
│   ├─ Deserialize ClashZone objects                                  │
│   └─ Store in existingClashZones.ClashZones                         │
│                                                                       │
│ 📋 Result: List of ClashZone objects with GUID, flags, coordinates  │
└─────────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────────┐
│ STEP 2: Sync Flags from Global XML (RefreshService)                 │
│                                                                       │
│ ⚠️ WHY THIS STEP IS NEEDED:                                          │
│    - Global XML is FILTER-INDEPENDENT (stores flags for ALL filters)│
│    - Filter XML is FILTER-SPECIFIC (stores data for ONE filter)     │
│    - When switching filters, new Filter XML doesn't have flags     │
│    - Example: Place sleeves in Filter A → Global XML has flags      │
│               Switch to Filter B → Filter B XML has NO flags        │
│               Sync from Global XML → Filter B XML gets flags        │
│                                                                       │
│ For each ClashZone in existingClashZones:                           │
│   ├─ GlobalIndexService.LoadOrCreate(doc, category)                │
│   ├─ Find entry by ClashZone.Id (GUID)                             │
│   ├─ Sync IsResolved, IsClusterResolved flags                       │
│   ├─ Sync SleeveInstanceId, ClusterSleeveInstanceId                │
│   └─ **KEEP clash zone in Filter XML** (for OK button logic)        │
│                                                                       │
│ 🔄 Purpose: Sync FROM Global XML (authoritative) TO Filter XML     │
│    This ensures Filter XML shows correct flags even if:             │
│    - You switched filters                                            │
│    - You refreshed after placing sleeves in another filter          │
│    - Filter XML was manually edited or corrupted                     │
└─────────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────────┐
│ STEP 3: 3-POINT VALIDATION (RefreshService) ⚠️ FIRST VALIDATION    │
│                                                                       │
│ ⚠️ DATA SOURCE: Filter XML (NOT Global XML)                         │
│    Filter XML has: MEP Element ID, Structural Element ID,           │
│                    Intersection Point                               │
│                                                                       │
│ For each ClashZone in existingClashZones (from Filter XML):         │
│   ├─ Point 1: MEP Element Exists in Revit?                          │
│   │   └─ GetElementFromDocumentOrLinked(MepElementId)              │
│   │   └─ Checks BOTH linked documents AND active document           │
│   │   └─ Uses: existingZone.MepElementId (from Filter XML)         │
│   ├─ Point 2: Structural Element Exists in Revit?                   │
│   │   └─ GetElementFromDocumentOrLinked(StructuralElementId)         │
│   │   └─ Checks BOTH linked documents AND active document           │
│   │   └─ Uses: existingZone.StructuralElementId (from Filter XML)   │
│   └─ Point 3: Elements Still Intersect?                             │
│       └─ CalculateIntersectionPoint(mep, structural)                │
│       └─ Uses: MEP & Structural elements from Revit API              │
│                                                                       │
│ ✅ VALID (All 3 points pass):                                        │
│   ├─ Keep clash zone in Filter XML                                  │
│   ├─ Update intersection point if moved (>1mm)                      │
│   ├─ Preserve flags (IsResolved, IsClusterResolved)                 │
│   └─ **NO Global XML changes** (GUID remains, flags preserved)      │
│                                                                       │
│ ❌ INVALID (Any point fails):                                        │
│   ├─ Remove clash zone from Filter XML                              │
│   ├─ Clear entry from Global XML (by GUID)                          │
│   └─ Log removal reason (MEP deleted / Structural deleted / No longer intersect)
│                                                                       │
│ 🔧 Global XML Cleanup (Only for invalid clash zones):               │
│   ├─ GlobalIndexService.LoadOrCreate(doc, category)                  │
│   ├─ globalIndex.Entries.RemoveAll(by GUID)                        │
│   │   └─ Uses GUID to find entry (Global XML doesn't have MEP/Host)│
│   └─ GlobalIndexService.Save(doc, globalIndex)                       │
└─────────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────────┐
│ STEP 4: Detect New Clash Zones (ClashZoneService)                    │
│                                                                       │
│ DetectNewClashZones(currentIntersections, document, ...)            │
│   ├─ For each intersection from IntersectionDetectionService:        │
│   │   ├─ Check if clash zone already exists (FindExistingClashZone) │
│   │   │   └─ Matches by MEP ID + Structural ID (not GUID)           │
│   │   ├─ If NOT found → Create new ClashZone with new GUID          │
│   │   └─ If found → Update existing (preserve GUID)                  │
│   │                                                                   │
│   └─ **INSIDE DetectNewClashZones:**                                 │
│       └─ ResetResolvedFlagForDeletedSleeves() ⚠️ SLEEVE CHECK        │
│                                                                       │
│ 📋 Result: List of new clash zones created                          │
└─────────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────────┐
│ STEP 5: Reset Flags for Deleted Sleeves ⚠️ SLEEVE VALIDATION        │
│         (Called INSIDE DetectNewClashZones, line 753)                │
│                                                                       │
│ ResetResolvedFlagForDeletedSleeves(document, selectedCategories)     │
│                                                                       │
│ ⚠️ CRITICAL: Checks sleeves in **ACTIVE DOCUMENT** Revit only       │
│    (NOT linked documents - sleeves are always in active document)   │
│                                                                       │
│ For each ClashZone in _clashZoneStorage.ClashZones:                 │
│   ├─ Check Global XML FIRST (read flags)                            │
│   │   └─ If Global XML says resolved → trust it (sleeve may be in linked file)
│   │                                                                   │
│   ├─ **FLAG HIERARCHY: Check cluster FIRST**                         │
│   │   ├─ If IsClusterResolved = true:                                │
│   │   │   ├─ Check cluster sleeve in ACTIVE DOCUMENT Revit          │
│   │   │   │   └─ GetElement(ClusterSleeveInstanceId)                 │
│   │   │   ├─ If cluster sleeve found → SKIP (keep flags=true)       │
│   │   │   └─ If cluster sleeve NOT found → Reset ALL flags to false  │
│   │   │                                                                 │
│   │   └─ If IsClusterResolved = false:                                │
│   │       └─ Check individual sleeve (see below)                     │
│   │                                                                   │
│   ├─ **FLAG HIERARCHY: Then check individual**                       │
│   │   ├─ If IsResolved = true (and cluster is false):                │
│   │   │   ├─ Check individual sleeve in ACTIVE DOCUMENT Revit       │
│   │   │   │   └─ GetElement(SleeveInstanceId)                         │
│   │   │   ├─ If individual sleeve found → SKIP (keep flags=true)    │
│   │   │   └─ If individual sleeve NOT found → Reset IsResolved=false │
│   │   │                                                                 │
│   │   └─ If both flags false → No action needed                       │
│   │                                                                   │
│   └─ **Save reset flags to BOTH Global XML and Filter XML**         │
│       ├─ SaveResetFlagsToBothXmls(document, selectedCategories)     │
│       ├─ GlobalIndexService.UpsertFlagsWithIds()                    │
│       └─ Filter XML saved by RefreshService later                    │
│                                                                       │
│ ✅ If cluster sleeve found → Sleeves already placed, disable OK    │
│ ❌ If cluster sleeve NOT found → Reset flags, enable OK              │
│                                                                       │
│ 📋 Result: Flags reset for deleted sleeves, flags preserved for existing sleeves
└─────────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────────┐
│ STEP 6: Save Updated Clash Zones (RefreshService)                   │
│                                                                       │
│ For each category:                                                    │
│   ├─ Combine existingClashZones + newClashZones                     │
│   ├─ Save to Filter XML ({filter}_{category}.xml)                   │
│   ├─ EnsureEntries in Global XML ({category}_global.xml)             │
│   │   └─ Creates entries for new clash zones (with flags=false)       │
│   └─ **DOES NOT** overwrite existing entries                        │
│                                                                       │
│ 📋 Result: Filter XML and Global XML updated                       │
└─────────────────────────────────────────────────────────────────────┘
```

---

## Key Principles

### 1. Two-Stage Validation Process

**Stage 1: Clash Zone Validity (3-Point Validation)**
- **When:** During refresh, BEFORE sleeve check
- **Purpose:** Ensure clash zone is still valid (MEP/Host exist, still intersect)
- **Data Source:** Filter XML (has all 3 data points)
- **Checks:** Linked documents AND active document (elements can be in either)
- **Result:** Invalid clash zones removed, valid ones kept

**Stage 2: Sleeve Existence (ResetResolvedFlagForDeletedSleeves)**
- **When:** During refresh, INSIDE DetectNewClashZones
- **Purpose:** Check if sleeves were manually deleted
- **Data Source:** Global XML flags + ACTIVE DOCUMENT Revit
- **Checks:** ACTIVE DOCUMENT ONLY (sleeves are always in active document)
- **Result:** Flags reset if sleeves deleted, flags preserved if sleeves exist

### 2. Flag Hierarchy

```
IsClusterResolved (HIGHEST PRIORITY)
    ↓
    If true → Check cluster sleeve in ACTIVE DOCUMENT
    │
    ├─ If cluster sleeve found → Keep flags=true, SKIP individual check
    └─ If cluster sleeve NOT found → Reset ALL flags to false
    
IsResolved (LOWER PRIORITY)
    ↓
    If true (and cluster is false) → Check individual sleeve in ACTIVE DOCUMENT
    │
    ├─ If individual sleeve found → Keep IsResolved=true
    └─ If individual sleeve NOT found → Reset IsResolved=false
```

### 3. OK Button Logic

The Main UI checks flags from Filter XML (which was synced from Global XML):
- If `IsClusterResolved=true` OR `IsResolved=true` → Disable OK button (sleeves already placed)
- If both flags=false → Enable OK button (sleeves need placement)

---

## ⚠️ CRITICAL: Global XML Does NOT Store 3-Point Data

**Global XML stores:**
- GUID (Id)
- Flags (IsResolved, IsClusterResolved)
- Sleeve IDs (SleeveInstanceId, ClusterSleeveInstanceId)

**Global XML does NOT store:**
- ❌ MEP Element ID
- ❌ Structural Element ID
- ❌ Intersection Point

**Therefore:** 3-Point Validation **MUST** use Filter XML, not Global XML!

---

## Flow Summary

### Refresh Click Flow:
1. **Load Filter XML** → Get existing clash zones
2. **Sync Global XML flags** → Read flags FROM Global XML TO Filter XML (filter-independent state)
3. **3-Point Validation** → Validate clash zones (uses Filter XML + Revit API)
   - Invalid → Remove from Filter XML, clear Global XML entry
   - Valid → Keep in Filter XML, preserve Global XML entry
4. **Detect New Clash Zones** → Create new clash zones for new intersections
   - **Inside this:** Reset flags for deleted sleeves (uses Global XML flags + ACTIVE DOCUMENT Revit)
5. **Save Updated Clash Zones** → Save to Filter XML, ensure entries in Global XML (for new clash zones only)

### OK Click (Placement) Flow:
1. **Place Individual Sleeves** → `UniversalSleevePlacerService`
   - Place sleeves in Revit
   - Update Filter XML (flags + SleeveInstanceId)
   - **Update Global XML** (flags + SleeveInstanceId) ⚠️ WRITES TO GLOBAL XML

2. **Place Cluster Sleeves** → `UniversalClusterService`
   - Place cluster sleeves in Revit
   - Delete individual sleeves
   - Update Filter XML (flags + ClusterSleeveInstanceId)
   - **Update Global XML** (flags + ClusterSleeveInstanceId) ⚠️ WRITES TO GLOBAL XML

### ⚠️ CRITICAL: Why Sync at Refresh?
**Global XML is the AUTHORITATIVE SOURCE** for flags because:
- It's filter-independent (works across all filters)
- It's updated AFTER every placement (OK click)
- Filter XML is filter-specific (only has data for current filter)

**Sync is needed when:**
- Switching between filters (new Filter XML doesn't have flags from previous filter)
- Refreshing after placing sleeves in another filter
- Filter XML was corrupted or manually edited

**Without sync:** Filter XML would have stale or missing flags, causing OK button to enable incorrectly.

---

## Why This Order Matters

1. **3-Point Validation FIRST** ensures we're only working with valid clash zones
2. **Flag Reset SECOND** ensures flags are accurate based on actual sleeve existence
3. **Global XML sync** ensures filter-independent flag management
4. **Filter XML save** ensures all data persists correctly

---

## Reference Code Locations

- **3-Point Validation:** `RefreshService.cs` lines 943-1083
- **ResetResolvedFlagForDeletedSleeves:** `ClashZoneService.cs` lines 1407-1605
- **DetectNewClashZones:** `ClashZoneService.cs` lines 157-757
- **Global XML Sync:** `RefreshService.cs` lines 858-914
- **Global XML Save:** `RefreshService.cs` line 2113

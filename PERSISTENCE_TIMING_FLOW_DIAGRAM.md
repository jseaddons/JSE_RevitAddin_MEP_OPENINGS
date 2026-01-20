# 🔄 PERSISTENCE TIMING FLOW DIAGRAM

## Purpose of `ClashZonePersistenceService.UpdateGlobalEntry`

**What it does**: Updates Global XML entries with clash zone data (MEP+Host+Point, flags, sleeve IDs) to maintain cross-filter state tracking.

**Why we need it**: 
- Global XML tracks flags and sleeve IDs across ALL filters for a category
- When clash zones are saved (after refresh or placement), we need to sync their state to Global XML
- This enables cross-filter detection (e.g., if a sleeve was placed in "Electrical" filter, "Plumbing" filter knows about it)

---

## 🔧 What "Structural Updates" Means

**`allowStructuralUpdates` parameter controls:**

### **When `allowStructuralUpdates = true`**:
1. **In `UpdateGlobalEntry`** (Global XML):
   - Updates `MepElementId` and `StructuralElementId`
   - Updates `IntersectionPointX`, `IntersectionPointY`, `IntersectionPointZ`
   - **Purpose**: Sync coordinates from Filter XML to Global XML for cross-filter tracking

2. **In `MergeZone`** (Filter XML):
   - Updates `SourceDocKey` and `HostDocKey`
   - Updates `MepParameterValues` and `HostParameterValues`
   - Updates `StructuralElementThickness`, `WallThickness`, `FramingThickness`
   - Updates `StructuralElementNormal`
   - Updates `IntersectionPointX/Y/Z`
   - Updates MEP orientation properties (`MepElementOrientation`, `MepElementRotationAngle`, etc.)
   - **Purpose**: Merge new detection results with existing zones

### **When `allowStructuralUpdates = false`**:
- Skips all coordinate and structural property updates
- Only updates flags and sleeve IDs
- **Purpose**: Preserve existing coordinates after placement (replay-only mode)

---

## ❓ Do We Need Structural Updates for PATH 1?

### **Answer: NO**

**PATH 1 (Replay-Only Mode):**
- ❌ **NO structural updates** (`allowStructuralUpdates = false`)
- ✅ **What it does**:
  - Checks if sleeves exist in Revit
  - Resets flags and instance IDs if sleeves are deleted
  - Places sleeves using existing Filter XML data
- ❌ **DOES NOT need 3-point validation** because we're not detecting new intersections

**PATH 2 (Fresh Mode - First Time Adding Filter):**
- ✅ **NEEDS structural updates** (`allowStructuralUpdates = true`)
- ✅ **What it does**:
  - Full intersection detection
  - Dump all data to DB and XML
  - Place sleeves → set flags to true
- ❌ **NO 3-point validation** (no existing data to validate)
- ❌ **NO flag reset** (no existing flags)
- ❌ **NO sync** (no existing Global XML entries)
- ✅ **Complexity**: SIMPLE - no timing issues

**PATH 3 (Non-Fresh Mode - Filter Exists, Refresh Again):**
- ✅ **NEEDS structural updates** (`allowStructuralUpdates = true`)
- ✅ **What it does**:
  - Full intersection detection
  - 3-point validation
  - Zone splitting: validated zones → PATH 1, non-validated zones → PATH 3 full logic
  - GUID checking, flag reset, flag sync
  - Updates coordinates, IDs, properties from new detection
- ✅ **NEEDS 3-point validation** to verify existing zones are still valid
- ⚠️ **Complexity**: COMPLEX - timing issues (only for non-validated zones)

### **Key Differences:**
- **PATH 1**: No structural updates, only flags/IDs
- **PATH 2 (Fresh)**: Dump all data, simple, no validation/sync needed
- **PATH 3 (Non-Fresh)**: Full validation with zone splitting - validated zones use PATH 1, non-validated zones use full PATH 3 logic

---

## ⚠️ THE TIMING ISSUE: Race Condition Flow

```
┌─────────────────────────────────────────────────────────────────────────┐
│              REFRESH REFACTORED OPERATION WORKFLOW                       │
│              TWO PATHS: Based on "Adopt to Document" Setting             │
│              (RefreshServiceRefactored - Clean Orchestrator Pattern)     │
└─────────────────────────────────────────────────────────────────────────┘

DECISION POINT: enableThreePointValidation (Adopt to Document checkbox)
  ├─ CHECKED (true)  → PATH 2: Full Detection + Validation
  └─ UNCHECKED (false) → PATH 1: Replay-Only (Skip Detection if Possible)

NOTE: This document focuses on RefreshServiceRefactored, not the legacy RefreshService.
      The refactored service uses a phased approach with RefreshContext and helper services.

═══════════════════════════════════════════════════════════════════════════
PATH 1: "Adopt to Document" UNCHECKED (enableThreePointValidation = false)
═══════════════════════════════════════════════════════════════════════════

STEP 1A: CHECK IF DETECTION CAN BE SKIPPED
────────────────────────────────────────────
RefreshServiceRefactored.ExecuteRefresh()
  ↓
Creates RefreshContext (holds all state)
  ↓
PHASE 2: Load XML once (XmlCacheManager.LoadAll)
  ↓
PHASE 3: Load existing clash zones from XML cache
  ↓
PHASE 4: Sync flags from Global XML (FlagManager.SyncFlagsFromGlobal)
  ↓
PHASE 6: IntersectionProcessor.PrepareExistingZones()
  ↓
IntersectionProcessor.DetermineDetectionDecision()
  - Checks: Are all file combos already processed?
  - Checks: Does Filter XML have placement data?
  - Checks: enableThreePointValidation setting
  ↓
IF enableThreePointValidation = false AND combos processed → Skip detection (PATH 1A)
IF enableThreePointValidation = false AND Filter XML has data → Skip detection (PATH 1A)
IF enableThreePointValidation = true OR combos NOT processed → Run detection (PATH 1B/PATH 2)

STEP 1A: SKIP DETECTION (Replay-Only Mode)
──────────────────────────────────────────
  ↓
IntersectionProcessor.PrepareExistingZones()
  - Decision.Mode = RefreshMode.Replay
  - Decision.ShouldRunDetection = false
  ↓
IntersectionProcessor.RunDetectionIfNeeded(decision)
  - Returns existing zones from context.ExistingClashZones
  - No new detection performed
  ↓
[Memory State: context.AllClashZones = existing zones from Filter XML]
  - Flags synced from Global XML (via FlagManager.SyncFlagsFromGlobal in PHASE 4)
  - Placement data from Filter XML
  ↓
SKIP: No new intersection detection
SKIP: No 3-point validation (if enableThreePointValidation = false)
KEEP: ALL zones (append-only, even invalid ones)

STEP 1B: RUN DETECTION (New File Combos or PATH 2)
───────────────────────────────────────────────────
  ↓
IntersectionProcessor.PrepareExistingZones()
  - Decision.Mode = RefreshMode.FullDetection
  - Decision.ShouldRunDetection = true
  ↓
IntersectionProcessor.RunDetectionIfNeeded(decision)
  - Calls ClashZoneService.DetectClashes()
  - Creates NEW ClashZone objects in memory
  ↓
Creates NEW ClashZone objects in memory
  - ClashZone objects have DEFAULT values:
    - IsResolved = false
    - IsClusterResolved = false
    - SleeveInstanceId = -1
    - ClusterSleeveInstanceId = -1
  ↓
[Memory State: context.NewClashZones = new zones, context.AllClashZones = existing + new]

═══════════════════════════════════════════════════════════════════════════
PATH 2: "Adopt to Document" CHECKED (enableThreePointValidation = true)
═══════════════════════════════════════════════════════════════════════════

STEP 1: ALWAYS RUN INTERSECTION DETECTION
──────────────────────────────────────────
RefreshServiceRefactored.ExecuteRefresh()
  ↓
PHASE 6: IntersectionProcessor.PrepareExistingZones()
  - Decision.Mode = RefreshMode.FullDetection
  - Decision.ShouldRunDetection = true (always when enableThreePointValidation = true)
  ↓
IntersectionProcessor.RunDetectionIfNeeded(decision)
  - Calls ClashZoneService.DetectClashes()
  - Runs intersection detection (always, even if combos processed)
  ↓
Creates NEW ClashZone objects in memory
  - ClashZone objects have DEFAULT values:
    - IsResolved = false
    - IsClusterResolved = false
    - SleeveInstanceId = -1
    - ClusterSleeveInstanceId = -1
  ↓
[Memory State: context.NewClashZones = new zones, context.AllClashZones = existing + new]

PHASE 5: 3-POINT VALIDATION (PATH 2 ONLY)
──────────────────────────────────────────
  ↓
ValidationService.ValidateClashZones(context.ExistingClashZones)
  - Uses ThreePointValidator internally
  - Checks if MEP element still exists
  - Checks if structural element still exists
  - Checks if intersection point still valid
  ↓
IF validation fails:
  - Mark as invalid
  - IF enableThreePointValidation = true → REMOVE from Global XML (via ValidationService.RemoveInvalidZonesFromGlobal)
  - IF enableThreePointValidation = false → KEEP (append-only)
  ↓
IF intersection point moved:
  - Update intersection point coordinates
  - Delete old sleeve (if moved beyond tolerance)
  - Reset flags (sleeve needs to be re-placed)
  ↓
[Memory State: context.ExistingClashZones = valid zones only, invalid zones removed]


═══════════════════════════════════════════════════════════════════════════
COMMON STEP: FLAG RESET (Both Paths)
═══════════════════════════════════════════════════════════════════════════

NOTE: Flag Reset in Refactored Service
─────────────────────────────────────────
  ↓
⚠️ IMPORTANT: RefreshServiceRefactored does NOT call FlagManager.ResetFlagsForDeletedSleeves()
  directly. Instead, flag reset happens in ValidationService during 3-point validation.
  ↓
PHASE 5: ValidationService.ValidateClashZones()
  - If intersection point moved → Flags reset automatically
  - If MEP/structural element deleted → Zone marked invalid (removed if enableThreePointValidation = true)
  ↓
For manually deleted sleeves (not detected by validation):
  - Flag reset should happen BEFORE refresh (via separate flag reset command)
  - OR: Flag reset happens in legacy RefreshService path (not in refactored)
  ↓
[Global XML State: Flags may be reset during validation]
[Database State: Flags may be reset during validation]
[Memory State: context.ExistingClashZones updated with validation results]
  ↓
NOTE: The timing issue still applies if flag reset happens separately before refresh


═══════════════════════════════════════════════════════════════════════════
COMMON STEP: PERSISTENCE (Both Paths - The Problem!)
═══════════════════════════════════════════════════════════════════════════

PHASE 8: PARAMETER CAPTURE
──────────────────────────
  ↓
ParameterCaptureService.CaptureParametersParallel(context.NewClashZones)
  - Captures MEP and structural element parameters
  - Updates context.NewClashZones with parameter values
  ↓
Updates context.AllClashZones to include parameters from new zones
  ↓
[Memory State: context.AllClashZones includes all zones with parameters]

PHASE 9: PERSISTENCE (The Problem!)
────────────────────────────────────
RefreshServiceRefactored.MergeAndSave(context)
  ↓
Determine allowStructuralUpdates:
  - PATH 1 (replay mode): allowStructuralUpdates = false ✅
  - PATH 2 (full detection): allowStructuralUpdates = true ✅
  - First run (no data): allowStructuralUpdates = true ✅
  ↓
ClashZonePersistenceService.SaveClashZones(
    context.AllClashZones,  // ✅ ALL zones at once (not per category)
    normalizedBaseName,
    enabledFilter,
    allowStructuralUpdates)  // ✅ Based on path (false for PATH 1, true for PATH 2/first run)
  ↓
For each ClashZone in context.AllClashZones:
  ↓
  ClashZonePersistenceService.UpdateGlobalEntry(entry, clashZone, allowStructuralUpdates)
    ↓
    [PROBLEM]: clashZone object has STALE values:
      - clashZone.IsResolved = false (default, not from Global XML)
      - clashZone.IsClusterResolved = false (default, not from Global XML)
      - clashZone.SleeveInstanceId = -1 (default)
      - clashZone.ClusterSleeveInstanceId = -1 (default)
    ↓
    Checks: entryWasResetByFlagManager?
      - entry.IsResolved = false ✅ (FlagManager just reset it)
      - entry.IsClusterResolved = false ✅ (FlagManager just reset it)
      - entry.SleeveInstanceId = -1 ✅
      - entry.ClusterSleeveInstanceId = -1 ✅
      → entryWasResetByFlagManager = TRUE ✅
    ↓
    [FIXED]: If entryWasResetByFlagManager = true:
      - PRESERVE entry flags (don't overwrite)
      - Return early ✅
    ↓
    [BUT]: If entryWasResetByFlagManager = false:
      - Line 1274-1277: OVERWRITES with clashZone values
      - entry.IsResolved = zone.IsResolved (stale false)
      - entry.IsClusterResolved = zone.IsClusterResolved (stale false)
      - This is CORRECT if entry wasn't reset, but WRONG if timing is off
    ↓
    [STRUCTURAL UPDATES]: 
      - If allowStructuralUpdates = true (PATH 2/first run):
        - Updates MEP+Host+Point coordinates in Global XML
        - Updates MEP/Structural Element IDs in Global XML
        - Updates structural element properties (thickness, normal, orientation)
      - If allowStructuralUpdates = false (PATH 1):
        - Only updates flags and instance IDs ✅
  ↓
[Global XML State: May be overwritten with stale values if timing is wrong]


STEP 4: SAVE TO DISK
────────────────────
GlobalIndexService.Save()
  ↓
Writes Global XML to file
  ↓
[File State: May have stale flags if Step 3 overwrote them]
```

---

## 🔴 THE ROOT CAUSE

### **Problem**: 
`allClashZones` list contains **NEW** `ClashZone` objects created during clash detection. These objects have **DEFAULT** flag values (`false`, `-1`), NOT the values from Global XML that `FlagManager` just updated.

### **Timing Issue**:
1. `FlagManager.ResetFlagsForDeletedSleeves()` updates Global XML with correct reset flags ✅
2. `ClashZonePersistenceService.SaveClashZones()` runs AFTER with stale `ClashZone` objects ❌
3. `UpdateGlobalEntry()` tries to overwrite Global XML with stale values from `ClashZone` objects ❌

### **The Fix** (Already Implemented):
`UpdateGlobalEntry()` checks if entry was reset by `FlagManager`:
- If `entryWasResetByFlagManager = true` → **PRESERVE** reset flags, don't overwrite ✅
- If `entryWasResetByFlagManager = false` → Update with `ClashZone` values (for new entries or moved MEP elements) ✅

---

## 📊 DETAILED FLOW WITH CODE REFERENCES

### **Scenario A: Sleeve Deleted, Flag Reset, Then Persistence (Both Paths)**

```
TIME T0: User deletes cluster sleeve from Revit
  ↓
TIME T1: RefreshServiceRefactored.ExecuteRefresh() starts
  ↓
  Creates RefreshContext with enableThreePointValidation setting
  ↓
  DECISION: enableThreePointValidation?
    ├─ PATH 1 (false): Skip detection if combos processed
    └─ PATH 2 (true): Always run detection
  ↓
TIME T2: PHASE 4: FlagManager.SyncFlagsFromGlobal(context.ExistingClashZones)
  - Syncs flags from Global XML to in-memory clash zones
  - clashZone.IsClusterResolved = true (from Global XML, before reset)
  ↓
TIME T3A (PATH 1): IntersectionProcessor.PrepareExistingZones()
  - Decision.Mode = RefreshMode.Replay
  - Decision.ShouldRunDetection = false
  - Returns existing zones from context.ExistingClashZones
  ↓
TIME T3B (PATH 2): IntersectionProcessor.PrepareExistingZones()
  - Decision.Mode = RefreshMode.FullDetection
  - Decision.ShouldRunDetection = true
  - Calls ClashZoneService.DetectClashes()
  - Creates NEW ClashZone objects with default flags
  ↓
TIME T4: PHASE 5: ValidationService.ValidateClashZones()
  - Checks if sleeves exist in Revit
  - If intersection point moved → Resets flags
  - If element deleted → Marks invalid (removed if enableThreePointValidation = true)
  ↓
TIME T5: PHASE 9: ClashZonePersistenceService.SaveClashZones(context.AllClashZones, allowStructuralUpdates)
  - PATH 1: allowStructuralUpdates = false ✅ (replay mode)
  - PATH 2: allowStructuralUpdates = true ✅ (full detection)
  - PATH 1: context.AllClashZones from Filter XML (may have stale flags from T2)
  - PATH 2: context.AllClashZones contains NEW ClashZone objects (from T3B)
  - clashZone.IsClusterResolved = false (default or stale, not from Global XML)
  ↓
TIME T6: UpdateGlobalEntry(entry, clashZone, allowStructuralUpdates)
  - entry.IsClusterResolved = false (from T4, validation reset)
  - clashZone.IsClusterResolved = false (from T3A/T3B, default or stale value)
  - entryWasResetByFlagManager = true ✅ (entry has flags=false, IDs=-1)
  - PRESERVES entry flags ✅ (doesn't overwrite)
  - PATH 1: Only updates flags/IDs ✅
  - PATH 2: Updates MEP+Host+Point (allowStructuralUpdates = true) ✅
  ↓
RESULT: Global XML keeps IsClusterResolved=false ✅ CORRECT (Both Paths)
```

### **Scenario B: New Clash Zone (No Previous Entry)**

```
TIME T0: New clash detected
  ↓
TIME T1: RefreshServiceRefactored.ExecuteRefresh()
  ↓
  Creates RefreshContext
  ↓
  DECISION: enableThreePointValidation?
    ├─ PATH 1 (false): Skip detection if combos processed
    └─ PATH 2 (true): Always run detection
  ↓
TIME T2A (PATH 1): IntersectionProcessor.PrepareExistingZones()
  - Decision.Mode = RefreshMode.Replay
  - No new clash zones (using existing from Filter XML)
  - Scenario B doesn't apply (no new clashes in replay mode)
  ↓
TIME T2B (PATH 2): IntersectionProcessor.RunDetectionIfNeeded(decision)
  - Calls ClashZoneService.DetectClashes()
  - Creates NEW ClashZone object
  - clashZone.IsResolved = false (default)
  - clashZone.IsClusterResolved = false (default)
  - Added to context.NewClashZones
  ↓
TIME T3: PHASE 5: ValidationService.ValidateClashZones()
  - No entry in Global XML (new clash)
  - Nothing to validate (new zone)
  ↓
TIME T4: PHASE 9: ClashZonePersistenceService.SaveClashZones(context.AllClashZones, allowStructuralUpdates: true)
  - PATH 2: context.AllClashZones contains NEW ClashZone object
  - allowStructuralUpdates = true ✅ (full detection mode)
  ↓
TIME T5: UpdateGlobalEntry(entry, clashZone, allowStructuralUpdates: true)
  - entry doesn't exist (new entry)
  - entryExists = false
  - Creates new entry with clashZone values ✅
  - entry.IsResolved = false ✅
  - entry.IsClusterResolved = false ✅
  - Updates MEP+Host+Point (allowStructuralUpdates = true) ✅
  ↓
RESULT: New entry created with correct default values ✅ CORRECT
```

### **Scenario C: Sleeve Placed, Then Persistence (Both Paths)**

```
TIME T0: Sleeve placed via UniversalSleevePlacerService
  ↓
TIME T1: FlagManager.UpdateFlagsForPlacement(clashZone, sleeveId, isCluster: false)
  - Sets clashZone.IsResolved = true ✅
  - Sets clashZone.SleeveInstanceId = sleeveId ✅
  - Updates Database ✅
  - Updates Global XML ✅
  ↓
TIME T2: RefreshServiceRefactored.ExecuteRefresh() runs (either PATH 1 or PATH 2)
  ↓
TIME T3: PHASE 4: FlagManager.SyncFlagsFromGlobal(context.ExistingClashZones)
  - Syncs flags from Global XML to in-memory clash zones
  - clashZone.IsResolved = true ✅ (synced from Global XML)
  - clashZone.SleeveInstanceId = sleeveId ✅ (synced from Global XML)
  ↓
TIME T4: PHASE 9: ClashZonePersistenceService.SaveClashZones(context.AllClashZones, allowStructuralUpdates)
  - PATH 1: allowStructuralUpdates = false ✅ (replay mode)
  - PATH 2: allowStructuralUpdates = true ✅ (full detection mode)
  - PATH 1: context.AllClashZones from Filter XML (flags synced from Global XML in T3)
  - PATH 2: context.AllClashZones from detection (flags synced from Global XML in T3)
  - clashZone.IsResolved = true ✅ (synced from Global XML)
  - clashZone.SleeveInstanceId = sleeveId ✅ (synced from Global XML)
  ↓
TIME T5: UpdateGlobalEntry(entry, clashZone, allowStructuralUpdates)
  - entry exists
  - zoneHasSleeve = true (clashZone.SleeveInstanceId > 0)
  - entryWasResetByFlagManager = false (entry has flags=true, IDs>0)
  - Updates entry.IsResolved = true ✅
  - Updates entry.SleeveInstanceId = sleeveId ✅
  - PATH 1: Only updates flags/IDs ✅
  - PATH 2: Updates MEP+Host+Point (allowStructuralUpdates = true) ✅
  ↓
RESULT: Global XML updated with correct placed sleeve state ✅ CORRECT (Both Paths)
```

### **Scenario D: MEP Element Moved (PATH 2 Only - "Adopt Document" Enabled)**

```
TIME T0: MEP element moved in Revit (intersection point changed)
  ↓
TIME T1: RefreshServiceRefactored.ExecuteRefresh() with enableThreePointValidation = true
  ↓
TIME T2: PHASE 6: IntersectionProcessor.RunDetectionIfNeeded(decision)
  - Calls ClashZoneService.DetectClashes()
  - Detects NEW intersection point
  - Creates NEW ClashZone object with new coordinates
  - clashZone.IsResolved = false (default)
  - clashZone.SleeveInstanceId = -1 (default)
  - Added to context.NewClashZones
  ↓
TIME T3: PHASE 5: ValidationService.ValidateClashZones(context.ExistingClashZones)
  - Uses ThreePointValidator internally
  - Finds existing clash zone with OLD intersection point
  - Detects intersection point moved beyond tolerance
  - Deletes old sleeve (if exists)
  - Updates clash zone with NEW intersection point
  - Resets flags (sleeve needs to be re-placed at new location)
  - Updates context.ExistingClashZones with validated zones
  ↓
TIME T4: PHASE 9: ClashZonePersistenceService.SaveClashZones(context.AllClashZones, allowStructuralUpdates: true)
  - PATH 2 only (full detection mode)
  - context.AllClashZones contains updated ClashZone with NEW intersection point
  - clashZone.IsResolved = false (reset from T3)
  - clashZone.SleeveInstanceId = -1 (reset from T3)
  ↓
TIME T5: UpdateGlobalEntry(entry, clashZone, allowStructuralUpdates: true)
  - entry exists
  - entryWasResetByFlagManager = true (entry has flags=false, IDs=-1)
  - PRESERVES reset flags ✅ (doesn't overwrite)
  - Updates MEP+Host+Point with NEW coordinates ✅ (allowStructuralUpdates = true)
  ↓
RESULT: Global XML has reset flags + new intersection point ✅ CORRECT
```

---

## 🎯 WHY `UpdateGlobalEntry` EXISTS

### **Purpose**:
1. **Sync Clash Zone Data**: When clash zones are saved (after refresh or placement), sync their MEP+Host+Point data to Global XML
2. **Maintain Flag Consistency**: Ensure Global XML flags match the actual state of clash zones
3. **Cross-Filter Tracking**: Enable other filters to know about sleeves placed in different filters

### **What It Updates**:
- **MEP+Host+Point**: Intersection coordinates, MEP element ID, structural element ID
- **Flags**: `IsResolved`, `IsClusterResolved` (with preservation logic)
- **Sleeve IDs**: `SleeveInstanceId`, `ClusterSleeveInstanceId` (with preservation logic)
- **Filter Name**: Which filter XML file contains the placement data

### **Why We Need It**:
Without `UpdateGlobalEntry`, Global XML would become stale:
- New clash zones wouldn't be tracked
- Flag changes wouldn't be persisted
- Cross-filter detection wouldn't work

---

## 🔧 THE PRESERVATION LOGIC (The Fix)

```csharp
// Line 1236-1237: Check if entry was reset by FlagManager
bool entryWasResetByFlagManager = entryExists && 
                                   !entry.IsResolved && 
                                   !entry.IsClusterResolved && 
                                   entry.SleeveInstanceId <= 0 && 
                                   entry.ClusterSleeveInstanceId <= 0;

// Line 1241-1250: If reset, PRESERVE the reset state
if (entryWasResetByFlagManager)
{
    // Don't overwrite with stale clashZone values
    // Keep entry flags as false, IDs as -1
    return; // Exit early
}
```

**This prevents**: Stale `ClashZone` objects from overwriting correctly reset flags in Global XML.

---

## 📋 SUMMARY: Three Paths Comparison

| Step | PATH 1 (Replay) | PATH 2 (Fresh) | PATH 3 (Non-Fresh) |
|------|----------------|----------------|-------------------|
| **When Used** | "Adopt to Document" UNCHECKED<br>Filter already processed | First time adding filter<br>"Adopt to Document" CHECKED | Filter exists, refresh again<br>"Adopt to Document" CHECKED |
| **Intersection Detection** | ⏭️ **SKIPPED**<br>✅ Uses Filter XML data | ✅ **FULL DETECTION**<br>✅ Detects all intersections | ✅ **FULL DETECTION**<br>✅ Detects all intersections |
| **3-Point Validation** | ❌ **DISABLED**<br>✅ Keeps ALL zones | ❌ **NOT NEEDED**<br>✅ No existing data to validate | ✅ **ENABLED**<br>✅ Validates existing zones |
| **Zone Splitting** | N/A | N/A | ✅ **Validated zones** → PATH 1 behavior<br>✅ **Non-validated zones** → PATH 3 full logic |
| **GUID Checking** | ❌ **NOT NEEDED** | ❌ **NOT NEEDED** | ✅ **REQUIRED**<br>✅ Match existing entries |
| **Flag Reset** | ✅ **ENABLED**<br>✅ Checks deleted sleeves | ❌ **NOT NEEDED**<br>✅ No existing flags | ✅ **ENABLED**<br>✅ Checks deleted sleeves |
| **Flag Sync** | ❌ **NOT NEEDED** | ❌ **NOT NEEDED** | ✅ **REQUIRED**<br>✅ Sync from Global XML |
| **Structural Updates** | ❌ **DISABLED** (`allowStructuralUpdates=false`)<br>❌ NO coordinates, NO IDs, NO properties<br>✅ ONLY flags/IDs | ✅ **ENABLED** (`allowStructuralUpdates=true`)<br>✅ Dump all data to DB/XML | ✅ **ENABLED** (`allowStructuralUpdates=true`)<br>✅ Updates coordinates, IDs, properties |
| **Persistence** | ✅ Only flags/IDs reset<br>❌ NO structural data updates<br>✅ Places sleeves using existing Filter XML data | ✅ Dump to DB and XML<br>✅ Place sleeves → set flags true<br>✅ Simple, no timing issues | ✅ Updates structural data (coordinates, IDs, properties)<br>⚠️ Complex timing issues |
| **Complexity** | ✅ **SIMPLE**<br>✅ No timing issues | ✅ **SIMPLE**<br>✅ No timing issues | ⚠️ **COMPLEX**<br>⚠️ Timing issues (only for non-validated zones) |

## 🎯 KEY DIFFERENCES

### **PATH 1: Replay Mode ("Adopt to Document" UNCHECKED)**
- **Purpose**: Replay-only mode - use existing placement data
- **When Used**: File combos already processed, user trusts model unchanged
- **Behavior**: 
  - Skips intersection detection (uses Filter XML via IntersectionProcessor)
  - Keeps ALL zones (even invalid ones) - validation skipped
  - ✅ **Only flags and instance IDs updated** (`allowStructuralUpdates = false`)
  - ✅ **Places sleeves using existing Filter XML data**
  - ❌ **3-Point Validation**: NOT needed (no new detection)
  - ✅ **Complexity**: SIMPLE - no timing issues
  - Faster (no detection overhead)

### **PATH 2: Fresh Mode (First Time Adding Filter)**
- **Purpose**: Initial detection and data dump
- **When Used**: User adds a new filter for the first time → clicks refresh
- **Behavior**:
  - ✅ **Full intersection detection** (all intersections)
  - ❌ **NO 3-point validation** (no existing data to validate)
  - ❌ **NO GUID checking** (no existing entries)
  - ❌ **NO flag reset** (no existing flags)
  - ❌ **NO flag sync** (no existing Global XML entries)
  - ✅ **Dump all data to DB and XML** (`allowStructuralUpdates = true`)
  - ✅ **Place sleeves → set flags to true**
  - ✅ **Complexity**: SIMPLE - no timing issues
  - Straightforward: detect → dump → place → done

### **PATH 3: Non-Fresh Mode (Filter Exists, Refresh Again)**
- **Purpose**: Full validation mode - detect changes and validate existing zones
- **When Used**: Filter already processed before, user clicks refresh again
- **Behavior**:
  - ✅ **Full intersection detection** (all intersections)
  - ✅ **3-Point Validation**: ENABLED
  - ✅ **Zone Splitting After Validation**:
    - **Validated zones** (no change detected) → **PATH 1 behavior** (simple, no timing issues)
    - **Non-validated zones** (moved/deleted/invalid) → **PATH 3 full logic** (complex timing issues)
  - ✅ **GUID checking** (match existing entries)
  - ✅ **Flag reset** (check if sleeves deleted)
  - ✅ **Flag sync** (sync from Global XML)
  - ✅ **Updates structural data** (`allowStructuralUpdates = true`)
  - ⚠️ **Complexity**: COMPLEX - timing issues (only for non-validated zones)

## ⚠️ THE TIMING ISSUE (PATH 3 Non-Validated Zones Only)

| Step | Service | What It Does | State After |
|------|---------|--------------|-------------|
| 1 | `RefreshServiceRefactored` | Creates RefreshContext, loads XML cache | Context: All state held in RefreshContext |
| 2 | `FlagManager.SyncFlagsFromGlobal` | Syncs flags from Global XML to in-memory zones | Memory: Flags synced from Global XML ✅ |
| 3 | `IntersectionProcessor` | Determines detection mode, runs detection if needed | Context: New zones added to context.NewClashZones |
| 4 | `ValidationService` | Validates existing zones, resets flags if needed | Context: Valid zones only, flags reset if moved/deleted ✅ |
| 5 | `ParameterCaptureService` | Captures parameters for new zones | Context: AllClashZones includes parameters |
| 6 | `ClashZonePersistenceService` | Saves ALL clash zones at once | Calls `UpdateGlobalEntry` for each zone |
| 7 | `UpdateGlobalEntry` | Updates Global XML entry | **PRESERVES** reset flags if `entryWasResetByFlagManager=true` ✅<br>**OVERWRITES** with clashZone values if new entry or sleeve placed ✅<br>**PATH 1**: Only flags/IDs ✅<br>**PATH 2**: Updates structural data (coordinates, IDs, properties) ✅ |

**The timing issue is HANDLED by the preservation logic**, but it's a delicate balance that can break if the logic isn't perfect.


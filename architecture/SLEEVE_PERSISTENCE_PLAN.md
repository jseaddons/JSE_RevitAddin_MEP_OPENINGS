# Sleeve Persistence Architecture Plan

## 0. Objective
Build a persistence pipeline that:
- Treats **Global XML** as the authoritative flag ledger across all filters and categories, guaranteeing cross-filter consistency and reliable skip logic.
- Uses **Filter XML** (per-filter master) to store the complete placement dataset that drives recalculation, clustering, and diagnostics.
- Uses **Filter\_cat XML** (category snapshots such as `Plumbing_pipes.xml`) strictly as the payload for sleeve placement and cluster formation, never as a flag authority.
- Eliminates duplicated entries (especially in Global XML) so flag resets and placement decisions remain deterministic and idempotent.
- Preserves clear sequencing: data flows from detection → persistence → serialization → cleanup, with no conflicting writes.

## 1. Guiding Principles
- **Single write pipeline.** Only `ClashZonePersistenceService` mutates Filter XML and Global XML. All higher-level services (refresh, placement, clustering, coordinate updates) hand it sanitized snapshots.
- **Base-name discipline.** Always pass the base filter name (e.g. `Plumbing`) into persistence. Let the service derive `Plumbing_pipes`, etc. Misnamed inputs (like `Plumbing_pipes` passed as base) create ghost branches and duplicated entries.
- **Save before scrub.** Never call `ClearRevitApiObjects()` until every persistence method has run. Clearing first wipes placement and bounding-box data, so the next cycle falls back to host centroids.
- **Clone after placement.** Clone `ClashZone` objects only after placement has populated sleeve IDs, placement points, and bounding boxes. Passing pre-placement clones leads to empty XML but “success” logs.
- **Idempotent saves.** Persistence must detect and trim duplicate `FileCombo` / GUID entries so re-runs do not grow the files.

## 2. Roles of Each XML
- **Global XML (`{category}_global.xml`)**
  - Owns flags (`IsResolved`, `IsClusterResolved`, `SleeveInstanceId`, `ClusterSleeveInstanceId`) for every clash across all filters.
  - Stores the canonical `(MEP Id, Host Id, Intersection Point)` key used for cross-filter skip checks and flag resets.
  - Never depends on per-filter state to make decisions; filters synchronize to it, not vice versa.
- **Filter XML (`{filter}.xml`)**
  - The complete placement archive for a filter (all categories, all file combos).
  - Used to pre-calculate placement geometry, clustering, and diagnostics when running the same filter later.
  - Mirrors flags for transparency but should not be used for skip decisions.
- **Filter-category XML (`{filter}_{category}.xml`)**
  - Optimized payload for placement commands; contains only that category’s clash zones grouped by file combo.
  - Drives multi-category clustering by supplying bounding boxes and placement points.
  - No flag authority; relies on Global XML to know which zones should be processed.

## 2. End-to-End Save/Load Flow

| Stage | Key Steps | Persistence Call | Files Persisted | Cleanup |
|-------|-----------|------------------|-----------------|---------|
| **Refresh (Process Clash Zones)** | Detect intersections → build/merge clash zones | `SaveClashZones(allZones, baseFilter, targetFilterTree)` | `Plumbing.xml` (master) + each category snapshot (e.g. `Plumbing_pipes.xml`) via `FilterManagementService.SaveFilterToXmlFile` | Only after all saves, call `ClearRevitApiObjects()` |
| **Placement (UniversalSleevePlacerService)** | Load per-category XML → merge placement results | One `SaveClashZones` per category (updates Global + tree) | Serialize category file; serialize master filter **without** re-calling `SaveClashZones` | Clear API objects after logging |
| **Coordinate refresh (SleeveCoordinateService)** | Update bounding boxes | Single `SaveClashZones` | Serialize updated filter snapshots | Optional cleanup |
| **Clustering (UniversalClusterService)** | Update cluster IDs + bounding boxes | Single `SaveClashZones` | Serialize category + master filters | Clear cached cluster data |
| **Reset (FlagManager)** | Compare Revit sleeve IDs vs Global XML | Does not call persistence; rewrites entries directly in Global XML | Global XML rewritten in-place | n/a |

## 3. Naming & Timing Guardrails
- **Base name handling.** Always derive category-specific filenames inside persistence using helper methods; never pass suffixed names (`*_pipes`) into `SaveClashZones`.
- **Single persistence invocation per logical update.** Per-category loops in placement/refresh invoke `SaveClashZones` once. Subsequent serialization uses the already-updated in-memory tree.
- **Log critical steps.** Use direct `File.AppendAllText` logging around `[XML-MERGE-*]`, `[PERSIST-*]`, and flag reset points to validate the pipeline.
- **Ensure serialization after merge.** Once persistence updates the in-memory filter tree, immediately call `FilterManagementService.SaveFilterToXmlFile` for the master filter and each category snapshot.
- **Clear objects last.** Only after serialization and logging should we call `ClearRevitApiObjects()` on clash zones.

## 4. Scenarios & Required Hooks

| Scenario | Expected Behaviour | Notes |
|----------|-------------------|-------|
| **New UI combo (new linked/host selection)** | Refresh detects new `ProcessedFileCombo`, runs detection, persistence adds new filter group once | Ensure `DetectNewUiSelections` cross-checks Global XML keys |
| **Deleted sleeve** | Next refresh: `FlagManager` sees missing ElementId, resets `IsResolved=false`, `SleeveInstanceId=-1`; persistence rewrites Global XML, placement reuses Filter XML data on next run | Requires deduped Global XML so reset hits the correct entry |
| **Crash recovery** | Master filter (`Plumbing.xml`) retains latest placement data; Global XML holds flag state → restart reads both, no manual cleanup | Avoid double-saving to keep timestamps honest |
| **Mixed linked files (FP + PH)** | Filter tree stores both combos; refresh iterates only UI-selected combos, but flags for unselected remain untouched | Document expectation that UI selection drives processing scope |

## 4A. Sleeve Placement Paths (Post-Refactor)

To make placement decisions predictable while supporting targeted recomputation, the pipeline now splits into three explicit services. Each path consumes a different slice of persisted data and sets `allowStructuralUpdates` appropriately when calling `ClashZonePersistenceService`.

| Path | Trigger | Service | Data Read | Persistence Behaviour | Notes |
|------|---------|---------|-----------|------------------------|-------|
| **Path 1 – Replay Placement** | Default refresh/placement when neither global configuration nor UI overrides changed | `SleevePlacementReplayService` | Uses the latest `{filter}_{category}.xml` snapshot (intersection point, sleeve size, offsets, cluster payload) | After placement, calls `SaveClashZones(..., allowStructuralUpdates: false)` so only flags/IDs flow back to Global XML | Guarantees we place exactly what the snapshot describes. Zones with zeroed coordinates are skipped and flagged for detection. |
| **Path 2 – Clearance/Type Recalc** | User edits UI clearance/type controls without adopting the document | `SleeveSizingService` (composes Path 1) | Starts from the same snapshot as Path 1, recomputes sleeve width/height/diameter, opening family/type, placement offsets | Persists the new sizing via `SaveClashZones(..., false)` and immediately hands the refreshed snapshot to Path 1 for placement | Detection geometry is untouched; only size-related fields change so replay still honours cached intersection points. |
| **Path 3 – Detection Rebuild** | “Adopt to modified document” or global configuration change (tolerance, level strategy, etc.) | `SleeveDetectionService` | Runs full clash detection and rebuilds the master tree (`{filter}.xml` + category snapshots) | Calls `SaveClashZones(..., true)` to overwrite intersection geometry, host IDs, and placement payload | Only Path 3 is allowed to mutate detection-derived data. Placement falls back to Path 1 immediately afterwards. |

### Persistence Impact
- **Snapshot discipline:** Path 1 defines the canonical replay snapshot. Any service that adjusts data (Path 2 or Path 3) must run Path 1 immediately after so sleeve IDs and flags stay synchronized with Global XML.
- **`allowStructuralUpdates`:** Path 1 and Path 2 always pass `false`; they never rewrite intersection geometry. Path 3 passes `true` because new detection legitimately changes those fields.
- **Conditions files:** UI overrides (clearance/type) map to Path 2; global configuration edits drive Path 3. Documenting the trigger makes it obvious why a run is “lightweight” versus “heavy.”
- **Global XML ledger:** All paths still push flag state through `SaveClashZones`, but only Path 3 updates the `(MEP Id, Host Id, Intersection Point)` key that flag reset relies on.

### 4B. Anticipated Risks and Mitigations

Splitting placement into three services introduces new failure modes. The table below lists the main risks we need to watch while implementing the refactor.

| Risk | Description | Mitigation |
|------|-------------|------------|
| **Wrong path selection** | Orchestrator misreads state (e.g. hash/timestamp drift) and picks the wrong path, causing stale sleeve sizes or unnecessary detection reruns. | Centralize path decision in one helper that compares persisted fingerprints (UI overrides vs global config). Unit-test the decision matrix and log `[PLACEMENT-PATH]` for each run. |
| **Snapshot divergence** | Path 2 recalculates sizing but Path 1 reloads an old snapshot from disk, so placement uses stale data. | Treat the sizing output as canonical: persist the updated snapshot before replay and pass the same in-memory instance to Path 1. Add assertions/logging that replay sees refreshed dimensions. |
| **Misuse of `allowStructuralUpdates`** | A new persistence call forgets to pass `false` for replay/sizing and overwrites intersection points, reintroducing host-centre placement. | Wrap persistence entry points per path (`SaveReplaySnapshot`, `SaveSizingSnapshot`, etc.) so callers never touch the flag directly. Consider a lint/test that forbids raw `SaveClashZones(..., true)` outside detection. |
| **Conditions misclassification** | UI overrides vs global settings aren’t differentiated correctly; detection fails to run when geometry actually changed (or runs too often). | Maintain separate hashes/timestamps for “global” vs “session/UI” settings, store them with the snapshot, and log the reason a path fired. |
| **Cluster incompatibility** | Path 2 changes member sizing but forgets to refresh derived cluster payload, so cluster placement uses stale bounding boxes. | After sizing, recompute cluster metadata (or flag clusters for regeneration) before handing off to replay. Tests should verify cluster snapshots update correctly. |
| **Flag persistence gaps** | Replay routes skip persistence on error, leaving Global XML with stale flags and causing double placement. | Keep `SaveClashZones` calls in `finally` blocks or ensure failures surface with explicit error logging. Track success per path to confirm flags always reach Global XML. |
| **Legacy callers bypass coordinator** | Other services (cluster, coordinate updates) continue calling old placement entry points, bypassing the new path logic. | Provide a single public coordinator API and migrate every caller to it. Back-stop with logging that warns when legacy code paths are invoked. |
| **Logging blind spots** | Without distinct log markers per path, diagnosing misbehaviour becomes guesswork. | Emit consistent markers (`[PLACEMENT-PATH] Replay/Sizing/Detection`) and include snapshot fingerprints (hashes, timestamps) for traceability. |

## 5. Failure Modes & Mitigations
- **Duplicate GUIDs in Global XML.** Cause: multiple `SaveClashZones` calls or incorrect naming. Mitigation: one call per save, enforce base-name normalization, add dedupe helper for existing data.
- **Global XML and Filter XML disagree on flags.** Cause: stale duplicates or missing resets. Mitigation: run dedupe, ensure reset logic updates all matching entries, log mismatches with GUID context.
- **Zeroed placement data after persistence.** Cause: clones created before placement fields populated, or `ClearRevitApiObjects()` executed too early. Mitigation: adjust clone timing, enforce “save before scrub”.
- **“Unknown” filter groups.** Cause: persistence unable to derive base name. Mitigation: sanitize `_filterName` before saving, log + reject fallbacks.
- **Silent save failures.** Cause: logging wrapped in deployment guards or swallowed exceptions. Mitigation: direct logging, diagnostic file (`save_xml_diagnostic.log`), rethrow or surface errors.

## 6. Lessons Learned
- Logging must be unconditional for critical stages; any `if (!DeploymentMode)` wrappers masked failures.
- Filter XML can mirror flags for diagnostics, but Global XML remains authoritative—reset logic should always read/write there first.
- Automated refactors (e.g. `.ClashZones → .AllZones`) need manual cleanup of read-only assignments; unit compilation is the safety net.
- Removing duplicate refresh saves without reinstating serialization left `Plumbing.xml` empty—always persist after merging.
- Base-name normalization is non-negotiable; incorrect naming cascades into duplicate branches and stale flags.

## 7. Immediate Action Items
1. Remove the redundant `SaveClashZones` call during the master-filter serialization in placement.
2. Add a one-off dedupe helper in `ClashZonePersistenceService` to collapse duplicate Global XML entries by GUID.
   - Load Global XML, group by GUID, retain the newest entry (based on `ProcessedAt` or file order) per linked/host combo.
   - Drop extra nodes, re-save, and emit a summary log with before/after counts.
3. Harden `SaveClashZones` to prevent duplication going forward:
   - Normalize base name at entry, map normalized file-combo keys, merge updates into existing entries, skip when source equals target.
   - Emit `[PERSIST-CONSOLIDATE]` logs with counts to monitor.
4. Add assertions or guard logs in persistence to verify base filter names and clone counts before writing.
5. Keep this plan linked from `SLEEVE_PLACEMENT_METHODOLOGY.md` and update code comments to reference the rules above.

## 8. Duplication Mitigation Strategy
1. **Prevent new duplicates**
   - Enforce one `SaveClashZones` invocation per logical update.
   - Within `SaveClashZones`, build a dictionary keyed by `(GUID, normalized combo)` and update in place; refuse to append when data matches existing entry.
   - Ensure `GetBaseFilterName` always returns the canonical filter name before persistence work begins.
2. **Clean existing duplicates**
   - Implement `DeduplicateGlobalEntries` utility:
     - Iterate Global XML entries grouped by GUID.
     - For each group, prefer the entry whose `FilterName` matches the base filter and whose `Linked/Host` match normalized keys; if multiple remain, keep the newest `ProcessedAt`.
     - Remove any redundant entries and compact the tree.
   - Run once as part of the next refresh or via a maintenance command; log results to `placement_debug.log`.
3. **Validate after cleanup**
   - During flag reset, log per-category counts (`updates`, `resetCount`) and warn when mismatched duplicates are detected.
   - Add regression tests (or scripted checks) that load Global XML and assert unique GUID per combo.

## 9. Global XML Creation – Current Problems & Remediation Plan

### Pain Points Observed
- **Duplicate entries per GUID.** Multiple `SaveClashZones` invocations (or malformed base names) create parallel branches in Global XML (`Plumbing` vs `Plumbing_pipes`), leading to conflicting flag states.
- **Stale file-combo status.** When new linked files are selected, the pipeline processes them, but previously processed combos (still selected) may be skipped because Global XML already marks them processed. Deleted sleeves in those combos are never reset.
- **Flag reset misses.** Duplicated Global entries mean `FlagManager.ResetFlagsForDeletedSleeves` encounters a stale copy with `SleeveInstanceId = -1` and stops before updating the valid entry, leaving the deleted sleeve unresolved.
- **Processed key drift.** Normalization mismatches between detection and persistence produce divergent keys (e.g., path vs display name), so `DetectNewUiSelections` or `HasNewFileCombosInGlobal` misclassifies combos as already handled.
- **Missing audit trail.** Without explicit logging of combo keys and GUID counts per save, diagnosing gaps becomes guesswork.

### Remediation Plan
1. **Enforce normalized keys everywhere.**
   - Extend `ProcessedFileCombo.GetNormalizedKey()` usage across detection, persistence, and Global index updates.
   - Add guard logs inside persistence showing `Linked`, `Host`, `NormalizedKey` so mismatches surface quickly.
2. **Single data source per combo.**
   - During `SaveClashZones`, consolidate existing `FileCombo` entries by normalized key before merging new zones. If a combo exists, update in place; never append a second copy.
   - When saving category snapshots, reuse the same normalized key lookups to keep Filter XML and Global XML aligned.
3. **Dedupe Global XML on load.**
   - On each persistence call, load and prune duplicate entries before writing new ones (lightweight in-memory dedupe).
   - Provide a `RunGlobalIndexMaintenance()` helper for one-time cleanup of existing files (callable from refresh).
4. **Track processed combos explicitly.**
   - After persistence, update a per-category `ProcessedComboRegistry` (e.g., within Global XML metadata) listing all normalized keys touched in this run.
   - Refresh logic compares current UI selections against this registry rather than inferring from raw entries.
5. **Strengthen flag reset.**
   - In `FlagManager.ResetFlagsForDeletedSleeves`, iterate **all** entries for a GUID and ensure every copy is updated before saving.
   - Emit warning logs when a GUID has >1 entry, prompting maintenance.
6. **UI selection auditing.**
   - Enhance `DetectNewUiSelections` to log both the “requested” combo keys and the “processed” keys found in Global XML. When the sets diverge, force a detection run even if the combo was previously processed.
7. **Regression validation.**
   - Add diagnostics after refresh summarizing, per combo: clash count, resolved flag count, and number of sleeves in Revit. When a combo shows deleted sleeves without resets, highlight it in `Refresh_*.log`.
8. **Future-proofing.**
   - Consider moving Global index generation behind a unit-testable builder that accepts normalized combos, flags, and placement metadata, returning a deduped structure before serialization.

With the above changes, new UI selections will still be processed, while previously processed linked files remain tracked and their deletions detected. Global XML will act as a stable, deduplicated ledger, enabling accurate flag resets and preventing re-placement gaps.

## 10. Implementation Roadmap

### Phase A – Stabilize Persistence Calls (Code-Level)
1. **`Services/UniversalSleevePlacerService.cs`**
   - Target method: `SaveUpdatedXmlFiles`.
   - Delete block that re-invokes `persistenceService.SaveClashZones` for the main filter (lines ≈3748–3810). Replace with:
     ```csharp
     var mainFilterPath = Path.Combine(filtersDirectory, $"{baseFilterName}.xml");
     if (File.Exists(mainFilterPath))
     {
         using var reader = new StreamReader(mainFilterPath);
         var mainFilter = (OpeningFilter)serializer.Deserialize(reader);
         // merge clones into mainFilter in-memory only
         MergeZonesIntoFilter(mainFilter, updatedClashZones);
         filterManagementService.SaveFilterToXmlFile(mainFilter, mainFilterPath);
     }
     ```
   - Add assertion/log before serialization:
     ```csharp
     if (string.IsNullOrWhiteSpace(baseFilterName))
         throw new InvalidOperationException("[PERSIST] Base filter name missing.");
     ```
2. **`Services/ClashZonePersistenceService.cs`**
   - At top of `SaveClashZones`, insert:
     ```csharp
     baseFilterName = NormalizeBaseFilterName(baseFilterName);
     ```
   - Implement `NormalizeBaseFilterName` near helper section:
     ```csharp
     private static string NormalizeBaseFilterName(string rawName)
     {
         if (string.IsNullOrWhiteSpace(rawName)) return "Unknown";
         var name = Path.GetFileNameWithoutExtension(rawName.Trim());
         return name;
     }
     ```
   - Create `DeduplicateInMemory(OpeningFilter filter)`:
     - Iterate `filter.ClashZoneStorage.Filters`.
     - Group by `FilterFileComboGroup.GetNormalizedKey()` (add helper).
     - Merge clash zones, drop duplicates, update `ProcessedAt` to latest.
   - Before grouping `clashZonesByFileCombo`, call `DeduplicateInMemory(targetFilter)`.
   - Update loop merging zones:
     ```csharp
     if (existingByGuid.TryGetValue(newZone.Id, out var targetZone))
     {
         if (!ZonesEqual(targetZone, newZone))
             MergeZone(targetZone, newZone, logPath);
     }
     else
     {
         if (!ShouldSkipAdd(newZone))
             existingCombo.ClashZones.Add(newZone);
     }
     ```
   - Add `[PERSIST-KEY]` log:
     ```csharp
     File.AppendAllText(logPath,
         $"[...][PERSIST-KEY] Combo={normalizedKey} Before={beforeCount} After={afterCount}");
     ```

### Phase B – Global XML Dedupe & Cleanup
1. **New helper class** `Services/GlobalIndexMaintenance.cs`
   ```csharp
   public static class GlobalIndexMaintenance
   {
       public static bool Deduplicate(Document doc, string category, StringBuilder report = null)
       {
           var index = GlobalIndexService.LoadCategoryIndex(doc, category);
           var map = new Dictionary<string, CategoryGlobalIndexEntry>(StringComparer.OrdinalIgnoreCase);
           int removed = 0;

           foreach (var filter in index.Filters)
           {
               foreach (var combo in filter.FileCombos)
               {
                   var normalizedKey = combo.GetNormalizedKey();
                   combo.ClashZones = combo.ClashZones?
                       .GroupBy(z => z.Id)
                       .Select(g => g.OrderByDescending(z => combo.ProcessedAt).First())
                       .ToList();
               }
           }
           // flatten GUIDs and remove duplicates
           // ...
       }
   }
   ```
   - For each GUID group, keep latest entry (based on `ProcessedAt` or XML order).
   - Return true if file mutated; caller re-saves via `GlobalIndexService.Save`.
2. **Integration:**
   - Add call in `ClashZonePersistenceService.SaveClashZones` before writing Global XML:
     ```csharp
     GlobalIndexMaintenance.Deduplicate(_document, category, diagnosticLog);
     ```
   - Log summary when duplicates removed.

### Phase C – Flag Reset Hardening
1. **`Services/FlagManager.cs`**
   - In `ResetFlagsForDeletedSleeves`, replace dictionary lookup with:
     ```csharp
     var entries = GetEntriesForGuid(globalEntries, guid);
     foreach (var entry in entries)
     {
         entry.IsResolved = false;
         entry.SleeveInstanceId = -1;
         // ...
     }
     ```
   - Add log when `entries.Count > 1`:
     ```csharp
     LogToRefresh($"[FLAG-MANAGER] GUID {guid} had {entries.Count} entries; all reset.");
     ```

### Phase D – Processed Combo Tracking
1. **Schema update** (`CategoryGlobalIndex`):
   ```xml
   <ProcessedCombos>
       <Combo Key="plumbing::fp-00001|st-00001" ProcessedAt="2025-11-11T12:24:00Z" />
   </ProcessedCombos>
   ```
2. **`GlobalIndexService.UpsertFlagsWithIdsAndClashZoneData`**
   - After handling entries, ensure combo key exists in `ProcessedCombos` with `ProcessedAt = DateTime.Now`.
3. **`RefreshService.DetectNewUiSelections`**
   - Replace direct entry lookup with:
     ```csharp
     var processedCombos = GlobalIndexService.GetProcessedComboKeys(document, category);
     if (!processedCombos.Contains(requestedKey))
         return true;
     ```

### Phase E – Diagnostics & Regression Checks
1. Add method `LogPersistenceSummary(category, comboKey, zonesBefore, zonesAfter, resolvedCount)` inside `ClashZonePersistenceService`.
2. Create script `scripts/validate_global_duplicates.py`:
   ```python
   import xml.etree.ElementTree as ET
   # scan for duplicate GUIDs per combo and report
   ```
3. Update docs (`SLEEVE_PLACEMENT_METHODOLOGY.md`) referencing new combo tracking and dedupe requirements.

### Phase F – Optional Enhancements
- Add unit tests under `Tests/Persistence/ClashZonePersistenceTests.cs` verifying dedupe logic.
- Extract Global index build/update to `GlobalIndexBuilder` class to simplify future maintenance.

## 11. SQLite Migration & Application Simplification Plan

### 11.1 Target State Overview
- **Single SQLite database file** per project (e.g. `SleevePersistence.db`).
- Schema captures both flag state and placement data—no separate Global/Filter XML.
- All persistence operations occur inside ACID transactions; WAL mode enabled for crash recovery.
- XML files become optional exports (for compatibility/debugging) generated from the database on demand.
- Detection and placement operate on a shared persistence context; UI reflects database state live.

### 11.2 Proposed Schema (Initial Draft)
```sql
CREATE TABLE Filters (
    FilterId      INTEGER PRIMARY KEY AUTOINCREMENT,
    FilterName    TEXT NOT NULL UNIQUE,
    Category      TEXT NOT NULL,
    CreatedAt     DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
    UpdatedAt     DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE FileCombos (
    ComboId       INTEGER PRIMARY KEY AUTOINCREMENT,
    FilterId      INTEGER NOT NULL REFERENCES Filters(FilterId) ON DELETE CASCADE,
    LinkedFileKey TEXT NOT NULL,
    HostFileKey   TEXT NOT NULL,
    ProcessedAt   DATETIME,
    UNIQUE(FilterId, LinkedFileKey, HostFileKey)
);

CREATE TABLE ClashZones (
    ClashZoneId   INTEGER PRIMARY KEY AUTOINCREMENT,
    ComboId       INTEGER NOT NULL REFERENCES FileCombos(ComboId) ON DELETE CASCADE,
    MepElementId  INTEGER NOT NULL,
    HostElementId INTEGER NOT NULL,
    IntersectionX REAL NOT NULL,
    IntersectionY REAL NOT NULL,
    IntersectionZ REAL NOT NULL,
    SleeveState   INTEGER NOT NULL,          -- enum (0=Unprocessed, 1=IndividualPlaced, ...)
    SleeveInstanceId INTEGER,
    ClusterInstanceId INTEGER,
    SleeveWidth   REAL,
    SleeveHeight  REAL,
    SleeveDiameter REAL,
    SleevePlacementX REAL,
    SleevePlacementY REAL,
    SleevePlacementZ REAL,
    BoundingBoxMinX REAL,
    BoundingBoxMinY REAL,
    BoundingBoxMinZ REAL,
    BoundingBoxMaxX REAL,
    BoundingBoxMaxY REAL,
    BoundingBoxMaxZ REAL,
    PlacementSource TEXT,                    -- XML snapshot, recalculated, etc.
    UpdatedAt     DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
    UNIQUE(ComboId, MepElementId, HostElementId, IntersectionX, IntersectionY, IntersectionZ)
);

CREATE TABLE SleeveEvents (
    EventId       INTEGER PRIMARY KEY AUTOINCREMENT,
    ClashZoneId   INTEGER NOT NULL REFERENCES ClashZones(ClashZoneId) ON DELETE CASCADE,
    EventType     TEXT NOT NULL,             -- "Placed", "Deleted", "Clustered"
    Payload       TEXT,                      -- JSON for diagnostics
    CreatedAt     DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE Conditions (
    ConditionId   INTEGER PRIMARY KEY AUTOINCREMENT,
    FilterId      INTEGER NOT NULL REFERENCES Filters(FilterId) ON DELETE CASCADE,
    Category      TEXT NOT NULL,
    RectNormal    REAL,
    RectInsulated REAL,
    RoundNormal   REAL,
    RoundInsulated REAL,
    PipesNormal   REAL,
    PipesInsulated REAL,
    CableTrayTop  REAL,
    CableTrayOther REAL,
    OpeningPrefs  TEXT,                      -- JSON blob capturing opening type preferences
    UpdatedAt     DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
    UNIQUE(FilterId, Category)
);
```
- **Indexes** on `SleeveState`, `MepElementId`, `HostElementId` for fast lookups.
- Additional tables can track configuration, clearances, or UI selections as needed.

### 11.3 Migration Phases

#### Phase SQLite-1 – Prototype (1–2 sprints)
1. Create new data access layer (`SleeveDbContext`) using `System.Data.SQLite`.
2. Implement schema creation & migrations (build simple migration runner).
3. Mirror existing XML persistence into database (write-through mode): keep writing XML but also insert rows.
4. Build internal tooling to inspect DB (simple WPF/CLI viewer or SQLite browser instructions).

#### Phase SQLite-2 – Cutover (2–3 sprints)
1. Refactor `ClashZonePersistenceService` to operate on `SleeveDbContext` transactions instead of XML.
2. Update `RefreshService`, `UniversalSleevePlacerService`, `UniversalClusterService`, and `SleeveCoordinateService` to query/update the database:
   - Detection results → `InsertOrUpdateClashZones`.
   - Placement → update `SleeveState`, dimensions, bounding boxes.
   - Cluster placement → mark related zones, insert cluster events.
3. Replace flag reset logic with SQL queries (e.g. `UPDATE ClashZones SET SleeveState = ? WHERE SleeveInstanceId NOT IN (...)`).
4. Introduce write-ahead logging (WAL) and checkpoint strategy for reliability.
5. Ensure all operations happen inside `using var tx = db.BeginTransaction();` blocks; commit once per high-level operation.

#### Phase SQLite-3 – Legacy Decommission (1 sprint)
1. Remove direct XML reads/writes from runtime. Provide explicit “Export XML” and “Import XML” utilities for interoperability.
2. Strip dedupe/normalization logic specific to XML; rely on database constraints.
3. Update documentation (`SLEEVE_PLACEMENT_METHODOLOGY.md`) to describe the new persistence model.
4. Provide data migration tool `XmlToSqlite.exe` to convert existing projects.

### 11.4 Application Simplifications Enabled
- **Single source of truth:** one DB file eliminates cross-file sync.
- **Query flexibility:** dynamic filtering, aggregations, and diagnostics via SQL.
- **Event logging:** `SleeveEvents` table becomes audit trail; can power real-time UI updates.
- **Conditions inline:** per-filter/per-category clearances and opening preferences reside in the `Conditions` table, so placement logic reads sizes directly from SQLite instead of `*_CONDITIONS.xml`.
- **State enums, strong typing:** map directly to integer columns; enforced through DAL.
- **Performance:** WAL + indexed queries remove need for in-memory dedupe. Batch updates (e.g. cluster placements) run in sub-second durations even at scale.

### 11.5 Supporting Changes
- Introduce repository/service layer abstractions (e.g. `IClashZoneRepository`) to isolate SQLite from business logic.
- Add automated tests using in-memory SQLite (`DataSource = ":memory:";`) to validate persistence operations.
- Implement migration/versioning metadata table to track schema upgrades.
- Provide fallback export command `SleevePersistenceExporter --format xml --output ...` for clients still expecting XML payloads.

### 11.6 Risks & Mitigations
- **Migration complexity:** Use dual-write period (Phase SQLite-1) to verify parity before cutover.
- **File locking on network shares:** Configure SQLite to use `PRAGMA journal_mode=WAL;` and document best practices.
- **Corruption concerns:** Enable automatic backups (copy DB before major operations) and provide `sqlite3 .recover` guidelines.
- **Learning curve:** Create developer docs/tutorial for using `SleeveDbContext`, including sample queries.
- **Conditions parity:** During dual-write, keep writing `*_CONDITIONS.xml` from `Conditions` rows so legacy tooling continues to function until XML export is officially deprecated.

### 11.7 Success Criteria
- All existing workflows (refresh, placement, clustering, flag reset) operate exclusively against SQLite.
- No duplicate or conflicting clash zone entries (enforced by unique constraints).
- Revit sessions can recover gracefully after crash (thanks to WAL and transaction boundaries).
- Exported XML (if requested) matches legacy format for downstream consumers.
- Performance benchmarks show reduced runtime and smaller persistence footprint compared to XML.

## 12. Near-Term Stabilization vs Migration Strategy

### 12.1 Stabilize XML for Current Production
- Implement **Phase A** items immediately:
  - Remove redundant `SaveClashZones` invocation in `UniversalSleevePlacerService`.
  - Normalize base names and deduplicate in-memory filter structures before persistence.
  - Add guard logs/assertions to catch malformed input early.
- Run **Phase B** Global XML dedupe utility once to clean existing files so refresh/reset logic behaves predictably.
- Wrap persistence calls in transaction-style try/finally blocks to ensure “save before scrub” happens reliably even before SQLite.
- Defer deeper XML refactors (state enums, append-only overhaul) to avoid double work.

### 12.2 Begin SQLite Migration in Parallel
- As soon as XML writes are stable (no new duplicates), start **Phase SQLite-1**: dual-write to both XML and SQLite to validate schema and data parity.
- Keep XML as the operational store during this phase; treat SQLite as a mirror for verification.
- Once parity is confirmed, progress through **Phase SQLite-2/3** to cut over and retire XML writers.

### 12.3 Guiding Principle
- **Fix forward:** apply only minimal necessary fixes to keep XML reliable for day-to-day use, channel further effort into the SQLite path where long-term complexity is resolved.

## 13. Conditions XML Lifecycle (from UI to Placement)

### 13.1 When `*_CONDITIONS.xml` Is Created or Updated
1. **UI interaction (`EmergencyMainDialog.SaveConditionsToXml`)**
   - Trigger: user clicks **Place Sleeves** (or any action that persists the current configuration).
   - Collects:
     - Selected filter(s) from the UI list.
     - Selected categories (horizontal/vertical host types).
     - Current clearance settings and opening type preferences.
   - For each `(filterName, category)` pair, builds a **combined key** (e.g. `Plumbing_pipes`) and instantiates `ConditionsService` with the active project’s `Filters` directory:  
     `var conditionsService = new ConditionsService(projectFiltersDir, …);`
   - Creates an `OpeningConditions` object populated from the UI state and calls:  
     `conditionsService.SaveConditions(conditions, combinedKey);`
   - This writes `${combinedKey}_CONDITIONS.xml` alongside `Plumbing_pipes.xml`.

2. **File format**
   - `<OpeningConditions>` XML contains:
     - `FilterName`, `Category` (the combined key).
     - `ClearanceSettings` (rectangular/round/pipes/cable tray values in mm).
     - `OpeningTypePreferences` (per category decisions such as circular vs rectangular).
     - Timestamps for auditing.

### 13.2 Loading Conditions Before Placement
1. **Command entry (`UniversalSleevePlacementCommand.Execute`)**
   - The orchestrator gathers selected filters and linked/host files, then instantiates `UniversalSleevePlacementCommand`.
2. **`LoadConditionsFromXml()`**
   - Reconstructs the combined key using the first selected filter name and the current category:  
     `string combinedKey = $"{filterName}_{NormalizeCategoryName(_category)}";`
   - Creates `ConditionsService` pointing at the same project `Filters` directory.
   - Invokes `conditionsService.LoadConditions(combinedKey);`
     - If the file exists, it deserializes the stored clearances and opening preferences.
     - If missing, it returns defaults and immediately calls `SaveConditions` so the file is created for future runs.
   - Diagnostic logging records the path (`CONDITIONS_DIR`, `CONDITIONS_CREATED`, `CONDITIONS_EXISTS`) in `orchestrator_debug.log` for traceability.
3. **Propagation to the placer**
   - The loaded `OpeningConditions` instance is passed into `UniversalSleevePlacerService` via its constructor.
   - During placement (`GetClearanceFromConditions`), the service reads either:
     - UI overrides provided at runtime (`_clearanceSettings` dictionary), or
     - The persisted XML values via `_conditions`.
   - This ensures every sleeve placement uses the same clearances and opening preferences the user set before clicking **Place Sleeves**.

### 13.3 Summary Flow
```
UI (EmergencyMainDialog)
   └─ SaveConditionsToXml()  → ConditionsService.SaveConditions()
                                 ↓
                           <Filter>CATEGORY_CONDITIONS.xml

Placement Command (UniversalSleevePlacementCommand)
   └─ LoadConditionsFromXml() → ConditionsService.LoadConditions()
                                 ↓
                         OpeningConditions injected into UniversalSleevePlacerService

UniversalSleevePlacerService
   └─ GetClearanceFromConditions() / GetClearanceFromXmlConditions()
      → uses OpeningConditions for sizing and placement decisions
```

This lifecycle guarantees the placer always reflects the latest UI state, while also creating the `*_CONDITIONS.xml` if it is missing so subsequent runs operate with the same configuration.



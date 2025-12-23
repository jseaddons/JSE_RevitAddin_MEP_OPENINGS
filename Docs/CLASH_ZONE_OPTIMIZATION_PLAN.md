# Optimization Plan: ClashZoneService Lookup & Logging

## Goal Description
Drastically improve `Refesh` command performance (currently taking ~37s for ~200 items) by optimizing the `ClashZoneService` lookup mechanism and removing excessive logging.
Currently: `FindExistingClashZone` allocates a new List (`.ToList()`) and performs linear search (`O(N)`) for **every** detected intersection.
Proposed: Pre-index existing ClashZones into a `Dictionary` (Key: `MepId-StructId`) once per batch, enabling `O(1)` lookup. Also reduce logging verbosity.

## User Review Required
None.

## Proposed Changes

### [MODIFY] Services/ClashZoneService.cs

1.  **Introduce Lookup Cache**:
    *   Add `private Dictionary<(int MepId, int StructId), List<ClashZone>> _clashZoneLookup;`
    *   Add method `BuildClashZoneLookup()` which clears and repopulates this dictionary from `_clashZoneStorage.ClashZones`.

2.  **Optimize `FindExistingClashZone`**:
    *   Replace `var clashZonesList = _clashZoneStorage?.ClashZones?.ToList()` with a check against `_clashZoneLookup`.
    *   If lookup is null or empty, rebuild it (lazy initialization or explicit build at start of detection).
    *   Use the dictionary to retrieve candidates for `(mepId, structId)`.
    *   Perform specific checks (Guid, Point Match) only on this small candidate list.
    *   REMOVE `.ToList()` call which allocates memory 1000s of times.
    *   **Disable/Reduce Logging**: Comment out or wrap non-critical logs (like "Checking ... against X existing zones") in `if (OptimizationFlags.UseDiagnosticMode)`.

3.  **Update `DetectNewClashZones` (or equivalent batch method)**:
    *   Call `BuildClashZoneLookup()` at the very beginning of the batch process (before the loop over intersections).

### [MODIFY] Services/GuidManager.cs (if accessible/used)
*   Ensure methods like `FindByRevitSleeveGuid` accept a pre-filtered list or index, rather than iterating the whole storage.

## Verification Plan

### Manual Verification
1.  **Performance Test**: Run `Refresh` command again.
2.  **Expectation**: "Intersection Processing" time should drop from ~37s to <5s.
3.  **Log Check**: Ensure "Checking ... against 262 existing clash zones" lines are gone/reduced.

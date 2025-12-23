# Batch Mode Optimization Plan

## Problem Statement
Current performance: **24 zones/sec** (target: 500+)

### Critical Bottlenecks
1. **Element Collection: 13.2s** (49% of Intersection Processing time)
2. **Database Save: 18.5s** (49% of total time)
3. **Clash Zone Creation: 9.2s** (34% of Intersection Processing time)

---

## Bottleneck #1: Element Collection (13.2s → <1s)

### Root Cause Analysis
The `6a. Element Collection` operation in `IntersectionProcessor.RunDetectionWithCollectorLevelFilters` is taking 13.2 seconds to collect MEP and structural elements.

**Potential Causes:**
1. **Section Box Filtering Overhead**: `GetSectionBoxOutline` may be slow
2. **Inefficient Filters**: Multiple filters being applied sequentially
3. **Large Element Count**: Collecting too many elements before filtering
4. **Revit API Overhead**: FilteredElementCollector performance degradation

### Proposed Solution

#### Option A: Add Performance Tracking (Diagnostic)
Add sub-timings to identify which collection phase is slow:
- MEP element collection time
- Host element collection time
- Section box outline creation time

#### Option B: Optimize Collection Strategy
1. **Cache Section Box Outline**: Calculate once, reuse for both MEP and host collections
2. **Parallel Collection**: Collect MEP and host elements in parallel (if safe)
3. **Optimize Filters**: Ensure filters are applied at collector level, not post-filtering

### Implementation

**File**: `refresh refactor/intersection_processor.cs`

**Changes**:
1. Add performance sub-timings for MEP vs Host collection
2. Cache section box outline
3. Investigate if parallel collection is safe

---

## Bottleneck #2: Database Save (18.5s → <2s)

### Root Cause Analysis
Even with `UseBulkSqliteUpdates = true`, the save operation takes 18.5 seconds because:

1. **File Combo Creation Loop** (Lines 163-176 in `ClashZoneRepository.cs`):
   - Loops through 888 zones
   - Calls `GetOrCreateFileCombo` for each zone
   - Each call executes 1-2 SQL queries
   - **Estimated time**: ~10-15 seconds

2. **Individual INSERT Loop** (Lines 294-297):
   - Loops through new zones
   - Calls `InsertClashZone` for each zone
   - Each call executes 1 INSERT statement
   - **Estimated time**: ~5-10 seconds

### Proposed Solution

#### Phase 1: Batch File Combo Creation
**Current**: One query per zone (888 queries)
```csharp
foreach (var zone in zonesList) {
    var comboId = GetOrCreateFileCombo(...); // 1-2 queries per zone
}
```

**Optimized**: Single batch query
```csharp
// 1. Extract all unique (LinkedFile, HostFile, Category) tuples
// 2. Single SELECT to find existing combos
// 3. Single multi-row INSERT for new combos
// 4. Map zones to comboIds
```

**Expected Improvement**: 10-15s → <1s

#### Phase 2: Batch INSERT
**Current**: One INSERT per new zone
```csharp
foreach (var zone in actuallyNew) {
    InsertClashZone(comboId, zone, transaction); // 1 INSERT
}
```

**Optimized**: Single multi-row INSERT
```sql
INSERT INTO ClashZones (col1, col2, ...) VALUES
    (val1, val2, ...),
    (val1, val2, ...),
    ...
```

**Expected Improvement**: 5-10s → <0.5s

### Implementation

**File**: `Data/Repositories/ClashZoneRepository.cs`

**New Methods**:
1. `BatchGetOrCreateFileCombos(zones, filterId, transaction)` → Dictionary<Guid, int>
2. `BulkInsertClashZones(zones, comboMap, transaction)`

**Changes to `InsertOrUpdateClashZonesBulk`**:
- Replace lines 163-176 with `BatchGetOrCreateFileCombos`
- Replace lines 294-297 with `BulkInsertClashZones`

---

## Bottleneck #3: Clash Zone Creation (9.2s → <2s)

### Root Cause Analysis
The `foreach` loop in `ClashZoneService.DetectNewClashZones` (line 377) processes 888 zones one-by-one, taking ~10ms per zone.

**Per-Zone Overhead**:
- Wall centerline calculation
- Category lookups
- Damper proximity checks
- Geometry calculations

### Proposed Solution

#### Option A: Pre-calculate Common Data
1. **Build Wall Lookup Cache**: Map wall IDs to centerline calculation data
2. **Build Category Cache**: Pre-fetch all MEP categories
3. **Build Damper Location Cache**: Already implemented (line 242)

#### Option B: Parallelize Creation (Risky)
Use `Parallel.ForEach` to create zones concurrently.
- **Risk**: Thread safety issues with Revit API
- **Benefit**: 2-4x speedup on multi-core systems

### Implementation

**File**: `Services/ClashZoneService.cs`

**Controlled by**: `OptimizationFlags.UseBatchClashZoneCreation`

**Changes**:
1. Pre-calculate wall centerline data for all walls in section box
2. Cache MEP element categories
3. Consider parallelization if safe

---

## Expected Performance Improvements

| Operation | Current | Optimized | Improvement |
|-----------|---------|-----------|-------------|
| Element Collection | 13.2s | <1s | **13x faster** |
| Database Save | 18.5s | <2s | **9x faster** |
| Clash Zone Creation | 9.2s | <2s | **4.5x faster** |
| **Total Time** | **37.6s** | **<8s** | **4.7x faster** |
| **Zones/Second** | 24 | **111+** | **4.6x faster** |

---

## Implementation Priority

### Phase 1: Database Save Optimization (Highest Impact)
1. ✅ Add `UseBulkSqliteUpdates` flag (already exists, set to true)
2. Implement `BatchGetOrCreateFileCombos`
3. Implement `BulkInsertClashZones`
4. Test with 888 zones

**Expected Result**: 18.5s → <2s

### Phase 2: Element Collection Investigation
1. Add performance sub-timings
2. Identify slow phase (MEP vs Host vs Section Box)
3. Optimize based on findings

**Expected Result**: 13.2s → <1s

### Phase 3: Clash Zone Creation Optimization
1. Add `UseBatchClashZoneCreation` flag
2. Implement pre-calculation caches
3. Test with 888 zones

**Expected Result**: 9.2s → <2s

---

## Risk Assessment

### Low Risk
- ✅ Database save batch optimization (uses transactions, can rollback)
- ✅ Element collection performance tracking (diagnostic only)

### Medium Risk
- ⚠️ Clash zone creation pre-calculation (needs validation)

### High Risk
- ❌ Parallel clash zone creation (Revit API thread safety)

---

## Next Steps

1. **Implement Phase 1**: Database save batch optimization
2. **Test**: Run "Process Clash Zones" with 888 zones
3. **Validate**: Verify data integrity (no missing zones, correct parameters)
4. **Implement Phase 2**: Element collection investigation
5. **Implement Phase 3**: Clash zone creation optimization

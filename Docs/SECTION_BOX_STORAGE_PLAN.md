# Phase 2 Optimization: DB-First Architecture

## Philosophy: "Dump Once, Use Many Times"

```
REFRESH (Main Thread)          →  Store in DB  →  All Services (Multi-Threaded)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Read Section Box              →  SessionContext   →  Filter sleeves by bounds
Read Family Names             →  ClashZones       →  Identify sleeve types
Read MEP Parameters           →  SleeveSnapshots  →  Parameter Transfer
Calculate Marks               →  (in memory)       →  Batch Write to Revit
```

---

## Benefits

| Current | Optimized |
|---------|-----------|
| Multiple Revit API calls per service | READ: DB only (fast, multi-threaded) |
| Sequential processing | Parallel processing |
| Multiple transactions | Single batch transaction |
| ~1000ms per operation | ~100ms per operation (10x faster) |

---

## Data to Store in DB During Refresh

### 1. SessionContext Table (New)
```sql
CREATE TABLE SessionContext (
    Key TEXT PRIMARY KEY,
    Value TEXT,
    UpdatedAt DATETIME
);

-- Section Box Bounds
INSERT INTO SessionContext (Key, Value) VALUES 
('SectionBoxMinX', '-100.5'), ('SectionBoxMinY', '-50.0'), ('SectionBoxMinZ', '0.0'),
('SectionBoxMaxX', '100.5'),  ('SectionBoxMaxY', '50.0'),  ('SectionBoxMaxZ', '20.0'),
('SectionBoxIsActive', '1');
```

### 2. ClashZones Table (Existing - Add Columns)
```sql
-- Already exists: SleeveInstanceId, PlacementLocationX/Y/Z, MepCategory
-- ADD: FamilyName, IsCombinedSleeve flag
ALTER TABLE ClashZones ADD FamilyName TEXT;
ALTER TABLE ClashZones ADD IsCombinedSleeve INTEGER DEFAULT 0;
```

### 3. CombinedSleeves Table (Existing)
```sql
-- Already exists: CombinedInstanceId, PlacementLocationX/Y/Z
-- ADD: FamilyName for consistency
ALTER TABLE CombinedSleeves ADD FamilyName TEXT;
```

---

## Service Flow: GET from DB, SET to Revit Once

### Apply Marks Flow
```
1. READ (DB, Multi-threaded):
   - GetSleevesInSectionBox(min, max) → List<SleeveId>
   - GetFamilyNames(sleeveIds) → no Revit lookup needed
   - GetCombinedSleeveFlag(sleeveIds) → identify combined sleeves

2. CALCULATE (Memory, Parallel):
   - Sort by Level → Y → X
   - Assign sequential marks with prefix

3. WRITE (Revit, Single Transaction):
   - Batch write all marks in one transaction
   - doc.GetElement() only for final write
```

### Parameter Transfer Flow
```
1. READ (DB, Multi-threaded):
   - GetSnapshotsBySleeveIds() → parameters from DB
   - No Revit API calls

2. CALCULATE (Memory, Parallel):
   - Aggregate parameters for combined sleeves
   - Normalize Size values

3. WRITE (Revit, Single Transaction):
   - Batch write all parameters
```

---

## Implementation Checklist

- [ ] Add `SessionContext` table to SleeveDbContext
- [ ] Add `FamilyName` column to ClashZones
- [ ] Add `IsCombinedSleeve` column to ClashZones
- [ ] Create `ISectionBoxService` interface
- [ ] Implement `SectionBoxService.CaptureAndStore()`
- [ ] Store family names during Refresh
- [ ] Update `GetAllCombinedSleeves()` to query DB instead of Revit
- [ ] Update `GetAllSleevesForCategory()` to query DB
- [ ] Implement batch write for Apply Marks
- [ ] Implement batch write for Parameter Transfer

---

## Summary

**REFRESH**: Dump all data to DB (section box, family names, parameters)  
**ALL SERVICES**: Read from DB only (fast, parallel)  
**FINAL WRITE**: Single batch transaction to Revit (minimizes BIM 360 sync)

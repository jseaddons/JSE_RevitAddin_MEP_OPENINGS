# Three Paths - Correct Explanation (DB-Driven)

## Overview

The refresh process has **3 distinct paths** based on:
1. **IsFilterComboNew** flag from **FileCombos table (DATABASE)**
2. **enableThreePointValidation** setting (Adopt to Modified Document)

**All operations are now DB-driven, not XML-driven.**

## Path Decision Logic

```
IF IsFilterComboNew = 0 (from DB)
  → PATH 1: Replay Mode
ELSE IF IsFilterComboNew = 1 (from DB) AND enableThreePointValidation = false
  → PATH 2: Fresh Placement Mode (NO validation)
ELSE IF enableThreePointValidation = true
  → PATH 3: Full Detection with Validation
```

---

## PATH 1: REPLAY MODE (Red)

### Trigger:
- **IsFilterComboNew = 0** (from FileCombos table in DB)
- File combos already processed in database

### Actions (All DB-Driven):
1. ✅ **Skip intersection detection** - No new detection needed
2. ✅ **Load existing clash zones from Database** (PRIMARY source)
3. ✅ **ResetFlagsForDeletedSleeves** - Update flags in DB for deleted sleeves
4. ✅ **Skip MergeAndSave** - No database save of new zones
5. ✅ **Go directly to Place Sleeves** - Only for deleted sleeves

### Database Operations:
- **READ**: Load clash zones from `ClashZones` table
- **UPDATE**: Update flags in `ClashZones` table for deleted sleeves
- **NO INSERT**: No new clash zones created

### When Used:
- Files haven't changed since last refresh
- User just wants to refresh flags for deleted sleeves
- Fast path - minimal processing

---

## PATH 2: FRESH PLACEMENT MODE (Orange)

### Trigger:
- **IsFilterComboNew = 1** (from FileCombos table in DB - new file combos)
- **AND enableThreePointValidation = false** (Adopt to Modified Document = OFF)

### Actions (All DB-Driven):
1. ✅ **Run intersection detection** - Find all MEP vs Structural intersections
2. ✅ **SKIP 3-Point Validation** - No validation required (fresh placement)
3. ✅ **MergeAndSave to Database** - Save all clash zones to DB
4. ✅ **Database save sequence**:
   - Get/Create Filter (DB)
   - Get/Create FileCombo (DB, IsFilterComboNew=1)
   - Insert/Update ClashZones (DB)
   - Insert/Update SleeveSnapshots (DB)
   - Commit Transaction

### Database Operations:
- **INSERT**: Create new FileCombo with IsFilterComboNew=1
- **INSERT/UPDATE**: Save clash zones to `ClashZones` table
- **INSERT/UPDATE**: Save sleeve snapshots to `SleeveSnapshots` table

### When Used:
- New file combinations detected
- "Adopt to modified document" is disabled
- Fresh placement - no validation needed
- First time processing a file combo

### Key Distinction:
- **NO 3-point validation** - This is the difference from PATH 3
- Fresh placement assumes all zones are valid
- Faster processing (no validation overhead)

---

## PATH 3: FULL DETECTION WITH VALIDATION (Green)

### Trigger:
- **enableThreePointValidation = true** (Adopt to Modified Document = ON)
- **OR** validation is required for existing zones

### Actions (All DB-Driven):
1. ✅ **Run intersection detection** - Find all MEP vs Structural intersections
2. ✅ **3-Point Validation ENABLED** - Validate:
   - MEP Element exists
   - Structural Element exists
   - Elements still intersect
3. ✅ **Validate existing zones** - Check all existing clash zones from DB
4. ✅ **MergeAndSave to Database** - Save validated clash zones to DB
5. ✅ **Database save sequence**:
   - Get/Create Filter (DB)
   - Get/Create FileCombo (DB, IsFilterComboNew=1)
   - Insert/Update ClashZones (DB)
   - Insert/Update SleeveSnapshots (DB)
   - Commit Transaction

### Database Operations:
- **READ**: Load existing clash zones from `ClashZones` table for validation
- **INSERT**: Create new FileCombo with IsFilterComboNew=1
- **INSERT/UPDATE**: Save validated clash zones to `ClashZones` table
- **UPDATE**: Update existing zones if intersection points changed
- **INSERT/UPDATE**: Save sleeve snapshots to `SleeveSnapshots` table

### When Used:
- "Adopt to modified document" is enabled
- Need to validate existing zones (elements may have moved)
- Need to update intersection points if elements changed
- Full validation required for accuracy

### Key Distinction:
- **WITH 3-point validation** - This is the difference from PATH 2
- Validates all zones before saving
- Updates intersection points if elements moved
- More thorough but slower processing

---

## Summary Table

| Path | Trigger | Validation | Database Save | Use Case |
|------|---------|------------|---------------|----------|
| **PATH 1** | IsFilterComboNew = 0 (DB) | ❌ Skip | ❌ No | Replay existing zones |
| **PATH 2** | IsFilterComboNew = 1 (DB) AND enableThreePointValidation = false | ❌ Skip | ✅ Yes | Fresh placement, no validation |
| **PATH 3** | enableThreePointValidation = true | ✅ Yes | ✅ Yes | Full detection with validation |

## Key Differences

### PATH 2 vs PATH 3:
- **PATH 2**: Fresh placement, **NO validation** (enableThreePointValidation = false)
- **PATH 3**: Full detection, **WITH validation** (enableThreePointValidation = true)

### All Paths are DB-Driven:
- ✅ FileCombos table checked from DB (not XML)
- ✅ Clash zones loaded from DB (not XML)
- ✅ Flags updated in DB (not XML)
- ✅ All saves go to DB (not XML)
- ✅ XML is fallback only (if DB has no data)

## Flowchart File

Import `THREE_PATHS_CORRECT.drawio` into draw.io to see the complete visual flow with DB-driven operations.


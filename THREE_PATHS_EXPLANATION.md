# Three Paths Flowchart - Complete Explanation

## Overview

The refresh process has **3 distinct paths** based on two key triggers:
1. **Adopt to Modified Document** setting (ON/OFF)
2. **IsFilterComboNew** flag in FileCombos table (0/1)

## Path Decision Logic

```
IF Adopt = ON
  → PATH 3: Full Detection (always)
ELSE IF Adopt = OFF
  → Check IsFilterComboNew flag:
    - IF IsFilterComboNew = 0 → PATH 1: Replay Mode
    - IF IsFilterComboNew = 0 AND no new combos → PATH 2: Replace Mode
    - IF IsFilterComboNew = 1 → PATH 3: Full Detection
```

## PATH 1: REPLAY MODE (Red)

### Trigger:
- **IsFilterComboNew = 0** (File combos already processed)
- Files haven't changed since last refresh

### Actions:
1. ✅ **Skip intersection detection** - No new detection needed
2. ✅ **Load existing clash zones** from Database/XML
3. ✅ **ResetFlagsForDeletedSleeves** - Mark deleted sleeves as unresolved
4. ✅ **Skip MergeAndSave** - No database save of new zones
5. ✅ **Go directly to Place Sleeves** - Only for deleted sleeves

### When Used:
- Files haven't changed
- User just wants to refresh flags for deleted sleeves
- Fast path - minimal processing

### Database Impact:
- **No new clash zones saved**
- **No FileCombo creation**
- Only flag updates for deleted sleeves

---

## PATH 2: REPLACE MODE (Orange)

### Trigger:
- **Adopt to Modified Document = OFF**
- **AND IsFilterComboNew = 0** (No new file combos)
- **AND** no new files added

### Actions:
1. ✅ **Skip Filter XML and detection** - No clash zone processing
2. ✅ **Sync Global XML flags only** - Update resolved flags
3. ✅ **ResetFlagsForDeletedSleeves** - Mark deleted sleeves as unresolved
4. ✅ **No clash zone processing** - Minimal work

### When Used:
- "Adopt to modified document" is disabled
- All file combos already processed
- User wants flag sync only (no detection)

### Database Impact:
- **No clash zones saved**
- **No FileCombo creation**
- Only Global XML flag updates

---

## PATH 3: FULL DETECTION MODE (Green)

### Trigger:
- **Adopt to Modified Document = ON** (always triggers PATH 3)
- **OR IsFilterComboNew = 1** (New file combos detected)
- **OR** new files added to project

### Actions:
1. ✅ **Full intersection detection** - Find all MEP vs Structural intersections
2. ✅ **Validate clash zones** - 3-Point Validation (if enabled)
3. ✅ **MergeAndSave** - Process and save all clash zones
4. ✅ **Database save sequence**:
   - Get/Create Filter
   - Get/Create FileCombo (IsFilterComboNew=1 if new)
   - Insert/Update ClashZones
   - Insert/Update SleeveSnapshots
   - Commit Transaction

### When Used:
- New files added to project
- File combinations changed
- "Adopt to modified document" is enabled
- First time processing a filter

### Database Impact:
- **All tables populated**:
  - ✅ Filters
  - ✅ FileCombos (with IsFilterComboNew=1)
  - ✅ ClashZones
  - ✅ SleeveSnapshots

---

## Key Relationships

### Filter → FileCombo Relationship:
- **Filter** is created/retrieved first
- **FileCombo** is checked for `IsFilterComboNew` flag
- This flag determines PATH 1 vs PATH 3

### Flag Reset After Placement:
- After successful sleeve placement, `IsFilterComboNew` is reset to `0`
- This is done via `OpeningCommandOrchestrator.ResetFilterComboFlagAfterPlacement()`
- Next refresh will use PATH 1 (Replay) instead of PATH 3

### Adopt Setting Impact:
- **Adopt = ON**: Always uses PATH 3 (Full Detection)
- **Adopt = OFF**: Uses PATH 1 or PATH 2 based on IsFilterComboNew flag

## Flowchart File

Import `THREE_PATHS_FLOWCHART.drawio` into draw.io to see the complete visual flow.

## Summary Table

| Path | Trigger | Detection | Database Save | Use Case |
|------|---------|-----------|---------------|----------|
| **PATH 1** | IsFilterComboNew = 0 | ❌ Skip | ❌ No | Replay existing zones |
| **PATH 2** | Adopt=OFF AND IsFilterComboNew=0 | ❌ Skip | ❌ No | Flag sync only |
| **PATH 3** | Adopt=ON OR IsFilterComboNew=1 | ✅ Full | ✅ Yes | New detection needed |


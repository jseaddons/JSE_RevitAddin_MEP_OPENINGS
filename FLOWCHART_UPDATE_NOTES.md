# Flowchart Update - PATH 1 vs PATH 2 Decision

## Critical Missing Flow Added

The updated flowchart now includes the **crucial decision point** that was missing:

### After "Mark Filter as Enabled":

1. **Check FileCombos Table for IsFilterComboNew Flag**
   - Query: `SELECT IsFilterComboNew FROM FileCombos WHERE FilterId = @FilterId`
   - This determines which path to take

2. **PATH Decision:**
   - **PATH 1 (Replay Mode)**: If `IsFilterComboNew = 0`
     - Skip clash zone processing entirely
     - Load existing clash zones from database/XML
     - Reset flags for deleted sleeves only
     - Go directly to sleeve placement (for deleted sleeves only)
     - **No new clash zone detection or database save**
   
   - **PATH 2 (Full Detection Mode)**: If `IsFilterComboNew = 1`
     - Process clash zones (full detection)
     - Save to database (Filters → FileCombos → ClashZones → SleeveSnapshots)
     - Continue with normal flow

## Key Relationships Shown:

1. **Filter → FileCombo Relationship:**
   - Filter is created/retrieved first
   - FileCombo is checked for `IsFilterComboNew` flag
   - This flag determines the refresh path

2. **PATH 1 Flow (IsFilterComboNew = 0):**
   - Load existing zones
   - Reset flags for deleted sleeves
   - Skip to placement (no database save of new zones)

3. **PATH 2 Flow (IsFilterComboNew = 1):**
   - Full detection and processing
   - Database save sequence (all 4 steps)
   - New FileCombo created with `IsFilterComboNew = 1`

## Why This Matters:

- **PATH 1** is for **replay/refresh** scenarios where files haven't changed
- **PATH 2** is for **new file combinations** that need full detection
- The flag is reset to `0` after successful placement (via `OpeningCommandOrchestrator`)
- This prevents unnecessary re-detection when files haven't changed

## File: `DATABASE_SAVE_FLOWCHART_UPDATED.drawio`

Import this file into draw.io to see the complete flow with PATH 1 vs PATH 2 decision.


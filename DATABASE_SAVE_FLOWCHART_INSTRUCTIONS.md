# Database Save Flowchart - Instructions

## Quick Access Links

### Option 1: Mermaid Live Editor (Easiest)
1. Go to: **https://mermaid.live**
2. Copy the entire contents of `DATABASE_SAVE_FLOWCHART.mmd`
3. Paste into the editor
4. The flowchart will render automatically!

### Option 2: Draw.io (More Control)
1. Go to: **https://app.diagrams.net**
2. Create a new diagram
3. Use the flowchart structure below as a guide

## Flowchart Structure

### Main Flow:
```
START: User Clicks Refresh
  ↓
OnRefreshClick (EmergencyMainDialog)
  ↓
RefreshServiceRefactored.ExecuteRefresh
  ↓
LoadEnabledFilter
  ├─→ Filter Exists? NO → CreateFilterFromCurrentUIState
  │                        ↓
  │                    RegisterFilterInDatabase
  │                        ↓
  │                    SaveFilterUIState (SelectedHostElementTypes, OpeningSettings)
  │                        ↓
  └─→ Filter Exists? YES → Mark Filter as Enabled
                            ↓
                    MergeAndSave
                            ↓
                    ClashZonePersistenceService.SaveClashZones
                            ↓
                    Check: SQLite Repository Exists?
                    ├─→ NO → ❌ ERROR: Log and Skip
                    └─→ YES → Group Clash Zones by Category
                                ↓
                        For Each Category:
                          ↓
                        Validate Clash Zones
                          ↓
                        Valid Zones?
                        ├─→ NO → Skip Category
                        └─→ YES → Group by File Combo
                                    ↓
                            ClashZoneRepository.InsertOrUpdateClashZones
                                    ↓
                            BEGIN TRANSACTION
                                    ↓
                            STEP 1: Get or Create Filter
                            ├─→ Filter Exists? NO → INSERT INTO Filters
                            └─→ Filter Exists? YES → SELECT FilterId
                                    ↓
                            FilterId > 0?
                            ├─→ NO → ❌ Rollback
                            └─→ YES → STEP 2: Process Each Clash Zone
                                        ↓
                                For Each Clash Zone:
                                  ↓
                                Extract File Keys
                                  ↓
                                Keys Valid?
                                ├─→ NO → ❌ Skip Zone
                                └─→ YES → Get or Create File Combo
                                            ├─→ Combo Exists? NO → INSERT INTO FileCombos
                                            └─→ Combo Exists? YES → SELECT ComboId
                                                    ↓
                                            ComboId > 0?
                                            ├─→ NO → ❌ Skip Zone
                                            └─→ YES → Insert or Update Clash Zone
                                                        ├─→ Exists by GUID? → UPDATE
                                                        ├─→ Exists by MEP+Host+Point? → UPDATE
                                                        └─→ Not Found → INSERT INTO ClashZones
                                ↓
                            STEP 3: Insert/Update Sleeve Snapshots
                            (For zones with SleeveInstanceId > 0)
                                ↓
                            STEP 4: COMMIT TRANSACTION
                                ├─→ Success → ✅ All Tables Populated
                                └─→ Failure → ❌ Rollback
```

## Tables Populated (in order):
1. **Filters** - Created first
2. **FileCombos** - Created for each unique file combination
3. **ClashZones** - Inserted/updated for each clash zone
4. **SleeveSnapshots** - Inserted/updated for placed sleeves

## Key Decision Points:
- 🔴 **Red (Error)**: Repository NULL, Rollback, Invalid Keys
- 🟡 **Yellow (Warning)**: Skip Zone, Skip Category
- 🟢 **Green (Success)**: All operations completed
- 🔵 **Blue (Create)**: New records being created

## Color Coding:
- **Green**: Start/End/Success nodes
- **Red**: Error/Rollback nodes
- **Yellow**: Warning/Skip nodes
- **Blue**: Create/Insert operations


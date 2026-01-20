# Placement Issues Diagnostic Report

## Issues Identified

### 1. Placement Using XML Instead of Database
**Problem**: Placement logs show "Source=XML snapshot" even though SQLite loaded 36 zones.

**Root Cause**: 
- `LoadClashZonesForCategory` is called with `unresolvedOnly=true` by default
- After flag reset, all zones have `IsResolved=false` and `SleeveInstanceId=-1`
- DB query filters for `SleeveState = 0` (unresolved), but the query might not match reset flags correctly
- Returns 0 zones → falls back to XML

**Solution**: 
- Check if placement service is calling with `unresolvedOnly=true` when it should use `unresolvedOnly=false`
- Verify DB query logic for `SleeveState` calculation matches flag reset logic
- Ensure placement service filters by flags in memory, not in DB query

### 2. MEP Orientation and Structural Thickness Not Retrieved from Linked Files
**Problem**: MEP orientation and structural thickness values are 0 in XML/DB.

**Root Cause**:
- Methods `GetMepElementOrientation`, `GetElementThickness`, `GetWallThickness`, `GetFramingThickness` use `element.Document`
- For linked file elements, `element.Document` should be the linked document, which should work
- However, values might not be saved correctly during clash zone creation
- Or values are being loaded as 0 from XML/DB

**Solution**:
- Verify elements are retrieved from linked files correctly before calling these methods
- Add logging to confirm linked file document is being used
- Check if values are being saved correctly to DB/XML
- Verify values are being loaded correctly from DB/XML

### 3. Duplicate Placement Warning After Clustering
**Problem**: Revit warning "identical instances in the same place" appears after clustering, even though there are no actual clusters.

**Root Cause**:
- After clustering, individual sleeves are still being placed when they should be skipped
- Flag state might be inconsistent: `IsClusterResolved` might not be set correctly after clustering
- Placement service checks `IsClusterResolved` and `ClusterSleeveInstanceId > 0`, but if flag state is wrong, it might still place individual sleeves
- Individual sleeves are being placed at the same location as cluster sleeves

**Solution**:
- Verify flag state is correctly set after clustering (DB first, then XML)
- Ensure placement service skips zones with `IsClusterResolved=true` even if `ClusterSleeveInstanceId` is not set
- Add additional check: if a zone has `IsClusterResolved=true`, skip individual placement regardless of `ClusterSleeveInstanceId`
- Verify cluster sleeve deletion logic doesn't leave orphaned flags

## Fixes Implemented

### ✅ Fix 1: Placement Data Source (COMPLETED)
- **Changed**: `LoadClashZonesFromDatabase` now uses `unresolvedOnly=false` to load ALL zones from DB
- **Added**: In-memory filtering by flags: skip zones with `IsResolved=true` OR `IsClusterResolved=true`
- **Added**: Logging to show data source (DB vs XML) in placement logs
- **Location**: `Services/OpeningCommandOrchestrator.cs` line 689-715

### ✅ Fix 2: Linked File Element Access (IN PROGRESS - Logging Added)
- **Added**: Logging to verify elements are from linked files before calling orientation/thickness methods
- **Added**: Logging to show which document (active vs linked) is being used
- **Added**: Diagnostic logging for orientation/thickness values during clash zone creation
- **Location**: `Services/ClashZoneService.cs` lines 1969-1985, 2156-2158
- **Note**: Methods use `element.Document` which should work for linked elements. Logging will help diagnose if issue is with retrieval or save/load.

### ✅ Fix 3: Flag State After Clustering (COMPLETED)
- **Added**: Check in placement service: skip if `IsClusterResolved=true` regardless of `ClusterSleeveInstanceId`
- **Added**: Flag state validation before placement
- **Location**: `Services/UniversalSleevePlacerService.cs` lines 618-655


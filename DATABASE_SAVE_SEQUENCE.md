# Database Save Sequence - Complete Flow

This document outlines the complete sequence of operations when saving clash zones to the database, ensuring all tables are populated.

## Complete Save Sequence

### 1. **Entry Point: `ClashZonePersistenceService.SaveClashZones()`**
   - **Location**: `Services/ClashZonePersistenceService.cs`
   - **Called from**: `RefreshServiceRefactored.MergeAndSave()`
   - **Parameters**: 
     - `allClashZones`: List of clash zones to save
     - `baseFilterName`: Base filter name (e.g., "Plumbing")
     - `targetFilter`: OpeningFilter object (must not be null)
   - **Actions**:
     - ✅ Logs entry with zone count, filter name, and SQLite repository status
     - ✅ Groups clash zones by category (`MepElementCategory`)
     - ✅ Processes each category separately

### 2. **Category Processing: `SaveCategory()`**
   - **Location**: `Services/ClashZonePersistenceService.cs`
   - **Actions**:
     - ✅ Validates clash zones using `IsValidClashZone()`
     - ✅ Groups valid zones by file combo (LinkedFile + HostFile)
     - ✅ Calls `_sqliteRepository.InsertOrUpdateClashZones()` for database save

### 3. **Database Save: `ClashZoneRepository.InsertOrUpdateClashZones()`**
   - **Location**: `Data/Repositories/ClashZoneRepository.cs`
   - **Transaction**: All operations are wrapped in a single SQLite transaction
   - **Sequence**:

#### **STEP 1: Get or Create Filter**
   - **Method**: `GetOrCreateFilter(filterName, category, transaction)`
   - **Table**: `Filters`
   - **Actions**:
     - ✅ Checks if filter exists: `SELECT FilterId FROM Filters WHERE FilterName = @FilterName AND Category = @Category`
     - ✅ If exists: Returns existing `FilterId`
     - ✅ If not exists: Creates new filter:
       ```sql
       INSERT INTO Filters (FilterName, Category, IsFilterComboNew, CreatedAt, UpdatedAt)
       VALUES (@FilterName, @Category, 0, CURRENT_TIMESTAMP, CURRENT_TIMESTAMP);
       SELECT last_insert_rowid();
       ```
     - ✅ Returns `FilterId` (must be > 0 to continue)

#### **STEP 2: Process Each Clash Zone**
   For each clash zone in the list:

   **2a. Get or Create File Combo**
   - **Method**: `GetOrCreateFileCombo(filterId, clashZone, transaction)`
   - **Table**: `FileCombos`
   - **Actions**:
     - ✅ Extracts file keys:
       - `LinkedFileKey`: From `SourceDocKey` or `DocumentPath` (normalized)
       - `HostFileKey`: From `HostDocKey` or `StructuralElementDocumentTitle` (normalized)
     - ✅ Checks if combo exists:
       ```sql
       SELECT ComboId FROM FileCombos 
       WHERE FilterId = @FilterId 
         AND LinkedFileKey = @LinkedFileKey 
         AND HostFileKey = @HostFileKey
       ```
     - ✅ If exists: Returns existing `ComboId`
     - ✅ If not exists: Creates new file combo:
       ```sql
       INSERT INTO FileCombos (FilterId, LinkedFileKey, HostFileKey, IsFilterComboNew, ProcessedAt, CreatedAt, UpdatedAt)
       VALUES (@FilterId, @LinkedFileKey, @HostFileKey, 1, CURRENT_TIMESTAMP, CURRENT_TIMESTAMP, CURRENT_TIMESTAMP);
       SELECT last_insert_rowid();
       ```
     - ✅ Returns `ComboId` (must be > 0 to continue)
     - ⚠️ **CRITICAL**: If `ComboId <= 0`, the clash zone is skipped

   **2b. Insert or Update Clash Zone**
   - **Method**: `InsertOrUpdateClashZone(comboId, clashZone, transaction)`
   - **Table**: `ClashZones`
   - **Actions**:
     - ✅ Checks if clash zone exists by GUID:
       ```sql
       SELECT ClashZoneId FROM ClashZones 
       WHERE UPPER(ClashZoneGuid) = @ClashZoneGuid 
         AND ClashZoneGuid != '' AND ClashZoneGuid IS NOT NULL
       LIMIT 1
       ```
     - ✅ If found by GUID: Updates existing clash zone (`UpdateClashZone()`)
     - ✅ If not found: Checks by MEP+Host+Point:
       ```sql
       SELECT ClashZoneId FROM ClashZones 
       WHERE ComboId = @ComboId 
         AND MepElementId = @MepElementId 
         AND HostElementId = @HostElementId
         AND ABS(IntersectionX - @IntersectionX) < 0.001
         AND ABS(IntersectionY - @IntersectionY) < 0.001
         AND ABS(IntersectionZ - @IntersectionZ) < 0.001
       LIMIT 1
       ```
     - ✅ If found by MEP+Host+Point: Updates existing clash zone
     - ✅ If not found: Inserts new clash zone (`InsertClashZone()`)
       - Inserts all clash zone fields including:
         - `ComboId` (foreign key to `FileCombos`)
         - `MepElementId`, `HostElementId`
         - `IntersectionX`, `IntersectionY`, `IntersectionZ`
         - `SleeveState`, `SleeveInstanceId`, `ClusterInstanceId`
         - Bounding box coordinates
         - All other clash zone properties

#### **STEP 3: Insert or Update Sleeve Snapshots**
   - **Method**: `InsertOrUpdateSleeveSnapshots(filterId, processedZones, transaction)`
   - **Table**: `SleeveSnapshots`
   - **Actions**:
     - ✅ For each processed clash zone with `SleeveInstanceId > 0`:
       - Checks if snapshot exists
       - If exists: Updates snapshot
       - If not exists: Inserts new snapshot

#### **STEP 4: Commit Transaction**
   - **Action**: `transaction.Commit()`
   - ⚠️ **CRITICAL**: If any step fails, the entire transaction is rolled back
   - ✅ All tables are updated atomically

## Table Population Order

1. **`Filters`** - Created/retrieved first (required for all other operations)
2. **`FileCombos`** - Created/retrieved for each unique file combination (required for clash zones)
3. **`ClashZones`** - Inserted/updated for each clash zone (references `FileCombos.ComboId`)
4. **`SleeveSnapshots`** - Inserted/updated for placed sleeves (references `ClashZones.ClashZoneId`)

## Diagnostic Logging

All operations are logged to:
- **Console**: `DebugLogger.Info/Warning/Error()`
- **File**: `save_db_diagnostic.log` (in logs directory)
- **Refresh Log**: `Refresh_YYYY-MM-DD_HH-MM-SS.log`

### Key Log Messages to Check:

1. **Filter Creation**:
   ```
   [SQLite] ✅ Created filter 'FilterName' (Category='Category') in database (FilterId=X)
   ```

2. **File Combo Creation**:
   ```
   [SQLite] ✅ Created new FileCombo: ComboId=X, FilterId=Y, Linked='...', Host='...'
   ```

3. **Clash Zone Insert/Update**:
   ```
   Zone {Guid}: ✅ INSERTED (ComboId=X)
   Zone {Guid}: ✅ UPDATED (ComboId=X)
   ```

4. **Transaction Commit**:
   ```
   Step 4 SUCCESS: Transaction committed
   ✅ Database save complete: X zones, Duration=Yms
   ```

## Common Issues and Solutions

### Issue 1: FileCombos Table Not Populated
**Symptoms**: `Filters` table has data, but `FileCombos` is empty

**Possible Causes**:
1. ❌ `GetOrCreateFileCombo()` returning `-1` (invalid file keys)
2. ❌ Transaction rollback due to error in clash zone insert
3. ❌ `comboId <= 0` check causing zones to be skipped

**Diagnosis**:
- Check `save_db_diagnostic.log` for:
  - `❌ CRITICAL: Failed to get/create file combo`
  - `⚠️ Invalid file combo keys`
  - `❌ Failed to create FileCombo: INSERT returned NULL`

**Solution**:
- Ensure `SourceDocKey`/`DocumentPath` and `HostDocKey`/`StructuralElementDocumentTitle` are populated in clash zones
- Check that `NormalizeDocumentKey()` is working correctly

### Issue 2: ClashZones Table Not Populated
**Symptoms**: `FileCombos` has data, but `ClashZones` is empty

**Possible Causes**:
1. ❌ All clash zones failing validation (`IsValidClashZone()`)
2. ❌ `InsertOrUpdateClashZone()` throwing exceptions
3. ❌ Transaction rollback

**Diagnosis**:
- Check `save_db_diagnostic.log` for:
  - `❌ Zone {Guid} failed: {Exception}`
  - `⚠️ Skipped null clash zone`
  - `⚠️ No valid zones to save`

**Solution**:
- Verify clash zones have valid:
  - `MepElementId > 0`
  - `HostElementId > 0`
  - Non-zero intersection point
  - Valid `SourceDocKey` or `DocumentPath`

### Issue 3: Transaction Rollback
**Symptoms**: No data in any table after save attempt

**Possible Causes**:
1. ❌ Foreign key constraint violation
2. ❌ UNIQUE constraint violation
3. ❌ Database connection error

**Diagnosis**:
- Check `save_db_diagnostic.log` for:
  - `❌ ERROR: {Exception}`
  - `Stack trace: ...`
  - `=== SaveToDatabase END (FAILED) ===`

**Solution**:
- Check database schema matches code expectations
- Verify foreign key relationships are correct
- Check for duplicate GUIDs or constraint violations

## Verification Checklist

After running a refresh, verify:

- [ ] `Filters` table has entries (one per filter/category combination)
- [ ] `FileCombos` table has entries (one per unique LinkedFile+HostFile combination per filter)
- [ ] `ClashZones` table has entries (one per clash zone)
- [ ] `SleeveSnapshots` table has entries (one per placed sleeve, if any)
- [ ] Foreign key relationships are correct:
  - `FileCombos.FilterId` → `Filters.FilterId`
  - `ClashZones.ComboId` → `FileCombos.ComboId`
  - `SleeveSnapshots.ClashZoneId` → `ClashZones.ClashZoneId`

## Next Steps

If tables are still not populated:
1. Check `save_db_diagnostic.log` for detailed error messages
2. Verify clash zones have valid file keys (`SourceDocKey`, `HostDocKey`)
3. Check that `_sqliteRepository` is not null in `ClashZonePersistenceService`
4. Verify transaction is being committed (not rolled back)


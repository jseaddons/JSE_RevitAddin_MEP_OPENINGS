# DB Filter UI State Migration Plan

**Objective:** Migrate filter UI state persistence (`SelectedHostCategories` + `OpeningSettings`) from XML to SQLite DB as the single source of truth.

**Scope:** UI state only — does not include clash zones or clustering data (separate migrations).

**Timeline:** 3 days (dev → test → production).

---

## 1. Scope & Prerequisites

- This migration now focuses only on adopting the opening filter's document-refresh checkbox state (the single UI flag indicating whether the filter should refresh/apply to the current document). XML backups and broad schema checks are NOT required — the project already contains opening filters and the `Filters.OpeningSettings` column is present.
- Assumption: `Filters.OpeningSettings` exists and stores UI state as JSON. No XML migration or backups will be performed.
- Dry-run support is still available for safety: `dryRun = true` will only report actions without writing changes.

---

## 2. Implementation: `MigrateFiltersXmlToDb` Method

**Location:** `Services/FilterManagementService.cs`

**Signature:**
```csharp
public (int migratedCount, int skippedCount, int failedCount) MigrateFiltersXmlToDb(
    bool dryRun = true,
    bool preserveXml = true,
    string logFileName = "filter_migration.log")
```

**Implementation outline (adopt Document-Refresh flag only):**

1. Enumerate filters from DB using existing repository APIs (no XML reads).
2. For each filter record:
   a. Deserialize `Filters.OpeningSettings` JSON into an `OpeningSettings` object (or null if empty).
   b. Inspect the integer `DocumentRefresh` value in the `OpeningSettings` object. Semantics: `DocumentRefresh = 1` means checkbox was checked; if the value/key is absent the box was unchecked.
   c. If `DocumentRefresh` is present and equals `1`, mark as skipped (already adopted).
   d. If `DocumentRefresh` is missing but the in-memory or UI-facing `OpeningFilter` (from current session) has the checkbox checked, adopt that value:
      - Read the current `OpeningFilter` instance via `FilterManagementService.LoadFilterAuto(filterName)` or active in-memory provider.
      - Extract the document-refresh checkbox value: `bool docRefreshChecked = uiFilter.OpeningSettings?.DocumentRefresh == 1`.
      - If `docRefreshChecked` is true and NOT `dryRun`: set `OpeningSettings.DocumentRefresh = 1` and call `UseFilterRepository(repo => repo.SaveFilterUIState(filterName, category, existingHostCategories, newOpeningSettings))` to update only that integer value.
      - If `docRefreshChecked` is false: do not add the `DocumentRefresh` key (leave omitted).
      - Log migration/adoption per filter.
3. Continue through all filters; do not alter other UI state fields.
4. Summary log: counts for adopted/skipped/failed and elapsed time.

**Key behaviors:**
- **Minimal & Idempotent:** Only the `DocumentRefresh` integer is adopted. If `OpeningSettings` already contains `DocumentRefresh = 1`, skip that filter.
- **Storage semantics:** `DocumentRefresh = 1` when the checkbox is checked; if the checkbox is unchecked the `DocumentRefresh` key is omitted (no key/value written).
- **No XML touches:** The code will not read, write, or backup XML files.
- **Error handling:** Log and skip individual filters on error; do not abort the overall run.
- **Logging:** Write per-filter summary + final summary to log file + console.

---

## 3. Idempotency & Safety

1. **Check before overwrite:**
   - Only overwrite/add the `DocumentRefreshChecked` flag if it is missing. Do not replace or clear other `OpeningSettings` fields.

2. **Record-run marker (optional):**
   - Optionally record a lightweight entry in `SchemaMigrations` or a dedicated `MigrationRuns` table noting the adoption run ID and timestamp. This is optional given the minimal scope.

3. **No XML preservation needed:**
   - Per agreed scope, XML files are not used or modified.

---

## 4. Verification & Tests

### Unit Tests

1. **`Adopt_DocumentRefreshFlag_DryRun_DoesNotChangeDb()`**
   - Setup: Ensure a filter record exists with empty or missing `DocumentRefresh` in `OpeningSettings`.
   - Call: `MigrateFiltersXmlToDb(dryRun: true)`.
   - Assert: DB `Filters.OpeningSettings` remains unchanged.

2. **`Adopt_DocumentRefreshFlag_Writes_Flag()`**
   - Setup: Filter exists without the `DocumentRefresh` key.
   - Call: `MigrateFiltersXmlToDb(dryRun: false)`.
   - Assert: `Filters.OpeningSettings` now contains the `DocumentRefresh` key with value `1` for adopted filters; filters with checkbox unchecked remain without the key.

3. **`Adopt_DocumentRefreshFlag_Skips_If_Present()`**
   - Setup: Filter already has `DocumentRefresh = 1` in DB.
   - Call: `MigrateFiltersXmlToDb(dryRun: false)`.
   - Assert: Skipped, no modifications done.

### Integration Tests

1. **`LoadFilterAuto_Reads_OpeningSettings_From_Db()`**
   - Migrate XML filter to DB.
   - Load via `FilterManagementService.LoadFilterAuto(filterName)`.
   - Assert: Returned filter has `OpeningSettings` from DB.

2. **`UI_Apply_Filter_Restores_HostCategories_From_Db()`**
   - Migrate filter to DB.
   - Call `FilterUiStateProvider.ApplyFilterToUi(filter)`.
   - Assert: UI checkboxes for host categories are checked correctly.

### Manual Verification Queries

Run these after adoption:

```sql
-- 1. Count filters missing the document-refresh adoption (DocumentRefresh=1)
SELECT COUNT(*) AS MissingDocRefreshAdoption
FROM Filters
WHERE OpeningSettings IS NULL
   OR OpeningSettings NOT LIKE '%"DocumentRefresh":1%';

-- 2. Inspect a specific filter's UI state preview
SELECT FilterId, FilterName, Category, substr(OpeningSettings,1,200) AS OpeningSettings_Preview
FROM Filters
WHERE FilterName = 'Electrical' AND Category = 'Cable Trays';
```

**PowerShell wrapper (example):**
```powershell
$dbPath = "C:\Users\jse2084\AppData\Roaming\JSE_MEP_Openings\Projects\Shared Test Model -00001\Filters\Shared_Test_Model_-00001_SleevePersistence.db"
sqlite3 $dbPath "SELECT COUNT(*) FROM Filters WHERE OpeningSettings IS NULL OR OpeningSettings NOT LIKE '%DocumentRefreshChecked%';"
```

---

## 5. Toggle & Cleanup

1. **After verification:**
   ```csharp
   // In DeploymentConfiguration or project settings
   public static bool UseSqliteAsPrimary { get; set; } = true;
   public static bool DisableXmlCreation { get; set; } = true; // XML creation remains optional
   ```

2. **Rollback option:**
   - If needed: `FilterManagementService.RegenerateXmlFromDb(string filterName, string category)`.
   - Re-creates XML from DB `OpeningSettings` using `SaveFilterToXmlFile`.

3. **Admin CLI (optional):**
   - Add a command: `MigrateFiltersToDb --dryRun --log filter_migration.log`.
   - Display: `Migrated: 5, Skipped: 2, Failed: 0`.

---

## 6. Documentation & Rollout

1. **Update architecture doc:**
   - Add section to `COMPREHENSIVE_MEP_OPENINGS_IMPLEMENTATION_PLAN.md` under "Database Persistence."
   - Note: UI state (SelectedHostCategories, OpeningSettings) is now stored in `Filters.OpeningSettings` column (JSON).

2. **Migration README (this file):**
   - Include all steps above + verification queries + troubleshooting.

3. **Example migration run log:**
   ```
   [2025-11-18 14:30:00] [MIGRATION] Starting UI state migration (dryRun=false)
   [2025-11-18 14:30:01] [MIGRATION] Found 12 XML filter files
   [2025-11-18 14:30:02] [MIGRATION] Filter=Electrical, FilterId=1, Category=Ducts, Status=MIGRATED
   [2025-11-18 14:30:03] [MIGRATION] Filter=Plumbing, FilterId=2, Category=Pipes, Status=MIGRATED
   [2025-11-18 14:30:04] [MIGRATION] Filter=HVAC, FilterId=3, Category=Cable Trays, Status=SKIPPED (already migrated)
   [2025-11-18 14:30:05] [MIGRATION-SUMMARY] Migrated=11, Skipped=1, Failed=0, ElapsedMs=5000
   [2025-11-18 14:30:05] [MIGRATION-VERIFY] Run SQL: SELECT COUNT(*) FROM Filters WHERE OpeningSettings IS NULL;
   ```

---

## 7. Dev Implementation Checklist

- [ ] Add `AdoptDocumentRefreshFlag(...)` method to `FilterManagementService`.
- [ ] Add logging wrapper (use existing `_log` delegate).
- [ ] Write unit tests (dry-run, write, skip focused on the flag).
- [ ] Write integration tests (LoadFilterAuto reads flag; UI apply restores behavior).
- [ ] Add SQL verification queries to migration readme.
- [ ] Run adoption on test project; verify DB.
- [ ] Toggle `DisableXmlCreation = true` in production config if desired.
- [ ] Update `COMPREHENSIVE_MEP_OPENINGS_IMPLEMENTATION_PLAN.md` to reflect the narrowed scope.

---

## 8. Troubleshooting

| Issue | Solution |
|-------|----------|
| Adoption did not set document-refresh flag | Verify `FilterManagementService.LoadFilterAuto` exposes the current UI state; ensure `AdoptDocumentRefreshFlag` reads the UI/provider value correctly. |
| UI does not restore document-refresh behavior after adoption | Verify `FilterUiStateProvider.ApplyFilterToUi` reads `OpeningSettings.DocumentRefreshChecked` and applies refresh behavior. |
| DB connection fails | Verify `SleeveDbContext` can access DB file; check WAL mode enabled. |

---

## 9. Rollout Timeline

| Day | Task | Owner |
|-----|------|-------|
| Day 0 | Write `MigrateFiltersXmlToDb` + unit tests | Dev |
| Day 1 | Integration tests + run on test project | QA/Dev |
| Day 2 | Manual DB verification + sign-off | QA |
| Day 3 | Deploy to production + toggle flags | DevOps/Dev |

---

## 10. Rollback Plan

If migration fails or causes issues:

1. **Restore from backup:**
   ```powershell
   Remove-Item "C:\path\to\Filters" -Recurse
   Copy-Item -Path $backupDir -Destination "C:\path\to\Filters" -Recurse
   ```

2. **Revert DB:**
   ```sql
   UPDATE Filters SET OpeningSettings = NULL WHERE FilterId > 0;
   DELETE FROM SchemaMigrations WHERE Version = 'UI_STATE_MIGRATION_001';
   ```

3. **Toggle flags:**
   ```csharp
   DeploymentConfiguration.DisableXmlCreation = false;
   ```

4. **Restart application and reload filters from XML.**

---

**End of Plan**

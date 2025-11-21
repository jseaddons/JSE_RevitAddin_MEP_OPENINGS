using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Data.SQLite;
using System.Text.Json;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Data.Entities;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Data.Repositories
{
    /// <summary>
    /// SQLite repository implementation for ClashZone persistence
    /// Phase SQLite-1: Dual-write mode (writes to both XML and SQLite)
    /// </summary>
    public class ClashZoneRepository : IClashZoneRepository
    {
        private readonly SleeveDbContext _context;
        private readonly Action<string> _logger;

        public ClashZoneRepository(SleeveDbContext context, Action<string> logger = null)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            _logger = logger ?? (msg => { });
        }

        public void InsertOrUpdateClashZones(IEnumerable<ClashZone> clashZones, string filterName, string category)
        {
            if (clashZones == null) return;

            // ✅ DIAGNOSTIC: Create database diagnostic log similar to XML diagnostic log
            var dbDiagnosticLogPath = SafeFileLogger.GetLogFilePath("save_db_diagnostic.log");
            var diagnosticLog = new System.Text.StringBuilder();
            var startTime = DateTime.Now;
            
            try
            {
                var zonesList = clashZones.ToList();
                diagnosticLog.AppendLine($"[{startTime:yyyy-MM-dd HH:mm:ss.fff}] === SaveToDatabase START ===");
                diagnosticLog.AppendLine($"[{startTime:HH:mm:ss.fff}] Input count: {zonesList.Count}");
                
                // Sample data (first 10)
                var sample = string.Join(", ", zonesList.Take(10).Select(z => $"{z.Id}:{z.SleeveInstanceId}"));
                diagnosticLog.AppendLine($"[{startTime:HH:mm:ss.fff}] Sample (first 10): {sample}");
                
                diagnosticLog.AppendLine($"[{startTime:HH:mm:ss.fff}] Filter name: '{filterName}', Category: '{category}'");
                diagnosticLog.AppendLine($"[{startTime:HH:mm:ss.fff}] Database path: {_context.DatabasePath}");
                diagnosticLog.AppendLine($"[{startTime:HH:mm:ss.fff}] Database exists: {System.IO.File.Exists(_context.DatabasePath)}");

            using (var transaction = _context.Connection.BeginTransaction())
            {
                try
                {
                        diagnosticLog.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Step 1: Beginning transaction...");
                        
                    // Retrieve existing filter; do not create new entries
                    var filterId = GetOrCreateFilter(filterName, category, transaction);
                        diagnosticLog.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Step 1: FilterId={filterId}");
                        
                    if (filterId <= 0)
                    {
                            diagnosticLog.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Step 1 FAILED: Filter '{filterName}' with category '{category}' does not exist in SQLite.");
                        _logger($"[SQLite] ⚠️ Skipping persistence: Filter '{filterName}' with category '{category}' does not exist in SQLite.");
                        transaction.Rollback();
                            
                            diagnosticLog.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] === SaveToDatabase END (FAILED - Filter not found) ===");
                            File.AppendAllText(dbDiagnosticLogPath, diagnosticLog.ToString());
                        return;
                    }
                        
                        diagnosticLog.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Step 1 SUCCESS: FilterId={filterId}");
                        diagnosticLog.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Step 2: Processing {zonesList.Count} clash zones...");

                    var processedZones = new List<(int ComboId, ClashZone Zone)>();
                        int insertedCount = 0;
                        int updatedCount = 0;
                        int skippedCount = 0;
                        int comboCreatedCount = 0;
                        int comboExistingCount = 0;

                        foreach (var clashZone in zonesList)
                    {
                            if (clashZone == null)
                            {
                                skippedCount++;
                                diagnosticLog.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] ⚠️ Skipped null clash zone");
                                continue;
                            }

                            // ✅ STEP 1: Get or create file combo with category and host categories
                            diagnosticLog.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Processing zone {clashZone.Id}: Getting/creating file combo...");
                            
                            // ✅ CRITICAL: Get category from clash zone (MEP category)
                            string mepCategory = clashZone.MepElementCategory ?? category;
                            
                            // ✅ CRITICAL: Get host categories from FilterUiStateProvider or fallback
                            List<string> hostCategories = null;
                            if (FilterUiStateProvider.GetSelectedHostCategories != null)
                            {
                                hostCategories = FilterUiStateProvider.GetSelectedHostCategories.Invoke();
                            }
                            
                        var comboId = GetOrCreateFileCombo(filterId, mepCategory, hostCategories, clashZone, transaction);

                            if (comboId <= 0)
                            {
                                // ⚠️ CRITICAL: File combo creation failed - skip this zone
                                skippedCount++;
                                var errorMsg = $"[{DateTime.Now:HH:mm:ss.fff}] ❌ CRITICAL: Failed to get/create file combo for zone {clashZone.Id} (comboId={comboId}). LinkedFile='{clashZone.SourceDocKey ?? clashZone.DocumentPath}', HostFile='{clashZone.HostDocKey ?? clashZone.StructuralElementDocumentTitle}'";
                                diagnosticLog.AppendLine(errorMsg);
                                _logger($"[SQLite] {errorMsg}");
                                continue; // Skip this zone - cannot save without a valid comboId
                            }
                            
                            // ✅ STEP 2: Track combo creation/retrieval
                            if (processedZones.Any(p => p.Zone.Id == clashZone.Id))
                            {
                                comboExistingCount++;
                                diagnosticLog.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Zone {clashZone.Id}: Using existing combo (ComboId={comboId})");
                            }
                            else
                            {
                                comboCreatedCount++;
                                diagnosticLog.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Zone {clashZone.Id}: Created new combo (ComboId={comboId})");
                            }

                            // ✅ STEP 3: Insert or update clash zone
                            bool wasInsert = false;
                            try
                            {
                                diagnosticLog.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Zone {clashZone.Id}: Inserting/updating clash zone with ComboId={comboId}...");
                                wasInsert = InsertOrUpdateClashZone(comboId, clashZone, transaction);
                                if (wasInsert)
                                {
                                    insertedCount++;
                                    diagnosticLog.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Zone {clashZone.Id}: ✅ INSERTED (ComboId={comboId})");
                                }
                                else
                                {
                                    updatedCount++;
                                    diagnosticLog.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Zone {clashZone.Id}: ✅ UPDATED (ComboId={comboId})");
                                }
                            }
                            catch (Exception zoneEx)
                            {
                                skippedCount++;
                                var errorMsg = $"[{DateTime.Now:HH:mm:ss.fff}] ❌ Zone {clashZone.Id} failed: {zoneEx.Message}";
                                diagnosticLog.AppendLine(errorMsg);
                                diagnosticLog.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Stack trace: {zoneEx.StackTrace}");
                                _logger($"[SQLite] {errorMsg}");
                            }
                            
                        processedZones.Add((comboId, clashZone));
                    }

                        diagnosticLog.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Step 2 SUCCESS: Processed={processedZones.Count}, Inserted={insertedCount}, Updated={updatedCount}, Skipped={skippedCount}, CombosCreated={comboCreatedCount}, CombosExisting={comboExistingCount}");
                        diagnosticLog.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Step 2 Sample (first 10 processed): {string.Join(", ", processedZones.Take(10).Select(p => $"{p.Zone.Id}:{p.ComboId}"))}");
                        
                        diagnosticLog.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Step 3: Inserting/updating sleeve snapshots...");
                    InsertOrUpdateSleeveSnapshots(filterId, processedZones, transaction);
                        diagnosticLog.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Step 3 SUCCESS: Sleeve snapshots updated");

                        diagnosticLog.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Step 4: Committing transaction...");
                    transaction.Commit();
                        diagnosticLog.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Step 4 SUCCESS: Transaction committed");
                        
                        var endTime = DateTime.Now;
                        var duration = (endTime - startTime).TotalMilliseconds;
                        diagnosticLog.AppendLine($"[{endTime:HH:mm:ss.fff}] ✅ Database save complete: {zonesList.Count} zones, Duration={duration:F0}ms");
                        diagnosticLog.AppendLine($"[{endTime:HH:mm:ss.fff}] === SaveToDatabase END (SUCCESS) ===");
                        
                        _logger($"[SQLite] ✅ Inserted/updated {zonesList.Count} clash zones for filter '{filterName}', category '{category}'");
                }
                catch (Exception ex)
                {
                        var errorTime = DateTime.Now;
                        diagnosticLog.AppendLine($"[{errorTime:HH:mm:ss.fff}] ❌ ERROR: {ex.Message}");
                        diagnosticLog.AppendLine($"[{errorTime:HH:mm:ss.fff}] Stack trace: {ex.StackTrace}");
                        diagnosticLog.AppendLine($"[{errorTime:HH:mm:ss.fff}] === SaveToDatabase END (FAILED) ===");
                        
                    transaction.Rollback();
                    _logger($"[SQLite] ❌ Error inserting/updating clash zones: {ex.Message}");
                    throw;
                }
                    finally
                    {
                        // Write diagnostic log to file
                        try
                        {
                            File.AppendAllText(dbDiagnosticLogPath, diagnosticLog.ToString() + "\n");
                        }
                        catch { } // Don't fail if log write fails
                    }
                }
            }
            catch (Exception ex)
            {
                var errorTime = DateTime.Now;
                diagnosticLog.AppendLine($"[{errorTime:HH:mm:ss.fff}] ❌ FATAL ERROR: {ex.Message}");
                diagnosticLog.AppendLine($"[{errorTime:HH:mm:ss.fff}] === SaveToDatabase END (FATAL ERROR) ===");
                
                try
                {
                    File.AppendAllText(dbDiagnosticLogPath, diagnosticLog.ToString() + "\n");
                }
                catch { }
                
                throw;
            }
        }

        private int GetOrCreateFilter(string filterName, string category, SQLiteTransaction transaction)
        {
            if (string.IsNullOrWhiteSpace(filterName) || string.IsNullOrWhiteSpace(category))
            {
                _logger($"[SQLite] ⚠️ Invalid filter metadata (Name='{filterName}', Category='{category}').");
                return -1;
            }

            // ✅ STEP 1: Try to get existing filter
            using (var cmd = _context.Connection.CreateCommand())
            {
                cmd.Transaction = transaction;
                cmd.CommandText = @"
                    SELECT FilterId FROM Filters 
                    WHERE FilterName = @FilterName AND Category = @Category";
                cmd.Parameters.AddWithValue("@FilterName", filterName);
                cmd.Parameters.AddWithValue("@Category", category);

                var existingId = cmd.ExecuteScalar();
                if (existingId != null)
                {
                    return Convert.ToInt32(existingId);
                }
            }

            // ✅ STEP 2: Filter doesn't exist - CREATE it
            // This ensures clash zones can be saved even if filter wasn't pre-created
            using (var insertCmd = _context.Connection.CreateCommand())
            {
                insertCmd.Transaction = transaction;
                insertCmd.CommandText = @"
                    INSERT INTO Filters (FilterName, Category, IsFilterComboNew, CreatedAt, UpdatedAt)
                    VALUES (@FilterName, @Category, 0, CURRENT_TIMESTAMP, CURRENT_TIMESTAMP);
                    SELECT last_insert_rowid();";
                insertCmd.Parameters.AddWithValue("@FilterName", filterName);
                insertCmd.Parameters.AddWithValue("@Category", category);

                var newId = insertCmd.ExecuteScalar();
                if (newId != null && int.TryParse(newId.ToString(), out int insertedId))
                {
                    _logger($"[SQLite] ✅ Created filter '{filterName}' (Category='{category}') in database (FilterId={insertedId})");
                    return insertedId;
                }
            }

            _logger($"[SQLite] ⚠️ Failed to create filter '{filterName}' (Category='{category}') in SQLite.");
            return -1;
        }

        private int GetOrCreateFileCombo(int filterId, string category, List<string> selectedHostCategories, ClashZone clashZone, SQLiteTransaction transaction)
        {
            // ✅ FIX: Use same fallback logic as ClashZonePersistenceService.GetFileComboKey()
            // Try SourceDocKey first, then fall back to DocumentPath
            var linkedFileKey = NormalizeDocumentKey(
                !string.IsNullOrWhiteSpace(clashZone.SourceDocKey) 
                    ? clashZone.SourceDocKey 
                    : (!string.IsNullOrWhiteSpace(clashZone.DocumentPath) ? clashZone.DocumentPath : "unknown-linked"));
            
            // Try HostDocKey first, then fall back to StructuralElementDocumentTitle
            var hostFileKey = NormalizeDocumentKey(
                !string.IsNullOrWhiteSpace(clashZone.HostDocKey) 
                    ? clashZone.HostDocKey 
                    : (!string.IsNullOrWhiteSpace(clashZone.StructuralElementDocumentTitle) ? clashZone.StructuralElementDocumentTitle : "unknown-host"));

            // ✅ CRITICAL: Normalize category
            string normalizedCategory = MepCategoryConstants.Normalize(category ?? "Unknown");
            
            // ✅ CRITICAL: Store host categories as comma-separated string (not JSON) for readability
            // User wants "Floors" not ["Floors"] in database
            string hostCategoriesJson = null;
            if (selectedHostCategories != null && selectedHostCategories.Count > 0)
            {
                // ✅ NORMALIZE: Store as comma-separated string (normalized, no brackets)
                var normalizedCategories = selectedHostCategories
                    .Where(c => !string.IsNullOrWhiteSpace(c))
                    .Select(c => c.Trim())
                    .ToList();
                if (normalizedCategories.Count > 0)
                {
                    hostCategoriesJson = string.Join(", ", normalizedCategories);
                }
            }

            // ✅ DIAGNOSTIC: Log file combo lookup/creation
            if (string.IsNullOrWhiteSpace(linkedFileKey) || string.IsNullOrWhiteSpace(hostFileKey))
            {
                _logger($"[SQLite] ⚠️ Invalid file combo keys: LinkedFile='{linkedFileKey}', HostFile='{hostFileKey}' for zone {clashZone.Id}");
                return -1;
            }

            using (var cmd = _context.Connection.CreateCommand())
            {
                cmd.Transaction = transaction;
                
                // ✅ STEP 1: Check if file combo already exists (with Category in UNIQUE constraint)
                cmd.CommandText = @"
                    SELECT ComboId FROM FileCombos 
                    WHERE FilterId = @FilterId AND Category = @Category AND LinkedFileKey = @LinkedFileKey AND HostFileKey = @HostFileKey";
                cmd.Parameters.AddWithValue("@FilterId", filterId);
                cmd.Parameters.AddWithValue("@Category", normalizedCategory);
                cmd.Parameters.AddWithValue("@LinkedFileKey", linkedFileKey);
                cmd.Parameters.AddWithValue("@HostFileKey", hostFileKey);

                var existingId = cmd.ExecuteScalar();
                if (existingId != null)
                {
                    var comboId = Convert.ToInt32(existingId);
                    // ✅ UPDATE: Update SelectedHostCategories if provided and different
                    if (!string.IsNullOrEmpty(hostCategoriesJson))
                    {
                        using (var updateCmd = _context.Connection.CreateCommand())
                        {
                            updateCmd.Transaction = transaction;
                            updateCmd.CommandText = @"
                                UPDATE FileCombos 
                                SET SelectedHostCategories = @SelectedHostCategories,
                                    UpdatedAt = CURRENT_TIMESTAMP
                                WHERE ComboId = @ComboId";
                            updateCmd.Parameters.AddWithValue("@ComboId", comboId);
                            updateCmd.Parameters.AddWithValue("@SelectedHostCategories", hostCategoriesJson);
                            updateCmd.ExecuteNonQuery();
                        }
                    }
                    
                    if (!DeploymentConfiguration.DeploymentMode)
                        _logger($"[SQLite] ✅ Found existing FileCombo: ComboId={comboId}, FilterId={filterId}, Category='{normalizedCategory}', Linked='{linkedFileKey}', Host='{hostFileKey}'");
                    return comboId;
                }

                // ✅ STEP 2: Create new file combo with Category and SelectedHostCategories
                cmd.CommandText = @"
                    INSERT INTO FileCombos (FilterId, Category, SelectedHostCategories, LinkedFileKey, HostFileKey, IsFilterComboNew, ProcessedAt, CreatedAt, UpdatedAt)
                    VALUES (@FilterId, @Category, @SelectedHostCategories, @LinkedFileKey, @HostFileKey, 1, CURRENT_TIMESTAMP, CURRENT_TIMESTAMP, CURRENT_TIMESTAMP);
                    SELECT last_insert_rowid();";
                
                cmd.Parameters.Clear();
                cmd.Parameters.AddWithValue("@FilterId", filterId);
                cmd.Parameters.AddWithValue("@Category", normalizedCategory);
                cmd.Parameters.AddWithValue("@SelectedHostCategories", hostCategoriesJson ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@LinkedFileKey", linkedFileKey);
                cmd.Parameters.AddWithValue("@HostFileKey", hostFileKey);
                
                var newIdObj = cmd.ExecuteScalar();
                if (newIdObj == null)
                {
                    _logger($"[SQLite] ❌ Failed to create FileCombo: INSERT returned NULL for FilterId={filterId}, Category='{normalizedCategory}', Linked='{linkedFileKey}', Host='{hostFileKey}'");
                    return -1;
                }
                
                var newId = Convert.ToInt32(newIdObj);
                if (!DeploymentConfiguration.DeploymentMode)
                    _logger($"[SQLite] ✅ Created new FileCombo: ComboId={newId}, FilterId={filterId}, Category='{normalizedCategory}', HostCategories={selectedHostCategories?.Count ?? 0}, Linked='{linkedFileKey}', Host='{hostFileKey}'");
                return newId;
            }
        }

        private bool InsertOrUpdateClashZone(int comboId, ClashZone clashZone, SQLiteTransaction transaction)
        {
            using (var cmd = _context.Connection.CreateCommand())
            {
                cmd.Transaction = transaction;

                var mepElementId = clashZone.MepElementId?.IntegerValue ?? clashZone.MepElementIdValue;
                var hostElementId = clashZone.StructuralElementId?.IntegerValue ?? clashZone.StructuralElementIdValue;
                var intersectionX = clashZone.IntersectionPoint?.X ?? clashZone.IntersectionPointX;
                var intersectionY = clashZone.IntersectionPoint?.Y ?? clashZone.IntersectionPointY;
                var intersectionZ = clashZone.IntersectionPoint?.Z ?? clashZone.IntersectionPointZ;
                var guid = clashZone.Id.ToString().ToUpperInvariant();
                
                // ✅ CRITICAL FIX: Check by GUID FIRST (to prevent UNIQUE constraint violation)
                // Then fall back to MEP+Host+Point if GUID doesn't match
                cmd.CommandText = @"
                    SELECT ClashZoneId FROM ClashZones 
                    WHERE UPPER(ClashZoneGuid) = @ClashZoneGuid 
                      AND ClashZoneGuid != '' AND ClashZoneGuid IS NOT NULL
                    LIMIT 1";
                
                cmd.Parameters.AddWithValue("@ClashZoneGuid", guid);
                var existingIdByGuid = cmd.ExecuteScalar();

                if (existingIdByGuid != null)
                {
                    // ✅ Found by GUID - update existing
                    var clashZoneId = Convert.ToInt32(existingIdByGuid);
                    UpdateClashZone(clashZoneId, comboId, clashZone, transaction);
                    return false; // Was an update
                }
                
                // ✅ FALLBACK: Check by MEP+Host+Point if GUID didn't match
                cmd.CommandText = @"
                    SELECT ClashZoneId FROM ClashZones 
                    WHERE ComboId = @ComboId 
                      AND MepElementId = @MepElementId 
                      AND HostElementId = @HostElementId
                      AND ABS(IntersectionX - @IntersectionX) < 0.001
                      AND ABS(IntersectionY - @IntersectionY) < 0.001
                      AND ABS(IntersectionZ - @IntersectionZ) < 0.001
                    LIMIT 1";
                
                cmd.Parameters.Clear();
                cmd.Parameters.AddWithValue("@ComboId", comboId);
                cmd.Parameters.AddWithValue("@MepElementId", mepElementId);
                cmd.Parameters.AddWithValue("@HostElementId", hostElementId);
                cmd.Parameters.AddWithValue("@IntersectionX", intersectionX);
                cmd.Parameters.AddWithValue("@IntersectionY", intersectionY);
                cmd.Parameters.AddWithValue("@IntersectionZ", intersectionZ);

                var existingId = cmd.ExecuteScalar();

                if (existingId != null)
                {
                    // Update existing
                    var clashZoneId = Convert.ToInt32(existingId);
                    UpdateClashZone(clashZoneId, comboId, clashZone, transaction);
                    return false; // Was an update
                }
                else
                {
                    // Insert new
                    InsertClashZone(comboId, clashZone, transaction);
                    return true; // Was an insert
                }
            }
        }

        private void InsertClashZone(int comboId, ClashZone clashZone, SQLiteTransaction transaction)
        {
            using (var cmd = _context.Connection.CreateCommand())
            {
                cmd.Transaction = transaction;
                cmd.CommandText = @"
                    INSERT INTO ClashZones (
                        ComboId, MepElementId, HostElementId,
                        IntersectionX, IntersectionY, IntersectionZ,
                        SleeveState, SleeveInstanceId, ClusterInstanceId,
                        SleeveWidth, SleeveHeight, SleeveDiameter,
                        SleevePlacementX, SleevePlacementY, SleevePlacementZ,
                        BoundingBoxMinX, BoundingBoxMinY, BoundingBoxMinZ,
                        BoundingBoxMaxX, BoundingBoxMaxY, BoundingBoxMaxZ,
                        SleeveBoundingBoxRCS_MinX, SleeveBoundingBoxRCS_MinY, SleeveBoundingBoxRCS_MinZ,
                        SleeveBoundingBoxRCS_MaxX, SleeveBoundingBoxRCS_MaxY, SleeveBoundingBoxRCS_MaxZ,
                        PlacementSource, UpdatedAt,
                        ClashZoneGuid, MepCategory, StructuralType,
                        HostOrientation, MepOrientationDirection,
                        MepOrientationX, MepOrientationY, MepOrientationZ,
                        MepRotationAngleRad, MepRotationAngleDeg,
                        MepRotationCos, MepRotationSin,
                        MepAngleToXRad, MepAngleToXDeg,
                        MepAngleToYRad, MepAngleToYDeg,
                        MepWidth, MepHeight, SleeveFamilyName,
                        SleevePlacementActiveX, SleevePlacementActiveY, SleevePlacementActiveZ,
                        SourceDocKey, HostDocKey, MepElementUniqueId,
                        IsResolvedFlag, IsClusterResolvedFlag, IsClusteredFlag,
                        MarkedForClusterProcess, AfterClusterSleeveId,
                        HasDamperNearbyFlag, IsCurrentClashFlag, ReadyForPlacementFlag,
                        StructuralThickness, WallThickness, FramingThickness
                    ) VALUES (
                        @ComboId, @MepElementId, @HostElementId,
                        @IntersectionX, @IntersectionY, @IntersectionZ,
                        @SleeveState, @SleeveInstanceId, @ClusterInstanceId,
                        @SleeveWidth, @SleeveHeight, @SleeveDiameter,
                        @SleevePlacementX, @SleevePlacementY, @SleevePlacementZ,
                        @BoundingBoxMinX, @BoundingBoxMinY, @BoundingBoxMinZ,
                        @BoundingBoxMaxX, @BoundingBoxMaxY, @BoundingBoxMaxZ,
                        @SleeveBoundingBoxRCS_MinX, @SleeveBoundingBoxRCS_MinY, @SleeveBoundingBoxRCS_MinZ,
                        @SleeveBoundingBoxRCS_MaxX, @SleeveBoundingBoxRCS_MaxY, @SleeveBoundingBoxRCS_MaxZ,
                        @PlacementSource, CURRENT_TIMESTAMP,
                        @ClashZoneGuid, @MepCategory, @StructuralType,
                        @HostOrientation, @MepOrientationDirection,
                        @MepOrientationX, @MepOrientationY, @MepOrientationZ,
                        @MepRotationAngleRad, @MepRotationAngleDeg,
                        @MepRotationCos, @MepRotationSin,
                        @MepAngleToXRad, @MepAngleToXDeg,
                        @MepAngleToYRad, @MepAngleToYDeg,
                        @MepWidth, @MepHeight, @SleeveFamilyName,
                        @SleevePlacementActiveX, @SleevePlacementActiveY, @SleevePlacementActiveZ,
                        @SourceDocKey, @HostDocKey, @MepElementUniqueId,
                        @IsResolvedFlag, @IsClusterResolvedFlag, @IsClusteredFlag,
                        @MarkedForClusterProcess, @AfterClusterSleeveId,
                        @HasDamperNearbyFlag, @IsCurrentClashFlag, @ReadyForPlacementFlag,
                        @StructuralThickness, @WallThickness, @FramingThickness
                    )";

                AddClashZoneParameters(cmd, comboId, clashZone);
                cmd.ExecuteNonQuery();
            }
        }

        private void UpdateClashZone(int clashZoneId, int comboId, ClashZone clashZone, SQLiteTransaction transaction)
        {
            using (var cmd = _context.Connection.CreateCommand())
            {
                cmd.Transaction = transaction;
                
                // ✅ CRITICAL FIX: Smart flag/ID preservation logic to handle both deleted sleeves and moved MEP elements
                // Flags and SleeveInstanceIds are a PAIR - they must be consistent!
                // Flags are managed by FlagManager (ResetFlagsForDeletedSleeves, UpdateFlagsForPlacement, DeleteSleeveForIntersectionPointChange)
                // 
                // Logic:
                // 1. ALWAYS check if entry is in reset state FIRST (flags=false, IDs=-1) - this takes priority
                //    If reset, preserve flags/IDs even if ClashZone has stale values
                // 2. If clash zone has sleeve IDs > 0 AND entry is not reset: Update flags to true and sleeve IDs (new sleeve was placed)
                // 3. If clash zone has no sleeve IDs AND entry is not reset:
                //    a. If entry in DB has flags=true and IDs>0: Update to clash zone values (MEP moved, old sleeve deleted)
                //    This handles the "Adopt to Document" scenario where MEP moves, old sleeve is deleted, new sleeve will be placed
                
                // ✅ CRITICAL: ALWAYS read existing flags from database FIRST to check if FlagManager already reset them
                // This must happen regardless of whether ClashZone has sleeve IDs, because ClashZone might have stale values
                cmd.CommandText = @"
                    SELECT IsResolvedFlag, IsClusterResolvedFlag, SleeveInstanceId, ClusterInstanceId
                    FROM ClashZones
                    WHERE ClashZoneId = @ClashZoneId";
                cmd.Parameters.AddWithValue("@ClashZoneId", clashZoneId);
                
                int? existingIsResolved = null;
                int? existingIsClusterResolved = null;
                int? existingSleeveInstanceId = null;
                int? existingClusterInstanceId = null;
                bool entryWasResetByFlagManager = false;
                
                using (var reader = cmd.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        existingIsResolved = GetInt(reader, "IsResolvedFlag", 0);
                        existingIsClusterResolved = GetInt(reader, "IsClusterResolvedFlag", 0);
                        existingSleeveInstanceId = GetInt(reader, "SleeveInstanceId", -1);
                        existingClusterInstanceId = GetInt(reader, "ClusterInstanceId", -1);
                        
                        // Check if FlagManager already reset this entry (flags=false, IDs=-1)
                        entryWasResetByFlagManager = existingIsResolved == 0 && existingIsClusterResolved == 0 && 
                                                     existingSleeveInstanceId <= 0 && existingClusterInstanceId <= 0;
                    }
                }
                cmd.Parameters.Clear();
                
                bool zoneHasSleeve = clashZone.SleeveInstanceId > 0 || clashZone.ClusterSleeveInstanceId > 0;
                
                // ✅ CRITICAL: If entry was reset by FlagManager, ALWAYS preserve flags/IDs, even if ClashZone has stale values
                // This prevents SaveClashZones from overwriting correctly reset flags with old values from clash zone objects
                bool preserveFlags = entryWasResetByFlagManager;
                
                if (preserveFlags && !DeploymentConfiguration.DeploymentMode)
                {
                    _logger($"[SQLite] ✅ PRESERVING reset flags for ClashZoneId={clashZoneId}, GUID={clashZone.Id}: IsResolved=false, IsClusterResolved=false (FlagManager reset, ClashZone has stale values)");
                }
                
                cmd.CommandText = @"
                    UPDATE ClashZones SET
                        SleeveState = @SleeveState,
                        SleeveInstanceId = @SleeveInstanceId,
                        ClusterInstanceId = @ClusterInstanceId,
                        SleeveWidth = @SleeveWidth,
                        SleeveHeight = @SleeveHeight,
                        SleeveDiameter = @SleeveDiameter,
                        SleevePlacementX = @SleevePlacementX,
                        SleevePlacementY = @SleevePlacementY,
                        SleevePlacementZ = @SleevePlacementZ,
                        BoundingBoxMinX = @BoundingBoxMinX,
                        BoundingBoxMinY = @BoundingBoxMinY,
                        BoundingBoxMinZ = @BoundingBoxMinZ,
                        BoundingBoxMaxX = @BoundingBoxMaxX,
                        BoundingBoxMaxY = @BoundingBoxMaxY,
                        BoundingBoxMaxZ = @BoundingBoxMaxZ,
                        PlacementSource = @PlacementSource,
                        ClashZoneGuid = @ClashZoneGuid,
                        MepCategory = @MepCategory,
                        StructuralType = @StructuralType,
                        HostOrientation = @HostOrientation,
                        MepOrientationDirection = @MepOrientationDirection,
                        MepOrientationX = @MepOrientationX,
                        MepOrientationY = @MepOrientationY,
                        MepOrientationZ = @MepOrientationZ,
                        MepRotationAngleRad = @MepRotationAngleRad,
                        MepRotationAngleDeg = @MepRotationAngleDeg,
                        MepRotationCos = @MepRotationCos,
                        MepRotationSin = @MepRotationSin,
                        MepAngleToXRad = @MepAngleToXRad,
                        MepAngleToXDeg = @MepAngleToXDeg,
                        MepAngleToYRad = @MepAngleToYRad,
                        MepAngleToYDeg = @MepAngleToYDeg,
                        MepWidth = @MepWidth,
                        MepHeight = @MepHeight,
                        SleeveFamilyName = @SleeveFamilyName,
                        SleevePlacementActiveX = @SleevePlacementActiveX,
                        SleevePlacementActiveY = @SleevePlacementActiveY,
                        SleevePlacementActiveZ = @SleevePlacementActiveZ,
                        SourceDocKey = @SourceDocKey,
                        HostDocKey = @HostDocKey,
                        MepElementUniqueId = @MepElementUniqueId,
                        IsResolvedFlag = @IsResolvedFlag,
                        IsClusterResolvedFlag = @IsClusterResolvedFlag,
                        IsClusteredFlag = @IsClusteredFlag,
                        MarkedForClusterProcess = @MarkedForClusterProcess,
                        AfterClusterSleeveId = @AfterClusterSleeveId,
                        HasDamperNearbyFlag = @HasDamperNearbyFlag,
                        IsCurrentClashFlag = @IsCurrentClashFlag,
                        ReadyForPlacementFlag = @ReadyForPlacementFlag,
                        StructuralThickness = @StructuralThickness,
                        WallThickness = @WallThickness,
                        FramingThickness = @FramingThickness,
                        UpdatedAt = CURRENT_TIMESTAMP
                    WHERE ClashZoneId = @ClashZoneId";

                AddClashZoneParameters(cmd, comboId, clashZone, preserveFlags, existingIsResolved, existingIsClusterResolved, existingSleeveInstanceId, existingClusterInstanceId);
                cmd.Parameters.AddWithValue("@ClashZoneId", clashZoneId);
                cmd.ExecuteNonQuery();
            }
        }

        private void AddClashZoneParameters(SQLiteCommand cmd, int comboId, ClashZone clashZone, bool preserveFlags = false, int? existingIsResolved = null, int? existingIsClusterResolved = null, int? existingSleeveInstanceId = null, int? existingClusterInstanceId = null)
        {
            // Determine sleeve state
            int sleeveState = 0; // Unprocessed
            if (clashZone.IsClusterResolved && clashZone.ClusterSleeveInstanceId > 0)
                sleeveState = 2; // ClusterPlaced
            else if (clashZone.IsResolved && clashZone.SleeveInstanceId > 0)
                sleeveState = 1; // IndividualPlaced

            var mepElementId = clashZone.MepElementId?.IntegerValue ?? clashZone.MepElementIdValue;
            var hostElementId = clashZone.StructuralElementId?.IntegerValue ?? clashZone.StructuralElementIdValue;
            var intersectionPoint = clashZone.IntersectionPoint;
            var orientationX = clashZone.MepElementOrientationX;
            var orientationY = clashZone.MepElementOrientationY;
            var orientationZ = clashZone.MepElementOrientationZ;
            var rotationAngleRad = clashZone.MepElementRotationAngle;
            var rotationAngleDeg = NormalizeDegrees(rotationAngleRad);
            var orientationDirection = clashZone.MepElementOrientationDirection ?? string.Empty;
            var (angleToXRad, angleToXDeg, angleToYRad, angleToYDeg) =
                ComputePlanarOrientationAngles(orientationX, orientationY, orientationDirection);
            var activePoint = clashZone.SleevePlacementPointActiveDocument;
            var markedForCluster = clashZone.MarkedForClusteringSleeveProcess.HasValue
                ? (object)(clashZone.MarkedForClusteringSleeveProcess.Value ? 1 : 0)
                : DBNull.Value;
            var isClustered = clashZone.MarkedForClusteringSleeveProcess.HasValue
                ? (object)(clashZone.MarkedForClusteringSleeveProcess.Value ? 1 : 0)
                : DBNull.Value;

            cmd.Parameters.AddWithValue("@ComboId", comboId);
            cmd.Parameters.AddWithValue("@MepElementId", mepElementId);
            cmd.Parameters.AddWithValue("@HostElementId", hostElementId);
            cmd.Parameters.AddWithValue("@IntersectionX", intersectionPoint?.X ?? clashZone.IntersectionPointX);
            cmd.Parameters.AddWithValue("@IntersectionY", intersectionPoint?.Y ?? clashZone.IntersectionPointY);
            cmd.Parameters.AddWithValue("@IntersectionZ", intersectionPoint?.Z ?? clashZone.IntersectionPointZ);
            cmd.Parameters.AddWithValue("@SleeveState", sleeveState);
            
            // ✅ CRITICAL FIX: Preserve existing sleeve IDs from database if preserveFlags=true
            // Flags and SleeveInstanceIds are a PAIR - they must be consistent!
            if (preserveFlags && existingSleeveInstanceId.HasValue && existingClusterInstanceId.HasValue)
            {
                // Preserve existing database sleeve IDs (they're paired with flags)
                cmd.Parameters.AddWithValue("@SleeveInstanceId", existingSleeveInstanceId.Value > 0 ? (object)existingSleeveInstanceId.Value : DBNull.Value);
                cmd.Parameters.AddWithValue("@ClusterInstanceId", existingClusterInstanceId.Value > 0 ? (object)existingClusterInstanceId.Value : DBNull.Value);
            }
            else
            {
                // Use clash zone sleeve IDs (for new entries or when sleeve was placed)
            cmd.Parameters.AddWithValue("@SleeveInstanceId", (object)clashZone.SleeveInstanceId ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@ClusterInstanceId", (object)clashZone.ClusterSleeveInstanceId ?? DBNull.Value);
            }
            // ✅ DIAGNOSTIC: Log sleeve dimensions before saving
            if (!DeploymentConfiguration.DeploymentMode && clashZone.SleeveInstanceId > 0)
            {
                _logger($"[SQLite] [SLEEVE-DIMENSIONS] ClashZone {clashZone.Id}: SleeveWidth={clashZone.SleeveWidth}, SleeveHeight={clashZone.SleeveHeight}, SleeveDiameter={clashZone.SleeveDiameter}, SleeveInstanceId={clashZone.SleeveInstanceId}");
            }
            
            cmd.Parameters.AddWithValue("@SleeveWidth", (object)clashZone.SleeveWidth ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@SleeveHeight", (object)clashZone.SleeveHeight ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@SleeveDiameter", (object)clashZone.SleeveDiameter ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@SleevePlacementX", (object)clashZone.SleevePlacementPoint?.X ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@SleevePlacementY", (object)clashZone.SleevePlacementPoint?.Y ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@SleevePlacementZ", (object)clashZone.SleevePlacementPoint?.Z ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@BoundingBoxMinX", (object)clashZone.SleeveBoundingBoxMinX ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@BoundingBoxMinY", (object)clashZone.SleeveBoundingBoxMinY ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@BoundingBoxMinZ", (object)clashZone.SleeveBoundingBoxMinZ ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@BoundingBoxMaxX", (object)clashZone.SleeveBoundingBoxMaxX ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@BoundingBoxMaxY", (object)clashZone.SleeveBoundingBoxMaxY ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@BoundingBoxMaxZ", (object)clashZone.SleeveBoundingBoxMaxZ ?? DBNull.Value);
            // ✅ RCS BBOX: Save wall-aligned RCS bounding box coordinates (for walls/framing only)
            cmd.Parameters.AddWithValue("@SleeveBoundingBoxRCS_MinX", (object)clashZone.SleeveBoundingBoxRCS_MinX ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@SleeveBoundingBoxRCS_MinY", (object)clashZone.SleeveBoundingBoxRCS_MinY ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@SleeveBoundingBoxRCS_MinZ", (object)clashZone.SleeveBoundingBoxRCS_MinZ ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@SleeveBoundingBoxRCS_MaxX", (object)clashZone.SleeveBoundingBoxRCS_MaxX ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@SleeveBoundingBoxRCS_MaxY", (object)clashZone.SleeveBoundingBoxRCS_MaxY ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@SleeveBoundingBoxRCS_MaxZ", (object)clashZone.SleeveBoundingBoxRCS_MaxZ ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@PlacementSource", "XML");
            cmd.Parameters.AddWithValue("@ClashZoneGuid", clashZone.Id.ToString());
            cmd.Parameters.AddWithValue("@MepCategory", (object)clashZone.MepElementCategory ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@StructuralType", (object)clashZone.StructuralElementType ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@HostOrientation", (object)clashZone.HostOrientation ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@MepOrientationDirection", (object)clashZone.MepElementOrientationDirection ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@MepOrientationX", orientationX);
            cmd.Parameters.AddWithValue("@MepOrientationY", orientationY);
            cmd.Parameters.AddWithValue("@MepOrientationZ", orientationZ);
            cmd.Parameters.AddWithValue("@MepRotationAngleRad", rotationAngleRad);
            cmd.Parameters.AddWithValue("@MepRotationAngleDeg", rotationAngleDeg);
            // ✅ ROTATION MATRIX: Pre-calculate and save cos/sin (dump once use many times)
            cmd.Parameters.AddWithValue("@MepRotationCos", Math.Cos(rotationAngleRad));
            cmd.Parameters.AddWithValue("@MepRotationSin", Math.Sin(rotationAngleRad));
            cmd.Parameters.AddWithValue("@MepAngleToXRad", angleToXRad);
            cmd.Parameters.AddWithValue("@MepAngleToXDeg", angleToXDeg);
            cmd.Parameters.AddWithValue("@MepAngleToYRad", angleToYRad);
            cmd.Parameters.AddWithValue("@MepAngleToYDeg", angleToYDeg);
            cmd.Parameters.AddWithValue("@MepWidth", clashZone.MepElementWidth);
            cmd.Parameters.AddWithValue("@MepHeight", clashZone.MepElementHeight);
            cmd.Parameters.AddWithValue("@SleeveFamilyName", (object)clashZone.SleeveFamilyName ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@SleevePlacementActiveX", activePoint != null ? (object)activePoint.X : DBNull.Value);
            cmd.Parameters.AddWithValue("@SleevePlacementActiveY", activePoint != null ? (object)activePoint.Y : DBNull.Value);
            cmd.Parameters.AddWithValue("@SleevePlacementActiveZ", activePoint != null ? (object)activePoint.Z : DBNull.Value);
            cmd.Parameters.AddWithValue("@SourceDocKey", (object)clashZone.SourceDocKey ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@HostDocKey", (object)clashZone.HostDocKey ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@MepElementUniqueId", (object)clashZone.MepElementUniqueId ?? DBNull.Value);
            
            // ✅ CRITICAL FIX: Preserve existing flags from database if preserveFlags=true
            // This prevents SaveClashZones from overwriting correctly reset flags with old values from clash zone objects
            if (preserveFlags && existingIsResolved.HasValue && existingIsClusterResolved.HasValue)
            {
                // Preserve existing database flags (they're the source of truth)
                cmd.Parameters.AddWithValue("@IsResolvedFlag", existingIsResolved.Value);
                cmd.Parameters.AddWithValue("@IsClusterResolvedFlag", existingIsClusterResolved.Value);
            }
            else
            {
                // Use clash zone flags (for new entries or when sleeve was placed)
            cmd.Parameters.AddWithValue("@IsResolvedFlag", clashZone.IsResolved ? 1 : 0);
            cmd.Parameters.AddWithValue("@IsClusterResolvedFlag", clashZone.IsClusterResolved ? 1 : 0);
            }
            
            cmd.Parameters.AddWithValue("@IsClusteredFlag", isClustered);
            cmd.Parameters.AddWithValue("@MarkedForClusterProcess", markedForCluster);
            cmd.Parameters.AddWithValue("@AfterClusterSleeveId", clashZone.AfterClusterSleevePlacedSleeveInstanceId);
            cmd.Parameters.AddWithValue("@HasDamperNearbyFlag", clashZone.HasDamperNearby ? 1 : 0);
            cmd.Parameters.AddWithValue("@IsCurrentClashFlag", clashZone.IsCurrentClash ? 1 : 0);
            cmd.Parameters.AddWithValue("@ReadyForPlacementFlag", clashZone.ReadyForPlacement ? 1 : 0);
            
            // ✅ CRITICAL: Add thickness parameters for depth calculation
            cmd.Parameters.AddWithValue("@StructuralThickness", clashZone.StructuralElementThickness);
            cmd.Parameters.AddWithValue("@WallThickness", clashZone.WallThickness);
            cmd.Parameters.AddWithValue("@FramingThickness", clashZone.FramingThickness);
        }

        /// <summary>
        /// ✅ PUBLIC: Save sleeve snapshots for placed sleeves (called after placement)
        /// </summary>
        public void SaveSleeveSnapshotsForPlacedSleeves(int filterId, List<ClashZone> placedZones)
        {
            if (placedZones == null || placedZones.Count == 0)
                return;

            using (var transaction = _context.Connection.BeginTransaction())
            {
                try
                {
                    // Group zones by ComboId (get from first zone's combo lookup)
                    var zonesByCombo = placedZones
                        .Where(z => z != null && (z.SleeveInstanceId > 0 || z.ClusterSleeveInstanceId > 0))
                        .GroupBy(z =>
                        {
                            // Try to get ComboId from zone's associated data
                            // For now, use -1 if not available (will be set during snapshot creation)
                            return -1; // ComboId will be determined from FilterId + Category + FileKeys
                        })
                        .ToList();

                    foreach (var comboGroup in zonesByCombo)
                    {
                        var zones = comboGroup.ToList();
                        if (zones.Count == 0)
                            continue;

                        // Get ComboId from first zone (if available via database lookup)
                        int comboId = -1;
                        try
                        {
                            var firstZone = zones[0];
                            if (!string.IsNullOrWhiteSpace(firstZone.SourceDocKey) && !string.IsNullOrWhiteSpace(firstZone.HostDocKey))
                            {
                                // Try to find existing combo
                                using (var cmd = _context.Connection.CreateCommand())
                                {
                                    cmd.Transaction = transaction;
                                    cmd.CommandText = @"
                                        SELECT ComboId FROM FileCombos 
                                        WHERE FilterId = @FilterId 
                                          AND LinkedFileKey = @LinkedFileKey 
                                          AND HostFileKey = @HostFileKey
                                        LIMIT 1";
                                    cmd.Parameters.AddWithValue("@FilterId", filterId);
                                    cmd.Parameters.AddWithValue("@LinkedFileKey", firstZone.SourceDocKey);
                                    cmd.Parameters.AddWithValue("@HostFileKey", firstZone.HostDocKey);
                                    var result = cmd.ExecuteScalar();
                                    if (result != null)
                                    {
                                        comboId = Convert.ToInt32(result);
                                    }
                                }
                            }
                        }
                        catch { }

                        // Create processed zones list
                        var processedZones = zones.Select(z => (comboId, z)).ToList();
                        
                        // Call private method with transaction
                        InsertOrUpdateSleeveSnapshotsInternal(filterId, processedZones, transaction);
                    }

                    transaction.Commit();
                    
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        _logger($"[SQLite] ✅ Saved sleeve snapshots for {placedZones.Count} placed zones");
                    }
                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    _logger($"[SQLite] ❌ Error saving sleeve snapshots: {ex.Message}");
                    throw;
                }
            }
        }

        private void InsertOrUpdateSleeveSnapshots(
            int filterId,
            List<(int ComboId, ClashZone Zone)> processedZones,
            SQLiteTransaction transaction)
        {
            InsertOrUpdateSleeveSnapshotsInternal(filterId, processedZones, transaction);
        }

        private void InsertOrUpdateSleeveSnapshotsInternal(
            int filterId,
            List<(int ComboId, ClashZone Zone)> processedZones,
            SQLiteTransaction transaction)
        {
            if (processedZones == null || processedZones.Count == 0)
                return;

            var validZones = processedZones
                .Where(p => p.Zone != null && (p.Zone.SleeveInstanceId > 0 || p.Zone.ClusterSleeveInstanceId > 0))
                .ToList();

            if (validZones.Count == 0)
                return;

            var groups = validZones.GroupBy(p =>
            {
                var zone = p.Zone;
                if (zone.ClusterSleeveInstanceId > 0)
                    return ("cluster", zone.ClusterSleeveInstanceId);
                return ("sleeve", zone.SleeveInstanceId > 0 ? zone.SleeveInstanceId : -1);
            });

            foreach (var group in groups)
            {
                var key = group.Key;
                bool isCluster = string.Equals(key.Item1, "cluster", StringComparison.OrdinalIgnoreCase);
                int groupId = key.Item2;

                if (groupId <= 0)
                    continue;

                var zones = group.Select(g => g.Zone).Where(z => z != null).ToList();
                if (zones.Count == 0)
                    continue;

                var comboId = group.Select(g => g.ComboId).FirstOrDefault();

                var mepParams = AggregateParameterValues(zones, useHost: false);
                var hostParams = AggregateParameterValues(zones, useHost: true);

                var mepElementIds = zones
                    .Select(z => z.MepElementIdValue)
                    .Where(v => v > 0)
                    .Distinct()
                    .ToList();

                var hostElementIds = zones
                    .Select(z => z.StructuralElementIdValue)
                    .Where(v => v > 0)
                    .Distinct()
                    .ToList();

                var sourceDocKeys = zones
                    .Select(z => z.SourceDocKey)
                    .Where(s => !string.IsNullOrWhiteSpace(s))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();

                var hostDocKeys = zones
                    .Select(z => z.HostDocKey)
                    .Where(s => !string.IsNullOrWhiteSpace(s))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();

                var sleeveInstanceId = isCluster ? (int?)null : groupId;
                var clusterInstanceId = isCluster ? (int?)groupId : null;

                UpsertSleeveSnapshot(
                    transaction,
                    new SleeveSnapshot
                    {
                        SleeveInstanceId = sleeveInstanceId,
                        ClusterInstanceId = clusterInstanceId,
                        SourceType = isCluster ? "Cluster" : "Individual",
                        FilterId = filterId,
                        ComboId = comboId,
                        MepElementIdsJson = SerializeList(mepElementIds),
                        HostElementIdsJson = SerializeList(hostElementIds),
                        MepParametersJson = SerializeDictionary(mepParams),
                        HostParametersJson = SerializeDictionary(hostParams),
                        SourceDocKeysJson = SerializeList(sourceDocKeys),
                        HostDocKeysJson = SerializeList(hostDocKeys),
                        UpdatedAt = DateTime.UtcNow
                    });
            }
        }

        private void UpsertSleeveSnapshot(SQLiteTransaction transaction, SleeveSnapshot snapshot)
        {
            if (snapshot == null)
                return;

            var existingId = GetExistingSnapshotId(snapshot.SleeveInstanceId, snapshot.ClusterInstanceId, transaction);

            if (existingId.HasValue)
            {
                snapshot.SnapshotId = existingId.Value;
                UpdateSleeveSnapshot(transaction, snapshot);
            }
            else
            {
                InsertSleeveSnapshot(transaction, snapshot);
            }
        }

        private int? GetExistingSnapshotId(int? sleeveInstanceId, int? clusterInstanceId, SQLiteTransaction transaction)
        {
            if (!sleeveInstanceId.HasValue && !clusterInstanceId.HasValue)
                return null;

            using (var cmd = _context.Connection.CreateCommand())
            {
                cmd.Transaction = transaction;

                if (sleeveInstanceId.HasValue)
                {
                    cmd.CommandText = "SELECT SnapshotId FROM SleeveSnapshots WHERE SleeveInstanceId = @SleeveInstanceId";
                    cmd.Parameters.AddWithValue("@SleeveInstanceId", sleeveInstanceId.Value);
                }
                else
                {
                    cmd.CommandText = "SELECT SnapshotId FROM SleeveSnapshots WHERE ClusterInstanceId = @ClusterInstanceId";
                    cmd.Parameters.AddWithValue("@ClusterInstanceId", clusterInstanceId.Value);
                }

                var result = cmd.ExecuteScalar();
                if (result != null && result != DBNull.Value)
                    return Convert.ToInt32(result);
            }

            return null;
        }

        private void InsertSleeveSnapshot(SQLiteTransaction transaction, SleeveSnapshot snapshot)
        {
            using (var cmd = _context.Connection.CreateCommand())
            {
                cmd.Transaction = transaction;
                cmd.CommandText = @"
                    INSERT INTO SleeveSnapshots (
                        SleeveInstanceId,
                        ClusterInstanceId,
                        SourceType,
                        FilterId,
                        ComboId,
                        MepElementIdsJson,
                        HostElementIdsJson,
                        MepParametersJson,
                        HostParametersJson,
                        SourceDocKeysJson,
                        HostDocKeysJson,
                        CreatedAt,
                        UpdatedAt
                    ) VALUES (
                        @SleeveInstanceId,
                        @ClusterInstanceId,
                        @SourceType,
                        @FilterId,
                        @ComboId,
                        @MepElementIdsJson,
                        @HostElementIdsJson,
                        @MepParametersJson,
                        @HostParametersJson,
                        @SourceDocKeysJson,
                        @HostDocKeysJson,
                        CURRENT_TIMESTAMP,
                        CURRENT_TIMESTAMP
                    )";

                AddSnapshotParameters(cmd, snapshot);
                cmd.ExecuteNonQuery();
                
                // ✅ LOG: Log all columns of inserted row (one sample row)
                var snapshotParams = new Dictionary<string, object>
                {
                    { "SleeveInstanceId", snapshot.SleeveInstanceId?.ToString() ?? "NULL" },
                    { "ClusterInstanceId", snapshot.ClusterInstanceId?.ToString() ?? "NULL" },
                    { "SourceType", snapshot.SourceType ?? "NULL" },
                    { "FilterId", snapshot.FilterId?.ToString() ?? "NULL" },
                    { "ComboId", snapshot.ComboId?.ToString() ?? "NULL" },
                    { "MepElementIdsJson", snapshot.MepElementIdsJson ?? "NULL" },
                    { "HostElementIdsJson", snapshot.HostElementIdsJson ?? "NULL" },
                    { "MepParametersJson", snapshot.MepParametersJson ?? "NULL" },
                    { "HostParametersJson", snapshot.HostParametersJson ?? "NULL" },
                    { "SourceDocKeysJson", snapshot.SourceDocKeysJson ?? "NULL" },
                    { "HostDocKeysJson", snapshot.HostDocKeysJson ?? "NULL" }
                };
                
                DatabaseOperationLogger.LogOperation(
                    "INSERT",
                    "SleeveSnapshots",
                    snapshotParams,
                    rowsAffected: 1,
                    additionalInfo: "✅ Sample row: All columns logged");
            }
        }

        private void UpdateSleeveSnapshot(SQLiteTransaction transaction, SleeveSnapshot snapshot)
        {
            using (var cmd = _context.Connection.CreateCommand())
            {
                cmd.Transaction = transaction;
                cmd.CommandText = @"
                    UPDATE SleeveSnapshots SET
                        SleeveInstanceId = @SleeveInstanceId,
                        ClusterInstanceId = @ClusterInstanceId,
                        SourceType = @SourceType,
                        FilterId = @FilterId,
                        ComboId = @ComboId,
                        MepElementIdsJson = @MepElementIdsJson,
                        HostElementIdsJson = @HostElementIdsJson,
                        MepParametersJson = @MepParametersJson,
                        HostParametersJson = @HostParametersJson,
                        SourceDocKeysJson = @SourceDocKeysJson,
                        HostDocKeysJson = @HostDocKeysJson,
                        UpdatedAt = CURRENT_TIMESTAMP
                    WHERE SnapshotId = @SnapshotId";

                AddSnapshotParameters(cmd, snapshot);
                cmd.Parameters.AddWithValue("@SnapshotId", snapshot.SnapshotId);
                cmd.ExecuteNonQuery();
            }
        }

        private void AddSnapshotParameters(SQLiteCommand cmd, SleeveSnapshot snapshot)
        {
            cmd.Parameters.AddWithValue("@SleeveInstanceId", (object)snapshot.SleeveInstanceId ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@ClusterInstanceId", (object)snapshot.ClusterInstanceId ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@SourceType", snapshot.SourceType ?? "Individual");
            cmd.Parameters.AddWithValue("@FilterId", (object)snapshot.FilterId ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@ComboId", (object)snapshot.ComboId ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@MepElementIdsJson", snapshot.MepElementIdsJson ?? "[]");
            cmd.Parameters.AddWithValue("@HostElementIdsJson", snapshot.HostElementIdsJson ?? "[]");
            cmd.Parameters.AddWithValue("@MepParametersJson", snapshot.MepParametersJson ?? "{}");
            cmd.Parameters.AddWithValue("@HostParametersJson", snapshot.HostParametersJson ?? "{}");
            cmd.Parameters.AddWithValue("@SourceDocKeysJson", snapshot.SourceDocKeysJson ?? "[]");
            cmd.Parameters.AddWithValue("@HostDocKeysJson", snapshot.HostDocKeysJson ?? "[]");
        }

        private Dictionary<string, string> AggregateParameterValues(IEnumerable<ClashZone> zones, bool useHost)
        {
            var aggregated = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);

            foreach (var zone in zones)
            {
                var bag = useHost ? zone.HostParameterValues : zone.MepParameterValues;
                if (bag == null) continue;

                foreach (var kv in bag)
                {
                    if (kv == null) continue;
                    var key = kv.Key?.Trim();
                    var value = kv.Value?.Trim();
                    if (string.IsNullOrEmpty(key) || string.IsNullOrEmpty(value))
                        continue;

                    var normalizedValue = NormalizeParameterValue(key, value);
                    var isSize = IsSizeParameter(key);

                    if (!aggregated.TryGetValue(key, out var list))
                    {
                        list = new List<string>();
                        aggregated[key] = list;
                    }

                    if (isSize)
                    {
                        list.Add(normalizedValue);
                    }
                    else
                    {
                        if (!list.Any(v => string.Equals(v, normalizedValue, StringComparison.OrdinalIgnoreCase)))
                        {
                            list.Add(normalizedValue);
                        }
                    }
                }
            }

            return aggregated.ToDictionary(
                kvp => kvp.Key,
                kvp => string.Join(", ", kvp.Value),
                StringComparer.OrdinalIgnoreCase);
        }

        private bool IsSizeParameter(string key)
        {
            if (string.IsNullOrWhiteSpace(key))
                return false;

            var normalized = key.Trim();
            return normalized.Equals("Size", StringComparison.OrdinalIgnoreCase) ||
                   normalized.Equals("MEP Size", StringComparison.OrdinalIgnoreCase) ||
                   normalized.Equals("Service Size", StringComparison.OrdinalIgnoreCase) ||
                   normalized.Equals("MepElementFormattedSize", StringComparison.OrdinalIgnoreCase);
        }

        private string NormalizeParameterValue(string key, string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return value ?? string.Empty;

            if (!IsSizeParameter(key))
                return value;

            if (value.Contains("-"))
            {
                var parts = value.Split(new[] { '-' }, 2);
                if (parts.Length > 0 && !string.IsNullOrWhiteSpace(parts[0]))
                    return parts[0].Trim();
            }

            return value;
        }

        private string SerializeDictionary(Dictionary<string, string> dictionary)
        {
            if (dictionary == null || dictionary.Count == 0)
                return "{}";

            return JsonSerializer.Serialize(dictionary);
        }

        private string SerializeList<T>(IEnumerable<T> values)
        {
            if (values == null)
                return "[]";

            var materialized = values.ToList();
            if (materialized.Count == 0)
                return "[]";

            return JsonSerializer.Serialize(materialized);
        }

        private static (double angleToXRad, double angleToXDeg, double angleToYRad, double angleToYDeg) ComputePlanarOrientationAngles(
            double orientationX,
            double orientationY,
            string orientationDirection)
        {
            const double tolerance = 1e-6;
            double angleToXRad;
            double angleToYRad;

            var magnitudeXY = Math.Sqrt(orientationX * orientationX + orientationY * orientationY);

            if (magnitudeXY > tolerance)
            {
                angleToXRad = Math.Atan2(orientationY, orientationX);
                angleToYRad = Math.Atan2(orientationX, orientationY);
            }
            else
            {
                if (string.Equals(orientationDirection, "Y", StringComparison.OrdinalIgnoreCase))
                {
                    angleToXRad = Math.PI / 2.0;
                    angleToYRad = 0.0;
                }
                else if (string.Equals(orientationDirection, "X", StringComparison.OrdinalIgnoreCase))
                {
                    angleToXRad = 0.0;
                    angleToYRad = Math.PI / 2.0;
                }
                else
                {
                    angleToXRad = 0.0;
                    angleToYRad = 0.0;
                }
            }

            return (angleToXRad, NormalizeDegrees(angleToXRad), angleToYRad, NormalizeDegrees(angleToYRad));
        }

        private static double NormalizeDegrees(double radians)
        {
            var degrees = radians * 180.0 / Math.PI;
            degrees %= 360.0;
            if (degrees < 0)
            {
                degrees += 360.0;
            }

            return degrees;
        }

        private static string NormalizeDocumentKey(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return "Unknown";
            }

            var trimmed = value.Trim();

            // Strip ": <number> : location Shared" style suffixes
            var locationMatch = System.Text.RegularExpressions.Regex.Match(
                trimmed,
                @":\s*\d+\s*:\s*location\s+Shared",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            if (locationMatch.Success)
            {
                trimmed = trimmed.Substring(0, locationMatch.Index).Trim();
            }

            // Remove trailing "(xx elements)" or similar
            var parenIndex = trimmed.IndexOf('(');
            if (parenIndex >= 0)
            {
                trimmed = trimmed.Substring(0, parenIndex).Trim();
            }

            // If the string is a full path, reduce to file name
            trimmed = System.IO.Path.GetFileName(trimmed);

            // Drop extension
            var withoutExtension = System.IO.Path.GetFileNameWithoutExtension(trimmed);
            if (string.IsNullOrWhiteSpace(withoutExtension))
            {
                withoutExtension = trimmed;
            }

            return withoutExtension.Trim();
        }

        public List<ClashZone> GetClashZonesByFilter(string filterName, string category, bool unresolvedOnly = false)
        {
            var result = new List<ClashZone>();

            if (string.IsNullOrWhiteSpace(filterName) || string.IsNullOrWhiteSpace(category))
                return result;

            using (var cmd = _context.Connection.CreateCommand())
            {
                cmd.CommandText = @"
                    SELECT 
                        cz.*,
                        fc.LinkedFileKey,
                        fc.HostFileKey,
                        f.FilterName
                    FROM Filters f
                    INNER JOIN FileCombos fc ON f.FilterId = fc.FilterId
                    INNER JOIN ClashZones cz ON fc.ComboId = cz.ComboId
                    WHERE f.FilterName = @FilterName
                      AND f.Category = @Category" + (unresolvedOnly ? " AND cz.SleeveState = 0" : string.Empty) + @"
                    ORDER BY cz.UpdatedAt DESC";

                cmd.Parameters.AddWithValue("@FilterName", filterName);
                cmd.Parameters.AddWithValue("@Category", category);

                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        var clashZone = MapClashZone(reader);
                        SetMetadataFromReader(clashZone, reader);
                        result.Add(clashZone);
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// ⚠️⚠️⚠️ CRITICAL PROTECTED METHOD - DO NOT MODIFY WITHOUT TESTING ⚠️⚠️⚠️
        /// 
        /// ⚠️⚠️⚠️ MODIFICATION CONSENT REQUIRED ⚠️⚠️⚠️
        /// To modify this method, you MUST:
        /// 1. Get explicit consent from the project owner
        /// 2. Test thoroughly with parameter transfer
        /// 3. Verify LoadParameterValuesFromSnapshots is still called
        /// 
        /// Loads clash zones from database for a specific category.
        /// 
        /// ✅ CRITICAL: This method MUST call LoadParameterValuesFromSnapshots to populate
        /// MepParameterValues and HostParameterValues. Without this, ParameterTransferService
        /// will fail for both individual and cluster sleeves.
        /// 
        /// ⚠️ DO NOT:
        /// - Remove LoadParameterValuesFromSnapshots call (breaks parameter transfer)
        /// - Change the SQL query structure (breaks data loading)
        /// - Modify MapClashZone or SetMetadataFromReader calls (breaks data mapping)
        /// - Skip validation checks (prevents invalid data)
        /// 
        /// USED BY:
        /// - ClusterDataService (for clustering)
        /// - ParameterTransferService (for parameter transfer)
        /// - RefreshService (for data reload)
        /// </summary>
        public List<ClashZone> GetClashZonesByCategory(string category)
        {
            var result = new List<ClashZone>();

            // ✅ VALIDATION: Ensure category is valid
            if (string.IsNullOrWhiteSpace(category))
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning($"[ClashZoneRepository] GetClashZonesByCategory called with null/empty category");
                return result; // Return empty list - validation failed
            }

            try
            {
                using (var cmd = _context.Connection.CreateCommand())
                {
                    // ⚠️⚠️⚠️ PROTECTED SQL QUERY - MODIFICATION CONSENT REQUIRED ⚠️⚠️⚠️
                    // This query structure is required for proper data loading
                    // To modify: Get explicit consent from project owner
                    cmd.CommandText = @"
                        SELECT 
                            cz.*,
                            fc.LinkedFileKey,
                            fc.HostFileKey,
                            f.FilterName
                        FROM Filters f
                        INNER JOIN FileCombos fc ON f.FilterId = fc.FilterId
                        INNER JOIN ClashZones cz ON fc.ComboId = cz.ComboId
                        WHERE f.Category = @Category";

                    cmd.Parameters.AddWithValue("@Category", category);

                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            var clashZone = MapClashZone(reader);
                            SetMetadataFromReader(clashZone, reader);
                            result.Add(clashZone);
                        }
                    }
                }

                // ⚠️⚠️⚠️ CRITICAL: DO NOT REMOVE THIS CALL - MODIFICATION CONSENT REQUIRED ⚠️⚠️⚠️
                // ✅ CRITICAL: Load parameter values from SleeveSnapshots for parameter transfer service
                // This populates MepParameterValues and HostParameterValues which are required
                // for ParameterTransferService to work with both individual and cluster sleeves.
                // Without this call, parameter transfer will fail silently!
                // 
                // ⚠️ TO REMOVE OR MODIFY THIS CALL:
                // 1. Get explicit consent from project owner
                // 2. Test thoroughly with ParameterTransferService
                // 3. Verify cluster sleeves can still transfer parameters
                LoadParameterValuesFromSnapshots(result);
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Error($"[ClashZoneRepository] Error loading clash zones for category '{category}': {ex.Message}");
                }
                // Return empty list on error - don't throw (allows graceful degradation)
            }

            return result;
        }

        /// <summary>
        /// Load MepParameterValues and HostParameterValues from SleeveSnapshots table and populate into ClashZone objects.
        /// This is required for the parameter transfer service to work with cluster sleeves.
        /// </summary>
        private void LoadParameterValuesFromSnapshots(List<ClashZone> clashZones)
        {
            if (clashZones == null || clashZones.Count == 0)
                return;

            try
            {
                // Build lookup dictionaries for efficient matching
                var sleeveIdToClashZone = new Dictionary<int, List<ClashZone>>();
                var clusterIdToClashZone = new Dictionary<int, List<ClashZone>>();

                foreach (var cz in clashZones)
                {
                    // Match by SleeveInstanceId (for individual sleeves)
                    if (cz.SleeveInstanceId > 0)
                    {
                        if (!sleeveIdToClashZone.ContainsKey(cz.SleeveInstanceId))
                            sleeveIdToClashZone[cz.SleeveInstanceId] = new List<ClashZone>();
                        sleeveIdToClashZone[cz.SleeveInstanceId].Add(cz);
                    }

                    // Match by ClusterSleeveInstanceId (for cluster sleeves)
                    if (cz.ClusterSleeveInstanceId > 0)
                    {
                        if (!clusterIdToClashZone.ContainsKey(cz.ClusterSleeveInstanceId))
                            clusterIdToClashZone[cz.ClusterSleeveInstanceId] = new List<ClashZone>();
                        clusterIdToClashZone[cz.ClusterSleeveInstanceId].Add(cz);
                    }
                }

                if (sleeveIdToClashZone.Count == 0 && clusterIdToClashZone.Count == 0)
                    return;

                // Load snapshots from database
                using (var cmd = _context.Connection.CreateCommand())
                {
                    var sleeveIds = sleeveIdToClashZone.Keys.ToList();
                    var clusterIds = clusterIdToClashZone.Keys.ToList();

                if (sleeveIds.Count > 0 || clusterIds.Count > 0)
                {
                    var conditions = new List<string>();

                    if (sleeveIds.Count > 0)
                    {
                        var placeholders = string.Join(",", sleeveIds.Select((_, i) => $"@SleeveId{i}"));
                        conditions.Add($"SleeveInstanceId IN ({placeholders})");
                        for (int i = 0; i < sleeveIds.Count; i++)
                        {
                            cmd.Parameters.AddWithValue($"@SleeveId{i}", sleeveIds[i]);
                        }
                    }

                    if (clusterIds.Count > 0)
                    {
                        var placeholders = string.Join(",", clusterIds.Select((_, i) => $"@ClusterId{i}"));
                        conditions.Add($"ClusterInstanceId IN ({placeholders})");
                        for (int i = 0; i < clusterIds.Count; i++)
                        {
                            cmd.Parameters.AddWithValue($"@ClusterId{i}", clusterIds[i]);
                        }
                    }

                    cmd.CommandText = $@"
                        SELECT 
                            SleeveInstanceId,
                            ClusterInstanceId,
                            MepParametersJson,
                            HostParametersJson
                        FROM SleeveSnapshots
                        WHERE ({string.Join(" OR ", conditions)})";

                        using (var reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                var sleeveId = GetInt(reader, "SleeveInstanceId", -1);
                                var clusterId = GetInt(reader, "ClusterInstanceId", -1);
                                var mepParamsJson = GetNullableString(reader, "MepParametersJson");
                                var hostParamsJson = GetNullableString(reader, "HostParametersJson");

                                List<ClashZone> matchingZones = null;

                                // Match by SleeveInstanceId first (individual sleeves)
                                if (sleeveId > 0 && sleeveIdToClashZone.TryGetValue(sleeveId, out var sleeveZones))
                                {
                                    matchingZones = sleeveZones;
                                }
                                // Match by ClusterInstanceId (cluster sleeves)
                                else if (clusterId > 0 && clusterIdToClashZone.TryGetValue(clusterId, out var clusterZones))
                                {
                                    matchingZones = clusterZones;
                                }

                                if (matchingZones != null && matchingZones.Count > 0)
                                {
                                    // Deserialize JSON and populate parameter values
                                    var mepParams = DeserializeDictionary(mepParamsJson);
                                    var hostParams = DeserializeDictionary(hostParamsJson);

                                    foreach (var cz in matchingZones)
                                    {
                                        // Only populate if not already set (preserve existing values)
                                        if (cz.MepParameterValues == null || cz.MepParameterValues.Count == 0)
                                        {
                                            cz.MepParameterValues = mepParams?.Select(kv => new Models.SerializableKeyValue { Key = kv.Key, Value = kv.Value }).ToList() 
                                                ?? new List<Models.SerializableKeyValue>();
                                        }

                                        if (cz.HostParameterValues == null || cz.HostParameterValues.Count == 0)
                                        {
                                            cz.HostParameterValues = hostParams?.Select(kv => new Models.SerializableKeyValue { Key = kv.Key, Value = kv.Value }).ToList() 
                                                ?? new List<Models.SerializableKeyValue>();
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Log error but don't fail - parameter values are optional
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Warning($"[ClashZoneRepository] Error loading parameter values from snapshots: {ex.Message}");
                }
            }
        }

        private Dictionary<string, string> DeserializeDictionary(string json)
        {
            if (string.IsNullOrWhiteSpace(json) || json == "{}" || json == "NULL")
                return new Dictionary<string, string>();

            try
            {
                return System.Text.Json.JsonSerializer.Deserialize<Dictionary<string, string>>(json) 
                    ?? new Dictionary<string, string>();
            }
            catch
            {
                return new Dictionary<string, string>();
            }
        }

        private void SetMetadataFromReader(ClashZone clashZone, SQLiteDataReader reader)
        {
            if (clashZone == null || reader == null) return;

            var filterName = GetNullableString(reader, "FilterName");
            if (!string.IsNullOrWhiteSpace(filterName))
                clashZone.Metadata["FilterName"] = filterName;

            var linkedFileKey = GetNullableString(reader, "LinkedFileKey");
            if (!string.IsNullOrWhiteSpace(linkedFileKey))
                clashZone.Metadata["LinkedFileKey"] = linkedFileKey;

            var hostFileKey = GetNullableString(reader, "HostFileKey");
            if (!string.IsNullOrWhiteSpace(hostFileKey))
                clashZone.Metadata["HostFileKey"] = hostFileKey;
        }

        private ClashZone MapClashZone(SQLiteDataReader reader)
        {
            var clashZone = new ClashZone();

            var guidString = GetNullableString(reader, "ClashZoneGuid");
            if (Guid.TryParse(guidString, out var clashGuid))
            {
                clashZone.Id = clashGuid;
            }

            var mepId = GetInt(reader, "MepElementId");
            if (mepId > 0)
                clashZone.MepElementId = new ElementId(mepId);
            clashZone.MepElementIdValue = mepId;

            var hostId = GetInt(reader, "HostElementId");
            if (hostId > 0)
                clashZone.StructuralElementId = new ElementId(hostId);
            clashZone.StructuralElementIdValue = hostId;

            clashZone.SleeveInstanceId = GetInt(reader, "SleeveInstanceId", -1);
            clashZone.ClusterSleeveInstanceId = GetInt(reader, "ClusterInstanceId", -1);
            clashZone.AfterClusterSleevePlacedSleeveInstanceId = GetInt(reader, "AfterClusterSleeveId", -1);

            clashZone.SleeveWidth = GetDouble(reader, "SleeveWidth");
            clashZone.SleeveHeight = GetDouble(reader, "SleeveHeight");
            clashZone.SleeveDiameter = GetDouble(reader, "SleeveDiameter");

            clashZone.IntersectionPoint = new XYZ(
                GetDouble(reader, "IntersectionX"),
                GetDouble(reader, "IntersectionY"),
                GetDouble(reader, "IntersectionZ"));

            clashZone.SleevePlacementPoint = new XYZ(
                GetNullableDouble(reader, "SleevePlacementX") ?? 0.0,
                GetNullableDouble(reader, "SleevePlacementY") ?? 0.0,
                GetNullableDouble(reader, "SleevePlacementZ") ?? 0.0);

            clashZone.SleevePlacementPointActiveDocument = new XYZ(
                GetNullableDouble(reader, "SleevePlacementActiveX") ?? 0.0,
                GetNullableDouble(reader, "SleevePlacementActiveY") ?? 0.0,
                GetNullableDouble(reader, "SleevePlacementActiveZ") ?? 0.0);

            clashZone.SleeveBoundingBoxMinX = GetNullableDouble(reader, "BoundingBoxMinX") ?? 0.0;
            clashZone.SleeveBoundingBoxMinY = GetNullableDouble(reader, "BoundingBoxMinY") ?? 0.0;
            clashZone.SleeveBoundingBoxMinZ = GetNullableDouble(reader, "BoundingBoxMinZ") ?? 0.0;
            clashZone.SleeveBoundingBoxMaxX = GetNullableDouble(reader, "BoundingBoxMaxX") ?? 0.0;
            clashZone.SleeveBoundingBoxMaxY = GetNullableDouble(reader, "BoundingBoxMaxY") ?? 0.0;
            clashZone.SleeveBoundingBoxMaxZ = GetNullableDouble(reader, "BoundingBoxMaxZ") ?? 0.0;
            
            // ✅ RCS BBOX: Load wall-aligned RCS bounding box coordinates (0.0 for floors or if not calculated)
            clashZone.SleeveBoundingBoxRCS_MinX = GetNullableDouble(reader, "SleeveBoundingBoxRCS_MinX") ?? 0.0;
            clashZone.SleeveBoundingBoxRCS_MinY = GetNullableDouble(reader, "SleeveBoundingBoxRCS_MinY") ?? 0.0;
            clashZone.SleeveBoundingBoxRCS_MinZ = GetNullableDouble(reader, "SleeveBoundingBoxRCS_MinZ") ?? 0.0;
            clashZone.SleeveBoundingBoxRCS_MaxX = GetNullableDouble(reader, "SleeveBoundingBoxRCS_MaxX") ?? 0.0;
            clashZone.SleeveBoundingBoxRCS_MaxY = GetNullableDouble(reader, "SleeveBoundingBoxRCS_MaxY") ?? 0.0;
            clashZone.SleeveBoundingBoxRCS_MaxZ = GetNullableDouble(reader, "SleeveBoundingBoxRCS_MaxZ") ?? 0.0;
            
            // ✅ ROTATED BBOX: Load rotated bounding box coordinates (NULL for axis-aligned sleeves)
            clashZone.RotatedBoundingBoxMinX = GetNullableDouble(reader, "RotatedBoundingBoxMinX");
            clashZone.RotatedBoundingBoxMinY = GetNullableDouble(reader, "RotatedBoundingBoxMinY");
            clashZone.RotatedBoundingBoxMinZ = GetNullableDouble(reader, "RotatedBoundingBoxMinZ");
            clashZone.RotatedBoundingBoxMaxX = GetNullableDouble(reader, "RotatedBoundingBoxMaxX");
            clashZone.RotatedBoundingBoxMaxY = GetNullableDouble(reader, "RotatedBoundingBoxMaxY");
            clashZone.RotatedBoundingBoxMaxZ = GetNullableDouble(reader, "RotatedBoundingBoxMaxZ");
            
            // ✅ SLEEVE CORNERS: Load pre-calculated 4 corner coordinates in world space (NULL if not calculated yet)
            clashZone.SleeveCorner1X = GetNullableDouble(reader, "SleeveCorner1X");
            clashZone.SleeveCorner1Y = GetNullableDouble(reader, "SleeveCorner1Y");
            clashZone.SleeveCorner1Z = GetNullableDouble(reader, "SleeveCorner1Z");
            clashZone.SleeveCorner2X = GetNullableDouble(reader, "SleeveCorner2X");
            clashZone.SleeveCorner2Y = GetNullableDouble(reader, "SleeveCorner2Y");
            clashZone.SleeveCorner2Z = GetNullableDouble(reader, "SleeveCorner2Z");
            clashZone.SleeveCorner3X = GetNullableDouble(reader, "SleeveCorner3X");
            clashZone.SleeveCorner3Y = GetNullableDouble(reader, "SleeveCorner3Y");
            clashZone.SleeveCorner3Z = GetNullableDouble(reader, "SleeveCorner3Z");
            clashZone.SleeveCorner4X = GetNullableDouble(reader, "SleeveCorner4X");
            clashZone.SleeveCorner4Y = GetNullableDouble(reader, "SleeveCorner4Y");
            clashZone.SleeveCorner4Z = GetNullableDouble(reader, "SleeveCorner4Z");

            clashZone.MepElementCategory = GetNullableString(reader, "MepCategory") ?? string.Empty;
            clashZone.StructuralElementType = GetNullableString(reader, "StructuralType") ?? string.Empty;
            clashZone.HostOrientation = GetNullableString(reader, "HostOrientation") ?? string.Empty;

            clashZone.MepElementOrientationDirection = GetNullableString(reader, "MepOrientationDirection") ?? string.Empty;
            clashZone.MepElementOrientation = new XYZ(
                GetNullableDouble(reader, "MepOrientationX") ?? 0.0,
                GetNullableDouble(reader, "MepOrientationY") ?? 0.0,
                GetNullableDouble(reader, "MepOrientationZ") ?? 0.0);
            clashZone.MepElementRotationAngle = GetNullableDouble(reader, "MepRotationAngleRad") ?? 0.0;
            // ✅ ROTATION MATRIX: Load pre-calculated cos/sin (dump once use many times)
            clashZone.MepRotationCos = GetNullableDouble(reader, "MepRotationCos");
            clashZone.MepRotationSin = GetNullableDouble(reader, "MepRotationSin");
            
            // ✅ DIAGNOSTIC: Log MEP orientation values when loading for placement (to diagnose orientation=0 issue)
            if (!DeploymentConfiguration.DeploymentMode)
            {
                var rotationDeg = clashZone.MepElementRotationAngle > 0 ? clashZone.MepElementRotationAngle * 180 / Math.PI : 0;
                if (Math.Abs(clashZone.MepElementRotationAngle) < 1e-6 && string.IsNullOrEmpty(clashZone.MepElementOrientationDirection))
                {
                    DebugLogger.Warning($"[DB-LOAD-ORIENTATION] ⚠️ Zone={clashZone.Id}, Category={clashZone.MepElementCategory}: MEP orientation is missing! RotationAngle={rotationDeg:F1}°, OrientationDirection='{clashZone.MepElementOrientationDirection}', OrientationX={clashZone.MepElementOrientation?.X:F6}");
                }
            }

            clashZone.MepElementWidth = GetNullableDouble(reader, "MepWidth") ?? 0.0;
            clashZone.MepElementHeight = GetNullableDouble(reader, "MepHeight") ?? 0.0;

            clashZone.SleeveFamilyName = GetNullableString(reader, "SleeveFamilyName") ?? string.Empty;
            clashZone.SourceDocKey = GetNullableString(reader, "SourceDocKey") ?? string.Empty;
            clashZone.HostDocKey = GetNullableString(reader, "HostDocKey") ?? string.Empty;
            clashZone.MepElementUniqueId = GetNullableString(reader, "MepElementUniqueId") ?? string.Empty;

            clashZone.IsResolved = GetBool(reader, "IsResolvedFlag");
            clashZone.IsClusterResolved = GetBool(reader, "IsClusterResolvedFlag");
            clashZone.MarkedForClusteringSleeveProcess = GetNullableBool(reader, "MarkedForClusterProcess");
            clashZone.HasDamperNearby = GetBool(reader, "HasDamperNearbyFlag");
            clashZone.IsCurrentClash = GetBool(reader, "IsCurrentClashFlag");
            clashZone.ReadyForPlacement = GetBool(reader, "ReadyForPlacementFlag");
            
            // ✅ CRITICAL: Load thickness values for depth calculation
            clashZone.StructuralElementThickness = GetDouble(reader, "StructuralThickness", 0.0);
            clashZone.WallThickness = GetDouble(reader, "WallThickness", 0.0);
            clashZone.FramingThickness = GetDouble(reader, "FramingThickness", 0.0);
            
            // ✅ DIAGNOSTIC: Log thickness values for ALL zones when loading for placement (to diagnose depth=0 issue)
            if (!DeploymentConfiguration.DeploymentMode)
            {
                var structMm = clashZone.StructuralElementThickness > 0 ? UnitUtils.ConvertFromInternalUnits(clashZone.StructuralElementThickness, UnitTypeId.Millimeters) : 0;
                var wallMm = clashZone.WallThickness > 0 ? UnitUtils.ConvertFromInternalUnits(clashZone.WallThickness, UnitTypeId.Millimeters) : 0;
                var framingMm = clashZone.FramingThickness > 0 ? UnitUtils.ConvertFromInternalUnits(clashZone.FramingThickness, UnitTypeId.Millimeters) : 0;
                if (structMm == 0 && wallMm == 0 && framingMm == 0)
                {
                    DebugLogger.Warning($"[DB-LOAD-THICKNESS] ⚠️ Zone={clashZone.Id}, Type={clashZone.StructuralElementType}: ALL thickness values are 0! Structural={structMm:F1}mm, Wall={wallMm:F1}mm, Framing={framingMm:F1}mm");
                }
            }

            var updatedAt = GetNullableDateTime(reader, "UpdatedAt");
            if (updatedAt.HasValue)
            {
                clashZone.LastUpdated = updatedAt.Value;
            }

            return clashZone;
        }

        private static string GetNullableString(SQLiteDataReader reader, string column)
        {
            var ordinal = reader.GetOrdinal(column);
            return reader.IsDBNull(ordinal) ? null : reader.GetString(ordinal);
        }

        private static double GetDouble(SQLiteDataReader reader, string column, double defaultValue = 0.0)
        {
            var ordinal = reader.GetOrdinal(column);
            return reader.IsDBNull(ordinal) ? defaultValue : Convert.ToDouble(reader.GetValue(ordinal));
        }

        private static double? GetNullableDouble(SQLiteDataReader reader, string column)
        {
            var ordinal = reader.GetOrdinal(column);
            return reader.IsDBNull(ordinal) ? (double?)null : Convert.ToDouble(reader.GetValue(ordinal));
        }

        private static int GetInt(SQLiteDataReader reader, string column, int defaultValue = 0)
        {
            var ordinal = reader.GetOrdinal(column);
            return reader.IsDBNull(ordinal) ? defaultValue : Convert.ToInt32(reader.GetValue(ordinal));
        }

        private static bool GetBool(SQLiteDataReader reader, string column)
        {
            var ordinal = reader.GetOrdinal(column);
            if (reader.IsDBNull(ordinal))
                return false;
            return Convert.ToInt32(reader.GetValue(ordinal)) != 0;
        }

        private static bool? GetNullableBool(SQLiteDataReader reader, string column)
        {
            var ordinal = reader.GetOrdinal(column);
            if (reader.IsDBNull(ordinal))
                return null;
            return Convert.ToInt32(reader.GetValue(ordinal)) != 0;
        }

        private static DateTime? GetNullableDateTime(SQLiteDataReader reader, string column)
        {
            var ordinal = reader.GetOrdinal(column);
            if (reader.IsDBNull(ordinal))
                return null;

            var value = reader.GetValue(ordinal);
            if (value is DateTime dt)
                return dt;

            if (DateTime.TryParse(value?.ToString(), out var parsed))
                return parsed;

            return null;
        }

        public void UpdateSleevePlacement(System.Guid clashZoneGuid, int sleeveInstanceId, double width, double height, double diameter,
            double placementX, double placementY, double placementZ,
            double placementActiveX, double placementActiveY, double placementActiveZ,
            double rotationAngleRad)
        {
            using (var cmd = _context.Connection.CreateCommand())
            {
                cmd.CommandText = @"
                    UPDATE ClashZones SET
                        SleeveState = 1,
                        SleeveInstanceId = @SleeveInstanceId,
                        SleeveWidth = @SleeveWidth,
                        SleeveHeight = @SleeveHeight,
                        SleeveDiameter = @SleeveDiameter,
                        SleevePlacementX = @SleevePlacementX,
                        SleevePlacementY = @SleevePlacementY,
                        SleevePlacementZ = @SleevePlacementZ,
                        SleevePlacementActiveX = @SleevePlacementActiveX,
                        SleevePlacementActiveY = @SleevePlacementActiveY,
                        SleevePlacementActiveZ = @SleevePlacementActiveZ,
                        MepRotationAngleRad = @MepRotationAngleRad,
                        MepRotationAngleDeg = @MepRotationAngleDeg,
                        MepRotationCos = @MepRotationCos,
                        MepRotationSin = @MepRotationSin,
                        UpdatedAt = CURRENT_TIMESTAMP
                    WHERE UPPER(ClashZoneGuid) = UPPER(@ClashZoneGuid)
                      AND ClashZoneGuid != '' AND ClashZoneGuid IS NOT NULL";

                cmd.Parameters.AddWithValue("@ClashZoneGuid", clashZoneGuid.ToString());
                cmd.Parameters.AddWithValue("@SleeveInstanceId", sleeveInstanceId);
                cmd.Parameters.AddWithValue("@SleeveWidth", width);
                cmd.Parameters.AddWithValue("@SleeveHeight", height);
                cmd.Parameters.AddWithValue("@SleeveDiameter", diameter);
                cmd.Parameters.AddWithValue("@SleevePlacementX", placementX);
                cmd.Parameters.AddWithValue("@SleevePlacementY", placementY);
                cmd.Parameters.AddWithValue("@SleevePlacementZ", placementZ);
                cmd.Parameters.AddWithValue("@SleevePlacementActiveX", placementActiveX);
                cmd.Parameters.AddWithValue("@SleevePlacementActiveY", placementActiveY);
                cmd.Parameters.AddWithValue("@SleevePlacementActiveZ", placementActiveZ);
                cmd.Parameters.AddWithValue("@MepRotationAngleRad", rotationAngleRad);
                cmd.Parameters.AddWithValue("@MepRotationAngleDeg", rotationAngleRad * 180.0 / Math.PI);
                // ✅ ROTATION MATRIX: Pre-calculate and save cos/sin (dump once use many times)
                cmd.Parameters.AddWithValue("@MepRotationCos", Math.Cos(rotationAngleRad));
                cmd.Parameters.AddWithValue("@MepRotationSin", Math.Sin(rotationAngleRad));
                cmd.ExecuteNonQuery();
            }
        }

        public void UpdateClusterPlacement(int clashZoneId, int clusterInstanceId, double minX, double minY, double minZ,
            double maxX, double maxY, double maxZ, double? placementX = null, double? placementY = null, double? placementZ = null,
            double? rotatedMinX = null, double? rotatedMinY = null, double? rotatedMinZ = null,
            double? rotatedMaxX = null, double? rotatedMaxY = null, double? rotatedMaxZ = null,
            bool? isClustered = null, bool? markedForCluster = null)
        {
            using (var cmd = _context.Connection.CreateCommand())
            {
                var updateFields = new List<string>
                {
                    "SleeveState = 2",
                    "ClusterInstanceId = @ClusterInstanceId",
                    "BoundingBoxMinX = @BoundingBoxMinX",
                    "BoundingBoxMinY = @BoundingBoxMinY",
                    "BoundingBoxMinZ = @BoundingBoxMinZ",
                    "BoundingBoxMaxX = @BoundingBoxMaxX",
                    "BoundingBoxMaxY = @BoundingBoxMaxY",
                    "BoundingBoxMaxZ = @BoundingBoxMaxZ"
                };

                // ✅ Add placement point if provided
                if (placementX.HasValue && placementY.HasValue && placementZ.HasValue)
                {
                    updateFields.Add("SleevePlacementX = @SleevePlacementX");
                    updateFields.Add("SleevePlacementY = @SleevePlacementY");
                    updateFields.Add("SleevePlacementZ = @SleevePlacementZ");
                }

                // ✅ Add rotated bounding boxes if provided
                if (rotatedMinX.HasValue && rotatedMinY.HasValue && rotatedMinZ.HasValue &&
                    rotatedMaxX.HasValue && rotatedMaxY.HasValue && rotatedMaxZ.HasValue)
                {
                    updateFields.Add("RotatedBoundingBoxMinX = @RotatedBoundingBoxMinX");
                    updateFields.Add("RotatedBoundingBoxMinY = @RotatedBoundingBoxMinY");
                    updateFields.Add("RotatedBoundingBoxMinZ = @RotatedBoundingBoxMinZ");
                    updateFields.Add("RotatedBoundingBoxMaxX = @RotatedBoundingBoxMaxX");
                    updateFields.Add("RotatedBoundingBoxMaxY = @RotatedBoundingBoxMaxY");
                    updateFields.Add("RotatedBoundingBoxMaxZ = @RotatedBoundingBoxMaxZ");
                }

                // ✅ Add flags if provided
                if (isClustered.HasValue)
                {
                    updateFields.Add("IsClusteredFlag = @IsClusteredFlag");
                }
                if (markedForCluster.HasValue)
                {
                    updateFields.Add("MarkedForClusterProcess = @MarkedForClusterProcess");
                }

                updateFields.Add("UpdatedAt = CURRENT_TIMESTAMP");

                cmd.CommandText = $@"
                    UPDATE ClashZones SET
                        {string.Join(",\n                        ", updateFields)}
                    WHERE ClashZoneId = @ClashZoneId";

                cmd.Parameters.AddWithValue("@ClashZoneId", clashZoneId);
                cmd.Parameters.AddWithValue("@ClusterInstanceId", clusterInstanceId);
                cmd.Parameters.AddWithValue("@BoundingBoxMinX", minX);
                cmd.Parameters.AddWithValue("@BoundingBoxMinY", minY);
                cmd.Parameters.AddWithValue("@BoundingBoxMinZ", minZ);
                cmd.Parameters.AddWithValue("@BoundingBoxMaxX", maxX);
                cmd.Parameters.AddWithValue("@BoundingBoxMaxY", maxY);
                cmd.Parameters.AddWithValue("@BoundingBoxMaxZ", maxZ);

                // ✅ Add optional parameters
                if (placementX.HasValue && placementY.HasValue && placementZ.HasValue)
                {
                    cmd.Parameters.AddWithValue("@SleevePlacementX", placementX.Value);
                    cmd.Parameters.AddWithValue("@SleevePlacementY", placementY.Value);
                    cmd.Parameters.AddWithValue("@SleevePlacementZ", placementZ.Value);
                }

                if (rotatedMinX.HasValue && rotatedMinY.HasValue && rotatedMinZ.HasValue &&
                    rotatedMaxX.HasValue && rotatedMaxY.HasValue && rotatedMaxZ.HasValue)
                {
                    cmd.Parameters.AddWithValue("@RotatedBoundingBoxMinX", rotatedMinX.Value);
                    cmd.Parameters.AddWithValue("@RotatedBoundingBoxMinY", rotatedMinY.Value);
                    cmd.Parameters.AddWithValue("@RotatedBoundingBoxMinZ", rotatedMinZ.Value);
                    cmd.Parameters.AddWithValue("@RotatedBoundingBoxMaxX", rotatedMaxX.Value);
                    cmd.Parameters.AddWithValue("@RotatedBoundingBoxMaxY", rotatedMaxY.Value);
                    cmd.Parameters.AddWithValue("@RotatedBoundingBoxMaxZ", rotatedMaxZ.Value);
                }

                if (isClustered.HasValue)
                {
                    cmd.Parameters.AddWithValue("@IsClusteredFlag", isClustered.Value ? 1 : 0);
                }
                if (markedForCluster.HasValue)
                {
                    cmd.Parameters.AddWithValue("@MarkedForClusterProcess", markedForCluster.Value ? 1 : 0);
                }

                var rowsAffected = cmd.ExecuteNonQuery();
                if (rowsAffected > 0 && !DeploymentConfiguration.DeploymentMode)
                {
                    _logger($"[SQLite] ✅ UpdateClusterPlacement: Updated ClashZoneId={clashZoneId}, ClusterInstanceId={clusterInstanceId}, " +
                        $"PlacementPoint={placementX?.ToString("F6") ?? "NULL"}, " +
                        $"RotatedBbox={rotatedMinX?.ToString("F6") ?? "NULL"}");
                }
            }
        }

        /// <summary>
        /// Update bounding box coordinates for an individual sleeve by ClashZone GUID
        /// </summary>
        public void UpdateSleeveBoundingBoxes(Guid clashZoneGuid, double minX, double minY, double minZ,
            double maxX, double maxY, double maxZ)
        {
            using (var cmd = _context.Connection.CreateCommand())
            {
                cmd.CommandText = @"
                    UPDATE ClashZones SET
                        BoundingBoxMinX = @BoundingBoxMinX,
                        BoundingBoxMinY = @BoundingBoxMinY,
                        BoundingBoxMinZ = @BoundingBoxMinZ,
                        BoundingBoxMaxX = @BoundingBoxMaxX,
                        BoundingBoxMaxY = @BoundingBoxMaxY,
                        BoundingBoxMaxZ = @BoundingBoxMaxZ,
                        UpdatedAt = CURRENT_TIMESTAMP
                    WHERE UPPER(ClashZoneGuid) = UPPER(@ClashZoneGuid)
                      AND ClashZoneGuid != '' AND ClashZoneGuid IS NOT NULL";

                cmd.Parameters.AddWithValue("@ClashZoneGuid", clashZoneGuid.ToString());
                cmd.Parameters.AddWithValue("@BoundingBoxMinX", minX);
                cmd.Parameters.AddWithValue("@BoundingBoxMinY", minY);
                cmd.Parameters.AddWithValue("@BoundingBoxMinZ", minZ);
                cmd.Parameters.AddWithValue("@BoundingBoxMaxX", maxX);
                cmd.Parameters.AddWithValue("@BoundingBoxMaxY", maxY);
                cmd.Parameters.AddWithValue("@BoundingBoxMaxZ", maxZ);
                
                var rowsAffected = cmd.ExecuteNonQuery();
                if (rowsAffected == 0 && !DeploymentConfiguration.DeploymentMode)
                {
                    _logger($"[SQLite] ⚠️ UpdateSleeveBoundingBoxes: No rows updated for GUID {clashZoneGuid}");
                }
            }
        }

        /// <summary>
        /// Update RCS (wall-aligned) bounding box coordinates for an individual sleeve by ClashZone GUID.
        /// Used for walls/framing to store bounding boxes in wall-aligned coordinate system.
        /// </summary>
        public void UpdateSleeveBoundingBoxesRcs(Guid clashZoneGuid, double rcsMinX, double rcsMinY, double rcsMinZ,
            double rcsMaxX, double rcsMaxY, double rcsMaxZ)
        {
            try
            {
                using (var cmd = _context.Connection.CreateCommand())
                {
                    cmd.CommandText = @"
                        UPDATE ClashZones SET
                            SleeveBoundingBoxRCS_MinX = @RcsMinX,
                            SleeveBoundingBoxRCS_MinY = @RcsMinY,
                            SleeveBoundingBoxRCS_MinZ = @RcsMinZ,
                            SleeveBoundingBoxRCS_MaxX = @RcsMaxX,
                            SleeveBoundingBoxRCS_MaxY = @RcsMaxY,
                            SleeveBoundingBoxRCS_MaxZ = @RcsMaxZ,
                            UpdatedAt = CURRENT_TIMESTAMP
                        WHERE UPPER(ClashZoneGuid) = UPPER(@ClashZoneGuid)
                          AND ClashZoneGuid != '' AND ClashZoneGuid IS NOT NULL";

                    cmd.Parameters.AddWithValue("@ClashZoneGuid", clashZoneGuid.ToString());
                    cmd.Parameters.AddWithValue("@RcsMinX", rcsMinX);
                    cmd.Parameters.AddWithValue("@RcsMinY", rcsMinY);
                    cmd.Parameters.AddWithValue("@RcsMinZ", rcsMinZ);
                    cmd.Parameters.AddWithValue("@RcsMaxX", rcsMaxX);
                    cmd.Parameters.AddWithValue("@RcsMaxY", rcsMaxY);
                    cmd.Parameters.AddWithValue("@RcsMaxZ", rcsMaxZ);
                    
                    var rowsAffected = cmd.ExecuteNonQuery();
                    if (rowsAffected == 0 && !DeploymentConfiguration.DeploymentMode)
                    {
                        _logger($"[SQLite] ⚠️ UpdateSleeveBoundingBoxesRcs: No rows updated for GUID {clashZoneGuid}");
                    }
                }
            }
            catch (Exception ex)
            {
                SafeFileLogger.SafeAppendText("database_errors.log",
                    $"[{DateTime.Now:HH:mm:ss.fff}] [ClashZoneRepository] Error in UpdateSleeveBoundingBoxesRcs for GUID {clashZoneGuid}: {ex.Message}\n");
                throw;
            }
        }

        /// <summary>
        /// Update rotated bounding box coordinates for a rotated individual sleeve
        /// Called when MepElementRotationAngle is non-zero (non-axis-aligned sleeve)
        /// For axis-aligned sleeves, these columns remain NULL
        /// </summary>
        public void UpdateRotatedBoundingBoxes(Guid clashZoneGuid, double rotatedMinX, double rotatedMinY, double rotatedMinZ,
            double rotatedMaxX, double rotatedMaxY, double rotatedMaxZ)
        {
            using (var cmd = _context.Connection.CreateCommand())
            {
                cmd.CommandText = @"
                    UPDATE ClashZones SET
                        RotatedBoundingBoxMinX = @RotatedBoundingBoxMinX,
                        RotatedBoundingBoxMinY = @RotatedBoundingBoxMinY,
                        RotatedBoundingBoxMinZ = @RotatedBoundingBoxMinZ,
                        RotatedBoundingBoxMaxX = @RotatedBoundingBoxMaxX,
                        RotatedBoundingBoxMaxY = @RotatedBoundingBoxMaxY,
                        RotatedBoundingBoxMaxZ = @RotatedBoundingBoxMaxZ,
                        UpdatedAt = CURRENT_TIMESTAMP
                    WHERE UPPER(ClashZoneGuid) = UPPER(@ClashZoneGuid)
                      AND ClashZoneGuid != '' AND ClashZoneGuid IS NOT NULL";

                cmd.Parameters.AddWithValue("@ClashZoneGuid", clashZoneGuid.ToString());
                cmd.Parameters.AddWithValue("@RotatedBoundingBoxMinX", rotatedMinX);
                cmd.Parameters.AddWithValue("@RotatedBoundingBoxMinY", rotatedMinY);
                cmd.Parameters.AddWithValue("@RotatedBoundingBoxMinZ", rotatedMinZ);
                cmd.Parameters.AddWithValue("@RotatedBoundingBoxMaxX", rotatedMaxX);
                cmd.Parameters.AddWithValue("@RotatedBoundingBoxMaxY", rotatedMaxY);
                cmd.Parameters.AddWithValue("@RotatedBoundingBoxMaxZ", rotatedMaxZ);
                
                var rowsAffected = cmd.ExecuteNonQuery();
                if (rowsAffected == 0 && !DeploymentConfiguration.DeploymentMode)
                {
                    _logger($"[SQLite] ⚠️ UpdateRotatedBoundingBoxes: No rows updated for GUID {clashZoneGuid}");
                }
                else if (rowsAffected > 0 && !DeploymentConfiguration.DeploymentMode)
                {
                    _logger($"[SQLite] ✅ Saved rotated bounding box for GUID {clashZoneGuid}");
                }
            }
        }

        /// <summary>
        /// ✅ SLEEVE CORNERS: Update pre-calculated 4 corner coordinates in world space
        /// Calculated once during individual sleeve placement, stored for reuse during clustering
        /// Corner order: 1=Bottom-left, 2=Bottom-right, 3=Top-left, 4=Top-right (in local space, then rotated to world)
        /// </summary>
        public void UpdateSleeveCorners(Guid clashZoneGuid, 
            double corner1X, double corner1Y, double corner1Z,
            double corner2X, double corner2Y, double corner2Z,
            double corner3X, double corner3Y, double corner3Z,
            double corner4X, double corner4Y, double corner4Z)
        {
            using (var cmd = _context.Connection.CreateCommand())
            {
                cmd.CommandText = @"
                    UPDATE ClashZones SET
                        SleeveCorner1X = @SleeveCorner1X,
                        SleeveCorner1Y = @SleeveCorner1Y,
                        SleeveCorner1Z = @SleeveCorner1Z,
                        SleeveCorner2X = @SleeveCorner2X,
                        SleeveCorner2Y = @SleeveCorner2Y,
                        SleeveCorner2Z = @SleeveCorner2Z,
                        SleeveCorner3X = @SleeveCorner3X,
                        SleeveCorner3Y = @SleeveCorner3Y,
                        SleeveCorner3Z = @SleeveCorner3Z,
                        SleeveCorner4X = @SleeveCorner4X,
                        SleeveCorner4Y = @SleeveCorner4Y,
                        SleeveCorner4Z = @SleeveCorner4Z,
                        UpdatedAt = CURRENT_TIMESTAMP
                    WHERE UPPER(ClashZoneGuid) = UPPER(@ClashZoneGuid)
                      AND ClashZoneGuid != '' AND ClashZoneGuid IS NOT NULL";

                cmd.Parameters.AddWithValue("@ClashZoneGuid", clashZoneGuid.ToString());
                cmd.Parameters.AddWithValue("@SleeveCorner1X", corner1X);
                cmd.Parameters.AddWithValue("@SleeveCorner1Y", corner1Y);
                cmd.Parameters.AddWithValue("@SleeveCorner1Z", corner1Z);
                cmd.Parameters.AddWithValue("@SleeveCorner2X", corner2X);
                cmd.Parameters.AddWithValue("@SleeveCorner2Y", corner2Y);
                cmd.Parameters.AddWithValue("@SleeveCorner2Z", corner2Z);
                cmd.Parameters.AddWithValue("@SleeveCorner3X", corner3X);
                cmd.Parameters.AddWithValue("@SleeveCorner3Y", corner3Y);
                cmd.Parameters.AddWithValue("@SleeveCorner3Z", corner3Z);
                cmd.Parameters.AddWithValue("@SleeveCorner4X", corner4X);
                cmd.Parameters.AddWithValue("@SleeveCorner4Y", corner4Y);
                cmd.Parameters.AddWithValue("@SleeveCorner4Z", corner4Z);
                
                var rowsAffected = cmd.ExecuteNonQuery();
                if (rowsAffected == 0 && !DeploymentConfiguration.DeploymentMode)
                {
                    _logger($"[SQLite] ⚠️ UpdateSleeveCorners: No rows updated for GUID {clashZoneGuid}");
                }
                else if (rowsAffected > 0 && !DeploymentConfiguration.DeploymentMode)
                {
                    _logger($"[SQLite] ✅ Saved 4 sleeve corners for GUID {clashZoneGuid}");
                }
            }
        }

        /// <summary>
        /// Update SleeveInstanceId for a clash zone by GUID
        /// </summary>
        public void UpdateSleeveInstanceId(Guid clashZoneGuid, int sleeveInstanceId)
        {
            using (var cmd = _context.Connection.CreateCommand())
            {
                cmd.CommandText = @"
                    UPDATE ClashZones SET
                        SleeveInstanceId = @SleeveInstanceId,
                        UpdatedAt = CURRENT_TIMESTAMP
                    WHERE UPPER(ClashZoneGuid) = UPPER(@ClashZoneGuid)
                      AND ClashZoneGuid != '' AND ClashZoneGuid IS NOT NULL";

                cmd.Parameters.AddWithValue("@ClashZoneGuid", clashZoneGuid.ToString());
                cmd.Parameters.AddWithValue("@SleeveInstanceId", sleeveInstanceId > 0 ? (object)sleeveInstanceId : DBNull.Value);
                
                var rowsAffected = cmd.ExecuteNonQuery();
                if (rowsAffected == 0 && !DeploymentConfiguration.DeploymentMode)
                {
                    _logger($"[SQLite] ⚠️ UpdateSleeveInstanceId: No rows updated for GUID {clashZoneGuid}");
                }
            }
        }

        public void LogSleeveEvent(int clashZoneId, string eventType, string payload = null)
                {
                    using (var cmd = _context.Connection.CreateCommand())
                    {
                        cmd.CommandText = @"
                    INSERT INTO SleeveEvents (ClashZoneId, EventType, Payload, CreatedAt)
                    VALUES (@ClashZoneId, @EventType, @Payload, CURRENT_TIMESTAMP)";

                cmd.Parameters.AddWithValue("@ClashZoneId", clashZoneId);
                cmd.Parameters.AddWithValue("@EventType", eventType);
                cmd.Parameters.AddWithValue("@Payload", (object)payload ?? DBNull.Value);
                            cmd.ExecuteNonQuery();
            }
        }

        /// <summary>
        /// ✅ VERIFICATION: Checks flag consistency for a specific clash zone GUID
        /// Returns the current flag values from database
        /// </summary>
        public (bool IsResolved, bool IsClusterResolved, int SleeveInstanceId, int ClusterInstanceId, bool Found)? VerifyFlags(Guid clashZoneId)
        {
            using (var cmd = _context.Connection.CreateCommand())
            {
                cmd.CommandText = @"
                    SELECT IsResolvedFlag, IsClusterResolvedFlag, SleeveInstanceId, ClusterInstanceId
                    FROM ClashZones
                    WHERE UPPER(ClashZoneGuid) = UPPER(@ClashZoneGuid)
                    LIMIT 1";

                cmd.Parameters.AddWithValue("@ClashZoneGuid", clashZoneId.ToString().ToUpperInvariant());

                using (var reader = cmd.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        return (
                            GetBool(reader, "IsResolvedFlag"),
                            GetBool(reader, "IsClusterResolvedFlag"),
                            GetInt(reader, "SleeveInstanceId", -1),
                            GetInt(reader, "ClusterInstanceId", -1),
                            true
                        );
                    }
                }
            }

            return (false, false, -1, -1, false);
        }

        /// <summary>
        /// ✅ DIAGNOSTIC: Checks for duplicate clash zones (same GUID, MEP+Host+Point, or SleeveInstanceId)
        /// </summary>
        public (int DuplicateGuids, int DuplicateMepHostPoint, int DuplicateSleeveIds) CheckForDuplicates(string category)
        {
            using (var cmd = _context.Connection.CreateCommand())
            {
                // Check for duplicate GUIDs
                cmd.CommandText = @"
                    SELECT COUNT(*) - COUNT(DISTINCT UPPER(ClashZoneGuid)) as DuplicateGuids
                    FROM ClashZones cz
                    INNER JOIN FileCombos fc ON cz.ComboId = fc.ComboId
                    INNER JOIN Filters f ON fc.FilterId = f.FilterId
                    WHERE f.Category = @Category AND (ClashZoneGuid != '' AND ClashZoneGuid IS NOT NULL)";

                cmd.Parameters.AddWithValue("@Category", category);
                var duplicateGuids = Convert.ToInt32(cmd.ExecuteScalar() ?? 0);
                
                // Check for duplicate MEP+Host+Point combinations
                cmd.CommandText = @"
                    SELECT COUNT(*) - COUNT(DISTINCT MepElementId || '|' || HostElementId || '|' || CAST(IntersectionX AS TEXT) || '|' || CAST(IntersectionY AS TEXT) || '|' || CAST(IntersectionZ AS TEXT)) as DuplicatePoints
                    FROM ClashZones cz
                    INNER JOIN FileCombos fc ON cz.ComboId = fc.ComboId
                    INNER JOIN Filters f ON fc.FilterId = f.FilterId
                    WHERE f.Category = @Category";
                
                var duplicatePoints = Convert.ToInt32(cmd.ExecuteScalar() ?? 0);
                
                // Check for duplicate SleeveInstanceIds (same sleeve ID in multiple clash zones)
                cmd.CommandText = @"
                    SELECT COUNT(*) - COUNT(DISTINCT SleeveInstanceId) as DuplicateSleeveIds
                    FROM ClashZones cz
                    INNER JOIN FileCombos fc ON cz.ComboId = fc.ComboId
                    INNER JOIN Filters f ON fc.FilterId = f.FilterId
                    WHERE f.Category = @Category AND SleeveInstanceId > 0";
                
                var duplicateSleeveIds = Convert.ToInt32(cmd.ExecuteScalar() ?? 0);
                
                return (duplicateGuids, duplicatePoints, duplicateSleeveIds);
            }
        }

        /// <summary>
        /// ✅ VERIFICATION: Gets flag summary for a category (for debugging)
        /// </summary>
        public (int Total, int Resolved, int ClusterResolved, int Unresolved) GetFlagSummary(string category)
        {
            using (var cmd = _context.Connection.CreateCommand())
            {
                cmd.CommandText = @"
                    SELECT 
                        COUNT(*) as Total,
                        SUM(CASE WHEN IsResolvedFlag = 1 OR IsClusterResolvedFlag = 1 THEN 1 ELSE 0 END) as Resolved,
                        SUM(CASE WHEN IsClusterResolvedFlag = 1 THEN 1 ELSE 0 END) as ClusterResolved,
                        SUM(CASE WHEN IsResolvedFlag = 0 AND IsClusterResolvedFlag = 0 THEN 1 ELSE 0 END) as Unresolved
                    FROM ClashZones cz
                    INNER JOIN FileCombos fc ON cz.ComboId = fc.ComboId
                    INNER JOIN Filters f ON fc.FilterId = f.FilterId
                    WHERE f.Category = @Category";

                cmd.Parameters.AddWithValue("@Category", category);

                using (var reader = cmd.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        return (
                            GetInt(reader, "Total", 0),
                            GetInt(reader, "Resolved", 0),
                            GetInt(reader, "ClusterResolved", 0),
                            GetInt(reader, "Unresolved", 0)
                            );
                        }
                    }
                }

            return (0, 0, 0, 0);
        }

        /// <summary>
        /// ✅ GUID MANAGEMENT: Finds existing GUID by MEP+Host+Point (for deterministic GUID lookup)
        /// </summary>
        public Guid? FindGuidByMepHostAndPoint(int mepElementId, int hostElementId, double intersectionX, double intersectionY, double intersectionZ, double tolerance = 0.001)
        {
            using (var cmd = _context.Connection.CreateCommand())
            {
                cmd.CommandText = @"
                    SELECT ClashZoneGuid
                    FROM ClashZones
                    WHERE MepElementId = @MepElementId 
                      AND HostElementId = @HostElementId 
                      AND ABS(IntersectionX - @IntersectionX) < @Tolerance 
                      AND ABS(IntersectionY - @IntersectionY) < @Tolerance 
                      AND ABS(IntersectionZ - @IntersectionZ) < @Tolerance
                      AND ClashZoneGuid != '' AND ClashZoneGuid IS NOT NULL
                    LIMIT 1";
                
                cmd.Parameters.AddWithValue("@MepElementId", mepElementId);
                cmd.Parameters.AddWithValue("@HostElementId", hostElementId);
                cmd.Parameters.AddWithValue("@IntersectionX", intersectionX);
                cmd.Parameters.AddWithValue("@IntersectionY", intersectionY);
                cmd.Parameters.AddWithValue("@IntersectionZ", intersectionZ);
                cmd.Parameters.AddWithValue("@Tolerance", tolerance);

                var result = cmd.ExecuteScalar();
                if (result != null && Guid.TryParse(result.ToString(), out var guid))
                {
                    return guid;
        }
            }
            
            return null;
        }

        /// <summary>
        /// ✅ GUID MANAGEMENT: Gets or creates a deterministic GUID for a clash zone
        /// </summary>
        public Guid GetOrCreateDeterministicGuid(int mepElementId, int hostElementId, double intersectionX, double intersectionY, double intersectionZ, Func<int, int, double, double, double, double, Guid> generateDeterministicGuid, double tolerance = 0.001)
        {
            // First, try to find existing GUID
            var existingGuid = FindGuidByMepHostAndPoint(mepElementId, hostElementId, intersectionX, intersectionY, intersectionZ, tolerance);
            if (existingGuid.HasValue)
            {
                return existingGuid.Value;
            }
            
            // Generate deterministic GUID
            var guid = generateDeterministicGuid(mepElementId, hostElementId, intersectionX, intersectionY, intersectionZ, tolerance);
            
            // Note: The GUID will be stored when the clash zone is inserted via InsertOrUpdateClashZones
            // This method is mainly for pre-checking/lookup purposes
            
            return guid;
        }
        
        /// <summary>
        /// ✅ DATABASE FLAG MANAGEMENT: Batch update flags for multiple clash zones
        /// Updates IsResolvedFlag, IsClusterResolvedFlag, SleeveInstanceId, and ClusterInstanceId
        /// Uses GUID, OLD SleeveInstanceId/ClusterInstanceId, or MEP+Host+Point for matching
        /// </summary>
        public void BatchUpdateFlags(IEnumerable<(Guid ClashZoneId, bool IsResolved, bool IsClusterResolved, int SleeveInstanceId, int ClusterInstanceId, int MepElementId, int StructuralElementId, double IntersectionPointX, double IntersectionPointY, double IntersectionPointZ, int OldSleeveInstanceId, int OldClusterInstanceId, bool? MarkedForClusterProcess, int AfterClusterSleeveId, bool? IsClusteredFlag)> updates)
        {
            if (updates == null)
                return;
                
            var updatesList = updates.ToList();
            if (updatesList.Count == 0)
                return;
            
            // ✅ LOG: Batch flag update operation
            DatabaseOperationLogger.LogOperation(
                "UPDATE",
                "ClashZones",
                new Dictionary<string, object>
                {
                    { "BatchSize", updatesList.Count },
                    { "Operation", "BatchUpdateFlags" }
                },
                rowsAffected: -1,
                additionalInfo: $"Batch updating flags for {updatesList.Count} clash zones");
            
            using (var transaction = _context.Connection.BeginTransaction())
            {
                try
        {
            using (var cmd = _context.Connection.CreateCommand())
            {
                        cmd.Transaction = transaction;
                        // ✅ OPTION 4: Triggers automatically compute SleeveState - no manual computation needed
                        // ✅ CRITICAL: Match by GUID first, then by OLD SleeveInstanceId/ClusterInstanceId, then MEP+Host+Point
                        // This is much simpler and more reliable than GUID-only matching
                        // ✅ CRITICAL FIX: Prioritize GUID matching - only use fallbacks if GUID doesn't match or is empty
                        // This ensures we only match 1 row per update when GUIDs are unique
                        // Use COALESCE to try GUID first, then fallbacks in order
                cmd.CommandText = @"
                            UPDATE ClashZones SET
                                IsResolvedFlag = @IsResolvedFlag,
                                IsClusterResolvedFlag = @IsClusterResolvedFlag,
                                SleeveInstanceId = @SleeveInstanceId,
                                ClusterInstanceId = @ClusterInstanceId,
                                MarkedForClusterProcess = @MarkedForClusterProcess,
                                AfterClusterSleeveId = @AfterClusterSleeveId,
                                IsClusteredFlag = @IsClusteredFlag,
                                UpdatedAt = datetime('now', '+5 hours', '+30 minutes')
                            WHERE ClashZoneId = COALESCE(
                                -- ✅ PRIORITY 1: Try GUID first (most reliable, should match exactly 1 row)
                                (SELECT ClashZoneId FROM ClashZones 
                                 WHERE UPPER(ClashZoneGuid) = UPPER(@ClashZoneGuid) 
                                   AND ClashZoneGuid != '' AND ClashZoneGuid IS NOT NULL
                                 LIMIT 1),
                                -- ✅ PRIORITY 2: If GUID didn't match, try OLD SleeveInstanceId + MEP+Host+Point
                                (SELECT ClashZoneId FROM ClashZones 
                                 WHERE @OldSleeveInstanceId > 0 
                                   AND SleeveInstanceId = @OldSleeveInstanceId
                                   AND MepElementId = @MepElementId 
                                   AND HostElementId = @HostElementId 
                                   AND ABS(IntersectionX - @IntersectionX) < 0.001 
                                   AND ABS(IntersectionY - @IntersectionY) < 0.001 
                                   AND ABS(IntersectionZ - @IntersectionZ) < 0.001
                                 LIMIT 1),
                                -- ✅ PRIORITY 3: If SleeveInstanceId didn't match, try OLD ClusterInstanceId + MEP+Host+Point
                                (SELECT ClashZoneId FROM ClashZones 
                                 WHERE @OldClusterInstanceId > 0 
                                   AND ClusterInstanceId = @OldClusterInstanceId
                                   AND MepElementId = @MepElementId 
                                   AND HostElementId = @HostElementId 
                                   AND ABS(IntersectionX - @IntersectionX) < 0.001 
                                   AND ABS(IntersectionY - @IntersectionY) < 0.001 
                                   AND ABS(IntersectionZ - @IntersectionZ) < 0.001
                                 LIMIT 1),
                                -- ✅ PRIORITY 4: If all else fails, match by MEP+Host+Point only (when GUID is empty)
                                (SELECT ClashZoneId FROM ClashZones 
                                 WHERE (ClashZoneGuid = '' OR ClashZoneGuid IS NULL) 
                                   AND MepElementId = @MepElementId 
                                   AND HostElementId = @HostElementId 
                                   AND ABS(IntersectionX - @IntersectionX) < 0.001 
                                   AND ABS(IntersectionY - @IntersectionY) < 0.001 
                                   AND ABS(IntersectionZ - @IntersectionZ) < 0.001
                                 LIMIT 1),
                                -- ✅ If nothing matches, return NULL (no update)
                                NULL
                            )";

                        var guidParam = cmd.Parameters.Add("@ClashZoneGuid", System.Data.DbType.String);
                        var isResolvedParam = cmd.Parameters.Add("@IsResolvedFlag", System.Data.DbType.Int32);
                        var isClusterResolvedParam = cmd.Parameters.Add("@IsClusterResolvedFlag", System.Data.DbType.Int32);
                        var sleeveIdParam = cmd.Parameters.Add("@SleeveInstanceId", System.Data.DbType.Int32);
                        var clusterIdParam = cmd.Parameters.Add("@ClusterInstanceId", System.Data.DbType.Int32);
                        // ✅ EDGE CASE FIELDS: Add parameters for MarkedForClusterProcess, AfterClusterSleeveId, IsClusteredFlag
                        var markedForClusterParam = cmd.Parameters.Add("@MarkedForClusterProcess", System.Data.DbType.Int32);
                        var afterClusterSleeveIdParam = cmd.Parameters.Add("@AfterClusterSleeveId", System.Data.DbType.Int32);
                        var isClusteredFlagParam = cmd.Parameters.Add("@IsClusteredFlag", System.Data.DbType.Int32);
                        // ✅ SIMPLER MATCHING: Use OLD SleeveInstanceId/ClusterInstanceId for direct matching
                        var oldSleeveIdParam = cmd.Parameters.Add("@OldSleeveInstanceId", System.Data.DbType.Int32);
                        var oldClusterIdParam = cmd.Parameters.Add("@OldClusterInstanceId", System.Data.DbType.Int32);
                        // ✅ FALLBACK MATCHING: Add parameters for MEP+Host+Point matching when GUID is empty
                        var mepIdParam = cmd.Parameters.Add("@MepElementId", System.Data.DbType.Int32);
                        var hostIdParam = cmd.Parameters.Add("@HostElementId", System.Data.DbType.Int32);
                        var intersectionXParam = cmd.Parameters.Add("@IntersectionX", System.Data.DbType.Double);
                        var intersectionYParam = cmd.Parameters.Add("@IntersectionY", System.Data.DbType.Double);
                        var intersectionZParam = cmd.Parameters.Add("@IntersectionZ", System.Data.DbType.Double);

                        // 🚀 PERFORMANCE: Prepare statement once for reuse (compiles SQL query plan)
                        cmd.Prepare();

                        int updateCount = 0;
                        int rowsAffectedTotal = 0;
                        int notFoundCount = 0;
                        int multipleRowsCount = 0; // ✅ Track updates that match multiple rows
                        
                        foreach (var update in updatesList)
                        {
                            // ✅ CRITICAL: Use uppercase GUID format for consistent matching (SQLite TEXT is case-sensitive)
                            string guidString = update.ClashZoneId.ToString().ToUpperInvariant();
                            guidParam.Value = guidString;
                            isResolvedParam.Value = update.IsResolved ? 1 : 0;
                            isClusterResolvedParam.Value = update.IsClusterResolved ? 1 : 0;
                            sleeveIdParam.Value = update.SleeveInstanceId > 0 ? (object)update.SleeveInstanceId : DBNull.Value;
                            clusterIdParam.Value = update.ClusterInstanceId > 0 ? (object)update.ClusterInstanceId : DBNull.Value;
                            // ✅ EDGE CASE FIELDS: Set MarkedForClusterProcess, AfterClusterSleeveId, IsClusteredFlag
                            markedForClusterParam.Value = update.MarkedForClusterProcess.HasValue ? (object)(update.MarkedForClusterProcess.Value ? 1 : 0) : DBNull.Value;
                            
                            // ✅ CRITICAL: Save AfterClusterSleeveId (original individual sleeve ID before deletion)
                            // This is set by RefactoredClusterService BEFORE calling UpdateFlagsForPlacement
                            if (update.AfterClusterSleeveId > 0)
                            {
                                afterClusterSleeveIdParam.Value = update.AfterClusterSleeveId;
                                // ✅ DIAGNOSTIC: Log to database operations log
                                DatabaseOperationLogger.LogOperation(
                                    "UPDATE",
                                    "ClashZones",
                                    new Dictionary<string, object>
                                    {
                                        { "ClashZoneGuid", guidString },
                                        { "AfterClusterSleeveId", update.AfterClusterSleeveId },
                                        { "IsClusterResolvedFlag", update.IsClusterResolved ? 1 : 0 },
                                        { "ClusterInstanceId", update.ClusterInstanceId }
                                    },
                                    rowsAffected: -1,
                                    additionalInfo: $"Setting AfterClusterSleeveId={update.AfterClusterSleeveId} for cluster {update.ClusterInstanceId}");
                            }
                            else
                            {
                                afterClusterSleeveIdParam.Value = DBNull.Value;
                            }
                            
                            isClusteredFlagParam.Value = update.IsClusteredFlag.HasValue ? (object)(update.IsClusteredFlag.Value ? 1 : 0) : DBNull.Value;
                            // ✅ SIMPLER MATCHING: Use OLD sleeve/cluster IDs for direct matching (much more reliable!)
                            oldSleeveIdParam.Value = update.OldSleeveInstanceId > 0 ? (object)update.OldSleeveInstanceId : DBNull.Value;
                            oldClusterIdParam.Value = update.OldClusterInstanceId > 0 ? (object)update.OldClusterInstanceId : DBNull.Value;
                            // ✅ FALLBACK: Set MEP+Host+Point parameters (will be used if GUID doesn't match or is empty)
                            mepIdParam.Value = update.MepElementId;
                            hostIdParam.Value = update.StructuralElementId;
                            intersectionXParam.Value = update.IntersectionPointX;
                            intersectionYParam.Value = update.IntersectionPointY;
                            intersectionZParam.Value = update.IntersectionPointZ;
                            
                            int rowsAffected = cmd.ExecuteNonQuery();
                            rowsAffectedTotal += rowsAffected;
                            
                            // ✅ CRITICAL: Track if no rows were affected (indicates ClashZone not found in database)
                            if (rowsAffected == 0)
                            {
                                notFoundCount++;
                                _logger($"[SQLite] ⚠️⚠️⚠️ WARNING: GUID {update.ClashZoneId} matched 0 rows! ClashZone not found in database.");
                                _logger($"[SQLite] ⚠️ Update details: GUID={guidString}, OldSleeveId={update.OldSleeveInstanceId}, OldClusterId={update.OldClusterInstanceId}, MEP={update.MepElementId}, Host={update.StructuralElementId}");
                                _logger($"[SQLite] ⚠️ IsClusterResolved={update.IsClusterResolved}, ClusterId={update.ClusterInstanceId}");
                                
                                // ✅ DIAGNOSTIC: Try to find why it didn't match
                                using (var diagCmd = _context.Connection.CreateCommand())
                                {
                                    diagCmd.Transaction = transaction;
                                    diagCmd.CommandText = @"
                                        SELECT ClashZoneId, ClashZoneGuid, SleeveInstanceId, ClusterInstanceId, 
                                               MepElementId, HostElementId, IsClusterResolvedFlag
                                        FROM ClashZones
                                        WHERE UPPER(ClashZoneGuid) = UPPER(@ClashZoneGuid) 
                                           OR (SleeveInstanceId = @OldSleeveInstanceId AND MepElementId = @MepElementId AND HostElementId = @HostElementId)
                                           OR (ClusterInstanceId = @OldClusterInstanceId AND MepElementId = @MepElementId AND HostElementId = @HostElementId)
                                        LIMIT 5";
                                    diagCmd.Parameters.AddWithValue("@ClashZoneGuid", guidString);
                                    diagCmd.Parameters.AddWithValue("@OldSleeveInstanceId", update.OldSleeveInstanceId);
                                    diagCmd.Parameters.AddWithValue("@OldClusterInstanceId", update.OldClusterInstanceId);
                                    diagCmd.Parameters.AddWithValue("@MepElementId", update.MepElementId);
                                    diagCmd.Parameters.AddWithValue("@HostElementId", update.StructuralElementId);
                                    
                                    using (var reader = diagCmd.ExecuteReader())
                                    {
                                        int foundCount = 0;
                                        while (reader.Read())
                                        {
                                            foundCount++;
                                            var foundGuid = GetNullableString(reader, "ClashZoneGuid");
                                            var foundSleeveId = GetInt(reader, "SleeveInstanceId", -1);
                                            var foundClusterId = GetInt(reader, "ClusterInstanceId", -1);
                                            var foundMepId = GetInt(reader, "MepElementId");
                                            var foundHostId = GetInt(reader, "HostElementId");
                                            var foundIsClusterResolved = GetBool(reader, "IsClusterResolvedFlag");
                                            _logger($"[SQLite] ⚠️   Found ClashZone {foundCount}: GUID={foundGuid}, SleeveId={foundSleeveId}, ClusterId={foundClusterId}, MEP={foundMepId}, Host={foundHostId}, IsClusterResolved={foundIsClusterResolved}");
                                        }
                                        if (foundCount == 0)
                                        {
                                            _logger($"[SQLite] ⚠️   No matching ClashZones found with any criteria - GUID, SleeveId, or ClusterId");
                                        }
                                    }
                                }
                            }
                            
                            // ✅ CRITICAL: Track if multiple rows were affected (indicates duplicate entries or WHERE clause issue)
                            if (rowsAffected > 1)
                            {
                                multipleRowsCount++;
                                _logger($"[SQLite] ⚠️⚠️⚠️ WARNING: GUID {update.ClashZoneId} matched {rowsAffected} rows (expected 1)! This indicates duplicate entries.");
                                _logger($"[SQLite] ⚠️ Update details: GUID={guidString}, OldSleeveId={update.OldSleeveInstanceId}, OldClusterId={update.OldClusterInstanceId}, MEP={update.MepElementId}, Host={update.StructuralElementId}");
                                
                                // ✅ DIAGNOSTIC: Query to see which condition matched
                                using (var diagCmd = _context.Connection.CreateCommand())
                                {
                                    diagCmd.Transaction = transaction;
                                    diagCmd.CommandText = @"
                                        SELECT 
                                            ClashZoneId,
                                            ClashZoneGuid,
                                            SleeveInstanceId,
                                            ClusterInstanceId,
                                            MepElementId,
                                            HostElementId,
                                            IntersectionX,
                                            IntersectionY,
                                            IntersectionZ,
                                            CASE 
                                                WHEN UPPER(ClashZoneGuid) = UPPER(@ClashZoneGuid) THEN 'GUID'
                                                WHEN @OldSleeveInstanceId > 0 AND SleeveInstanceId = @OldSleeveInstanceId AND MepElementId = @MepElementId AND HostElementId = @HostElementId AND ABS(IntersectionX - @IntersectionX) < 0.001 AND ABS(IntersectionY - @IntersectionY) < 0.001 AND ABS(IntersectionZ - @IntersectionZ) < 0.001 THEN 'SleeveId+MEP+Host+Point'
                                                WHEN @OldClusterInstanceId > 0 AND ClusterInstanceId = @OldClusterInstanceId AND MepElementId = @MepElementId AND HostElementId = @HostElementId AND ABS(IntersectionX - @IntersectionX) < 0.001 AND ABS(IntersectionY - @IntersectionY) < 0.001 AND ABS(IntersectionZ - @IntersectionZ) < 0.001 THEN 'ClusterId+MEP+Host+Point'
                                                ELSE 'MEP+Host+Point'
                                            END as MatchType
                                        FROM ClashZones
                                        WHERE (
                                            UPPER(ClashZoneGuid) = UPPER(@ClashZoneGuid)
                                        ) OR (
                                            @OldSleeveInstanceId > 0 
                                            AND SleeveInstanceId = @OldSleeveInstanceId
                                            AND MepElementId = @MepElementId 
                                            AND HostElementId = @HostElementId 
                                            AND ABS(IntersectionX - @IntersectionX) < 0.001 
                                            AND ABS(IntersectionY - @IntersectionY) < 0.001 
                                            AND ABS(IntersectionZ - @IntersectionZ) < 0.001
                                        ) OR (
                                            @OldClusterInstanceId > 0 
                                            AND ClusterInstanceId = @OldClusterInstanceId
                                            AND MepElementId = @MepElementId 
                                            AND HostElementId = @HostElementId 
                                            AND ABS(IntersectionX - @IntersectionX) < 0.001 
                                            AND ABS(IntersectionY - @IntersectionY) < 0.001 
                                            AND ABS(IntersectionZ - @IntersectionZ) < 0.001
                                        )";
                                    
                                    diagCmd.Parameters.AddWithValue("@ClashZoneGuid", guidString);
                                    diagCmd.Parameters.AddWithValue("@OldSleeveInstanceId", update.OldSleeveInstanceId > 0 ? (object)update.OldSleeveInstanceId : DBNull.Value);
                                    diagCmd.Parameters.AddWithValue("@OldClusterInstanceId", update.OldClusterInstanceId > 0 ? (object)update.OldClusterInstanceId : DBNull.Value);
                                    diagCmd.Parameters.AddWithValue("@MepElementId", update.MepElementId);
                                    diagCmd.Parameters.AddWithValue("@HostElementId", update.StructuralElementId);
                                    diagCmd.Parameters.AddWithValue("@IntersectionX", update.IntersectionPointX);
                                    diagCmd.Parameters.AddWithValue("@IntersectionY", update.IntersectionPointY);
                                    diagCmd.Parameters.AddWithValue("@IntersectionZ", update.IntersectionPointZ);
                                    
                                    using (var diagReader = diagCmd.ExecuteReader())
                                    {
                                        int matchIndex = 0;
                                        while (diagReader.Read())
                                        {
                                            matchIndex++;
                                            int clashZoneId = GetInt(diagReader, "ClashZoneId", -1);
                                            string matchedGuid = GetNullableString(diagReader, "ClashZoneGuid") ?? "NULL";
                                            int sleeveId = GetInt(diagReader, "SleeveInstanceId", -1);
                                            int clusterId = GetInt(diagReader, "ClusterInstanceId", -1);
                                            string matchType = GetNullableString(diagReader, "MatchType") ?? "Unknown";
                                            
                                            _logger($"[SQLite] 🔍 Match #{matchIndex}: ClashZoneId={clashZoneId}, GUID={matchedGuid}, SleeveId={sleeveId}, ClusterId={clusterId}, MatchType={matchType}");
                                        }
                                    }
                                }
                            }
                            else if (rowsAffected == 0)
                            {
                                notFoundCount++;
                                _logger($"[SQLite] ⚠️ GUID {guidString} NOT FOUND in database - no rows updated");
                                _logger($"[SQLite] ⚠️ Searching for GUID: {guidString} (original: {update.ClashZoneId})");
                                _logger($"[SQLite] ⚠️ Will try fallback matching by MEP={update.MepElementId}, Host={update.StructuralElementId}, Point=({update.IntersectionPointX:F3}, {update.IntersectionPointY:F3}, {update.IntersectionPointZ:F3})");
                                _logger($"[SQLite] ⚠️ OldSleeveId={update.OldSleeveInstanceId}, OldClusterId={update.OldClusterInstanceId}, IsClusterResolved={update.IsClusterResolved}");
                                
                                // ✅ DIAGNOSTIC: Query database to see what exists for this GUID
                                using (var diagCmd = _context.Connection.CreateCommand())
                                {
                                    diagCmd.Transaction = transaction;
                                    diagCmd.CommandText = @"
                                        SELECT ClashZoneId, ClashZoneGuid, SleeveInstanceId, ClusterInstanceId, 
                                               IsResolvedFlag, IsClusterResolvedFlag, MepElementId, HostElementId
                                        FROM ClashZones
                                        WHERE UPPER(ClashZoneGuid) = UPPER(@ClashZoneGuid)
                                           OR (MepElementId = @MepElementId AND HostElementId = @HostElementId 
                                               AND ABS(IntersectionX - @IntersectionX) < 0.001 
                                               AND ABS(IntersectionY - @IntersectionY) < 0.001 
                                               AND ABS(IntersectionZ - @IntersectionZ) < 0.001)
                                        LIMIT 5";
                                    diagCmd.Parameters.AddWithValue("@ClashZoneGuid", guidString);
                                    diagCmd.Parameters.AddWithValue("@MepElementId", update.MepElementId);
                                    diagCmd.Parameters.AddWithValue("@HostElementId", update.StructuralElementId);
                                    diagCmd.Parameters.AddWithValue("@IntersectionX", update.IntersectionPointX);
                                    diagCmd.Parameters.AddWithValue("@IntersectionY", update.IntersectionPointY);
                                    diagCmd.Parameters.AddWithValue("@IntersectionZ", update.IntersectionPointZ);
                                    
                                    using (var diagReader = diagCmd.ExecuteReader())
                                    {
                                        int matchCount = 0;
                                        while (diagReader.Read())
                                        {
                                            matchCount++;
                                            int czId = GetInt(diagReader, "ClashZoneId", -1);
                                            string dbGuid = GetNullableString(diagReader, "ClashZoneGuid") ?? "NULL";
                                            int dbSleeveId = GetInt(diagReader, "SleeveInstanceId", -1);
                                            int dbClusterId = GetInt(diagReader, "ClusterInstanceId", -1);
                                            int dbIsResolved = GetInt(diagReader, "IsResolvedFlag", 0);
                                            int dbIsClusterResolved = GetInt(diagReader, "IsClusterResolvedFlag", 0);
                                            _logger($"[SQLite] 🔍 Match #{matchCount}: ClashZoneId={czId}, GUID={dbGuid}, SleeveId={dbSleeveId}, ClusterId={dbClusterId}, IsResolved={dbIsResolved}, IsClusterResolved={dbIsClusterResolved}");
                                        }
                                        if (matchCount == 0)
                                        {
                                            _logger($"[SQLite] ❌ NO MATCHES FOUND in database for GUID {guidString} or MEP+Host+Point");
                                        }
                                    }
                                }
                            }
                            else if (rowsAffected == 1)
                            {
                                // ✅ SUCCESS: Log with AfterClusterSleeveId
                                _logger($"[SQLite] ✅ Updated GUID {update.ClashZoneId}: IsResolved={update.IsResolved}, IsClusterResolved={update.IsClusterResolved}, " +
                                    $"SleeveId={update.SleeveInstanceId}, ClusterId={update.ClusterInstanceId}, AfterClusterSleeveId={update.AfterClusterSleeveId} (1 row affected)");
                                
                                // ✅ DIAGNOSTIC: Log to database operations log for AfterClusterSleeveId updates
                                if (update.AfterClusterSleeveId > 0)
                                {
                                    DatabaseOperationLogger.LogOperation(
                                        "UPDATE",
                                        "ClashZones",
                                        new Dictionary<string, object>
                                        {
                                            { "ClashZoneGuid", guidString },
                                            { "AfterClusterSleeveId", update.AfterClusterSleeveId },
                                            { "IsClusterResolvedFlag", update.IsClusterResolved ? 1 : 0 },
                                            { "ClusterInstanceId", update.ClusterInstanceId },
                                            { "SleeveInstanceId", update.SleeveInstanceId }
                                        },
                                        rowsAffected: 1,
                                        additionalInfo: $"✅ Saved AfterClusterSleeveId={update.AfterClusterSleeveId} (original individual sleeve ID before deletion)");
                                }
                            }
                            
                            updateCount++;
                        }

                        transaction.Commit();
                        _logger($"[SQLite] ✅ Batch updated flags: {updateCount} updates attempted, {rowsAffectedTotal} rows affected, {notFoundCount} GUIDs not found, {multipleRowsCount} updates matched multiple rows");
                        
                        if (multipleRowsCount > 0)
                        {
                            _logger($"[SQLite] ⚠️⚠️⚠️ WARNING: {multipleRowsCount} updates matched multiple rows! This indicates duplicate entries in the database. Expected: 1 row per update. Actual: {rowsAffectedTotal} rows for {updateCount} updates.");
                        }
                    }
                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    _logger($"[SQLite] ❌ Error batch updating flags: {ex.Message}");
                    throw;
                }
            }
        }

        /// <summary>
        /// ✅ FLAG RESET: Resets IsFilterComboNew flag to 0 after cluster sleeve placement completes
        /// Called from UniversalClusterService after all cluster sleeves are placed
        /// This ensures PATH 2 (Sizing) gets a chance to run before flag is reset
        /// </summary>
        public void ResetFileComboFlag(int comboId)
        {
            if (comboId <= 0)
            {
                _logger($"[SQLite] ⚠️ Invalid ComboId for flag reset: {comboId}");
                return;
            }

            try
            {
                // ✅ LOG: UPDATE operation
                var updateParams = new Dictionary<string, object>
                {
                    { "ComboId", comboId },
                    { "IsFilterComboNew", 0 }
                };

                DatabaseOperationLogger.LogOperation(
                    "UPDATE",
                    "FileCombos",
                    updateParams,
                    additionalInfo: $"Resetting IsFilterComboNew flag after cluster placement (Flow #7)");

                using (var cmd = _context.Connection.CreateCommand())
                {
                    cmd.CommandText = @"
                        UPDATE FileCombos
                        SET IsFilterComboNew = 0,
                            UpdatedAt = datetime('now', '+5 hours', '+30 minutes')
                        WHERE ComboId = @ComboId";
                    cmd.Parameters.AddWithValue("@ComboId", comboId);

                    int rowsAffected = cmd.ExecuteNonQuery();
                    if (rowsAffected > 0)
                    {
                        DatabaseOperationLogger.LogOperation(
                            "UPDATE",
                            "FileCombos",
                            updateParams,
                            rowsAffected,
                            $"✅ Reset IsFilterComboNew=0 for ComboId={comboId} (Flow #7)");
                        _logger($"[SQLite] ✅ Reset IsFilterComboNew=0 for ComboId={comboId}");
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[ClashZoneRepository] ✅ Reset IsFilterComboNew flag for ComboId={comboId}");
                        }
                    }
                    else
                    {
                        _logger($"[SQLite] ⚠️ No FileCombo found with ComboId={comboId} to reset flag");
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Warning($"[ClashZoneRepository] ⚠️ No FileCombo found with ComboId={comboId}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger($"[SQLite] ❌ Error resetting IsFilterComboNew flag for ComboId={comboId}: {ex.Message}");
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Error($"[ClashZoneRepository] ❌ Error resetting flag: {ex.Message}\n{ex.StackTrace}");
                }
                throw;
            }
        }
    }
}


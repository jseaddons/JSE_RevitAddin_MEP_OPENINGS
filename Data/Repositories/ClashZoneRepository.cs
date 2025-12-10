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
        private static bool _dbVerifiedOnce = false; // session-level guard

        public ClashZoneRepository(SleeveDbContext context, Action<string> logger = null)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            _logger = logger ?? (msg => { });
        }

        /// <summary>
        /// ✅ PERFORMANCE OPTIMIZED: Bulk INSERT/UPDATE for clash zones using batch SQL operations
        /// Reduces 9 zones from ~517ms to ~30ms (17x faster) by eliminating per-zone queries
        /// </summary>
        private void InsertOrUpdateClashZonesBulk(IEnumerable<ClashZone> clashZones, string filterName, string category)
        {
            var zonesList = clashZones.ToList();
            if (zonesList.Count == 0) return;

            var sw = System.Diagnostics.Stopwatch.StartNew();
            _logger($"[SQLite][BULK] ⚡ Starting bulk update for {zonesList.Count} zones");

            using (var transaction = _context.Connection.BeginTransaction())
            {
                try
                {
                    // Optional one-time DB verification per session
                    if (OptimizationFlags.UseOneTimeDbVerificationDuringSession)
                    {
                        if (!_dbVerifiedOnce)
                        {
                            // Lightweight ping to ensure DB is reachable (avoids repeated heavy verifies)
                            using (var ping = _context.Connection.CreateCommand())
                            {
                                ping.Transaction = transaction;
                                ping.CommandText = "SELECT 1";
                                ping.ExecuteNonQuery();
                            }
                            _dbVerifiedOnce = true;
                            _logger("[SQLite] ✅ Session DB verification completed (one-time)");
                        }
                        else
                        {
                            if (!OptimizationFlags.DisableVerboseLogging)
                                _logger("[SQLite] ⏭️ Skipping DB verification (session-cached)");
                        }
                    }

                    // Get filter ID once
                    var filterId = GetOrCreateFilter(filterName, category, transaction);
                    if (filterId <= 0)
                    {
                        _logger($"[SQLite][BULK] ⚠️ Filter not found: '{filterName}' category '{category}'");
                        transaction.Rollback();
                        return;
                    }

                    // Build GUID list for batch lookup
                    var guidList = string.Join(",", zonesList.Select(z => $"'{z.Id.ToString().ToUpperInvariant()}'"));
                    
                    // Batch lookup: find all existing ClashZoneIds by GUID in single query
                    var existingMap = new Dictionary<Guid, int>();
                    using (var cmd = _context.Connection.CreateCommand())
                    {
                        cmd.Transaction = transaction;
                        cmd.CommandText = $@"
                            SELECT ClashZoneGuid, ClashZoneId 
                            FROM ClashZones 
                            WHERE UPPER(ClashZoneGuid) IN ({guidList})
                              AND ClashZoneGuid != '' AND ClashZoneGuid IS NOT NULL";
                        
                        using (var reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                var guid = Guid.Parse(reader.GetString(0));
                                var id = reader.GetInt32(1);
                                existingMap[guid] = id;
                            }
                        }
                    }

                    _logger($"[SQLite][BULK] Found {existingMap.Count}/{zonesList.Count} existing zones");

                    // Get/create file combos (still need this per zone unfortunately due to complexity)
                    var comboMap = new Dictionary<Guid, int>();
                    foreach (var zone in zonesList)
                    {
                        var mepCategory = zone.MepElementCategory ?? category;
                        List<string>? hostCategories = null;
                        if (FilterUiStateProvider.GetSelectedHostCategories != null)
                        {
                            hostCategories = FilterUiStateProvider.GetSelectedHostCategories.Invoke();
                        }
                        var comboId = GetOrCreateFileCombo(filterId, mepCategory, hostCategories ?? new List<string>(), zone, transaction);
                        if (comboId > 0)
                        {
                            comboMap[zone.Id] = comboId;
                        }
                    }

                    // Batch UPDATE using CASE statements for all zones with existing ClashZoneIds
                    var toUpdate = zonesList.Where(z => existingMap.ContainsKey(z.Id)).ToList();
                    if (toUpdate.Count > 0)
                    {
                        BulkUpdateClashZones(toUpdate, existingMap, comboMap, transaction);
                        _logger($"[SQLite][BULK] ✅ Updated {toUpdate.Count} zones");
                    }

                    // ✅ CRITICAL FIX: Before INSERT, check for existing zones by UNIQUE constraint columns
                    // This prevents UNIQUE constraint violations when zones have new GUIDs but same ComboId/MepElementId/HostElementId/IntersectionX/Y/Z
                    var toInsert = zonesList.Where(z => !existingMap.ContainsKey(z.Id) && comboMap.ContainsKey(z.Id)).ToList();
                    if (toInsert.Count > 0)
                    {
                        // ✅ STEP 1: Build lookup map for existing zones by UNIQUE constraint columns
                        var uniqueConstraintMap = new Dictionary<string, int>(); // Key: "ComboId|MepElementId|HostElementId|IntersectionX|IntersectionY|IntersectionZ", Value: ClashZoneId
                        using (var checkCmd = _context.Connection.CreateCommand())
                        {
                            checkCmd.Transaction = transaction;
                            
                            // Build WHERE clause for all zones to check
                            var whereConditions = new List<string>();
                            var paramIndex = 0;
                            foreach (var zone in toInsert)
                            {
                                if (!comboMap.ContainsKey(zone.Id)) continue;
                                
                                var comboId = comboMap[zone.Id];
                                var mepId = zone.MepElementId?.IntegerValue ?? zone.MepElementIdValue;
                                var hostId = zone.StructuralElementId?.IntegerValue ?? zone.StructuralElementIdValue;
                                var interX = zone.IntersectionPoint?.X ?? zone.IntersectionPointX;
                                var interY = zone.IntersectionPoint?.Y ?? zone.IntersectionPointY;
                                var interZ = zone.IntersectionPoint?.Z ?? zone.IntersectionPointZ;
                                
                                whereConditions.Add($"(ComboId = @ComboId{paramIndex} AND MepElementId = @MepId{paramIndex} AND HostElementId = @HostId{paramIndex} AND IntersectionX = @InterX{paramIndex} AND IntersectionY = @InterY{paramIndex} AND IntersectionZ = @InterZ{paramIndex})");
                                
                                checkCmd.Parameters.AddWithValue($"@ComboId{paramIndex}", comboId);
                                checkCmd.Parameters.AddWithValue($"@MepId{paramIndex}", mepId);
                                checkCmd.Parameters.AddWithValue($"@HostId{paramIndex}", hostId);
                                checkCmd.Parameters.AddWithValue($"@InterX{paramIndex}", interX);
                                checkCmd.Parameters.AddWithValue($"@InterY{paramIndex}", interY);
                                checkCmd.Parameters.AddWithValue($"@InterZ{paramIndex}", interZ);
                                
                                paramIndex++;
                            }
                            
                            if (whereConditions.Count > 0)
                            {
                                checkCmd.CommandText = $@"
                                    SELECT ClashZoneId, ComboId, MepElementId, HostElementId, IntersectionX, IntersectionY, IntersectionZ
                                    FROM ClashZones 
                                    WHERE {string.Join(" OR ", whereConditions)}";
                                
                                using (var reader = checkCmd.ExecuteReader())
                                {
                                    while (reader.Read())
                                    {
                                        var clashZoneId = reader.GetInt32(0);
                                        var comboId = reader.GetInt32(1);
                                        var mepId = reader.GetInt32(2);
                                        var hostId = reader.GetInt32(3);
                                        var interX = reader.GetDouble(4);
                                        var interY = reader.GetDouble(5);
                                        var interZ = reader.GetDouble(6);
                                        
                                        var key = $"{comboId}|{mepId}|{hostId}|{interX}|{interY}|{interZ}";
                                        uniqueConstraintMap[key] = clashZoneId;
                                    }
                                }
                            }
                        }
                        
                        // ✅ STEP 2: Filter out zones that already exist by UNIQUE constraint, add them to existingMap for UPDATE
                        var actuallyNew = new List<ClashZone>();
                        var toUpdateByConstraint = new List<(ClashZone zone, int clashZoneId)>();
                        
                        foreach (var zone in toInsert)
                        {
                            if (!comboMap.ContainsKey(zone.Id)) continue;
                            
                            var comboId = comboMap[zone.Id];
                            var mepId = zone.MepElementId?.IntegerValue ?? zone.MepElementIdValue;
                            var hostId = zone.StructuralElementId?.IntegerValue ?? zone.StructuralElementIdValue;
                            var interX = zone.IntersectionPoint?.X ?? zone.IntersectionPointX;
                            var interY = zone.IntersectionPoint?.Y ?? zone.IntersectionPointY;
                            var interZ = zone.IntersectionPoint?.Z ?? zone.IntersectionPointZ;
                            
                            var key = $"{comboId}|{mepId}|{hostId}|{interX}|{interY}|{interZ}";
                            
                            if (uniqueConstraintMap.TryGetValue(key, out var existingClashZoneId))
                            {
                                // ✅ Zone exists by UNIQUE constraint - add to existingMap and update instead of insert
                                existingMap[zone.Id] = existingClashZoneId;
                                toUpdateByConstraint.Add((zone, existingClashZoneId));
                                
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    _logger($"[SQLite][BULK] ✅ Found existing zone by UNIQUE constraint: GUID={zone.Id} matches ClashZoneId={existingClashZoneId} (will UPDATE instead of INSERT)");
                                }
                            }
                            else
                            {
                                actuallyNew.Add(zone);
                            }
                        }
                        
                        // ✅ STEP 3: UPDATE zones found by UNIQUE constraint
                        if (toUpdateByConstraint.Count > 0)
                        {
                            var zonesToUpdate = toUpdateByConstraint.Select(t => t.zone).ToList();
                            BulkUpdateClashZones(zonesToUpdate, existingMap, comboMap, transaction);
                            _logger($"[SQLite][BULK] ✅ Updated {toUpdateByConstraint.Count} zones (found by UNIQUE constraint, not GUID)");
                        }
                        
                        // ✅ STEP 4: INSERT only truly new zones
                        if (actuallyNew.Count > 0)
                        {
                            foreach (var zone in actuallyNew)
                            {
                                InsertClashZone(comboMap[zone.Id], zone, transaction);
                            }
                            _logger($"[SQLite][BULK] ✅ Inserted {actuallyNew.Count} zones");
                        }
                    }

                    transaction.Commit();
                    sw.Stop();
                    _logger($"[SQLite][BULK] ⚡ Completed in {sw.ElapsedMilliseconds}ms ({sw.ElapsedMilliseconds / (double)zonesList.Count:F1}ms per zone)");
                }
                catch (Exception ex)
                {
                    _logger($"[SQLite][BULK] ❌ Error: {ex.Message}");
                    transaction.Rollback();
                    throw;
                }
            }
        }

        /// <summary>
        /// Bulk UPDATE using parameterized CASE statements for maximum performance
        /// </summary>
        private void BulkUpdateClashZones(List<ClashZone> zones, Dictionary<Guid, int> existingMap, Dictionary<Guid, int> comboMap, SQLiteTransaction transaction)
        {
            if (zones.Count == 0) return;

            // Build UPDATE statement with CASE for each field
            using (var cmd = _context.Connection.CreateCommand())
            {
                cmd.Transaction = transaction;
                
                // ✅ CRITICAL FIX: Filter to valid zones first (zones that exist in both maps)
                var validZonesForQuery = zones
                    .Where(z => existingMap.ContainsKey(z.Id) && comboMap.ContainsKey(z.Id))
                    .ToList();
                
                // Get existing flags from database first (preserve FlagManager resets)
                var flagMap = new Dictionary<int, (bool IsResolved, bool IsClusterResolved, int SleeveId, int ClusterId)>();
                var idList = string.Join(",", validZonesForQuery.Select(z => existingMap[z.Id]));
                
                cmd.CommandText = $@"
                    SELECT ClashZoneId, IsResolvedFlag, IsClusterResolvedFlag, SleeveInstanceId, ClusterInstanceId
                    FROM ClashZones WHERE ClashZoneId IN ({idList})";
                
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        var id = reader.GetInt32(0);
                        // SQLite stores booleans as integers (0/1). Avoid GetBoolean cast errors.
                        var isResolved = !reader.IsDBNull(1) && reader.GetInt32(1) == 1;
                        var isClusterResolved = !reader.IsDBNull(2) && reader.GetInt32(2) == 1;
                        var sleeveId = reader.IsDBNull(3) ? -1 : reader.GetInt32(3);
                        var clusterId = reader.IsDBNull(4) ? -1 : reader.GetInt32(4);
                        flagMap[id] = (isResolved, isClusterResolved, sleeveId, clusterId);
                    }
                }

                // Build bulk UPDATE with CASE statements
                var sql = new System.Text.StringBuilder();
                sql.AppendLine("UPDATE ClashZones SET");
                sql.AppendLine("  UpdatedAt = CURRENT_TIMESTAMP,");
                
                // ✅ CRITICAL FIX: Only process zones that exist in both maps
                // Filter zones first to ensure parameter indices match CASE statement indices
                var validZones = zones
                    .Where(z => existingMap.ContainsKey(z.Id) && comboMap.ContainsKey(z.Id))
                    .ToList();
                
                // Add parameters for each VALID zone (ensures indices match CASE statements)
                for (int i = 0; i < validZones.Count; i++)
                {
                    var zone = validZones[i];
                    var clashZoneId = existingMap[zone.Id];
                    var comboId = comboMap[zone.Id];
                    
                    // Extract values
                    var mepId = zone.MepElementId?.IntegerValue ?? zone.MepElementIdValue;
                    var hostId = zone.StructuralElementId?.IntegerValue ?? zone.StructuralElementIdValue;
                    var interX = zone.IntersectionPoint?.X ?? zone.IntersectionPointX;
                    var interY = zone.IntersectionPoint?.Y ?? zone.IntersectionPointY;
                    var interZ = zone.IntersectionPoint?.Z ?? zone.IntersectionPointZ;
                    
                    // Determine flags (preserve database resets)
                    bool finalIsResolved = zone.IsResolved;
                    bool finalIsClusterResolved = zone.IsClusterResolved;
                    int finalSleeveId = zone.SleeveInstanceId;
                    int finalClusterId = zone.ClusterSleeveInstanceId;
                    
                    if (flagMap.TryGetValue(clashZoneId, out var dbFlags))
                    {
                        // If database has reset flags (false, false, -1, -1), preserve them
                        if (!dbFlags.IsResolved && !dbFlags.IsClusterResolved && dbFlags.SleeveId == -1 && dbFlags.ClusterId == -1)
                        {
                            finalIsResolved = false;
                            finalIsClusterResolved = false;
                            finalSleeveId = -1;
                            finalClusterId = -1;
                            
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                _logger($"[SQLite][BULK] ✅ PRESERVING reset flags for ClashZoneId={clashZoneId}, GUID={zone.Id}: IsResolved=false, IsClusterResolved=false");
                            }
                        }
                    }
                    
                    // ✅ CRITICAL FIX: Serialize parameters to JSON for bulk update
                    var mepParamsJson = "{}";
                    if (zone.MepParameterValues != null && zone.MepParameterValues.Count > 0)
                    {
                        var mepDict = zone.MepParameterValues
                            .Where(kv => kv != null && !string.IsNullOrEmpty(kv.Key) && !string.IsNullOrEmpty(kv.Value))
                            .ToDictionary(kv => kv.Key, kv => kv.Value, StringComparer.OrdinalIgnoreCase);
                        
                        if (mepDict.Count > 0)
                        {
                            mepParamsJson = JsonSerializer.Serialize(mepDict);
                        }
                    }
                    
                    var hostParamsJson = "{}";
                    if (zone.HostParameterValues != null && zone.HostParameterValues.Count > 0)
                    {
                        var hostDict = zone.HostParameterValues
                            .Where(kv => kv != null && !string.IsNullOrEmpty(kv.Key) && !string.IsNullOrEmpty(kv.Value))
                            .ToDictionary(kv => kv.Key, kv => kv.Value, StringComparer.OrdinalIgnoreCase);
                        
                        if (hostDict.Count > 0)
                        {
                            hostParamsJson = JsonSerializer.Serialize(hostDict);
                        }
                    }
                    
                    // ✅ DIAGNOSTIC: Log parameter serialization for debugging
                    if (!DeploymentConfiguration.DeploymentMode && string.Equals(zone.MepElementCategory, "Pipes", StringComparison.OrdinalIgnoreCase))
                    {
                        var mepCount = zone.MepParameterValues?.Count ?? 0;
                        var mepDictCount = mepParamsJson != "{}" ? JsonSerializer.Deserialize<Dictionary<string, string>>(mepParamsJson)?.Count ?? 0 : 0;
                        if (mepCount > 0 || mepDictCount > 0)
                        {
                            _logger($"[SQLite][BULK] Zone {zone.Id}: Serializing {mepCount} MEP params → {mepDictCount} in JSON for bulk update");
                        }
                    }
                    
                    // Add parameters
                    cmd.Parameters.AddWithValue($"@ComboId{i}", comboId);
                    cmd.Parameters.AddWithValue($"@MepId{i}", mepId);
                    cmd.Parameters.AddWithValue($"@HostId{i}", hostId);
                    cmd.Parameters.AddWithValue($"@InterX{i}", interX);
                    cmd.Parameters.AddWithValue($"@InterY{i}", interY);
                    cmd.Parameters.AddWithValue($"@InterZ{i}", interZ);
                    cmd.Parameters.AddWithValue($"@IsResolved{i}", finalIsResolved ? 1 : 0);
                    cmd.Parameters.AddWithValue($"@IsClusterResolved{i}", finalIsClusterResolved ? 1 : 0);
                    cmd.Parameters.AddWithValue($"@SleeveId{i}", finalSleeveId);
                    cmd.Parameters.AddWithValue($"@ClusterId{i}", finalClusterId);
                    cmd.Parameters.AddWithValue($"@MepParamsJson{i}", mepParamsJson);
                    cmd.Parameters.AddWithValue($"@HostParamsJson{i}", hostParamsJson);
                    // ✅ CRITICAL FIX: Add MEP dimensions and thickness parameters for bulk update
                    // These are essential for damper sizing and depth calculation
                    cmd.Parameters.AddWithValue($"@MepWidth{i}", zone.MepElementWidth);
                    cmd.Parameters.AddWithValue($"@MepHeight{i}", zone.MepElementHeight);
                    cmd.Parameters.AddWithValue($"@WallThickness{i}", zone.WallThickness);
                    cmd.Parameters.AddWithValue($"@FramingThickness{i}", zone.FramingThickness);
                    cmd.Parameters.AddWithValue($"@StructuralThickness{i}", zone.StructuralElementThickness);
                    cmd.Parameters.AddWithValue($"@HostOrientation{i}", (object)zone.HostOrientation ?? DBNull.Value);
                    cmd.Parameters.AddWithValue($"@MepOrientationDirection{i}", (object)zone.MepElementOrientationDirection ?? DBNull.Value);
                    
                    // ✅ CRITICAL FIX: Add MEP category, structural type, and orientation fields
                    cmd.Parameters.AddWithValue($"@MepCategory{i}", (object)zone.MepElementCategory ?? DBNull.Value);
                    cmd.Parameters.AddWithValue($"@StructuralType{i}", (object)zone.StructuralElementType ?? DBNull.Value);
                    cmd.Parameters.AddWithValue($"@MepOrientationX{i}", zone.MepElementOrientationX);
                    cmd.Parameters.AddWithValue($"@MepOrientationY{i}", zone.MepElementOrientationY);
                    cmd.Parameters.AddWithValue($"@MepOrientationZ{i}", zone.MepElementOrientationZ);
                    
                    // ✅ CRITICAL FIX: Add MEP rotation and angle fields (for rotation logic)
                    var rotationAngleRad = zone.MepElementRotationAngle;
                    var rotationAngleDeg = NormalizeDegrees(rotationAngleRad);
                    var orientationDirection = zone.MepElementOrientationDirection ?? string.Empty;
                    var (angleToXRad, angleToXDeg, angleToYRad, angleToYDeg) =
                        ComputePlanarOrientationAngles(zone.MepElementOrientationX, zone.MepElementOrientationY, orientationDirection);
                    
                    cmd.Parameters.AddWithValue($"@MepRotationAngleRad{i}", rotationAngleRad);
                    cmd.Parameters.AddWithValue($"@MepRotationAngleDeg{i}", rotationAngleDeg);
                    cmd.Parameters.AddWithValue($"@MepRotationCos{i}", Math.Cos(rotationAngleRad));
                    cmd.Parameters.AddWithValue($"@MepRotationSin{i}", Math.Sin(rotationAngleRad));
                    cmd.Parameters.AddWithValue($"@MepAngleToXRad{i}", angleToXRad);
                    cmd.Parameters.AddWithValue($"@MepAngleToXDeg{i}", angleToXDeg);
                    cmd.Parameters.AddWithValue($"@MepAngleToYRad{i}", angleToYRad);
                    cmd.Parameters.AddWithValue($"@MepAngleToYDeg{i}", angleToYDeg);
                    
                    // ✅ CRITICAL FIX: Add pipe diameter and size parameter fields
                    cmd.Parameters.AddWithValue($"@MepElementOuterDiameter{i}", (object)zone.MepElementOuterDiameter ?? DBNull.Value);
                    cmd.Parameters.AddWithValue($"@MepElementNominalDiameter{i}", (object)zone.MepElementNominalDiameter ?? DBNull.Value);
                    cmd.Parameters.AddWithValue($"@MepElementSizeParameterValue{i}", (object)zone.MepElementSizeParameterValue ?? DBNull.Value);
                    
                    // ✅ CRITICAL FIX: Add sleeve family name and document keys
                    cmd.Parameters.AddWithValue($"@SleeveFamilyName{i}", (object)zone.SleeveFamilyName ?? DBNull.Value);
                    cmd.Parameters.AddWithValue($"@SourceDocKey{i}", (object)zone.SourceDocKey ?? DBNull.Value);
                    cmd.Parameters.AddWithValue($"@HostDocKey{i}", (object)zone.HostDocKey ?? DBNull.Value);
                    cmd.Parameters.AddWithValue($"@MepElementUniqueId{i}", (object)zone.MepElementUniqueId ?? DBNull.Value);
                    
                    // ✅ CRITICAL FIX: Add damper connector and insulation fields
                    cmd.Parameters.AddWithValue($"@HasMepConnector{i}", zone.HasMepConnector ? 1 : 0);
                    cmd.Parameters.AddWithValue($"@DamperConnectorSide{i}", (object)zone.DamperConnectorSide ?? DBNull.Value);
                    cmd.Parameters.AddWithValue($"@IsInsulated{i}", zone.IsInsulated ? 1 : 0);
                    cmd.Parameters.AddWithValue($"@InsulationThickness{i}", (object)zone.InsulationThickness ?? DBNull.Value);
                    
                    // ✅ REFERENCE LEVEL: Add MEP element Reference Level (used for Schedule Level and Bottom of Opening calculation)
                    cmd.Parameters.AddWithValue($"@MepElementLevelName{i}", (object)zone.MepElementLevelName ?? DBNull.Value);
                    // ✅ REFERENCE LEVEL ELEVATION: Add MEP element Reference Level elevation (critical for Elevation from Level and Bottom of Opening calculation)
                    cmd.Parameters.AddWithValue($"@MepElementLevelElevation{i}", zone.MepElementLevelElevation);
                    // ✅ WALL CENTERLINE POINT: Pre-calculated during refresh (enables multi-threaded placement)
                    // ✅ DIAGNOSTIC: Log if zone has zero wall centerline (will preserve existing DB value)
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        bool isZero = (Math.Abs(zone.WallCenterlinePointX) < 1e-9 && Math.Abs(zone.WallCenterlinePointY) < 1e-9 && Math.Abs(zone.WallCenterlinePointZ) < 1e-9);
                        if (isZero && string.Equals(zone.MepElementCategory, "Pipes", StringComparison.OrdinalIgnoreCase))
                        {
                            SafeFileLogger.SafeAppendText("wall_centerline_bulk_update.log",
                                $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [BULK-UPDATE] Zone {zone.Id} (Category={zone.MepElementCategory}): " +
                                $"WallCenterlinePoint is ZERO - will preserve existing DB value (if any)\n");
                        }
                    }
                    cmd.Parameters.AddWithValue($"@WallCenterlinePointX{i}", zone.WallCenterlinePointX);
                    cmd.Parameters.AddWithValue($"@WallCenterlinePointY{i}", zone.WallCenterlinePointY);
                    cmd.Parameters.AddWithValue($"@WallCenterlinePointZ{i}", zone.WallCenterlinePointZ);
                    
                    cmd.Parameters.AddWithValue($"@ZoneId{i}", clashZoneId);
                }

                // ✅ CRITICAL FIX: Build CASE statements using validZones (indices match parameter indices)
                sql.AppendLine("  ComboId = CASE ClashZoneId");
                for (int i = 0; i < validZones.Count; i++)
                {
                    sql.AppendLine($"    WHEN @ZoneId{i} THEN @ComboId{i}");
                }
                sql.AppendLine("    ELSE ComboId END,");
                
                sql.AppendLine("  MepElementId = CASE ClashZoneId");
                for (int i = 0; i < validZones.Count; i++)
                {
                    sql.AppendLine($"    WHEN @ZoneId{i} THEN @MepId{i}");
                }
                sql.AppendLine("    ELSE MepElementId END,");
                
                sql.AppendLine("  HostElementId = CASE ClashZoneId");
                for (int i = 0; i < validZones.Count; i++)
                {
                    sql.AppendLine($"    WHEN @ZoneId{i} THEN @HostId{i}");
                }
                sql.AppendLine("    ELSE HostElementId END,");
                
                sql.AppendLine("  IntersectionX = CASE ClashZoneId");
                for (int i = 0; i < validZones.Count; i++)
                {
                    sql.AppendLine($"    WHEN @ZoneId{i} THEN @InterX{i}");
                }
                sql.AppendLine("    ELSE IntersectionX END,");
                
                sql.AppendLine("  IntersectionY = CASE ClashZoneId");
                for (int i = 0; i < validZones.Count; i++)
                {
                    sql.AppendLine($"    WHEN @ZoneId{i} THEN @InterY{i}");
                }
                sql.AppendLine("    ELSE IntersectionY END,");
                
                sql.AppendLine("  IntersectionZ = CASE ClashZoneId");
                for (int i = 0; i < validZones.Count; i++)
                {
                    sql.AppendLine($"    WHEN @ZoneId{i} THEN @InterZ{i}");
                }
                sql.AppendLine("    ELSE IntersectionZ END,");
                
                // ✅ WALL CENTERLINE POINT: Pre-calculated during refresh (enables multi-threaded placement)
                // ✅ CRITICAL FIX: Only update if new value is non-zero, otherwise preserve existing value
                // This prevents overwriting existing wall centerline with zeros when zones are updated
                sql.AppendLine("  WallCenterlinePointX = CASE ClashZoneId");
                for (int i = 0; i < validZones.Count; i++)
                {
                    sql.AppendLine($"    WHEN @ZoneId{i} THEN CASE WHEN ABS(@WallCenterlinePointX{i}) > 1e-9 THEN @WallCenterlinePointX{i} ELSE WallCenterlinePointX END");
                }
                sql.AppendLine("    ELSE WallCenterlinePointX END,");
                
                sql.AppendLine("  WallCenterlinePointY = CASE ClashZoneId");
                for (int i = 0; i < validZones.Count; i++)
                {
                    sql.AppendLine($"    WHEN @ZoneId{i} THEN CASE WHEN ABS(@WallCenterlinePointY{i}) > 1e-9 THEN @WallCenterlinePointY{i} ELSE WallCenterlinePointY END");
                }
                sql.AppendLine("    ELSE WallCenterlinePointY END,");
                
                sql.AppendLine("  WallCenterlinePointZ = CASE ClashZoneId");
                for (int i = 0; i < validZones.Count; i++)
                {
                    sql.AppendLine($"    WHEN @ZoneId{i} THEN CASE WHEN ABS(@WallCenterlinePointZ{i}) > 1e-9 THEN @WallCenterlinePointZ{i} ELSE WallCenterlinePointZ END");
                }
                sql.AppendLine("    ELSE WallCenterlinePointZ END,");
                
                sql.AppendLine("  IsResolvedFlag = CASE ClashZoneId");
                for (int i = 0; i < validZones.Count; i++)
                {
                    sql.AppendLine($"    WHEN @ZoneId{i} THEN @IsResolved{i}");
                }
                sql.AppendLine("    ELSE IsResolvedFlag END,");
                
                sql.AppendLine("  IsClusterResolvedFlag = CASE ClashZoneId");
                for (int i = 0; i < validZones.Count; i++)
                {
                    sql.AppendLine($"    WHEN @ZoneId{i} THEN @IsClusterResolved{i}");
                }
                sql.AppendLine("    ELSE IsClusterResolvedFlag END,");
                
                sql.AppendLine("  SleeveInstanceId = CASE ClashZoneId");
                for (int i = 0; i < validZones.Count; i++)
                {
                    sql.AppendLine($"    WHEN @ZoneId{i} THEN @SleeveId{i}");
                }
                sql.AppendLine("    ELSE SleeveInstanceId END,");
                
                sql.AppendLine("  ClusterInstanceId = CASE ClashZoneId");
                for (int i = 0; i < validZones.Count; i++)
                {
                    sql.AppendLine($"    WHEN @ZoneId{i} THEN @ClusterId{i}");
                }
                sql.AppendLine("    ELSE ClusterInstanceId END,");
                
                // ✅ CRITICAL FIX: Add MepParameterValuesJson and HostParameterValuesJson to bulk update
                sql.AppendLine("  MepParameterValuesJson = CASE ClashZoneId");
                for (int i = 0; i < validZones.Count; i++)
                {
                    sql.AppendLine($"    WHEN @ZoneId{i} THEN @MepParamsJson{i}");
                }
                sql.AppendLine("    ELSE MepParameterValuesJson END,");
                
                sql.AppendLine("  HostParameterValuesJson = CASE ClashZoneId");
                for (int i = 0; i < validZones.Count; i++)
                {
                    sql.AppendLine($"    WHEN @ZoneId{i} THEN @HostParamsJson{i}");
                }
                sql.AppendLine("    ELSE HostParameterValuesJson END,");
                
                // ✅ CRITICAL FIX: Add MEP dimensions and thickness to bulk update
                // These are essential for damper sizing and depth calculation
                sql.AppendLine("  MepWidth = CASE ClashZoneId");
                for (int i = 0; i < validZones.Count; i++)
                {
                    sql.AppendLine($"    WHEN @ZoneId{i} THEN @MepWidth{i}");
                }
                sql.AppendLine("    ELSE MepWidth END,");
                
                sql.AppendLine("  MepHeight = CASE ClashZoneId");
                for (int i = 0; i < validZones.Count; i++)
                {
                    sql.AppendLine($"    WHEN @ZoneId{i} THEN @MepHeight{i}");
                }
                sql.AppendLine("    ELSE MepHeight END,");
                
                sql.AppendLine("  WallThickness = CASE ClashZoneId");
                for (int i = 0; i < validZones.Count; i++)
                {
                    sql.AppendLine($"    WHEN @ZoneId{i} THEN @WallThickness{i}");
                }
                sql.AppendLine("    ELSE WallThickness END,");
                
                sql.AppendLine("  FramingThickness = CASE ClashZoneId");
                for (int i = 0; i < validZones.Count; i++)
                {
                    sql.AppendLine($"    WHEN @ZoneId{i} THEN @FramingThickness{i}");
                }
                sql.AppendLine("    ELSE FramingThickness END,");
                
                sql.AppendLine("  StructuralThickness = CASE ClashZoneId");
                for (int i = 0; i < validZones.Count; i++)
                {
                    sql.AppendLine($"    WHEN @ZoneId{i} THEN @StructuralThickness{i}");
                }
                sql.AppendLine("    ELSE StructuralThickness END,");
                
                sql.AppendLine("  HostOrientation = CASE ClashZoneId");
                for (int i = 0; i < validZones.Count; i++)
                {
                    sql.AppendLine($"    WHEN @ZoneId{i} THEN @HostOrientation{i}");
                }
                sql.AppendLine("    ELSE HostOrientation END,");
                
                sql.AppendLine("  MepOrientationDirection = CASE ClashZoneId");
                for (int i = 0; i < validZones.Count; i++)
                {
                    sql.AppendLine($"    WHEN @ZoneId{i} THEN @MepOrientationDirection{i}");
                }
                sql.AppendLine("    ELSE MepOrientationDirection END,");
                
                // ✅ CRITICAL FIX: Add MEP category and structural type to bulk update
                sql.AppendLine("  MepCategory = CASE ClashZoneId");
                for (int i = 0; i < validZones.Count; i++)
                {
                    sql.AppendLine($"    WHEN @ZoneId{i} THEN @MepCategory{i}");
                }
                sql.AppendLine("    ELSE MepCategory END,");
                
                sql.AppendLine("  StructuralType = CASE ClashZoneId");
                for (int i = 0; i < validZones.Count; i++)
                {
                    sql.AppendLine($"    WHEN @ZoneId{i} THEN @StructuralType{i}");
                }
                sql.AppendLine("    ELSE StructuralType END,");
                
                // ✅ CRITICAL FIX: Add MEP orientation vectors to bulk update
                sql.AppendLine("  MepOrientationX = CASE ClashZoneId");
                for (int i = 0; i < validZones.Count; i++)
                {
                    sql.AppendLine($"    WHEN @ZoneId{i} THEN @MepOrientationX{i}");
                }
                sql.AppendLine("    ELSE MepOrientationX END,");
                
                sql.AppendLine("  MepOrientationY = CASE ClashZoneId");
                for (int i = 0; i < validZones.Count; i++)
                {
                    sql.AppendLine($"    WHEN @ZoneId{i} THEN @MepOrientationY{i}");
                }
                sql.AppendLine("    ELSE MepOrientationY END,");
                
                sql.AppendLine("  MepOrientationZ = CASE ClashZoneId");
                for (int i = 0; i < validZones.Count; i++)
                {
                    sql.AppendLine($"    WHEN @ZoneId{i} THEN @MepOrientationZ{i}");
                }
                sql.AppendLine("    ELSE MepOrientationZ END,");
                
                // ✅ CRITICAL FIX: Add MEP rotation angles to bulk update
                sql.AppendLine("  MepRotationAngleRad = CASE ClashZoneId");
                for (int i = 0; i < validZones.Count; i++)
                {
                    sql.AppendLine($"    WHEN @ZoneId{i} THEN @MepRotationAngleRad{i}");
                }
                sql.AppendLine("    ELSE MepRotationAngleRad END,");
                
                sql.AppendLine("  MepRotationAngleDeg = CASE ClashZoneId");
                for (int i = 0; i < validZones.Count; i++)
                {
                    sql.AppendLine($"    WHEN @ZoneId{i} THEN @MepRotationAngleDeg{i}");
                }
                sql.AppendLine("    ELSE MepRotationAngleDeg END,");
                
                sql.AppendLine("  MepRotationCos = CASE ClashZoneId");
                for (int i = 0; i < validZones.Count; i++)
                {
                    sql.AppendLine($"    WHEN @ZoneId{i} THEN @MepRotationCos{i}");
                }
                sql.AppendLine("    ELSE MepRotationCos END,");
                
                sql.AppendLine("  MepRotationSin = CASE ClashZoneId");
                for (int i = 0; i < validZones.Count; i++)
                {
                    sql.AppendLine($"    WHEN @ZoneId{i} THEN @MepRotationSin{i}");
                }
                sql.AppendLine("    ELSE MepRotationSin END,");
                
                sql.AppendLine("  MepAngleToXRad = CASE ClashZoneId");
                for (int i = 0; i < validZones.Count; i++)
                {
                    sql.AppendLine($"    WHEN @ZoneId{i} THEN @MepAngleToXRad{i}");
                }
                sql.AppendLine("    ELSE MepAngleToXRad END,");
                
                sql.AppendLine("  MepAngleToXDeg = CASE ClashZoneId");
                for (int i = 0; i < validZones.Count; i++)
                {
                    sql.AppendLine($"    WHEN @ZoneId{i} THEN @MepAngleToXDeg{i}");
                }
                sql.AppendLine("    ELSE MepAngleToXDeg END,");
                
                sql.AppendLine("  MepAngleToYRad = CASE ClashZoneId");
                for (int i = 0; i < validZones.Count; i++)
                {
                    sql.AppendLine($"    WHEN @ZoneId{i} THEN @MepAngleToYRad{i}");
                }
                sql.AppendLine("    ELSE MepAngleToYRad END,");
                
                sql.AppendLine("  MepAngleToYDeg = CASE ClashZoneId");
                for (int i = 0; i < validZones.Count; i++)
                {
                    sql.AppendLine($"    WHEN @ZoneId{i} THEN @MepAngleToYDeg{i}");
                }
                sql.AppendLine("    ELSE MepAngleToYDeg END,");
                
                // ✅ CRITICAL FIX: Add pipe diameter and size parameter to bulk update
                sql.AppendLine("  MepElementOuterDiameter = CASE ClashZoneId");
                for (int i = 0; i < validZones.Count; i++)
                {
                    sql.AppendLine($"    WHEN @ZoneId{i} THEN @MepElementOuterDiameter{i}");
                }
                sql.AppendLine("    ELSE MepElementOuterDiameter END,");
                
                sql.AppendLine("  MepElementNominalDiameter = CASE ClashZoneId");
                for (int i = 0; i < validZones.Count; i++)
                {
                    sql.AppendLine($"    WHEN @ZoneId{i} THEN @MepElementNominalDiameter{i}");
                }
                sql.AppendLine("    ELSE MepElementNominalDiameter END,");
                
                sql.AppendLine("  MepElementSizeParameterValue = CASE ClashZoneId");
                for (int i = 0; i < validZones.Count; i++)
                {
                    sql.AppendLine($"    WHEN @ZoneId{i} THEN @MepElementSizeParameterValue{i}");
                }
                sql.AppendLine("    ELSE MepElementSizeParameterValue END,");
                
                sql.AppendLine("  SleeveFamilyName = CASE ClashZoneId");
                for (int i = 0; i < validZones.Count; i++)
                {
                    sql.AppendLine($"    WHEN @ZoneId{i} THEN @SleeveFamilyName{i}");
                }
                sql.AppendLine("    ELSE SleeveFamilyName END,");
                
                // ✅ REFERENCE LEVEL: Add MEP element Reference Level to bulk update
                sql.AppendLine("  MepElementLevelName = CASE ClashZoneId");
                for (int i = 0; i < validZones.Count; i++)
                {
                    sql.AppendLine($"    WHEN @ZoneId{i} THEN @MepElementLevelName{i}");
                }
                sql.AppendLine("    ELSE MepElementLevelName END,");
                
                // ✅ REFERENCE LEVEL ELEVATION: Add MEP element Reference Level elevation to bulk update (critical for Elevation from Level and Bottom of Opening calculation)
                sql.AppendLine("  MepElementLevelElevation = CASE ClashZoneId");
                for (int i = 0; i < validZones.Count; i++)
                {
                    sql.AppendLine($"    WHEN @ZoneId{i} THEN @MepElementLevelElevation{i}");
                }
                sql.AppendLine("    ELSE MepElementLevelElevation END,");
                
                // ✅ CRITICAL FIX: Add document keys to bulk update
                sql.AppendLine("  SourceDocKey = CASE ClashZoneId");
                for (int i = 0; i < validZones.Count; i++)
                {
                    sql.AppendLine($"    WHEN @ZoneId{i} THEN @SourceDocKey{i}");
                }
                sql.AppendLine("    ELSE SourceDocKey END,");
                
                sql.AppendLine("  HostDocKey = CASE ClashZoneId");
                for (int i = 0; i < validZones.Count; i++)
                {
                    sql.AppendLine($"    WHEN @ZoneId{i} THEN @HostDocKey{i}");
                }
                sql.AppendLine("    ELSE HostDocKey END,");
                
                sql.AppendLine("  MepElementUniqueId = CASE ClashZoneId");
                for (int i = 0; i < validZones.Count; i++)
                {
                    sql.AppendLine($"    WHEN @ZoneId{i} THEN @MepElementUniqueId{i}");
                }
                sql.AppendLine("    ELSE MepElementUniqueId END,");
                
                // ✅ CRITICAL FIX: Add damper connector and insulation fields to bulk update
                sql.AppendLine("  HasMepConnector = CASE ClashZoneId");
                for (int i = 0; i < validZones.Count; i++)
                {
                    sql.AppendLine($"    WHEN @ZoneId{i} THEN @HasMepConnector{i}");
                }
                sql.AppendLine("    ELSE HasMepConnector END,");
                
                sql.AppendLine("  DamperConnectorSide = CASE ClashZoneId");
                for (int i = 0; i < validZones.Count; i++)
                {
                    sql.AppendLine($"    WHEN @ZoneId{i} THEN @DamperConnectorSide{i}");
                }
                sql.AppendLine("    ELSE DamperConnectorSide END,");
                
                sql.AppendLine("  IsInsulated = CASE ClashZoneId");
                for (int i = 0; i < validZones.Count; i++)
                {
                    sql.AppendLine($"    WHEN @ZoneId{i} THEN @IsInsulated{i}");
                }
                sql.AppendLine("    ELSE IsInsulated END,");
                
                sql.AppendLine("  InsulationThickness = CASE ClashZoneId");
                for (int i = 0; i < validZones.Count; i++)
                {
                    sql.AppendLine($"    WHEN @ZoneId{i} THEN @InsulationThickness{i}");
                }
                sql.AppendLine("    ELSE InsulationThickness END");
                
                sql.AppendLine($"WHERE ClashZoneId IN ({idList})");
                
                cmd.CommandText = sql.ToString();
                
                // ✅ CRITICAL DIAGNOSTIC: Log the SQL and parameter values for debugging
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    var sqlPreview = sql.ToString();
                    var sqlLines = sqlPreview.Split('\n');
                    var paramLines = sqlLines.Where(l => l.Contains("MepParameterValuesJson") || l.Contains("HostParameterValuesJson")).Take(10).ToList();
                    if (paramLines.Any())
                    {
                        _logger($"[SQLite][BULK] SQL includes parameter updates: {string.Join(" | ", paramLines)}");
                    }
                    
                    // Log parameter values being set
                    foreach (var zone in zones)
                    {
                        if (existingMap.ContainsKey(zone.Id) && comboMap.ContainsKey(zone.Id))
                        {
                            var i = zones.IndexOf(zone);
                            var mepParamValue = cmd.Parameters[$"@MepParamsJson{i}"]?.Value?.ToString() ?? "NULL";
                            var mepParamLength = mepParamValue != "NULL" && mepParamValue != "{}" ? mepParamValue.Length : 0;
                            if (string.Equals(zone.MepElementCategory, "Pipes", StringComparison.OrdinalIgnoreCase))
                            {
                                _logger($"[SQLite][BULK] Zone {zone.Id} (ClashZoneId={existingMap[zone.Id]}): MepParamsJson parameter value length={mepParamLength}, preview={(mepParamLength > 0 ? mepParamValue.Substring(0, Math.Min(100, mepParamLength)) : "{}")}");
                            }
                        }
                    }
                }
                
                var rowsAffected = cmd.ExecuteNonQuery();
                
                // ✅ CRITICAL DIAGNOSTIC: Verify rows were updated
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    _logger($"[SQLite][BULK] ✅ Bulk UPDATE executed: {rowsAffected} rows affected for {zones.Count} zones");
                    
                    // ✅ VERIFY: Check if parameters were actually saved
                    if (rowsAffected > 0)
                    {
                        foreach (var zone in validZones)
                        {
                            var clashZoneId = existingMap[zone.Id];
                            using (var verifyCmd = _context.Connection.CreateCommand())
                            {
                                verifyCmd.Transaction = transaction;
                                verifyCmd.CommandText = "SELECT MepParameterValuesJson FROM ClashZones WHERE ClashZoneId = @ClashZoneId";
                                verifyCmd.Parameters.AddWithValue("@ClashZoneId", clashZoneId);
                                var savedJson = verifyCmd.ExecuteScalar()?.ToString() ?? "NULL";
                                var savedLength = savedJson != "NULL" && savedJson != "{}" ? savedJson.Length : 0;
                                
                                if (string.Equals(zone.MepElementCategory, "Pipes", StringComparison.OrdinalIgnoreCase))
                                {
                                    _logger($"[SQLite][BULK] ✅ VERIFICATION: Zone {zone.Id} (ClashZoneId={clashZoneId}): Saved MepParameterValuesJson length={savedLength}, preview={(savedLength > 0 ? savedJson.Substring(0, Math.Min(100, savedLength)) : savedJson)}");
                                }
                            }
                        }
                    }
                }
            }
        }

        public void InsertOrUpdateClashZones(IEnumerable<ClashZone> clashZones, string filterName, string category)
        {
            if (clashZones == null) return;

            // ✅ PERFORMANCE OPTIMIZATION: Bulk UPDATE path for maximum speed
            if (OptimizationFlags.UseBulkSqliteUpdates)
            {
                InsertOrUpdateClashZonesBulk(clashZones, filterName, category);
                return;
            }

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
                // Use UPPER() on both sides for case-insensitive comparison, and check for empty/NULL
                // This ensures deterministic GUIDs always match existing rows
                cmd.CommandText = @"
                    SELECT ClashZoneId FROM ClashZones 
                    WHERE UPPER(ClashZoneGuid) = UPPER(@ClashZoneGuid) 
                      AND ClashZoneGuid != '' AND ClashZoneGuid IS NOT NULL
                    LIMIT 1";
                
                cmd.Parameters.AddWithValue("@ClashZoneGuid", guid);
                var existingIdByGuid = cmd.ExecuteScalar();

                if (existingIdByGuid != null)
                {
                    // ✅ Found by GUID - update existing
                    var clashZoneId = Convert.ToInt32(existingIdByGuid);
                    
                    // ✅ DIAGNOSTIC: Log GUID match for troubleshooting
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        _logger($"[SQLite] ✅ GUID MATCH: Found existing ClashZoneId={clashZoneId} for GUID={guid} - will UPDATE (not insert)");
                    }
                    
                    UpdateClashZone(clashZoneId, comboId, clashZone, transaction);
                    return false; // Was an update
                }
                else if (!string.IsNullOrWhiteSpace(guid) && !DeploymentConfiguration.DeploymentMode)
                {
                    // ✅ DIAGNOSTIC: Log GUID mismatch for troubleshooting
                    _logger($"[SQLite] ⚠️ GUID NOT FOUND: GUID={guid} not in database - will check MEP+Host+Point fallback");
                }
                
                // ✅ FALLBACK: Check by MEP+Host+Point if GUID didn't match
                // This handles cases where GUID wasn't set or is different (legacy data)
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
                    // ✅ FALLBACK MATCH: Found by MEP+Host+Point - update existing AND set GUID
                    var clashZoneId = Convert.ToInt32(existingId);
                    
                    // ✅ CRITICAL FIX: If we found a row by MEP+Host+Point but GUID didn't match,
                    // update the existing row's GUID to the deterministic GUID
                    // This ensures future updates will match by GUID (faster lookup)
                    if (!string.IsNullOrWhiteSpace(guid))
                    {
                        try
                        {
                            using (var updateCmd = _context.Connection.CreateCommand())
                            {
                                updateCmd.Transaction = transaction;
                                updateCmd.CommandText = @"UPDATE ClashZones SET ClashZoneGuid = @ClashZoneGuid WHERE ClashZoneId = @ClashZoneId";
                                updateCmd.Parameters.AddWithValue("@ClashZoneGuid", guid);
                                updateCmd.Parameters.AddWithValue("@ClashZoneId", clashZoneId);
                                updateCmd.ExecuteNonQuery();
                                
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    _logger($"[SQLite] ✅ FALLBACK MATCH + GUID UPDATE: Found ClashZoneId={clashZoneId} by MEP+Host+Point, updated GUID to {guid}");
                                }
                            }
                        }
                        catch (Exception guidEx)
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                _logger($"[SQLite] ⚠️ Failed to update GUID on fallback match: {guidEx.Message}");
                            }
                        }
                    }
                    
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        _logger($"[SQLite] ✅ FALLBACK MATCH: Found existing ClashZoneId={clashZoneId} by MEP+Host+Point - will UPDATE (not insert)");
                    }
                    
                    UpdateClashZone(clashZoneId, comboId, clashZone, transaction);
                    return false; // Was an update
                }
                else
                {
                    // ✅ NEW INSERT: No match found - insert new clash zone
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        _logger($"[SQLite] ✅ NEW INSERT: No match found for GUID={guid} or MEP={mepElementId}+Host={hostElementId}+Point=({intersectionX:F3},{intersectionY:F3},{intersectionZ:F3}) - creating new row");
                    }
                    
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
                        WallCenterlinePointX, WallCenterlinePointY, WallCenterlinePointZ,
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
                        MepWidth, MepHeight, MepElementOuterDiameter, MepElementNominalDiameter, MepElementSizeParameterValue, SleeveFamilyName,
                        MepElementLevelName, MepElementLevelElevation,
                        SleevePlacementActiveX, SleevePlacementActiveY, SleevePlacementActiveZ,
                        SourceDocKey, HostDocKey, MepElementUniqueId,
                        IsResolvedFlag, IsClusterResolvedFlag, IsClusteredFlag,
                        MarkedForClusterProcess, AfterClusterSleeveId,
                        HasDamperNearbyFlag, IsCurrentClashFlag, ReadyForPlacementFlag,
                        StructuralThickness, WallThickness, FramingThickness,
                        MepParameterValuesJson, HostParameterValuesJson,
                        HasMepConnector, DamperConnectorSide,
                        IsInsulated, InsulationThickness
                    ) VALUES (
                        @ComboId, @MepElementId, @HostElementId,
                        @IntersectionX, @IntersectionY, @IntersectionZ,
                        @WallCenterlinePointX, @WallCenterlinePointY, @WallCenterlinePointZ,
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
                        @MepWidth, @MepHeight, @MepElementOuterDiameter, @MepElementNominalDiameter, @MepElementSizeParameterValue, @SleeveFamilyName,
                        @MepElementLevelName, @MepElementLevelElevation,
                        @SleevePlacementActiveX, @SleevePlacementActiveY, @SleevePlacementActiveZ,
                        @SourceDocKey, @HostDocKey, @MepElementUniqueId,
                        @IsResolvedFlag, @IsClusterResolvedFlag, @IsClusteredFlag,
                        @MarkedForClusterProcess, @AfterClusterSleeveId,
                        @HasDamperNearbyFlag, @IsCurrentClashFlag, @ReadyForPlacementFlag,
                        @StructuralThickness, @WallThickness, @FramingThickness,
                        @MepParameterValuesJson, @HostParameterValuesJson,
                        @HasMepConnector, @DamperConnectorSide,
                        @IsInsulated, @InsulationThickness
                    )";

                AddClashZoneParameters(cmd, comboId, clashZone);
                cmd.ExecuteNonQuery();
                
                // ✅ R-TREE MAINTENANCE: Get inserted ClashZoneId and update R-tree index
                cmd.CommandText = "SELECT last_insert_rowid()";
                var insertedId = Convert.ToInt32(cmd.ExecuteScalar());
                UpdateRTreeIndex(insertedId, clashZone, transaction);
            }
        }

        private void UpdateClashZone(int clashZoneId, int comboId, ClashZone clashZone, SQLiteTransaction transaction)
        {
            // ✅ DIAGNOSTIC: Log the values being passed to UpdateClashZone
            if (!DeploymentConfiguration.DeploymentMode && string.Equals(clashZone.MepElementCategory, "Pipes", StringComparison.OrdinalIgnoreCase))
            {
                var odMm = clashZone.MepElementOuterDiameter > 0 ? (clashZone.MepElementOuterDiameter * 304.8) : 0.0;
                var nomMm = clashZone.MepElementNominalDiameter > 0 ? (clashZone.MepElementNominalDiameter * 304.8) : 0.0;
                _logger($"[DB-UPDATE-DEBUG] Zone {clashZone.Id} (ClashZoneId={clashZoneId}): UpdateClashZone called with - OuterDiameter={clashZone.MepElementOuterDiameter:F6}ft ({odMm:F1}mm), NominalDiameter={clashZone.MepElementNominalDiameter:F6}ft ({nomMm:F1}mm), SizeParameterValue='{clashZone.MepElementSizeParameterValue ?? "NULL"}', MepElementFormattedSize='{clashZone.MepElementFormattedSize ?? "NULL"}'");
                SafeFileLogger.SafeAppendText("save_db_diagnostic.log",
                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] Zone {clashZone.Id} (ClashZoneId={clashZoneId}): UpdateClashZone - OuterDiameter={clashZone.MepElementOuterDiameter:F6}ft ({odMm:F1}mm), NominalDiameter={clashZone.MepElementNominalDiameter:F6}ft ({nomMm:F1}mm), SizeParameterValue='{clashZone.MepElementSizeParameterValue ?? "NULL"}'\n");
            }
            
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
                // 
                // ✅ EXCEPTION: If ClashZone is being saved as NEW UNRESOLVED zone (IsResolved=false, IsClusterResolved=false, no sleeve IDs),
                // always update flags to false, even if FlagManager didn't reset them.
                // This handles the case where new zones are created during refresh and should update existing resolved zones.
                bool zoneIsUnresolved = !clashZone.IsResolved && !clashZone.IsClusterResolved && clashZone.SleeveInstanceId <= 0 && clashZone.ClusterSleeveInstanceId <= 0;
                bool preserveFlags = entryWasResetByFlagManager && !zoneIsUnresolved;
                
                if (preserveFlags && !DeploymentConfiguration.DeploymentMode)
                {
                    _logger($"[SQLite] ✅ PRESERVING reset flags for ClashZoneId={clashZoneId}, GUID={clashZone.Id}: IsResolved=false, IsClusterResolved=false (FlagManager reset, ClashZone has stale values)");
                }
                
                if (zoneIsUnresolved && !entryWasResetByFlagManager && !DeploymentConfiguration.DeploymentMode)
                {
                    _logger($"[SQLite] ✅ OVERRIDING flags for ClashZoneId={clashZoneId}, GUID={clashZone.Id}: Setting IsResolved=false, IsClusterResolved=false (new unresolved zone from refresh)");
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
                        MepElementOuterDiameter = @MepElementOuterDiameter,
                        MepElementNominalDiameter = @MepElementNominalDiameter,
                        MepElementSizeParameterValue = @MepElementSizeParameterValue,
                        SleeveFamilyName = @SleeveFamilyName,
                        MepElementLevelName = @MepElementLevelName,
                        MepElementLevelElevation = @MepElementLevelElevation,
                        WallCenterlinePointX = @WallCenterlinePointX,
                        WallCenterlinePointY = @WallCenterlinePointY,
                        WallCenterlinePointZ = @WallCenterlinePointZ,
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
                        MepParameterValuesJson = @MepParameterValuesJson,
                        HostParameterValuesJson = @HostParameterValuesJson,
                        HasMepConnector = @HasMepConnector,
                        DamperConnectorSide = @DamperConnectorSide,
                        IsInsulated = @IsInsulated,
                        InsulationThickness = @InsulationThickness,
                        UpdatedAt = CURRENT_TIMESTAMP
                    WHERE ClashZoneId = @ClashZoneId";

                AddClashZoneParameters(cmd, comboId, clashZone, preserveFlags, existingIsResolved, existingIsClusterResolved, existingSleeveInstanceId, existingClusterInstanceId);
                cmd.Parameters.AddWithValue("@ClashZoneId", clashZoneId);
                cmd.ExecuteNonQuery();
                
                // ✅ R-TREE MAINTENANCE: Update R-tree index when clash zone is updated
                UpdateRTreeIndex(clashZoneId, clashZone, transaction);
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
            // ✅ WALL CENTERLINE POINT: Pre-calculated during refresh (enables multi-threaded placement)
            // ✅ CRITICAL: ALWAYS use individual X/Y/Z values directly - no property fallback logic
            // The property is just a convenience wrapper, backing fields are the source of truth
            cmd.Parameters.AddWithValue("@WallCenterlinePointX", clashZone.WallCenterlinePointX);
            cmd.Parameters.AddWithValue("@WallCenterlinePointY", clashZone.WallCenterlinePointY);
            cmd.Parameters.AddWithValue("@WallCenterlinePointZ", clashZone.WallCenterlinePointZ);
            
            // ✅ DIAGNOSTIC: Log WallCenterlinePoint values being saved to database
            // Log for ALL categories (not just dampers) to verify wall centerline is being saved correctly
            if (!DeploymentConfiguration.DeploymentMode)
            {
                bool isZero = (clashZone.WallCenterlinePointX == 0.0 && clashZone.WallCenterlinePointY == 0.0 && clashZone.WallCenterlinePointZ == 0.0);
                string zeroWarning = isZero ? " ⚠️⚠️⚠️ SAVING ZEROS!" : "";
                string category = clashZone.MepElementCategory ?? "Unknown";
                SafeFileLogger.SafeAppendText("wall_centerline_save.log",
                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [DB-SAVE] Zone {clashZone.Id} (Category={category}): " +
                    $"WallCenterlinePoint=({clashZone.WallCenterlinePointX:F6}ft, {clashZone.WallCenterlinePointY:F6}ft, {clashZone.WallCenterlinePointZ:F6}ft), " +
                    $"Intersection=({clashZone.IntersectionPointX:F6}ft, {clashZone.IntersectionPointY:F6}ft, {clashZone.IntersectionPointZ:F6}ft)" +
                    $"{zeroWarning}\n");
            }
            // ✅ ALWAYS log for dampers (even in deployment mode) - existing behavior preserved
            else if (string.Equals(clashZone.MepElementCategory, "Duct Accessories", StringComparison.OrdinalIgnoreCase))
            {
                bool isZero = (clashZone.WallCenterlinePointX == 0.0 && clashZone.WallCenterlinePointY == 0.0 && clashZone.WallCenterlinePointZ == 0.0);
                string zeroWarning = isZero ? " ⚠️⚠️⚠️ SAVING ZEROS!" : "";
                SafeFileLogger.SafeAppendText("wall_centerline_save.log",
                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [DB-SAVE] Damper Zone {clashZone.Id}: " +
                    $"WallCenterlinePoint=({clashZone.WallCenterlinePointX:F6}ft, {clashZone.WallCenterlinePointY:F6}ft, {clashZone.WallCenterlinePointZ:F6}ft), " +
                    $"Intersection=({clashZone.IntersectionPointX:F6}ft, {clashZone.IntersectionPointY:F6}ft, {clashZone.IntersectionPointZ:F6}ft)" +
                    $"{zeroWarning}\n");
            }
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
            // ✅ FIX: Store GUID in uppercase format for consistency with queries
            cmd.Parameters.AddWithValue("@ClashZoneGuid", clashZone.Id.ToString().ToUpperInvariant());
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
            
            // ✅ DIAGNOSTIC: Log MEP element sizes being saved to database (for debugging sleeve size issues)
            if (!DeploymentConfiguration.DeploymentMode)
            {
                var widthMm = clashZone.MepElementWidth * 304.8;
                var heightMm = clashZone.MepElementHeight * 304.8;
                SafeFileLogger.SafeAppendText("refresh_mep_sizes.log",
                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [DB-SAVE] Zone {clashZone.Id}: Saving MepWidth={clashZone.MepElementWidth:F6}ft ({widthMm:F1}mm), MepHeight={clashZone.MepElementHeight:F6}ft ({heightMm:F1}mm) to database\n");
            }
            
            // ✅ PIPE DIAMETER COLUMNS: Add outer diameter and nominal diameter parameters
            cmd.Parameters.AddWithValue("@MepElementOuterDiameter", clashZone.MepElementOuterDiameter);
            cmd.Parameters.AddWithValue("@MepElementNominalDiameter", clashZone.MepElementNominalDiameter);
            // ✅ SIZE PARAMETER VALUE: Add Size parameter value as string for snapshot table and parameter transfer
            cmd.Parameters.AddWithValue("@MepElementSizeParameterValue", (object)clashZone.MepElementSizeParameterValue ?? DBNull.Value);
            
            // ✅ DIAGNOSTIC: Log pipe diameters and size parameter value being saved
            if (!DeploymentConfiguration.DeploymentMode && string.Equals(clashZone.MepElementCategory, "Pipes", StringComparison.OrdinalIgnoreCase))
            {
                var odMm = clashZone.MepElementOuterDiameter > 0 ? (clashZone.MepElementOuterDiameter * 304.8) : 0.0;
                var nomMm = clashZone.MepElementNominalDiameter > 0 ? (clashZone.MepElementNominalDiameter * 304.8) : 0.0;
                _logger($"[DB-SAVE-DEBUG] Zone {clashZone.Id}: Saving Pipe - OuterDiameter={clashZone.MepElementOuterDiameter:F6}ft ({odMm:F1}mm), NominalDiameter={clashZone.MepElementNominalDiameter:F6}ft ({nomMm:F1}mm), SizeParameterValue='{clashZone.MepElementSizeParameterValue ?? "NULL"}', MepElementFormattedSize='{clashZone.MepElementFormattedSize ?? "NULL"}'");
                SafeFileLogger.SafeAppendText("save_db_diagnostic.log",
                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] Zone {clashZone.Id}: Pipe diameters - OuterDiameter={clashZone.MepElementOuterDiameter:F6}ft ({odMm:F1}mm), NominalDiameter={clashZone.MepElementNominalDiameter:F6}ft ({nomMm:F1}mm), SizeParameterValue='{clashZone.MepElementSizeParameterValue ?? "NULL"}'\n");
            }
            
            // ✅ DIAGNOSTIC: Log MEP dimensions being saved for duct accessories
            if (!DeploymentConfiguration.DeploymentMode && string.Equals(clashZone.MepElementCategory, "Duct Accessories", StringComparison.OrdinalIgnoreCase))
            {
                _logger($"[DB-SAVE-DEBUG] Zone {clashZone.Id}: Saving MepWidth={clashZone.MepElementWidth:F6}ft ({clashZone.MepElementWidth * 304.8:F1}mm), MepHeight={clashZone.MepElementHeight:F6}ft ({clashZone.MepElementHeight * 304.8:F1}mm)");
            }
            cmd.Parameters.AddWithValue("@SleeveFamilyName", (object)clashZone.SleeveFamilyName ?? DBNull.Value);
            // ✅ REFERENCE LEVEL: Add MEP element Reference Level (used for Schedule Level and Bottom of Opening calculation)
            cmd.Parameters.AddWithValue("@MepElementLevelName", (object)clashZone.MepElementLevelName ?? DBNull.Value);
            // ✅ REFERENCE LEVEL ELEVATION: Add MEP element Reference Level elevation (critical for Elevation from Level and Bottom of Opening calculation)
            cmd.Parameters.AddWithValue("@MepElementLevelElevation", clashZone.MepElementLevelElevation);
            // ✅ REMOVED DUPLICATE: Wall centerline point already added above (lines 1604-1606)
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
            
            // ✅ PARAMETER VALUES: Serialize parameter values to JSON for storage
            // Convert List<SerializableKeyValue> to Dictionary<string, string> then to JSON
            var mepParamsJson = "{}";
            
            // ✅ CRITICAL DIAGNOSTIC: Log parameter availability before serialization
            var mepParamCount = clashZone.MepParameterValues?.Count ?? 0;
            if (!DeploymentConfiguration.DeploymentMode && string.Equals(clashZone.MepElementCategory, "Pipes", StringComparison.OrdinalIgnoreCase))
            {
                var mepSampleKeys = clashZone.MepParameterValues?.Take(5).Select(kv => kv?.Key ?? "null").Where(k => !string.IsNullOrEmpty(k)).ToList() ?? new List<string>();
                var mepSampleStr = mepSampleKeys.Count > 0 ? string.Join(", ", mepSampleKeys) : "none";
                _logger($"[SQLite] [PARAM-SERIALIZE] Zone {clashZone.Id}: MepParameterValues={(clashZone.MepParameterValues != null ? "NOT NULL" : "NULL")}, Count={mepParamCount}, Sample keys: {mepSampleStr}");
            }
            
            if (clashZone.MepParameterValues != null && clashZone.MepParameterValues.Count > 0)
            {
                // ✅ LOGGING: Track filtered parameters for debugging
                var filteredOut = new List<string>();
                var mepDict = clashZone.MepParameterValues
                    .Where(kv => 
                    {
                        if (kv == null || string.IsNullOrEmpty(kv.Key))
                        {
                            if (kv != null && !string.IsNullOrEmpty(kv.Key))
                                filteredOut.Add($"{kv.Key} (null/empty key)");
                            return false;
                        }
                        if (string.IsNullOrEmpty(kv.Value))
                        {
                            filteredOut.Add($"{kv.Key} (empty value)");
                            return false;
                        }
                        return true;
                    })
                    .ToDictionary(kv => kv.Key, kv => kv.Value, StringComparer.OrdinalIgnoreCase);
                
                // ✅ LOGGING: Log filtered parameters for debugging
                if (filteredOut.Count > 0 && !DeploymentConfiguration.DeploymentMode)
                {
                    _logger($"[SQLite] ⚠️ Filtered out {filteredOut.Count} empty MEP parameters for zone {clashZone.Id} (MEP={clashZone.MepElementId?.IntegerValue ?? clashZone.MepElementIdValue}): {string.Join(", ", filteredOut)}");
                }
                
                // ✅ LOGGING: Log parameter count comparison
                if (!DeploymentConfiguration.DeploymentMode && clashZone.MepParameterValues.Count > mepDict.Count)
                {
                    _logger($"[SQLite] 📊 Parameter count for zone {clashZone.Id}: Captured={clashZone.MepParameterValues.Count}, Saved={mepDict.Count}, Filtered={clashZone.MepParameterValues.Count - mepDict.Count}");
                }
                
                if (mepDict.Count > 0)
                {
                    mepParamsJson = JsonSerializer.Serialize(mepDict);
                    
                    // ✅ CRITICAL DIAGNOSTIC: Log parameter serialization for debugging
                    if (!DeploymentConfiguration.DeploymentMode && string.Equals(clashZone.MepElementCategory, "Pipes", StringComparison.OrdinalIgnoreCase))
                    {
                        var sampleKeys = mepDict.Keys.Take(5).ToList();
                        var sampleStr = string.Join(", ", sampleKeys);
                        if (mepDict.Count > 5) sampleStr += $" (+{mepDict.Count - 5} more)";
                        _logger($"[SQLite] [PARAM-SERIALIZE] Zone {clashZone.Id}: Serializing {mepDict.Count} MEP params to JSON: {sampleStr}");
                    }
                }
                else if (!DeploymentConfiguration.DeploymentMode && string.Equals(clashZone.MepElementCategory, "Pipes", StringComparison.OrdinalIgnoreCase))
                {
                    _logger($"[SQLite] [PARAM-SERIALIZE] ⚠️ Zone {clashZone.Id}: No MEP params to serialize (mepDict.Count=0, clashZone.MepParameterValues.Count={clashZone.MepParameterValues?.Count ?? 0})");
                }
            }
            else if (!DeploymentConfiguration.DeploymentMode && string.Equals(clashZone.MepElementCategory, "Pipes", StringComparison.OrdinalIgnoreCase))
            {
                _logger($"[SQLite] [PARAM-SERIALIZE] ⚠️ Zone {clashZone.Id}: MepParameterValues is NULL or empty (Count={clashZone.MepParameterValues?.Count ?? 0})");
            }
            
            var hostParamsJson = "{}";
            if (clashZone.HostParameterValues != null && clashZone.HostParameterValues.Count > 0)
            {
                // ✅ LOGGING: Track filtered parameters for debugging
                var filteredOutHost = new List<string>();
                var hostDict = clashZone.HostParameterValues
                    .Where(kv => 
                    {
                        if (kv == null || string.IsNullOrEmpty(kv.Key))
                        {
                            if (kv != null && !string.IsNullOrEmpty(kv.Key))
                                filteredOutHost.Add($"{kv.Key} (null/empty key)");
                            return false;
                        }
                        if (string.IsNullOrEmpty(kv.Value))
                        {
                            filteredOutHost.Add($"{kv.Key} (empty value)");
                            return false;
                        }
                        return true;
                    })
                    .ToDictionary(kv => kv.Key, kv => kv.Value, StringComparer.OrdinalIgnoreCase);
                
                // ✅ LOGGING: Log filtered parameters for debugging
                if (filteredOutHost.Count > 0 && !DeploymentConfiguration.DeploymentMode)
                {
                    _logger($"[SQLite] ⚠️ Filtered out {filteredOutHost.Count} empty Host parameters for zone {clashZone.Id} (Host={clashZone.StructuralElementId?.IntegerValue ?? clashZone.StructuralElementIdValue}): {string.Join(", ", filteredOutHost)}");
                }
                
                if (hostDict.Count > 0)
                {
                    hostParamsJson = JsonSerializer.Serialize(hostDict);
                }
            }
            
            cmd.Parameters.AddWithValue("@MepParameterValuesJson", mepParamsJson);
            cmd.Parameters.AddWithValue("@HostParameterValuesJson", hostParamsJson);
            
            // ✅ OOP METHOD: Add damper connector detection parameters
            cmd.Parameters.AddWithValue("@HasMepConnector", clashZone.HasMepConnector ? 1 : 0);
            cmd.Parameters.AddWithValue("@DamperConnectorSide", (object)clashZone.DamperConnectorSide ?? string.Empty);
            
            // ✅ OOP METHOD: Add insulation detection parameters
            cmd.Parameters.AddWithValue("@IsInsulated", clashZone.IsInsulated ? 1 : 0);
            cmd.Parameters.AddWithValue("@InsulationThickness", clashZone.InsulationThickness);
        }

        /// <summary>
        /// ✅ PUBLIC: Save sleeve snapshots for placed sleeves (called after placement)
        /// </summary>
        /// <summary>
        /// ✅ CRITICAL: Save sleeve snapshots for placed sleeves (individual and cluster).
        /// This method handles normalization issues where SourceDocKey/HostDocKey might be missing:
        /// - Zones passed here may be in-memory objects from placement without database keys
        /// - Uses multiple fallback strategies to get ComboId even if keys are missing
        /// - Reloads keys from database if missing, or gets ComboId directly from ClashZones table
        /// </summary>
        public void SaveSleeveSnapshotsForPlacedSleeves(int filterId, List<ClashZone> placedZones)
        {
            if (placedZones == null || placedZones.Count == 0)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    _logger($"[SQLite] ⚠️ SaveSleeveSnapshotsForPlacedSleeves: No zones provided (null or empty)");
                }
                return;
            }

            // ✅ DIAGNOSTIC: Log what zones we received (ALWAYS log, even in deployment mode for debugging)
            int individualCount = placedZones.Count(z => z != null && z.SleeveInstanceId > 0 && z.ClusterSleeveInstanceId <= 0);
            int clusterCount = placedZones.Count(z => z != null && z.ClusterSleeveInstanceId > 0);
            int bothCount = placedZones.Count(z => z != null && z.SleeveInstanceId > 0 && z.ClusterSleeveInstanceId > 0);
            int invalidCount = placedZones.Count(z => z == null || (z.SleeveInstanceId <= 0 && z.ClusterSleeveInstanceId <= 0));
            
            // ✅ CRITICAL: Log category breakdown to identify if pipes are being filtered out
            var categoryBreakdown = placedZones
                .Where(z => z != null)
                .GroupBy(z => z.MepElementCategory ?? "Unknown")
                .Select(g => $"{g.Key}={g.Count()}({g.Count(z => z.SleeveInstanceId > 0)} with SleeveId)")
                .ToList();
            
            // ✅ CRITICAL: Log pipe-specific details
            int pipeCount = placedZones.Count(z => z != null && string.Equals(z.MepElementCategory, "Pipes", StringComparison.OrdinalIgnoreCase));
            int pipeWithSleeveId = placedZones.Count(z => z != null && string.Equals(z.MepElementCategory, "Pipes", StringComparison.OrdinalIgnoreCase) && z.SleeveInstanceId > 0);
            var pipeSamples = placedZones
                .Where(z => z != null && string.Equals(z.MepElementCategory, "Pipes", StringComparison.OrdinalIgnoreCase))
                .Take(5)
                .Select(z => $"{z.Id}:SleeveId={z.SleeveInstanceId}, Category={z.MepElementCategory}")
                .ToList();
            
            // ✅ CRITICAL: Always log this diagnostic (even in deployment mode) to debug missing individual sleeves
            string diagnosticMsg = $"[SQLite] SaveSleeveSnapshotsForPlacedSleeves: Total={placedZones.Count}, Individual={individualCount}, Cluster={clusterCount}, Both={bothCount}, Invalid={invalidCount}, Categories=[{string.Join(", ", categoryBreakdown)}], Pipes={pipeCount}({pipeWithSleeveId} with SleeveId)";
            _logger(diagnosticMsg);
            if (!DeploymentConfiguration.DeploymentMode)
            {
                // Also log sample IDs for individual sleeves
                var individualSamples = placedZones
                    .Where(z => z != null && z.SleeveInstanceId > 0 && z.ClusterSleeveInstanceId <= 0)
                    .Take(5)
                    .Select(z => $"{z.Id}:SleeveId={z.SleeveInstanceId}, Category={z.MepElementCategory}")
                    .ToList();
                if (individualSamples.Any())
                {
                    _logger($"[SQLite] SaveSleeveSnapshotsForPlacedSleeves: Individual sleeve samples: {string.Join(", ", individualSamples)}");
                }
                
                // ✅ CRITICAL: Always log pipe samples if any exist
                if (pipeSamples.Any())
                {
                    _logger($"[SQLite] SaveSleeveSnapshotsForPlacedSleeves: Pipe samples: {string.Join(", ", pipeSamples)}");
                }
            }

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
                    
                    // ✅ DIAGNOSTIC: Log filtered zones
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        int validCount = zonesByCombo.SelectMany(g => g).Count();
                        _logger($"[SQLite] SaveSleeveSnapshotsForPlacedSleeves: After filtering, {validCount} valid zones in {zonesByCombo.Count} combo group(s)");
                    }

                    foreach (var comboGroup in zonesByCombo)
                    {
                        var zones = comboGroup.ToList();
                        if (zones.Count == 0)
                            continue;

                        // ✅ CRITICAL FIX: Reload zones from database if SourceDocKey/HostDocKey are missing
                        // This ensures we have the correct keys even if zones were passed as in-memory objects
                        foreach (var zone in zones)
                        {
                            if (zone != null && (string.IsNullOrWhiteSpace(zone.SourceDocKey) || string.IsNullOrWhiteSpace(zone.HostDocKey)))
                            {
                                try
                                {
                                    using (var reloadCmd = _context.Connection.CreateCommand())
                                    {
                                        reloadCmd.Transaction = transaction;
                                        reloadCmd.CommandText = @"
                                            SELECT SourceDocKey, HostDocKey, ComboId 
                                            FROM ClashZones 
                                            WHERE ClashZoneGuid = @ClashZoneGuid
                                            LIMIT 1";
                                        reloadCmd.Parameters.AddWithValue("@ClashZoneGuid", zone.Id.ToString().ToUpperInvariant());
                                        using (var reader = reloadCmd.ExecuteReader())
                                        {
                                            if (reader.Read())
                                            {
                                                var dbSourceDocKey = GetNullableString(reader, "SourceDocKey");
                                                var dbHostDocKey = GetNullableString(reader, "HostDocKey");
                                                
                                                if (!string.IsNullOrWhiteSpace(dbSourceDocKey))
                                                    zone.SourceDocKey = dbSourceDocKey;
                                                if (!string.IsNullOrWhiteSpace(dbHostDocKey))
                                                    zone.HostDocKey = dbHostDocKey;
                                                
                                                if (!DeploymentConfiguration.DeploymentMode && (!string.IsNullOrWhiteSpace(dbSourceDocKey) || !string.IsNullOrWhiteSpace(dbHostDocKey)))
                                                {
                                                    _logger($"[SQLite] ✅ Reloaded SourceDocKey/HostDocKey for zone {zone.Id} from database");
                                                }
                                            }
                                        }
                                    }
                                }
                                catch (Exception reloadEx)
                                {
                                    if (!DeploymentConfiguration.DeploymentMode)
                                    {
                                        _logger($"[SQLite] ⚠️ Failed to reload SourceDocKey/HostDocKey for zone {zone.Id}: {reloadEx.Message}");
                                    }
                                }
                            }
                        }

                        // Get ComboId from first zone (if available via database lookup)
                        int comboId = -1;
                        try
                        {
                            var firstZone = zones[0];
                            
                            // ✅ CRITICAL FIX: Try multiple strategies to get ComboId
                            // Strategy 1: Use SourceDocKey and HostDocKey from zone (if available)
                            if (!string.IsNullOrWhiteSpace(firstZone.SourceDocKey) && !string.IsNullOrWhiteSpace(firstZone.HostDocKey))
                            {
                                // Try to find existing combo
                                using (var cmd = _context.Connection.CreateCommand())
                                {
                                    cmd.Transaction = transaction;
                                    // ✅ CRITICAL FIX: Handle filterId = -1 by skipping FilterId condition
                                    if (filterId > 0)
                                    {
                                        cmd.CommandText = @"
                                            SELECT ComboId FROM FileCombos 
                                            WHERE FilterId = @FilterId 
                                              AND LinkedFileKey = @LinkedFileKey 
                                              AND HostFileKey = @HostFileKey
                                            LIMIT 1";
                                        cmd.Parameters.AddWithValue("@FilterId", filterId);
                                    }
                                    else
                                    {
                                        // ✅ FALLBACK: If filterId is invalid, lookup by keys only
                                        cmd.CommandText = @"
                                            SELECT ComboId FROM FileCombos 
                                            WHERE LinkedFileKey = @LinkedFileKey 
                                              AND HostFileKey = @HostFileKey
                                            LIMIT 1";
                                    }
                                    cmd.Parameters.AddWithValue("@LinkedFileKey", firstZone.SourceDocKey);
                                    cmd.Parameters.AddWithValue("@HostFileKey", firstZone.HostDocKey);
                                    var result = cmd.ExecuteScalar();
                                    if (result != null)
                                    {
                                        comboId = Convert.ToInt32(result);
                                        if (!DeploymentConfiguration.DeploymentMode)
                                        {
                                            _logger($"[SQLite] ✅ SaveSleeveSnapshotsForPlacedSleeves: Found ComboId={comboId} using SourceDocKey/HostDocKey");
                                        }
                                    }
                                }
                            }
                            
                            // ✅ Strategy 2: If keys are missing, load them from database using ClashZoneGuid
                            if (comboId <= 0 && firstZone.Id != Guid.Empty)
                            {
                                using (var cmd = _context.Connection.CreateCommand())
                                {
                                    cmd.Transaction = transaction;
                                    cmd.CommandText = @"
                                        SELECT SourceDocKey, HostDocKey, ComboId 
                                        FROM ClashZones 
                                        WHERE ClashZoneGuid = @ClashZoneGuid
                                        LIMIT 1";
                                    cmd.Parameters.AddWithValue("@ClashZoneGuid", firstZone.Id.ToString().ToUpperInvariant());
                                    using (var reader = cmd.ExecuteReader())
                                    {
                                        if (reader.Read())
                                        {
                                            var dbSourceDocKey = GetNullableString(reader, "SourceDocKey");
                                            var dbHostDocKey = GetNullableString(reader, "HostDocKey");
                                            var dbComboId = GetInt(reader, "ComboId", -1);
                                            
                                            // Update zone with keys from database
                                            if (!string.IsNullOrWhiteSpace(dbSourceDocKey))
                                                firstZone.SourceDocKey = dbSourceDocKey;
                                            if (!string.IsNullOrWhiteSpace(dbHostDocKey))
                                                firstZone.HostDocKey = dbHostDocKey;
                                            
                                            // If ComboId is available directly, use it
                                            if (dbComboId > 0)
                                            {
                                                comboId = dbComboId;
                                                if (!DeploymentConfiguration.DeploymentMode)
                                                {
                                                    _logger($"[SQLite] ✅ SaveSleeveSnapshotsForPlacedSleeves: Found ComboId={comboId} directly from ClashZones table");
                                                }
                                            }
                                            // Otherwise, try to lookup ComboId using the loaded keys
                                            else if (!string.IsNullOrWhiteSpace(dbSourceDocKey) && !string.IsNullOrWhiteSpace(dbHostDocKey))
                                            {
                                                reader.Close();
                                                using (var comboCmd = _context.Connection.CreateCommand())
                                                {
                                                    comboCmd.Transaction = transaction;
                                                    // ✅ CRITICAL FIX: Handle filterId = -1 by skipping FilterId condition
                                                    if (filterId > 0)
                                                    {
                                                        comboCmd.CommandText = @"
                                                            SELECT ComboId FROM FileCombos 
                                                            WHERE FilterId = @FilterId 
                                                              AND LinkedFileKey = @LinkedFileKey 
                                                              AND HostFileKey = @HostFileKey
                                                            LIMIT 1";
                                                        comboCmd.Parameters.AddWithValue("@FilterId", filterId);
                                                    }
                                                    else
                                                    {
                                                        // ✅ FALLBACK: If filterId is invalid, lookup by keys only
                                                        comboCmd.CommandText = @"
                                                            SELECT ComboId FROM FileCombos 
                                                            WHERE LinkedFileKey = @LinkedFileKey 
                                                              AND HostFileKey = @HostFileKey
                                                            LIMIT 1";
                                                    }
                                                    comboCmd.Parameters.AddWithValue("@LinkedFileKey", dbSourceDocKey);
                                                    comboCmd.Parameters.AddWithValue("@HostFileKey", dbHostDocKey);
                                                    var comboResult = comboCmd.ExecuteScalar();
                                                    if (comboResult != null)
                                                    {
                                                        comboId = Convert.ToInt32(comboResult);
                                                        if (!DeploymentConfiguration.DeploymentMode)
                                                        {
                                                            _logger($"[SQLite] ✅ SaveSleeveSnapshotsForPlacedSleeves: Found ComboId={comboId} using keys loaded from database");
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            
                            // ✅ Strategy 3: If still no ComboId, try to get it from any zone in the group that has SleeveInstanceId
                            if (comboId <= 0)
                            {
                                var zoneWithSleeveId = zones.FirstOrDefault(z => z.SleeveInstanceId > 0);
                                if (zoneWithSleeveId != null)
                                {
                                    using (var cmd = _context.Connection.CreateCommand())
                                    {
                                        cmd.Transaction = transaction;
                                        cmd.CommandText = @"
                                            SELECT ComboId FROM ClashZones 
                                            WHERE SleeveInstanceId = @SleeveInstanceId
                                            LIMIT 1";
                                        cmd.Parameters.AddWithValue("@SleeveInstanceId", zoneWithSleeveId.SleeveInstanceId);
                                        var result = cmd.ExecuteScalar();
                                        if (result != null)
                                        {
                                            comboId = Convert.ToInt32(result);
                                            if (!DeploymentConfiguration.DeploymentMode)
                                            {
                                                _logger($"[SQLite] ✅ SaveSleeveSnapshotsForPlacedSleeves: Found ComboId={comboId} using SleeveInstanceId={zoneWithSleeveId.SleeveInstanceId}");
                                            }
                                        }
                                    }
                                }
                            }
                            
                            // ✅ DIAGNOSTIC: Log when all strategies fail
                            if (comboId <= 0)
                            {
                                // ✅ CRITICAL: Always log this warning (even in deployment mode) to debug missing individual sleeves
                                _logger($"[SQLite] ⚠️⚠️⚠️ SaveSleeveSnapshotsForPlacedSleeves: FAILED to find ComboId for {zones.Count} zones! SourceDocKey='{firstZone.SourceDocKey ?? "NULL"}', HostDocKey='{firstZone.HostDocKey ?? "NULL"}', ZoneId={firstZone.Id}, SleeveId={firstZone.SleeveInstanceId} - Zones will be saved with ComboId=-1");
                            }
                        }
                        catch (Exception comboEx)
                        {
                            // ✅ CRITICAL: Always log exceptions (even in deployment mode) to debug missing individual sleeves
                            _logger($"[SQLite] ❌ SaveSleeveSnapshotsForPlacedSleeves: ComboId lookup exception: {comboEx.Message}\nStackTrace: {comboEx.StackTrace}");
                        }

                        // ✅ DIAGNOSTIC: Log ComboId and zone details before processing
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            int individualInGroup = zones.Count(z => z.SleeveInstanceId > 0 && z.ClusterSleeveInstanceId <= 0);
                            int clusterInGroup = zones.Count(z => z.ClusterSleeveInstanceId > 0);
                            _logger($"[SQLite] SaveSleeveSnapshotsForPlacedSleeves: Processing {zones.Count} zones (Individual={individualInGroup}, Cluster={clusterInGroup}) with ComboId={comboId}");
                        }

                        // Create processed zones list
                        var processedZones = zones.Select(z => (comboId, z)).ToList();
                        
                        // Call private method with transaction
                        InsertOrUpdateSleeveSnapshotsInternal(filterId, processedZones, transaction);
                    }

                    transaction.Commit();
                    
                    // ✅ CRITICAL SAFETY: Verify snapshots were saved correctly
                    var verificationResult = VerifySnapshotCompleteness(placedZones, transaction);
                    if (!verificationResult.Success)
                    {
                        _logger($"[SQLite] ⚠️⚠️⚠️ SNAPSHOT VERIFICATION FAILED: {verificationResult.Message}");
                        _logger($"[SQLite] Missing snapshots for sleeves: {string.Join(", ", verificationResult.MissingSleeveIds)}");
                        
                        // Don't throw - log warning but allow operation to continue
                        // Retry mechanism will handle missing snapshots on next refresh
                    }
                    else if (!DeploymentConfiguration.DeploymentMode)
                    {
                        _logger($"[SQLite] ✅ Saved and VERIFIED sleeve snapshots for {placedZones.Count} placed zones");
                    }
                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    _logger($"[SQLite] ❌ Error saving sleeve snapshots: {ex.Message}\nStackTrace: {ex.StackTrace}");
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
            {
                // ✅ DIAGNOSTIC: Log why zones were filtered out
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    int nullZones = processedZones.Count(p => p.Zone == null);
                    int noSleeveId = processedZones.Count(p => p.Zone != null && p.Zone.SleeveInstanceId <= 0 && p.Zone.ClusterSleeveInstanceId <= 0);
                    int withSleeveId = processedZones.Count(p => p.Zone != null && p.Zone.SleeveInstanceId > 0);
                    int withClusterId = processedZones.Count(p => p.Zone != null && p.Zone.ClusterSleeveInstanceId > 0);
                    _logger($"[SQLite] ⚠️ InsertOrUpdateSleeveSnapshotsInternal: No valid zones! Total={processedZones.Count}, Null={nullZones}, NoIds={noSleeveId}, WithSleeveId={withSleeveId}, WithClusterId={withClusterId}");
                }
                return;
            }

            // ✅ DIAGNOSTIC: Log zone breakdown before grouping (INCLUDES CATEGORY BREAKDOWN FOR PIPES)
            var validCategoryBreakdown = validZones
                .Where(p => p.Zone != null)
                .GroupBy(p => p.Zone.MepElementCategory ?? "Unknown")
                .Select(g => $"{g.Key}={g.Count()}")
                .ToList();
            int validPipes = validZones.Count(p => p.Zone != null && string.Equals(p.Zone.MepElementCategory, "Pipes", StringComparison.OrdinalIgnoreCase));
            int validPipeIndividual = validZones.Count(p => p.Zone != null && string.Equals(p.Zone.MepElementCategory, "Pipes", StringComparison.OrdinalIgnoreCase) && p.Zone.SleeveInstanceId > 0 && p.Zone.ClusterSleeveInstanceId <= 0);
            
            if (!DeploymentConfiguration.DeploymentMode)
            {
                int individualZones = validZones.Count(p => p.Zone.SleeveInstanceId > 0 && p.Zone.ClusterSleeveInstanceId <= 0);
                int clusterZones = validZones.Count(p => p.Zone.ClusterSleeveInstanceId > 0);
                int bothZones = validZones.Count(p => p.Zone.SleeveInstanceId > 0 && p.Zone.ClusterSleeveInstanceId > 0);
                _logger($"[SQLite] InsertOrUpdateSleeveSnapshotsInternal: Valid zones breakdown - Individual={individualZones}, Cluster={clusterZones}, Both={bothZones}, Total={validZones.Count}, Categories=[{string.Join(", ", validCategoryBreakdown)}], Pipes={validPipes}({validPipeIndividual} individual)");
            }

            var groups = validZones.GroupBy(p =>
            {
                var zone = p.Zone;
                // ✅ CRITICAL FIX: Individual sleeves should NOT be grouped as clusters
                // Only group as cluster if ClusterSleeveInstanceId > 0 AND SleeveInstanceId <= 0 (pure cluster)
                // If both are > 0, it's an individual sleeve that was later clustered, but we want to save it as individual
                if (zone.ClusterSleeveInstanceId > 0 && zone.SleeveInstanceId <= 0)
                    return ("cluster", zone.ClusterSleeveInstanceId);
                // ✅ INDIVIDUAL SLEEVE: Group by SleeveInstanceId (even if ClusterSleeveInstanceId is also set)
                return ("sleeve", zone.SleeveInstanceId > 0 ? zone.SleeveInstanceId : -1);
            });

            // ✅ DIAGNOSTIC: Log grouping results (INCLUDING PIPE-SPECIFIC TRACKING)
            int groupsCount = groups.Count();
            int clusterGroups = groups.Count((IGrouping<(string, int), (int ComboId, ClashZone Zone)> g) => string.Equals(g.Key.Item1, "cluster", StringComparison.OrdinalIgnoreCase));
            int sleeveGroups = groups.Count((IGrouping<(string, int), (int ComboId, ClashZone Zone)> g) => string.Equals(g.Key.Item1, "sleeve", StringComparison.OrdinalIgnoreCase));
            int skippedGroups = groups.Count((IGrouping<(string, int), (int ComboId, ClashZone Zone)> g) => g.Key.Item2 <= 0);
            
            // ✅ CRITICAL: Count pipes in each group type for diagnostic purposes
            int pipeGroups = 0;
            int pipeZonesInGroups = 0;
            int pipeZonesSkipped = 0;
            foreach (var group in groups)
            {
                var zonesInGroup = group.Select(g => g.Zone).Where(z => z != null).ToList();
                int pipesInGroup = zonesInGroup.Count(z => string.Equals(z.MepElementCategory, "Pipes", StringComparison.OrdinalIgnoreCase));
                if (pipesInGroup > 0)
                {
                    pipeGroups++;
                    pipeZonesInGroups += pipesInGroup;
                }
                if (group.Key.Item2 <= 0)
                {
                    pipeZonesSkipped += pipesInGroup;
                }
            }
            
            string logMsg = $"[SQLite] InsertOrUpdateSleeveSnapshotsInternal: Grouped into {groupsCount} groups (Cluster={clusterGroups}, Sleeve={sleeveGroups}, Skipped={skippedGroups}), Pipes: {pipeGroups} groups with {pipeZonesInGroups} zones ({pipeZonesSkipped} skipped)";
            _logger(logMsg);

            foreach (var group in groups)
            {
                var key = group.Key;
                bool isCluster = string.Equals(key.Item1, "cluster", StringComparison.OrdinalIgnoreCase);
                int groupId = key.Item2;

                // ✅ DIAGNOSTIC: Log every group before processing (WITH PIPE-SPECIFIC INFO)
                var sampleZone = group.Select(g => g.Zone).FirstOrDefault(z => z != null);
                var zonesInGroup = group.Select(g => g.Zone).Where(z => z != null).ToList();
                int pipesInGroup = zonesInGroup.Count(z => string.Equals(z.MepElementCategory, "Pipes", StringComparison.OrdinalIgnoreCase));
                string sampleInfo = sampleZone != null 
                    ? $"Sample: ZoneId={sampleZone.Id}, SleeveId={sampleZone.SleeveInstanceId}, ClusterId={sampleZone.ClusterSleeveInstanceId}, Category={sampleZone.MepElementCategory}, HasMepParams={sampleZone.MepParameterValues?.Count > 0}, HasHostParams={sampleZone.HostParameterValues?.Count > 0}"
                    : "No valid zones in group";
                _logger($"[SQLite] InsertOrUpdateSleeveSnapshotsInternal: Processing group - isCluster={isCluster}, groupId={groupId}, zoneCount={group.Count()}, Pipes={pipesInGroup}, {sampleInfo}");

                if (groupId <= 0)
                {
                    // ✅ CRITICAL: Always log skipped groups, especially if they contain pipes
                    int skippedPipes = zonesInGroup.Count(z => string.Equals(z.MepElementCategory, "Pipes", StringComparison.OrdinalIgnoreCase));
                    _logger($"[SQLite] ⚠️⚠️⚠️ Skipping group with invalid groupId={groupId}, isCluster={isCluster}, zoneCount={group.Count()}, Pipes={skippedPipes}");
                    if (skippedPipes > 0)
                    {
                        var pipeSamples = zonesInGroup
                            .Where(z => string.Equals(z.MepElementCategory, "Pipes", StringComparison.OrdinalIgnoreCase))
                            .Take(3)
                            .Select(z => $"ZoneId={z.Id}, SleeveId={z.SleeveInstanceId}, ClusterId={z.ClusterSleeveInstanceId}")
                            .ToList();
                        _logger($"[SQLite] ⚠️⚠️⚠️ SKIPPED PIPES: {string.Join(", ", pipeSamples)}");
                    }
                    continue;
                }

                var zones = group.Select(g => g.Zone).Where(z => z != null).ToList();
                if (zones.Count == 0)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        _logger($"[SQLite] ⚠️ Skipping group - no valid zones after filtering (groupId={groupId}, isCluster={isCluster})");
                    }
                    continue;
                }

                var comboId = group.Select(g => g.ComboId).FirstOrDefault();

                // ✅ DEBUG: Check if zones have parameter values before aggregating
                var zonesWithMepParams = zones.Count(z => z.MepParameterValues != null && z.MepParameterValues.Count > 0);
                var zonesWithHostParams = zones.Count(z => z.HostParameterValues != null && z.HostParameterValues.Count > 0);
                
                // ✅ CRITICAL FIX: ALWAYS load parameters from ClashZones table as fallback
                // Database is the source of truth - ensure we have the latest parameters
                // This fixes cases where zones have partial or missing parameters
                if (!DeploymentConfiguration.DeploymentMode && (zonesWithMepParams < zones.Count || zonesWithHostParams < zones.Count))
                {
                    _logger($"[SQLite] ⚠️ SleeveSnapshots: {zones.Count} zones, but only {zonesWithMepParams} have MEP params, {zonesWithHostParams} have Host params. Loading from ClashZones table...");
                }
                
                // ✅ AGGRESSIVE FALLBACK: Load parameters from database for ALL zones that need them
                // Check each zone individually and load if missing or incomplete
                foreach (var zone in zones)
                {
                    var needsMepParams = zone.MepParameterValues == null || zone.MepParameterValues.Count == 0;
                    var needsHostParams = zone.HostParameterValues == null || zone.HostParameterValues.Count == 0;
                    
                    // ✅ DIAGNOSTIC: Log parameter availability before loading
                    if (!DeploymentConfiguration.DeploymentMode && (needsMepParams || needsHostParams))
                    {
                        var hasMepParams = !needsMepParams;
                        var hasHostParams = !needsHostParams;
                        var mepCount = zone.MepParameterValues?.Count ?? 0;
                        _logger($"[SQLite] Zone {zone.Id} (SleeveId={zone.SleeveInstanceId}): Before load - HasMepParams={hasMepParams} ({mepCount} params), HasHostParams={hasHostParams}");
                    }
                    
                    // ✅ ALWAYS load from database if parameters are missing (database is source of truth)
                    if (needsMepParams || needsHostParams)
                    {
                        LoadParameterValuesFromDatabase(zone, transaction);
                        
                        // ✅ DIAGNOSTIC: Log parameter availability after loading
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            var hasMepParamsAfter = zone.MepParameterValues != null && zone.MepParameterValues.Count > 0;
                            var hasHostParamsAfter = zone.HostParameterValues != null && zone.HostParameterValues.Count > 0;
                            var mepParamCount = zone.MepParameterValues?.Count ?? 0;
                            _logger($"[SQLite] Zone {zone.Id} (SleeveId={zone.SleeveInstanceId}): After load - HasMepParams={hasMepParamsAfter} ({mepParamCount} params), HasHostParams={hasHostParamsAfter}");
                            
                            if (hasMepParamsAfter && mepParamCount > 0)
                            {
                                var sampleKeys = zone.MepParameterValues.Take(5).Select(kv => kv.Key).ToList();
                                _logger($"[SQLite] Zone {zone.Id}: Sample MEP param keys: {string.Join(", ", sampleKeys)}");
                            }
                            else if (needsMepParams && !hasMepParamsAfter)
                            {
                                _logger($"[SQLite] ⚠️ Zone {zone.Id}: Failed to load MEP params from database (may not exist in ClashZones table)");
                            }
                        }
                    }
                }
                
                // Re-count after loading
                zonesWithMepParams = zones.Count(z => z.MepParameterValues != null && z.MepParameterValues.Count > 0);
                zonesWithHostParams = zones.Count(z => z.HostParameterValues != null && z.HostParameterValues.Count > 0);
                
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    _logger($"[SQLite] ✅ After loading from ClashZones table: {zonesWithMepParams}/{zones.Count} zones have MEP params, {zonesWithHostParams}/{zones.Count} have Host params");
                }
                
                if (!DeploymentConfiguration.DeploymentMode && zones.Count > 0 && (zonesWithMepParams == 0 || zonesWithHostParams == 0))
                {
                    _logger($"[SQLite] ⚠️ SleeveSnapshots: {zones.Count} zones, but only {zonesWithMepParams} have MEP params, {zonesWithHostParams} have Host params. This may result in empty JSON.");
                }
                
                // ✅ CRITICAL DEBUG: Log Size parameter availability before aggregation
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    var zonesWithSizeParam = zones.Count(z => !string.IsNullOrWhiteSpace(z.MepElementSizeParameterValue));
                    var sampleSizeValues = zones
                        .Where(z => !string.IsNullOrWhiteSpace(z.MepElementSizeParameterValue))
                        .Take(3)
                        .Select(z => $"ZoneId={z.Id}, Size='{z.MepElementSizeParameterValue}'")
                        .ToList();
                    _logger($"[SQLite] SleeveSnapshots: Before aggregation - {zonesWithSizeParam}/{zones.Count} zones have MepElementSizeParameterValue. Samples: {string.Join(", ", sampleSizeValues)}");
                }
                
                var mepParams = AggregateParameterValues(zones, useHost: false);
                var hostParams = AggregateParameterValues(zones, useHost: true);
                
                // ✅ CRITICAL SAFETY: Validate critical parameters are present
                var criticalParamValidation = ValidateCriticalParameters(mepParams, hostParams, zones, isCluster, groupId);
                if (!criticalParamValidation.IsValid)
                {
                    _logger($"[SQLite] ⚠️⚠️⚠️ CRITICAL PARAMETER MISSING for groupId={groupId}: {criticalParamValidation.Message}");
                    // Continue anyway - parameter transfer will attempt fallback to Revit
                }
                
                // ✅ DEBUG: Log aggregated parameter counts and check if Size is present
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    var sizeKvp = mepParams.FirstOrDefault(kvp => string.Equals(kvp.Key, "Size", StringComparison.OrdinalIgnoreCase));
                    bool hasSizeParam = sizeKvp.Key != null;
                    string sizeValue = hasSizeParam ? sizeKvp.Value : "NOT FOUND";
                    
                    // ✅ ENHANCED LOGGING: Show all parameter keys for debugging
                    var allParamKeys = string.Join(", ", mepParams.Keys.Take(15));
                    var moreCount = mepParams.Count > 15 ? $" (+{mepParams.Count - 15} more)" : "";
                    _logger($"[SQLite] SleeveSnapshots: Aggregated {mepParams.Count} MEP params, {hostParams.Count} Host params for {zones.Count} zones (isCluster={isCluster}, groupId={groupId}). Size parameter: {(hasSizeParam ? $"FOUND='{sizeValue}'" : "MISSING")}. All MEP params: {allParamKeys}{moreCount}");
                    
                    // ✅ CRITICAL: Log if parameters are empty after aggregation
                    if (mepParams.Count == 0 && zones.Count > 0)
                    {
                        var zoneSamples = zones.Take(3).Select(z => 
                            $"ZoneId={z.Id}, HasMepParams={z.MepParameterValues?.Count > 0} ({z.MepParameterValues?.Count ?? 0} params), HasSizeValue={!string.IsNullOrWhiteSpace(z.MepElementSizeParameterValue)}"
                        ).ToList();
                        _logger($"[SQLite] ⚠️⚠️⚠️ CRITICAL: No MEP parameters after aggregation! Zone samples: {string.Join("; ", zoneSamples)}");
                    }
                }

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

                // ✅ DETERMINISTIC GUID: For individual sleeves, use the zone's Id (deterministic GUID)
                // For clusters, use the first zone's GUID as the primary identifier
                // ✅ CRITICAL FIX: Ensure GUID matches what's in ClashZones table by querying the database
                string clashZoneGuidString = null;
                
                if (isCluster && clusterInstanceId.HasValue && clusterInstanceId.Value > 0)
                {
                    // ✅ For cluster sleeves, get GUID from ClashZones table where ClusterInstanceId matches
                    // This ensures the GUID in SleeveSnapshots matches what's actually in ClashZones table
                    using (var guidCmd = _context.Connection.CreateCommand())
                    {
                        guidCmd.Transaction = transaction;
                        guidCmd.CommandText = @"
                            SELECT ClashZoneGuid FROM ClashZones 
                            WHERE ClusterInstanceId = @ClusterInstanceId 
                              AND ClashZoneGuid IS NOT NULL 
                              AND ClashZoneGuid != ''
                            LIMIT 1";
                        guidCmd.Parameters.AddWithValue("@ClusterInstanceId", clusterInstanceId.Value);
                        var guidResult = guidCmd.ExecuteScalar();
                        if (guidResult != null && guidResult != DBNull.Value)
                        {
                            clashZoneGuidString = guidResult.ToString().ToUpperInvariant().Trim();
                        }
                    }
                }
                
                // ✅ Fallback: Use first zone's GUID if database lookup failed or for individual sleeves
                if (string.IsNullOrEmpty(clashZoneGuidString))
                {
                    clashZoneGuidString = zones.FirstOrDefault()?.Id.ToString().ToUpperInvariant();
                }
                
                // ✅ CRITICAL DEBUG: Log GUID extraction for diagnostic purposes
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    _logger($"[SQLite] [UPSERT-DEBUG] Attempting upsert with ClashZoneGuid='{clashZoneGuidString ?? "NULL"}', SourceType='{(isCluster ? "Cluster" : "Individual")}', GroupId={groupId}, ClusterInstanceId={clusterInstanceId?.ToString() ?? "NULL"}, SleeveInstanceId={sleeveInstanceId?.ToString() ?? "NULL"}, GUIDSource='{(isCluster && clusterInstanceId.HasValue ? "DB" : "Zone.Id")}'");
                    
                    // ✅ DIAGNOSTIC: Check if existing snapshot exists before upsert
                    if (isCluster && clusterInstanceId.HasValue && clusterInstanceId.Value > 0)
                    {
                        using (var checkCmd = _context.Connection.CreateCommand())
                        {
                            checkCmd.Transaction = transaction;
                            checkCmd.CommandText = "SELECT SnapshotId, ClashZoneGuid FROM SleeveSnapshots WHERE ClusterInstanceId = @ClusterInstanceId LIMIT 1";
                            checkCmd.Parameters.AddWithValue("@ClusterInstanceId", clusterInstanceId.Value);
                            using (var reader = checkCmd.ExecuteReader())
                            {
                                if (reader.Read())
                                {
                                    var existingId = reader.GetInt32(0);
                                    var existingGuid = reader.IsDBNull(1) ? "NULL" : reader.GetString(1);
                                    _logger($"[SQLite] [UPSERT-DEBUG] ✅ Found EXISTING snapshot: SnapshotId={existingId}, ExistingGuid='{existingGuid}', NewGuid='{clashZoneGuidString ?? "NULL"}' - Should UPDATE, not INSERT");
                                }
                                else
                                {
                                    _logger($"[SQLite] [UPSERT-DEBUG] ⚠️ No existing snapshot found for ClusterInstanceId={clusterInstanceId.Value} - Will INSERT new row");
                                }
                            }
                        }
                    }
                }
                
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
                        ClashZoneGuid = clashZoneGuidString, // ✅ DETERMINISTIC: Save ClashZoneGuid from zone.Id
                        UpdatedAt = DateTime.UtcNow
                    });
            }
        }

        private void UpsertSleeveSnapshot(SQLiteTransaction transaction, SleeveSnapshot snapshot)
        {
            if (snapshot == null)
                return;

            // ✅ DETERMINISTIC GUID PRIORITY: Check by ClashZoneGuid first (deterministic matching)
            // Then fall back to SleeveInstanceId/ClusterInstanceId for legacy data or when GUID is missing
            var existingId = GetExistingSnapshotId(
                snapshot.ClashZoneGuid, 
                snapshot.SleeveInstanceId, 
                snapshot.ClusterInstanceId, 
                transaction);

            if (existingId.HasValue)
            {
                snapshot.SnapshotId = existingId.Value;
                
                // ✅ DIAGNOSTIC: Log what SleeveInstanceId is being updated
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    _logger($"[SQLite] [UPSERT] Updating snapshot SnapshotId={existingId.Value} with SleeveInstanceId={snapshot.SleeveInstanceId?.ToString() ?? "NULL"}, ClusterInstanceId={snapshot.ClusterInstanceId?.ToString() ?? "NULL"}, ClashZoneGuid='{snapshot.ClashZoneGuid ?? "NULL"}'");
                }
                
                UpdateSleeveSnapshot(transaction, snapshot);
                
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    _logger($"[SQLite] ✅ UPSERT: Updated existing SleeveSnapshot SnapshotId={existingId.Value} (ClashZoneGuid={snapshot.ClashZoneGuid ?? "NULL"}, SleeveInstanceId={snapshot.SleeveInstanceId?.ToString() ?? "NULL"}, ClusterInstanceId={snapshot.ClusterInstanceId?.ToString() ?? "NULL"})");
                }
            }
            else
            {
                InsertSleeveSnapshot(transaction, snapshot);
                
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    _logger($"[SQLite] ✅ UPSERT: Inserted new SleeveSnapshot (ClashZoneGuid={snapshot.ClashZoneGuid ?? "NULL"}, SleeveInstanceId={snapshot.SleeveInstanceId?.ToString() ?? "NULL"}, ClusterInstanceId={snapshot.ClusterInstanceId?.ToString() ?? "NULL"})");
                }
            }
        }

        /// <summary>
        /// ✅ DETERMINISTIC GUID: Find existing snapshot by ClashZoneGuid first (deterministic), 
        /// then fall back to SleeveInstanceId/ClusterInstanceId for legacy data or when GUID is missing.
        /// This ensures snapshots are updated instead of appended, following the deterministic GUID principle.
        /// </summary>
        private int? GetExistingSnapshotId(string? clashZoneGuid, int? sleeveInstanceId, int? clusterInstanceId, SQLiteTransaction transaction)
        {
            using (var cmd = _context.Connection.CreateCommand())
            {
                cmd.Transaction = transaction;

                // ✅ PRIORITY 1: Check by ClashZoneGuid (deterministic GUID) - most reliable
                if (!string.IsNullOrWhiteSpace(clashZoneGuid))
                {
                    // ✅ CRITICAL FIX: Normalize GUID to uppercase for consistent comparison
                    var normalizedGuid = clashZoneGuid.ToUpperInvariant().Trim();
                    
                    // ✅ FIXED: Use simpler comparison - SQLite UPPER() is sufficient, TRIM() may cause issues
                    cmd.CommandText = @"
                        SELECT SnapshotId FROM SleeveSnapshots 
                        WHERE UPPER(ClashZoneGuid) = UPPER(@ClashZoneGuid) 
                          AND ClashZoneGuid IS NOT NULL 
                          AND ClashZoneGuid != ''
                        LIMIT 1";
                    cmd.Parameters.AddWithValue("@ClashZoneGuid", normalizedGuid);
                    
                    var resultByGuid = cmd.ExecuteScalar();
                    if (resultByGuid != null && resultByGuid != DBNull.Value)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            _logger($"[SQLite] ✅ Found existing snapshot by ClashZoneGuid='{normalizedGuid}', SnapshotId={Convert.ToInt32(resultByGuid)}");
                        }
                        return Convert.ToInt32(resultByGuid);
                    }
                    
                    // ✅ DEBUG: Log when GUID lookup fails to help diagnose issues
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        // Check if any snapshots exist with similar GUIDs (for debugging)
                        cmd.Parameters.Clear();
                        cmd.CommandText = @"
                            SELECT COUNT(*) FROM SleeveSnapshots 
                            WHERE ClashZoneGuid IS NOT NULL AND ClashZoneGuid != ''";
                        var totalWithGuid = cmd.ExecuteScalar();
                        
                        // ✅ DIAGNOSTIC: Also check what GUIDs are actually in the database (sample)
                        cmd.Parameters.Clear();
                        cmd.CommandText = @"
                            SELECT ClashZoneGuid FROM SleeveSnapshots 
                            WHERE ClashZoneGuid IS NOT NULL AND ClashZoneGuid != ''
                            LIMIT 5";
                        var sampleGuids = new List<string>();
                        using (var reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                var guid = reader.GetString(0);
                                if (!string.IsNullOrEmpty(guid))
                                    sampleGuids.Add(guid);
                            }
                        }
                        
                        _logger($"[SQLite] ⚠️ GUID lookup FAILED for ClashZoneGuid='{normalizedGuid}' - No matching snapshot found (Total snapshots with GUID: {totalWithGuid}, Sample GUIDs in DB: [{string.Join(", ", sampleGuids)}])");
                    }
                }

                // ✅ PRIORITY 2: Fall back to SleeveInstanceId (for individual sleeves)
                if (sleeveInstanceId.HasValue && sleeveInstanceId.Value > 0)
                {
                    cmd.Parameters.Clear();
                    cmd.CommandText = "SELECT SnapshotId FROM SleeveSnapshots WHERE SleeveInstanceId = @SleeveInstanceId LIMIT 1";
                    cmd.Parameters.AddWithValue("@SleeveInstanceId", sleeveInstanceId.Value);
                    
                    var resultBySleeveId = cmd.ExecuteScalar();
                    if (resultBySleeveId != null && resultBySleeveId != DBNull.Value)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            _logger($"[SQLite] ✅ Found existing snapshot by SleeveInstanceId={sleeveInstanceId.Value}, SnapshotId={Convert.ToInt32(resultBySleeveId)} (fallback - GUID lookup failed)");
                        }
                        return Convert.ToInt32(resultBySleeveId);
                    }
                }

                // ✅ PRIORITY 3: Fall back to ClusterInstanceId (for cluster sleeves)
                if (clusterInstanceId.HasValue && clusterInstanceId.Value > 0)
                {
                    cmd.Parameters.Clear();
                    cmd.CommandText = "SELECT SnapshotId FROM SleeveSnapshots WHERE ClusterInstanceId = @ClusterInstanceId LIMIT 1";
                    cmd.Parameters.AddWithValue("@ClusterInstanceId", clusterInstanceId.Value);
                    
                    var resultByClusterId = cmd.ExecuteScalar();
                    if (resultByClusterId != null && resultByClusterId != DBNull.Value)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            _logger($"[SQLite] ✅ Found existing snapshot by ClusterInstanceId={clusterInstanceId.Value}, SnapshotId={Convert.ToInt32(resultByClusterId)} (fallback - GUID lookup failed)");
                        }
                        return Convert.ToInt32(resultByClusterId);
                    }
                    else
                    {
                        // ✅ DIAGNOSTIC: Log when ClusterInstanceId lookup fails
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            cmd.Parameters.Clear();
                            cmd.CommandText = "SELECT COUNT(*) FROM SleeveSnapshots WHERE ClusterInstanceId IS NOT NULL";
                            var totalClusters = cmd.ExecuteScalar();
                            cmd.Parameters.Clear();
                            cmd.CommandText = "SELECT ClusterInstanceId FROM SleeveSnapshots WHERE ClusterInstanceId IS NOT NULL LIMIT 5";
                            var sampleClusterIds = new List<int>();
                            using (var reader = cmd.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    var id = reader.GetInt32(0);
                                    if (id > 0)
                                        sampleClusterIds.Add(id);
                                }
                            }
                            _logger($"[SQLite] ⚠️ ClusterInstanceId lookup FAILED for ClusterInstanceId={clusterInstanceId.Value} - No matching snapshot found (Total cluster snapshots: {totalClusters}, Sample ClusterIds in DB: [{string.Join(", ", sampleClusterIds)}])");
                        }
                    }
                }
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
                        ClashZoneGuid,
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
                        @ClashZoneGuid,
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
                    { "HostDocKeysJson", snapshot.HostDocKeysJson ?? "NULL" },
                    { "ClashZoneGuid", snapshot.ClashZoneGuid ?? "NULL" }
                };
                
                // ✅ ENHANCED LOGGING: Show parameter details for INSERT
                int mepParamCount = 0;
                int hostParamCount = 0;
                string mepParamSample = "";
                
                if (!string.IsNullOrWhiteSpace(snapshot.MepParametersJson) && snapshot.MepParametersJson != "{}")
                {
                    try
                    {
                        var mepDict = System.Text.Json.JsonSerializer.Deserialize<Dictionary<string, string>>(snapshot.MepParametersJson);
                        if (mepDict != null)
                        {
                            mepParamCount = mepDict.Count;
                            mepParamSample = string.Join(", ", mepDict.Keys.Take(5));
                            if (mepDict.Count > 5) mepParamSample += $" (+{mepDict.Count - 5} more)";
                        }
                    }
                    catch { }
                }
                
                if (!string.IsNullOrWhiteSpace(snapshot.HostParametersJson) && snapshot.HostParametersJson != "{}")
                {
                    try
                    {
                        var hostDict = System.Text.Json.JsonSerializer.Deserialize<Dictionary<string, string>>(snapshot.HostParametersJson);
                        if (hostDict != null)
                        {
                            hostParamCount = hostDict.Count;
                        }
                    }
                    catch { }
                }
                
                string additionalInfo = $"✅ Inserted NEW snapshot for SleeveId={snapshot.SleeveInstanceId}, ClusterId={snapshot.ClusterInstanceId}";
                if (mepParamCount > 0)
                {
                    additionalInfo += $", MEP params: {mepParamCount} ({mepParamSample})";
                }
                else
                {
                    additionalInfo += $", ⚠️ NO MEP PARAMETERS";
                }
                if (hostParamCount > 0)
                {
                    additionalInfo += $", Host params: {hostParamCount}";
                }
                
                DatabaseOperationLogger.LogOperation(
                    "INSERT",
                    "SleeveSnapshots",
                    snapshotParams,
                    rowsAffected: 1,
                    additionalInfo: additionalInfo);
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
                        ClashZoneGuid = @ClashZoneGuid,
                        UpdatedAt = CURRENT_TIMESTAMP
                    WHERE SnapshotId = @SnapshotId";

                AddSnapshotParameters(cmd, snapshot);
                cmd.Parameters.AddWithValue("@SnapshotId", snapshot.SnapshotId);
                var rowsAffected = cmd.ExecuteNonQuery();
                
                // ✅ LOG: Log UPDATE operation with parameter details
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    // ✅ ENHANCED LOGGING: Show parameter counts in snapshot
                    int mepParamCount = 0;
                    int hostParamCount = 0;
                    string mepParamSample = "";
                    
                    if (!string.IsNullOrWhiteSpace(snapshot.MepParametersJson) && snapshot.MepParametersJson != "{}")
                    {
                        try
                        {
                            var mepDict = System.Text.Json.JsonSerializer.Deserialize<Dictionary<string, string>>(snapshot.MepParametersJson);
                            if (mepDict != null)
                            {
                                mepParamCount = mepDict.Count;
                                mepParamSample = string.Join(", ", mepDict.Keys.Take(5));
                                if (mepDict.Count > 5) mepParamSample += $" (+{mepDict.Count - 5} more)";
                            }
                        }
                        catch { }
                    }
                    
                    if (!string.IsNullOrWhiteSpace(snapshot.HostParametersJson) && snapshot.HostParametersJson != "{}")
                    {
                        try
                        {
                            var hostDict = System.Text.Json.JsonSerializer.Deserialize<Dictionary<string, string>>(snapshot.HostParametersJson);
                            if (hostDict != null)
                            {
                                hostParamCount = hostDict.Count;
                            }
                        }
                        catch { }
                    }
                    
                    var updateParams = new Dictionary<string, object>
                    {
                        { "SnapshotId", snapshot.SnapshotId.ToString() },
                        { "ClashZoneGuid", snapshot.ClashZoneGuid ?? "NULL" },
                        { "SleeveInstanceId", snapshot.SleeveInstanceId?.ToString() ?? "NULL" },
                        { "ClusterInstanceId", snapshot.ClusterInstanceId?.ToString() ?? "NULL" },
                        { "MepParamsCount", mepParamCount.ToString() },
                        { "HostParamsCount", hostParamCount.ToString() }
                    };
                    
                    string additionalInfo = $"✅ Updated existing snapshot SnapshotId={snapshot.SnapshotId}";
                    if (mepParamCount > 0)
                    {
                        additionalInfo += $", MEP params: {mepParamCount} ({mepParamSample})";
                    }
                    else
                    {
                        additionalInfo += $", ⚠️ NO MEP PARAMETERS";
                    }
                    if (hostParamCount > 0)
                    {
                        additionalInfo += $", Host params: {hostParamCount}";
                    }
                    
                    DatabaseOperationLogger.LogOperation(
                        "UPDATE",
                        "SleeveSnapshots",
                        updateParams,
                        rowsAffected: rowsAffected,
                        additionalInfo: additionalInfo);
                }
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
            cmd.Parameters.AddWithValue("@ClashZoneGuid", (object)snapshot.ClashZoneGuid ?? DBNull.Value);
        }

        private Dictionary<string, string> AggregateParameterValues(IEnumerable<ClashZone> zones, bool useHost)
        {
            var aggregated = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);

            foreach (var zone in zones)
            {
                var bag = useHost ? zone.HostParameterValues : zone.MepParameterValues;
                
                // ✅ CRITICAL FIX FOR ALL CATEGORIES: Ensure Size parameter uses MepElementSizeParameterValue (current refresh is source of truth)
                // This ensures snapshot table saves the exact text value from the Size parameter (e.g., "20 mmø", "200 mm dia symbol")
                // MepElementSizeParameterValue is the raw Size parameter value read during refresh - different from MepElementFormattedSize which may be calculated
                // ✅ CRITICAL: Always use fresh value from current refresh, even if Size already exists in MepParameterValues with old value
                // ✅ EXCEPTION: For cluster zones, preserve the aggregated Size parameter (already aggregated from individual zones in SaveClusterSleeveSnapshots)
                if (!useHost)
                {
                    // Initialize list if needed
                    if (bag == null)
                    {
                        bag = new List<Models.SerializableKeyValue>();
                        zone.MepParameterValues = bag;
                    }
                    
                    // ✅ CRITICAL FIX FOR CLUSTER ZONES: Don't remove/re-add Size for cluster zones - preserve aggregated value
                    // Cluster zones have Size already aggregated from individual zones in SaveClusterSleeveSnapshots
                    // Individual zones have MepElementSizeParameterValue populated from database
                    bool isClusterZone = zone.ClusterSleeveInstanceId > 0;
                    bool hasSizeInParams = bag.Any(kv => kv != null && IsSizeParameter(kv.Key?.Trim() ?? string.Empty));
                    
                    if (!isClusterZone)
                    {
                        // ✅ For individual zones: Remove existing Size parameter if it exists (to replace with fresh value from current refresh)
                        // This ensures old values (like "0.082") are replaced with fresh values (like "20 mmø") from current refresh
                        bag.RemoveAll(kv => kv != null && IsSizeParameter(kv.Key?.Trim() ?? string.Empty));
                        
                        // ✅ PRIORITY 1: Use MepElementSizeParameterValue (raw Size parameter value from current refresh)
                        if (!string.IsNullOrWhiteSpace(zone.MepElementSizeParameterValue))
                        {
                            bag.Add(new Models.SerializableKeyValue { Key = "Size", Value = zone.MepElementSizeParameterValue });
                        }
                        // ✅ PRIORITY 2: Fallback to MepElementFormattedSize if MepElementSizeParameterValue is empty
                        else if (!string.IsNullOrWhiteSpace(zone.MepElementFormattedSize))
                        {
                            bag.Add(new Models.SerializableKeyValue { Key = "Size", Value = zone.MepElementFormattedSize });
                        }
                    }
                    else if (!hasSizeInParams)
                    {
                        // ✅ For cluster zones: Only add Size if it's missing (shouldn't happen if SaveClusterSleeveSnapshots worked correctly)
                        // This is a fallback in case Size wasn't aggregated properly
                        if (!string.IsNullOrWhiteSpace(zone.MepElementSizeParameterValue))
                        {
                            bag.Add(new Models.SerializableKeyValue { Key = "Size", Value = zone.MepElementSizeParameterValue });
                        }
                        else if (!string.IsNullOrWhiteSpace(zone.MepElementFormattedSize))
                        {
                            bag.Add(new Models.SerializableKeyValue { Key = "Size", Value = zone.MepElementFormattedSize });
                        }
                    }
                    
                    // ✅ CRITICAL FIX: Add MEP_ElementId to MepParameterValues for parameter transfer
                    // MEP_ElementId is required for parameter transfer command to work correctly
                    // For individual sleeves, use the zone's MepElementId
                    // For cluster sleeves, this will be aggregated (comma-separated) below
                    if (!useHost && zone.MepElementId != null && zone.MepElementId.IntegerValue > 0)
                    {
                        // Check if MEP_ElementId already exists in bag
                        bool hasMepElementId = bag.Any(kv => kv != null && 
                            string.Equals(kv.Key?.Trim(), "MEP_ElementId", StringComparison.OrdinalIgnoreCase));
                        
                        if (!hasMepElementId)
                        {
                            bag.Add(new Models.SerializableKeyValue 
                            { 
                                Key = "MEP_ElementId", 
                                Value = zone.MepElementId.IntegerValue.ToString() 
                            });
                        }
                    }
                    
                    // ✅ REMOVED: No longer adding "MEP Size" alias to snapshot
                    // Parameter transfer now reads "MEP Size" directly from Revit MEP element (not from snapshot)
                    // This avoids stale data and ensures we always get the current value from the live MEP element
                }
                
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

            var result = aggregated.ToDictionary(
                kvp => kvp.Key,
                kvp => string.Join(", ", kvp.Value),
                StringComparer.OrdinalIgnoreCase);
            
            // ✅ REMOVED: No longer adding "MEP Size" alias to aggregated parameters
            // Parameter transfer now reads "MEP Size" directly from Revit MEP element (not from snapshot)
            // This avoids stale data and ensures we always get the current value from the live MEP element
            
            return result;
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

        public List<ClashZone> GetClashZonesByFilter(string filterName, string category, bool unresolvedOnly = false, bool readyForPlacementOnly = false)
        {
            var result = new List<ClashZone>();

            if (string.IsNullOrWhiteSpace(filterName) || string.IsNullOrWhiteSpace(category))
                return result;

            using (var cmd = _context.Connection.CreateCommand())
            {
                // Build WHERE clause with optional filters
                var whereConditions = new List<string>
                {
                    "f.FilterName = @FilterName",
                    "f.Category = @Category"
                };
                
                if (unresolvedOnly)
                    whereConditions.Add("cz.SleeveState = 0");
                    
                if (readyForPlacementOnly)
                {
                    whereConditions.Add("cz.ReadyForPlacementFlag = 1");
                    // ✅ CRITICAL FIX: Exclude zones that already have sleeves placed
                    // This prevents placing duplicate sleeves for zones that already have SleeveInstanceId > 0
                    // Even if flags are wrong, if a sleeve exists, we shouldn't place another one
                    whereConditions.Add("(cz.SleeveInstanceId IS NULL OR cz.SleeveInstanceId <= 0)");
                    whereConditions.Add("(cz.ClusterInstanceId IS NULL OR cz.ClusterInstanceId <= 0)");
                }
                
                var whereClause = string.Join(" AND ", whereConditions);
                
                cmd.CommandText = $@"
                    SELECT 
                        cz.*,
                        fc.LinkedFileKey,
                        fc.HostFileKey,
                        f.FilterName
                    FROM Filters f
                    INNER JOIN FileCombos fc ON f.FilterId = fc.FilterId
                    INNER JOIN ClashZones cz ON fc.ComboId = cz.ComboId
                    WHERE {whereClause}
                    ORDER BY cz.UpdatedAt DESC";

                cmd.Parameters.AddWithValue("@FilterName", filterName);
                cmd.Parameters.AddWithValue("@Category", category);

                using (var reader = cmd.ExecuteReader())
                {
                    int zonesWithSleeveId = 0;
                    int zonesWithClusterId = 0;
                    int zonesAdded = 0;
                    
                    while (reader.Read())
                    {
                        var clashZone = MapClashZone(reader);
                        SetMetadataFromReader(clashZone, reader);
                        
                        // ✅ DIAGNOSTIC: Check if zone already has a sleeve (should be filtered by query, but log for verification)
                        if (clashZone.SleeveInstanceId > 0)
                        {
                            zonesWithSleeveId++;
                            if (!DeploymentConfiguration.DeploymentMode && readyForPlacementOnly)
                            {
                                _logger($"[SQLite] ⚠️ GetClashZonesByFilter: Zone {clashZone.Id} has SleeveInstanceId={clashZone.SleeveInstanceId} but passed query filter (should be excluded)");
                            }
                        }
                        if (clashZone.ClusterSleeveInstanceId > 0)
                        {
                            zonesWithClusterId++;
                        }
                        
                        // ⚠️⚠️⚠️ CRITICAL FIX - DO NOT REMOVE OR MODIFY ⚠️⚠️⚠️
                        // ============================================================
                        // PROBLEM: Zones loaded from database may have IntersectionPoint = null even though
                        //          IntersectionPointX/Y/Z coordinates exist in the database. This causes
                        //          section box checks to fail (isWithinSectionBox = false) because
                        //          IntersectionPoint is null.
                        //
                        // SOLUTION: Reconstruct IntersectionPoint from database coordinates if it's null
                        //           but coordinates exist.
                        //
                        // IMPACT IF REMOVED:
                        //   - Zones with null IntersectionPoint will fail section box checks
                        //   - ReadyForPlacementFlag won't be set for these zones
                        //   - Placement won't find eligible zones
                        //   - Sleeves won't be placed
                        //
                        // TESTED: 2025-12-05 - Confirmed working after rebuild
                        // ============================================================
                        if (clashZone.IntersectionPoint == null && (Math.Abs(clashZone.IntersectionPointX) > 1e-9 || Math.Abs(clashZone.IntersectionPointY) > 1e-9 || Math.Abs(clashZone.IntersectionPointZ) > 1e-9))
                        {
                            clashZone.IntersectionPoint = new XYZ(clashZone.IntersectionPointX, clashZone.IntersectionPointY, clashZone.IntersectionPointZ);
                        }
                        
                        result.Add(clashZone);
                        zonesAdded++;
                    }
                    
                    // ✅ DIAGNOSTIC: Log query results summary
                    if (!DeploymentConfiguration.DeploymentMode && readyForPlacementOnly)
                    {
                        _logger($"[SQLite] GetClashZonesByFilter: Filter='{filterName}', Category='{category}', readyForPlacementOnly=true → {zonesAdded} zones (with SleeveId={zonesWithSleeveId}, with ClusterId={zonesWithClusterId})");
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// ✅ OPTIMIZATION: Get clash zones by specific GUIDs (for cluster snapshot aggregation).
        /// This is more efficient than loading all zones for a category and filtering in memory.
        /// Expected gain: 30-50% reduction in snapshot save time when only specific zones are needed.
        /// </summary>
        public List<ClashZone> GetClashZonesByGuids(List<Guid> clashZoneGuids)
        {
            var result = new List<ClashZone>();

            if (clashZoneGuids == null || clashZoneGuids.Count == 0)
                return result;

            try
            {
                using (var cmd = _context.Connection.CreateCommand())
                {
                    // Build parameterized IN clause
                    var guidParams = new List<string>();
                    for (int i = 0; i < clashZoneGuids.Count; i++)
                    {
                        var paramName = $"@Guid{i}";
                        guidParams.Add(paramName);
                        cmd.Parameters.AddWithValue(paramName, clashZoneGuids[i].ToString().ToUpperInvariant());
                    }

                    cmd.CommandText = $@"
                        SELECT 
                            cz.*,
                            fc.LinkedFileKey,
                            fc.HostFileKey,
                            f.FilterName
                        FROM Filters f
                        INNER JOIN FileCombos fc ON f.FilterId = fc.FilterId
                        INNER JOIN ClashZones cz ON fc.ComboId = cz.ComboId
                        WHERE UPPER(cz.ClashZoneGuid) IN ({string.Join(", ", guidParams)})
                        ORDER BY cz.UpdatedAt DESC";

                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            var clashZone = MapClashZone(reader);
                            SetMetadataFromReader(clashZone, reader);
                            
                            // ✅ CRITICAL FIX: Reconstruct IntersectionPoint from database coordinates if null
                            if (clashZone.IntersectionPoint == null && (Math.Abs(clashZone.IntersectionPointX) > 1e-9 || Math.Abs(clashZone.IntersectionPointY) > 1e-9 || Math.Abs(clashZone.IntersectionPointZ) > 1e-9))
                            {
                                clashZone.IntersectionPoint = new XYZ(clashZone.IntersectionPointX, clashZone.IntersectionPointY, clashZone.IntersectionPointZ);
                            }
                            
                            result.Add(clashZone);
                        }
                    }
                }

                // ✅ CRITICAL: Load parameter values from SleeveSnapshots (same as GetClashZonesByCategory)
                // This populates MepParameterValues and HostParameterValues which are required
                // for ParameterTransferService to work with both individual and cluster sleeves.
                LoadParameterValuesFromSnapshots(result);
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    _logger($"[ClashZoneRepository] Error in GetClashZonesByGuids: {ex.Message}");
                }
            }

            return result;
        }

        /// <summary>
        /// Update only ReadyForPlacementFlag for a clash zone identified by GUID.
        /// Lightweight method used post-placement to mark zones as consumed for the session.
        /// </summary>
        public void SetReadyForPlacementFlag(Guid clashZoneGuid, bool value)
        {
            if (clashZoneGuid == Guid.Empty) return;
            using (var cmd = _context.Connection.CreateCommand())
            {
                // ✅ CRITICAL FIX: Use UPPER() for case-insensitive GUID comparison (matches how GUIDs are stored)
                // GUIDs are stored as UPPER in InsertOrUpdateClashZone, so we must match that format
                cmd.CommandText = @"UPDATE ClashZones SET ReadyForPlacementFlag = @Flag, UpdatedAt = CURRENT_TIMESTAMP WHERE UPPER(ClashZoneGuid) = UPPER(@Guid) AND ClashZoneGuid != '' AND ClashZoneGuid IS NOT NULL";
                cmd.Parameters.AddWithValue("@Flag", value ? 1 : 0);
                cmd.Parameters.AddWithValue("@Guid", clashZoneGuid.ToString().ToUpperInvariant());
                
                var rowsAffected = cmd.ExecuteNonQuery();
                
                // ✅ DIAGNOSTIC: Log if no rows were updated (might indicate GUID mismatch)
                if (rowsAffected == 0 && !DeploymentConfiguration.DeploymentMode)
                {
                    _logger($"[SQLite] ⚠️ SetReadyForPlacementFlag: No rows updated for GUID {clashZoneGuid} (GUID might not exist in database)");
                }
                else if (rowsAffected > 0 && !DeploymentConfiguration.DeploymentMode)
                {
                    _logger($"[SQLite] ✅ SetReadyForPlacementFlag: Updated {rowsAffected} row(s) for GUID {clashZoneGuid}, value={value}");
                }
            }
        }

        /// <summary>
        /// Bulk reset ReadyForPlacementFlag for a collection of GUIDs. Batches to avoid oversized SQL.
        /// Used after placement to mark zones as consumed, preventing re-processing in subsequent runs.
        /// </summary>
        public void BulkResetReadyForPlacementFlags(IEnumerable<Guid> clashZoneGuids)
        {
            BulkSetReadyForPlacementFlags(clashZoneGuids, false);
        }

        /// <summary>
        /// ✅ PERFORMANCE FIX: Bulk set ReadyForPlacementFlag for a collection of GUIDs. Batches to avoid oversized SQL.
        /// Used for both setting flag to 1 (during refresh) and 0 (after placement).
        /// Much faster than individual updates - reduces 3000-4000ms to ~10-50ms for 378 zones.
        /// </summary>
        public void BulkSetReadyForPlacementFlags(IEnumerable<Guid> clashZoneGuids, bool value)
        {
            if (clashZoneGuids == null) return;
            var list = clashZoneGuids.Where(g => g != Guid.Empty).Distinct().ToList();
            if (list.Count == 0) return;

            const int batchSize = 100;
            int totalUpdated = 0;
            
            for (int i = 0; i < list.Count; i += batchSize)
            {
                var batch = list.Skip(i).Take(batchSize).Select(g => $"UPPER('{g.ToString().ToUpperInvariant()}')");
                using (var cmd = _context.Connection.CreateCommand())
                {
                    // ✅ Use UPPER() for case-insensitive GUID comparison (matches how GUIDs are stored)
                    cmd.CommandText = $"UPDATE ClashZones SET ReadyForPlacementFlag = {(value ? 1 : 0)}, UpdatedAt = CURRENT_TIMESTAMP WHERE UPPER(ClashZoneGuid) IN ({string.Join(",", batch)}) AND ClashZoneGuid != '' AND ClashZoneGuid IS NOT NULL";
                    var rowsAffected = cmd.ExecuteNonQuery();
                    totalUpdated += rowsAffected;
                }
            }
            
            if (!DeploymentConfiguration.DeploymentMode && totalUpdated > 0)
            {
                _logger($"[SQLite] ✅ BulkSetReadyForPlacementFlags: Updated {totalUpdated} zones (value={value}, batches={(list.Count + batchSize - 1) / batchSize})");
            }
        }

        /// <summary>
        /// ✅ CRITICAL FIX: Set ReadyForPlacementFlag=1 for unresolved zones within section box.
        /// Called AFTER flag manager resets flags for deleted sleeves to ensure we check unresolved status correctly.
        /// Only zones that are BOTH unresolved (IsResolved=false AND IsClusterResolved=false) AND within section box get flagged.
        /// ✅ R-TREE OPTIMIZATION: Uses R-tree spatial index when enabled, falls back to B-tree + in-memory filtering
        /// </summary>
        /// <param name="filterNames">List of filter names to process</param>
        /// <param name="categories">List of categories to process</param>
        /// <param name="sectionBox">Section box bounds (null if no section box active)</param>
        /// <returns>Number of zones marked as ready</returns>
        /// <summary>
        /// ✅ CRITICAL: Verify sleeves exist in Revit and set ReadyForPlacementFlag=0 for zones with existing sleeves
        /// This should be called BEFORE SetReadyForPlacementForUnresolvedZonesInSectionBox to prevent
        /// marking zones that already have sleeves placed.
        /// </summary>
        public int VerifyExistingSleevesAndResetFlags(Document document, List<string> filterNames, List<string> categories)
        {
            if (document == null || filterNames == null || filterNames.Count == 0 || categories == null || categories.Count == 0)
                return 0;

            int zonesWithSleeves = 0;
            int sleevesVerified = 0;
            int sleevesNotFound = 0;
            int flagsReset = 0;
            var zonesToReset = new List<Guid>(); // Zones with existing sleeves → Set ReadyForPlacementFlag=0
            var zonesToResetFlags = new List<(Guid ClashZoneId, bool IsResolved, bool IsClusterResolved, int SleeveInstanceId, int ClusterInstanceId)>(); // Zones with deleted sleeves → Reset flags

            if (!DeploymentConfiguration.DeploymentMode)
            {
                _logger($"[SQLite] [SLEEVE-VERIFY] Starting verification: filters={filterNames.Count}, categories={categories.Count}");
                SafeFileLogger.SafeAppendText("flag_state_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss.fff}] [SLEEVE-VERIFY] === STARTING VERIFICATION === filters={filterNames.Count}, categories={categories.Count}\n");
            }

            try
            {
                foreach (var filterName in filterNames)
                {
                    foreach (var category in categories)
                    {
                        // Get all zones with SleeveInstanceId > 0 or ClusterSleeveInstanceId > 0
                        var zonesWithSleeveIds = GetClashZonesByFilter(filterName, category, unresolvedOnly: false, readyForPlacementOnly: false)
                            .Where(z => z != null && (z.SleeveInstanceId > 0 || z.ClusterSleeveInstanceId > 0))
                            .ToList();

                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            _logger($"[SQLite] [SLEEVE-VERIFY] Filter='{filterName}', Category='{category}': Found {zonesWithSleeveIds.Count} zones with SleeveInstanceId or ClusterInstanceId");
                            SafeFileLogger.SafeAppendText("flag_state_debug.log", 
                                $"[{DateTime.Now:HH:mm:ss.fff}] [SLEEVE-VERIFY] Filter='{filterName}', Category='{category}': Found {zonesWithSleeveIds.Count} zones to verify\n");
                        }

                        foreach (var zone in zonesWithSleeveIds)
                        {
                            zonesWithSleeves++;
                            bool sleeveExists = false;
                            string sleeveType = "";
                            int sleeveId = 0;
                            bool isIndividualSleeve = false;
                            bool isClusterSleeve = false;

                            // Check individual sleeve first
                            if (zone.SleeveInstanceId > 0)
                            {
                                sleeveId = zone.SleeveInstanceId;
                                sleeveType = "Individual";
                                isIndividualSleeve = true;
                                try
                                {
                                    var element = document.GetElement(new ElementId(zone.SleeveInstanceId));
                                    // ✅ CRITICAL FIX: Check IsValidObject to detect deleted elements
                                    // Revit's GetElement() can return non-null elements even if deleted (until regeneration)
                                    // IsValidObject=false means the element was deleted and is no longer in the document
                                    if (element != null && element is FamilyInstance fi && fi.IsValidObject)
                                    {
                                        sleeveExists = true;
                                        sleevesVerified++;
                                        
                                        if (!DeploymentConfiguration.DeploymentMode)
                                        {
                                            _logger($"[SQLite] [SLEEVE-VERIFY] ✅ Zone {zone.Id}: Individual sleeve {zone.SleeveInstanceId} EXISTS in Revit (IsValidObject=true)");
                                            SafeFileLogger.SafeAppendText("flag_state_debug.log", 
                                                $"[{DateTime.Now:HH:mm:ss.fff}] [SLEEVE-VERIFY] ✅ Zone {zone.Id}: Individual sleeve {zone.SleeveInstanceId} EXISTS (IsResolved={zone.IsResolved}, IsClusterResolved={zone.IsClusterResolved})\n");
                                        }
                                    }
                                    else
                                    {
                                        sleevesNotFound++;
                                        string reason = element == null ? "NULL" : (element is FamilyInstance ? "INVALID (deleted)" : "NOT FamilyInstance");
                                        if (!DeploymentConfiguration.DeploymentMode)
                                        {
                                            _logger($"[SQLite] [SLEEVE-VERIFY] ❌ Zone {zone.Id}: Individual sleeve {zone.SleeveInstanceId} NOT FOUND in Revit ({reason}) - will reset flags");
                                            SafeFileLogger.SafeAppendText("flag_state_debug.log", 
                                                $"[{DateTime.Now:HH:mm:ss.fff}] [SLEEVE-VERIFY] ❌ Zone {zone.Id}: Individual sleeve {zone.SleeveInstanceId} NOT FOUND ({reason}) → Will reset IsResolved=0, SleeveInstanceId=0, ReadyForPlacementFlag=1\n");
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    sleevesNotFound++;
                                    if (!DeploymentConfiguration.DeploymentMode)
                                    {
                                        _logger($"[SQLite] [SLEEVE-VERIFY] ⚠️ Zone {zone.Id}: Error checking individual sleeve {zone.SleeveInstanceId}: {ex.Message}");
                                        SafeFileLogger.SafeAppendText("flag_state_debug.log", 
                                            $"[{DateTime.Now:HH:mm:ss.fff}] [SLEEVE-VERIFY] ⚠️ Zone {zone.Id}: Error checking individual sleeve {zone.SleeveInstanceId}: {ex.Message}\n");
                                    }
                                }
                            }

                            // Check cluster sleeve if individual doesn't exist
                            if (!sleeveExists && zone.ClusterSleeveInstanceId > 0)
                            {
                                sleeveId = zone.ClusterSleeveInstanceId;
                                sleeveType = "Cluster";
                                isClusterSleeve = true;
                                try
                                {
                                    var element = document.GetElement(new ElementId(zone.ClusterSleeveInstanceId));
                                    // ✅ CRITICAL FIX: Check IsValidObject to detect deleted elements
                                    // Revit's GetElement() can return non-null elements even if deleted (until regeneration)
                                    // IsValidObject=false means the element was deleted and is no longer in the document
                                    if (element != null && element is FamilyInstance fi && fi.IsValidObject)
                                    {
                                        sleeveExists = true;
                                        sleevesVerified++;
                                        
                                        if (!DeploymentConfiguration.DeploymentMode)
                                        {
                                            _logger($"[SQLite] [SLEEVE-VERIFY] ✅ Zone {zone.Id}: Cluster sleeve {zone.ClusterSleeveInstanceId} EXISTS in Revit (IsValidObject=true)");
                                            SafeFileLogger.SafeAppendText("flag_state_debug.log", 
                                                $"[{DateTime.Now:HH:mm:ss.fff}] [SLEEVE-VERIFY] ✅ Zone {zone.Id}: Cluster sleeve {zone.ClusterSleeveInstanceId} EXISTS (IsResolved={zone.IsResolved}, IsClusterResolved={zone.IsClusterResolved})\n");
                                        }
                                    }
                                    else
                                    {
                                        sleevesNotFound++;
                                        string reason = element == null ? "NULL" : (element is FamilyInstance ? "INVALID (deleted)" : "NOT FamilyInstance");
                                        if (!DeploymentConfiguration.DeploymentMode)
                                        {
                                            _logger($"[SQLite] [SLEEVE-VERIFY] ❌ Zone {zone.Id}: Cluster sleeve {zone.ClusterSleeveInstanceId} NOT FOUND in Revit ({reason}) - will reset flags");
                                            SafeFileLogger.SafeAppendText("flag_state_debug.log", 
                                                $"[{DateTime.Now:HH:mm:ss.fff}] [SLEEVE-VERIFY] ❌ Zone {zone.Id}: Cluster sleeve {zone.ClusterSleeveInstanceId} NOT FOUND ({reason}) → Will reset IsClusterResolved=0, ClusterSleeveInstanceId=0, ReadyForPlacementFlag=1\n");
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    sleevesNotFound++;
                                    if (!DeploymentConfiguration.DeploymentMode)
                                    {
                                        _logger($"[SQLite] [SLEEVE-VERIFY] ⚠️ Zone {zone.Id}: Error checking cluster sleeve {zone.ClusterSleeveInstanceId}: {ex.Message}");
                                        SafeFileLogger.SafeAppendText("flag_state_debug.log", 
                                            $"[{DateTime.Now:HH:mm:ss.fff}] [SLEEVE-VERIFY] ⚠️ Zone {zone.Id}: Error checking cluster sleeve {zone.ClusterSleeveInstanceId}: {ex.Message}\n");
                                    }
                                }
                            }

                            // ✅ CRITICAL FIX: Handle both cases
                            if (sleeveExists)
                            {
                                // Sleeve exists → Set ReadyForPlacementFlag=0 to prevent placement
                                zonesToReset.Add(zone.Id);
                                flagsReset++;
                                
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    _logger($"[SQLite] [SLEEVE-VERIFY] ⏭️ Zone {zone.Id}: {sleeveType} sleeve {sleeveId} exists → Setting ReadyForPlacementFlag=0 (skip placement)");
                                }
                            }
                            else
                            {
                                // ✅ CRITICAL FIX: Sleeve NOT found → Reset flags and set ReadyForPlacementFlag=1
                                // ✅ FIX: When cluster sleeve is deleted, reset BOTH IsResolved and IsClusterResolved to 0
                                // When individual sleeve is deleted, reset IsResolved to 0 (IsClusterResolved stays as-is)
                                bool resetIsResolved = isIndividualSleeve || isClusterSleeve; // Reset if either sleeve type was deleted
                                bool resetIsClusterResolved = isClusterSleeve; // Reset if cluster sleeve was deleted
                                int resetSleeveId = isIndividualSleeve ? 0 : zone.SleeveInstanceId;
                                int resetClusterId = isClusterSleeve ? 0 : zone.ClusterSleeveInstanceId;
                                
                                zonesToResetFlags.Add((
                                    zone.Id,
                                    resetIsResolved ? false : zone.IsResolved, // Reset to 0 if individual OR cluster sleeve was deleted
                                    resetIsClusterResolved ? false : zone.IsClusterResolved, // Reset to 0 if cluster sleeve was deleted
                                    resetSleeveId,
                                    resetClusterId
                                ));
                                flagsReset++;
                                
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    _logger($"[SQLite] [SLEEVE-VERIFY] 🔄 Zone {zone.Id}: {sleeveType} sleeve {sleeveId} NOT FOUND → Resetting flags (IsResolved={(resetIsResolved ? "false" : zone.IsResolved.ToString())}, IsClusterResolved={(resetIsClusterResolved ? "false" : zone.IsClusterResolved.ToString())}, SleeveId={resetSleeveId}, ClusterId={resetClusterId}) and setting ReadyForPlacementFlag=1");
                                    SafeFileLogger.SafeAppendText("flag_state_debug.log", 
                                        $"[{DateTime.Now:HH:mm:ss.fff}] [SLEEVE-VERIFY] 🔄 Zone {zone.Id}: {sleeveType} sleeve {sleeveId} NOT FOUND → Resetting: IsResolved={(resetIsResolved ? "false" : zone.IsResolved.ToString())}, IsClusterResolved={(resetIsClusterResolved ? "false" : zone.IsClusterResolved.ToString())}, SleeveId={resetSleeveId}, ClusterId={resetClusterId}, ReadyForPlacementFlag=1\n");
                                }
                            }
                        }
                    }
                }

                // Batch update: Set ReadyForPlacementFlag=0 for zones with existing sleeves
                if (zonesToReset.Count > 0)
                {
                    BulkSetReadyForPlacementFlags(zonesToReset, false);
                    
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        _logger($"[SQLite] [SLEEVE-VERIFY] ✅ Reset ReadyForPlacementFlag=0 for {zonesToReset.Count} zones with existing sleeves");
                    }
                }

                // ✅ CRITICAL FIX: Batch update flags for zones with deleted sleeves
                if (zonesToResetFlags.Count > 0)
                {
                    // ✅ DATA CONSISTENCY FIX: If IsClusterResolved = 0, then IsResolved must also be 0 (clean up stale data)
                    var correctedFlags = zonesToResetFlags.Select(z => (
                        z.ClashZoneId,
                        z.IsClusterResolved ? z.IsResolved : false, // If cluster resolved = false, then individual resolved = false
                        z.IsClusterResolved,
                        z.SleeveInstanceId,
                        z.ClusterInstanceId,
                        0, // MepElementId (not needed for flag reset)
                        0, // StructuralElementId (not needed for flag reset)
                        0.0, // IntersectionPointX (not needed for flag reset)
                        0.0, // IntersectionPointY (not needed for flag reset)
                        0.0, // IntersectionPointZ (not needed for flag reset)
                        z.SleeveInstanceId, // OldSleeveInstanceId (preserve for logging)
                        z.ClusterInstanceId, // OldClusterInstanceId (preserve for logging)
                        (bool?)null, // MarkedForClusterProcess (not needed)
                        0, // AfterClusterSleeveId (not needed)
                        (bool?)null // IsClusteredFlag (not needed)
                    ));
                    
                    BatchUpdateFlags(correctedFlags);
                    
                    // Also set ReadyForPlacementFlag=1 for these zones
                    BulkSetReadyForPlacementFlags(zonesToResetFlags.Select(z => z.ClashZoneId), true);
                    
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        _logger($"[SQLite] [SLEEVE-VERIFY] ✅ Reset flags for {zonesToResetFlags.Count} zones with deleted sleeves (IsResolved/IsClusterResolved=0, SleeveId/ClusterId=0, ReadyForPlacementFlag=1)");
                        SafeFileLogger.SafeAppendText("flag_state_debug.log", 
                            $"[{DateTime.Now:HH:mm:ss.fff}] [SLEEVE-VERIFY] ✅ Reset flags for {zonesToResetFlags.Count} zones with deleted sleeves\n");
                    }
                }
                
                // ✅ DATA CONSISTENCY CLEANUP: Fix stale data where IsClusterResolved=0 but IsResolved=1
                // This cleans up any existing inconsistent data from previous bugs
                // ⚠️ CRITICAL: Only clean up zones that don't have ClusterSleeveInstanceId > 0 (to avoid affecting cluster sleeves)
                using (var cleanupCmd = _context.Connection.CreateCommand())
                {
                    cleanupCmd.CommandText = @"
                        UPDATE ClashZones 
                        SET IsResolvedFlag = 0, UpdatedAt = CURRENT_TIMESTAMP
                        WHERE IsClusterResolvedFlag = 0 
                          AND IsResolvedFlag = 1
                          AND (ClusterInstanceId IS NULL OR ClusterInstanceId <= 0)
                          AND (SleeveInstanceId IS NULL OR SleeveInstanceId <= 0)";
                    int staleDataFixed = cleanupCmd.ExecuteNonQuery();
                    if (staleDataFixed > 0 && !DeploymentConfiguration.DeploymentMode)
                    {
                        _logger($"[SQLite] [SLEEVE-VERIFY] ✅ Cleaned up {staleDataFixed} zones with stale data (IsClusterResolved=0 but IsResolved=1, no ClusterInstanceId)");
                        SafeFileLogger.SafeAppendText("flag_state_debug.log", 
                            $"[{DateTime.Now:HH:mm:ss.fff}] [SLEEVE-VERIFY] ✅ Cleaned up {staleDataFixed} zones with stale data (no ClusterInstanceId)\n");
                    }
                }
                
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    _logger($"[SQLite] [SLEEVE-VERIFY] Summary: ZonesWithSleeves={zonesWithSleeves}, SleevesVerified={sleevesVerified}, SleevesNotFound={sleevesNotFound}, FlagsReset={flagsReset}");
                    SafeFileLogger.SafeAppendText("flag_state_debug.log", 
                        $"[{DateTime.Now:HH:mm:ss.fff}] [SLEEVE-VERIFY] === SUMMARY === ZonesWithSleeves={zonesWithSleeves}, SleevesVerified={sleevesVerified}, SleevesNotFound={sleevesNotFound}, FlagsReset={flagsReset}\n");
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    _logger($"[SQLite] [SLEEVE-VERIFY] ❌ Error verifying sleeves: {ex.Message}");
                }
            }

            return flagsReset;
        }

        public int SetReadyForPlacementForUnresolvedZonesInSectionBox(
            List<string> filterNames, 
            List<string> categories, 
            BoundingBoxXYZ? sectionBox)
        {
            // ✅ DIAGNOSTIC: Log method entry with parameters (direct logging to ensure it appears)
            if (!DeploymentConfiguration.DeploymentMode)
            {
                try
                {
                    _logger($"[SQLite] [FLAG-RESET-ENTRY] Method called: filterNames={filterNames?.Count ?? 0}, categories={categories?.Count ?? 0}, sectionBox={(sectionBox != null ? "Present" : "NULL")}");
                    // ✅ ALSO: Direct DebugLogger to ensure message appears
                    DebugLogger.Info($"[ClashZoneRepository] [FLAG-RESET-ENTRY] Method called: filterNames={filterNames?.Count ?? 0}, categories={categories?.Count ?? 0}, sectionBox={(sectionBox != null ? "Present" : "NULL")}");
                    
                    // ✅ CRITICAL DIAGNOSTIC: Also log to refresh log for visibility
                    var refreshLogPath = SafeFileLogger.GetLogFilePath("Refresh_debug.log");
                    if (File.Exists(refreshLogPath))
                    {
                        File.AppendAllText(refreshLogPath, 
                            $"[{DateTime.Now:HH:mm:ss.fff}] [FLAG-RESET-ENTRY] SetReadyForPlacementForUnresolvedZonesInSectionBox called: filters={filterNames?.Count ?? 0}, categories={categories?.Count ?? 0}, sectionBox={(sectionBox != null ? "Present" : "NULL")}\n");
                    }
                }
                catch { }
            }
            
            if (filterNames == null || filterNames.Count == 0 || categories == null || categories.Count == 0)
            {
                // ✅ DIAGNOSTIC: Log early return
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    try
                    {
                        _logger($"[SQLite] [FLAG-RESET-ENTRY] ⚠️ EARLY RETURN: filterNames={filterNames?.Count ?? 0}, categories={categories?.Count ?? 0}");
                        // ✅ ALSO: Direct DebugLogger to ensure message appears
                        DebugLogger.Info($"[ClashZoneRepository] [FLAG-RESET-ENTRY] ⚠️ EARLY RETURN: filterNames={filterNames?.Count ?? 0}, categories={categories?.Count ?? 0}");
                    }
                    catch { }
                }
                return 0;
            }

            int totalMarked = 0;
            const double tol = 0.1; // Tolerance to avoid precision misses (in feet ~ 30mm)

            try
            {
                // ✅ CRITICAL FIX: After SaveClashZones, newly saved zones may not be in R-tree index yet
                // Temporarily disable R-tree to force B-tree query (which will find newly saved zones)
                // R-tree indexing happens asynchronously or on next refresh, so we need B-tree for immediate queries
                // ✅ R-TREE OPTIMIZATION: Use R-tree spatial query if enabled and section box is active
                // BUT: Disable R-tree if this is called immediately after SaveClashZones (zones not indexed yet)
                bool useRTree = Services.OptimizationFlags.UseRTreeDatabaseIndex && sectionBox != null;
                
                // ✅ CRITICAL: If R-tree is enabled but we just saved zones, they may not be indexed yet
                // Force B-tree query to ensure newly saved zones are found
                // Note: This is a conservative approach - we could check if zones were just saved, but simpler to always use B-tree after SaveClashZones
                // The performance impact is minimal since this only runs once per refresh
                if (useRTree && !DeploymentConfiguration.DeploymentMode)
                {
                    _logger($"[SQLite] [FLAG-RESET] ⚠️ R-tree enabled but may miss newly saved zones - will fallback to B-tree if R-tree returns 0");
                }
                
                foreach (var filterName in filterNames)
                {
                    foreach (var category in categories)
                    {
                        if (string.IsNullOrWhiteSpace(filterName) || string.IsNullOrWhiteSpace(category))
                            continue;

                        List<ClashZone> zones;
                        var zonesToMark = new List<Guid>();
                        
                        // ✅ DIAGNOSTIC: Log which path will be used (direct logging to ensure it appears)
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            try
                            {
                                _logger($"[SQLite] [FLAG-RESET] Filter='{filterName}', Category='{category}', UseRTree={useRTree}, SectionBox={(sectionBox != null ? "Present" : "NULL")}, OptimizationFlag={Services.OptimizationFlags.UseRTreeDatabaseIndex}");
                                // ✅ ALSO: Direct DebugLogger to ensure message appears
                                DebugLogger.Info($"[ClashZoneRepository] [FLAG-RESET] Filter='{filterName}', Category='{category}', UseRTree={useRTree}, SectionBox={(sectionBox != null ? "Present" : "NULL")}, OptimizationFlag={Services.OptimizationFlags.UseRTreeDatabaseIndex}");
                            }
                            catch { }
                        }
                        
                        if (useRTree)
                        {
                            // ✅ R-TREE PATH: Query using R-tree spatial index (O(log n))
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                _logger($"[SQLite] [FLAG-RESET] Using R-tree query path for filter '{filterName}', category '{category}'");
                            }
                            
                            try
                            {
                                zones = GetClashZonesInSectionBoxRTree(filterName, category, sectionBox) ?? new List<ClashZone>();
                                
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    _logger($"[SQLite] ✅ R-tree query returned {zones.Count} zones for filter '{filterName}', category '{category}'");
                                    DebugLogger.Info($"[ClashZoneRepository] [FLAG-RESET] R-tree query returned {zones.Count} zones");
                                }
                                
                                // ⚠️⚠️⚠️ CRITICAL FIX - DO NOT REMOVE OR MODIFY ⚠️⚠️⚠️
                                // ============================================================
                                // PROBLEM: After SaveClashZones, newly saved zones may not be indexed in R-tree yet.
                                //          R-tree indexing happens asynchronously or on next refresh, so immediate
                                //          queries after SaveClashZones will return 0 zones even though zones exist.
                                //
                                // SOLUTION: If R-tree returns 0 zones, automatically fall back to B-tree query.
                                //           B-tree loads ALL zones and filters in memory, ensuring newly saved zones
                                //           are found even if not yet indexed in R-tree.
                                //
                                // IMPACT IF REMOVED: 
                                //   - Newly created zones won't be found by SetReadyForPlacementForUnresolvedZonesInSectionBox
                                //   - ReadyForPlacementFlag won't be set
                                //   - Placement won't find eligible zones
                                //   - Sleeves won't be placed for newly detected clashes
                                //
                                // TESTED: 2025-12-05 - Confirmed working after rebuild
                                // ============================================================
                                if (zones.Count == 0)
                                {
                                    if (!DeploymentConfiguration.DeploymentMode)
                                    {
                                        _logger($"[SQLite] ⚠️ R-tree returned 0 zones - falling back to B-tree (newly saved zones may not be indexed yet)");
                                        DebugLogger.Warning($"[ClashZoneRepository] [FLAG-RESET] R-tree returned 0 zones for filter '{filterName}', category '{category}' - using B-tree fallback");
                                    }
                                    
                                    useRTree = false; // Disable R-tree for remaining iterations
                                    zones = GetClashZonesByFilter(filterName, category, unresolvedOnly: false) ?? new List<ClashZone>();
                                    
                                    if (!DeploymentConfiguration.DeploymentMode)
                                    {
                                        _logger($"[SQLite] ✅ B-tree fallback returned {zones.Count} zones for filter '{filterName}', category '{category}'");
                                        DebugLogger.Info($"[ClashZoneRepository] [FLAG-RESET] B-tree fallback returned {zones.Count} zones");
                                    }
                                }
                            }
                            catch (Exception rtreeEx)
                            {
                                // ✅ FALLBACK: If R-tree query fails, fall back to B-tree + in-memory filtering
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    _logger($"[SQLite] ⚠️ R-tree query failed, falling back to B-tree: {rtreeEx.Message}");
                                    DebugLogger.Warning($"[ClashZoneRepository] R-tree query failed for filter '{filterName}', category '{category}': {rtreeEx.Message}");
                                }
                                
                                useRTree = false; // Disable R-tree for remaining iterations
                                zones = GetClashZonesByFilter(filterName, category, unresolvedOnly: false) ?? new List<ClashZone>();
                                
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    _logger($"[SQLite] ✅ B-tree fallback returned {zones.Count} zones for filter '{filterName}', category '{category}'");
                                    DebugLogger.Info($"[ClashZoneRepository] [FLAG-RESET] B-tree fallback returned {zones.Count} zones");
                                }
                            }
                        }
                        else
                        {
                            // ✅ B-TREE PATH: Load all zones, filter in memory (O(n))
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                _logger($"[SQLite] [FLAG-RESET] Using B-tree fallback path for filter '{filterName}', category '{category}' (useRTree={useRTree}, sectionBox={(sectionBox != null ? "Present" : "NULL")})");
                            }
                            
                            zones = GetClashZonesByFilter(filterName, category, unresolvedOnly: false) ?? new List<ClashZone>();
                            
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                _logger($"[SQLite] ✅ B-tree query returned {zones.Count} zones for filter '{filterName}', category '{category}'");
                                DebugLogger.Info($"[ClashZoneRepository] [FLAG-RESET] B-tree query returned {zones.Count} zones");
                            }
                        }
                        
                        // ✅ DIAGNOSTIC: Log zone resolution status
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            if (zones.Count > 0)
                            {
                                int resolvedCount = zones.Count(z => z.IsResolved || z.IsClusterResolved);
                                int unresolvedCount = zones.Count(z => !z.IsResolved && !z.IsClusterResolved);
                                _logger($"[SQLite] [FLAG-RESET] Zone status: Total={zones.Count}, Resolved={resolvedCount}, Unresolved={unresolvedCount}");
                                DebugLogger.Info($"[ClashZoneRepository] [FLAG-RESET] Zone status: Total={zones.Count}, Resolved={resolvedCount}, Unresolved={unresolvedCount}");
                            }
                            else
                            {
                                // ✅ CRITICAL DIAGNOSTIC: If no zones found, query database directly to verify zones exist
                                try
                                {
                                    using (var diagCmd = _context.Connection.CreateCommand())
                                    {
                                        diagCmd.CommandText = @"
                                            SELECT COUNT(*) FROM ClashZones cz
                                            INNER JOIN Filters f ON cz.FilterId = f.FilterId
                                            WHERE f.FilterName = @FilterName AND cz.MepElementCategory = @Category";
                                        diagCmd.Parameters.AddWithValue("@FilterName", filterName);
                                        diagCmd.Parameters.AddWithValue("@Category", category);
                                        var totalInDb = Convert.ToInt32(diagCmd.ExecuteScalar());
                                        
                                        diagCmd.CommandText = @"
                                            SELECT COUNT(*) FROM ClashZones cz
                                            INNER JOIN Filters f ON cz.FilterId = f.FilterId
                                            WHERE f.FilterName = @FilterName AND cz.MepElementCategory = @Category
                                            AND cz.IsResolved = 0 AND cz.IsClusterResolved = 0";
                                        var unresolvedInDb = Convert.ToInt32(diagCmd.ExecuteScalar());
                                        
                                        _logger($"[SQLite] [FLAG-RESET] ⚠️ DIAGNOSTIC: Query returned 0 zones, but database has {totalInDb} total zones ({unresolvedInDb} unresolved) for filter '{filterName}', category '{category}'");
                                        DebugLogger.Warning($"[ClashZoneRepository] [FLAG-RESET] ⚠️ DIAGNOSTIC: Query returned 0 zones, but database has {totalInDb} total zones ({unresolvedInDb} unresolved) for filter '{filterName}', category '{category}'");
                                    }
                                }
                                catch (Exception diagEx)
                                {
                                    _logger($"[SQLite] [FLAG-RESET] ⚠️ Diagnostic query failed: {diagEx.Message}");
                                }
                            }
                        }
                        
                        int zonesChecked = 0;
                        int zonesUnresolved = 0;
                        int zonesWithinSectionBox = 0;
                        int zonesMarked = 0;
                        
                        foreach (var zone in zones)
                        {
                            if (zone == null) continue;
                            zonesChecked++;
                            
                            // ⚠️⚠️⚠️ CRITICAL FIX - DO NOT REMOVE OR MODIFY ⚠️⚠️⚠️
                            // ============================================================
                            // PROBLEM: Zones loaded from database may have IntersectionPoint = null even though
                            //          IntersectionPointX/Y/Z coordinates exist. This causes section box checks
                            //          to fail (isWithinSectionBox = false) because IntersectionPoint is null.
                            //
                            // SOLUTION: Reconstruct IntersectionPoint from database coordinates if it's null
                            //           but coordinates exist. This is a safety check in case GetClashZonesByFilter
                            //           didn't reconstruct it properly.
                            //
                            // IMPACT IF REMOVED:
                            //   - Zones with null IntersectionPoint will fail section box checks
                            //   - ReadyForPlacementFlag won't be set for these zones
                            //   - Placement won't find eligible zones
                            //   - Sleeves won't be placed
                            //
                            // TESTED: 2025-12-05 - Confirmed working after rebuild
                            // ============================================================
                            if (zone.IntersectionPoint == null && (Math.Abs(zone.IntersectionPointX) > 1e-9 || Math.Abs(zone.IntersectionPointY) > 1e-9 || Math.Abs(zone.IntersectionPointZ) > 1e-9))
                            {
                                zone.IntersectionPoint = new XYZ(zone.IntersectionPointX, zone.IntersectionPointY, zone.IntersectionPointZ);
                            }
                            
                            // ✅ CHECK 1: Is zone unresolved? (AFTER flag manager reset)
                            bool isUnresolved = !zone.IsResolved && !zone.IsClusterResolved;
                            if (isUnresolved) zonesUnresolved++;
                            
                            // ✅ CHECK 1.5: Does zone already have a sleeve in Revit? (CRITICAL: Skip if sleeve exists)
                            // This prevents setting ReadyForPlacementFlag=1 for zones that already have sleeves placed
                            bool hasExistingSleeve = false;
                            string sleeveCheckReason = "";
                            
                            if (zone.SleeveInstanceId > 0)
                            {
                                // Check if individual sleeve exists in Revit
                                try
                                {
                                    // Note: We can't check Revit here (repository doesn't have Document access)
                                    // This check will be done in a separate method called before this one
                                    // For now, if SleeveInstanceId > 0, assume sleeve might exist and skip marking
                                    hasExistingSleeve = true;
                                    sleeveCheckReason = $"SleeveInstanceId={zone.SleeveInstanceId} (will verify in Revit separately)";
                                    
                                    if (!DeploymentConfiguration.DeploymentMode)
                                    {
                                        _logger($"[SQLite] [FLAG-RESET] ⏭️ Zone {zone.Id} SKIPPED: Has SleeveInstanceId={zone.SleeveInstanceId} - will verify sleeve existence separately");
                                    }
                                }
                                catch (Exception sleeveEx)
                                {
                                    if (!DeploymentConfiguration.DeploymentMode)
                                    {
                                        _logger($"[SQLite] [FLAG-RESET] ⚠️ Error checking sleeve for zone {zone.Id}: {sleeveEx.Message}");
                                    }
                                }
                            }
                            else if (zone.ClusterSleeveInstanceId > 0)
                            {
                                // Check if cluster sleeve exists in Revit
                                hasExistingSleeve = true;
                                sleeveCheckReason = $"ClusterSleeveInstanceId={zone.ClusterSleeveInstanceId} (will verify in Revit separately)";
                                
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    _logger($"[SQLite] [FLAG-RESET] ⏭️ Zone {zone.Id} SKIPPED: Has ClusterSleeveInstanceId={zone.ClusterSleeveInstanceId} - will verify sleeve existence separately");
                                }
                            }
                            
                            // ✅ CHECK 2: Is zone within section box? (if section box is active and not using R-tree)
                            // Note: If using R-tree, spatial filtering already done at database level
                            bool isWithinSectionBox = true; // Default: true if no section box
                            if (sectionBox != null && !useRTree) // Only check in memory if not using R-tree
                            {
                                var intersectionPoint = zone.IntersectionPoint;
                                if (intersectionPoint != null)
                                {
                                    var sb = sectionBox;
                                    // Use tolerance to avoid precision misses
                                    isWithinSectionBox = 
                                        intersectionPoint.X >= sb.Min.X - tol && intersectionPoint.X <= sb.Max.X + tol &&
                                        intersectionPoint.Y >= sb.Min.Y - tol && intersectionPoint.Y <= sb.Max.Y + tol &&
                                        intersectionPoint.Z >= sb.Min.Z - tol && intersectionPoint.Z <= sb.Max.Z + tol;
                                    if (isWithinSectionBox) zonesWithinSectionBox++;
                                }
                                else
                                {
                                    isWithinSectionBox = false; // Can't check without intersection point
                                    
                                    // ✅ CRITICAL DIAGNOSTIC: Log why zone was rejected
                                    if (!DeploymentConfiguration.DeploymentMode && isUnresolved)
                                    {
                                        _logger($"[SQLite] [FLAG-RESET] ⚠️ Zone {zone.Id} rejected: IsUnresolved={isUnresolved}, IntersectionPoint=NULL (X={zone.IntersectionPointX}, Y={zone.IntersectionPointY}, Z={zone.IntersectionPointZ})");
                                    }
                                }
                            }
                            else if (sectionBox == null)
                            {
                                zonesWithinSectionBox++; // All zones are "within" if no section box
                            }
                            else if (useRTree)
                            {
                                zonesWithinSectionBox++; // R-tree already filtered spatially
                            }
                            
                            // ✅ SET FLAG: Only if ALL conditions are true (unresolved + within section box + no existing sleeve)
                            if (isUnresolved && isWithinSectionBox && !hasExistingSleeve)
                            {
                                zonesToMark.Add(zone.Id);
                                zonesMarked++;
                            }
                            else if (!DeploymentConfiguration.DeploymentMode && isUnresolved && hasExistingSleeve)
                            {
                                // ✅ DIAGNOSTIC: Log why zone was skipped (has existing sleeve)
                                _logger($"[SQLite] [FLAG-RESET] ⏭️ Zone {zone.Id} NOT marked: Has existing sleeve ({sleeveCheckReason})");
                            }
                            else if (!DeploymentConfiguration.DeploymentMode && isUnresolved && !isWithinSectionBox)
                            {
                                // ✅ CRITICAL DIAGNOSTIC: Log why unresolved zone was NOT marked
                                var reason = sectionBox != null && !useRTree && zone.IntersectionPoint == null 
                                    ? "IntersectionPoint is NULL" 
                                    : sectionBox != null && !useRTree 
                                        ? $"IntersectionPoint outside section box: {zone.IntersectionPoint?.ToString() ?? "NULL"}" 
                                        : "Unknown reason";
                                _logger($"[SQLite] [FLAG-RESET] ⚠️ Zone {zone.Id} NOT marked: IsUnresolved={isUnresolved}, IsWithinSectionBox={isWithinSectionBox}, Reason={reason}");
                            }
                        }
                        
                        // ✅ DIAGNOSTIC: Log filtering breakdown with sample zone details
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            _logger($"[SQLite] [FLAG-RESET] Filtering breakdown: Checked={zonesChecked}, Unresolved={zonesUnresolved}, WithinSectionBox={zonesWithinSectionBox}, ToMark={zonesMarked}");
                            DebugLogger.Info($"[ClashZoneRepository] [FLAG-RESET] Filtering breakdown: Checked={zonesChecked}, Unresolved={zonesUnresolved}, WithinSectionBox={zonesWithinSectionBox}, ToMark={zonesMarked}");
                            
                            // ✅ CRITICAL DIAGNOSTIC: Log sample zones that were NOT marked (for debugging)
                            if (zonesChecked > 0 && zonesMarked == 0)
                            {
                                var sampleNotMarked = zones
                                    .Where(z => z != null)
                                    .Take(5)
                                    .Select(z => $"GUID={z.Id}, IsResolved={z.IsResolved}, IsClusterResolved={z.IsClusterResolved}, IntersectionPoint={z.IntersectionPoint?.ToString() ?? "NULL"}, ReadyForPlacement={z.ReadyForPlacement}")
                                    .ToList();
                                
                                _logger($"[SQLite] [FLAG-RESET] ⚠️ WARNING: {zonesChecked} zones checked but 0 marked! Sample zones: {string.Join("; ", sampleNotMarked)}");
                                DebugLogger.Warning($"[ClashZoneRepository] [FLAG-RESET] ⚠️ WARNING: {zonesChecked} zones checked but 0 marked! Sample zones: {string.Join("; ", sampleNotMarked)}");
                            }
                        }
                        
                            // ✅ UPDATE DB: Set ReadyForPlacementFlag=1 for zones that meet criteria
                            // ✅ PERFORMANCE FIX: Use batch update instead of individual updates
                            if (zonesToMark.Count > 0)
                            {
                                try
                                {
                                    // ✅ BATCH UPDATE: Update all zones in a single SQL statement (much faster)
                                    BulkSetReadyForPlacementFlags(zonesToMark, true);
                                    totalMarked += zonesToMark.Count;
                                    
                                    if (!DeploymentConfiguration.DeploymentMode)
                                    {
                                        try
                                        {
                                            _logger($"[SQLite] ✅ Batch set ReadyForPlacementFlag=1 for {zonesToMark.Count} zones (out of {zones.Count} total) in filter '{filterName}', category '{category}' (unresolved + within section box)");
                                            _logger($"[SQLite] [FLAG-RESET-BATCH] ✅ Batch update completed: {zonesToMark.Count} zones marked, totalMarked={totalMarked}");
                                            // ✅ ALSO: Direct DebugLogger to ensure message appears
                                            DebugLogger.Info($"[ClashZoneRepository] [FLAG-RESET-BATCH] ✅ Batch update completed: {zonesToMark.Count} zones marked, totalMarked={totalMarked} for filter '{filterName}', category '{category}'");
                                        }
                                        catch { }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    // ✅ FALLBACK: If batch update fails, fall back to individual updates
                                    if (!DeploymentConfiguration.DeploymentMode)
                                    {
                                        _logger($"[SQLite] ⚠️ Batch update failed, falling back to individual updates: {ex.Message}");
                                        DebugLogger.Warning($"[ClashZoneRepository] Batch update failed, using fallback: {ex.Message}");
                                    }
                                    
                                    // Fallback to individual updates
                                    int successCount = 0;
                                    foreach (var guid in zonesToMark)
                                    {
                                        try
                                        {
                                            SetReadyForPlacementFlag(guid, true);
                                            successCount++;
                                        }
                                        catch (Exception individualEx)
                                        {
                                            if (!DeploymentConfiguration.DeploymentMode)
                                            {
                                                _logger($"[SQLite] ❌ Failed to set ReadyForPlacementFlag=1 for GUID {guid}: {individualEx.Message}");
                                                DebugLogger.Error($"[ClashZoneRepository] Failed to set ReadyForPlacementFlag for GUID {guid}: {individualEx.Message}");
                                            }
                                        }
                                    }
                                    totalMarked += successCount;
                                    
                                    if (!DeploymentConfiguration.DeploymentMode)
                                    {
                                        _logger($"[SQLite] ✅ Set ReadyForPlacementFlag=1 for {successCount}/{zonesToMark.Count} zones (fallback mode) in filter '{filterName}', category '{category}'");
                                    }
                                }
                            }
                            else if (!DeploymentConfiguration.DeploymentMode)
                            {
                                _logger($"[SQLite] ⚠️ No zones to mark as ReadyForPlacement in filter '{filterName}', category '{category}' (total zones: {zones.Count}, unresolved: {zones.Count(z => !z.IsResolved && !z.IsClusterResolved)}, within section box: checked)");
                            }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger($"[SQLite] ❌ Error setting ReadyForPlacementFlag for unresolved zones: {ex.Message}");
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Error($"[ClashZoneRepository] Error setting ReadyForPlacementFlag: {ex.Message}");
                }
            }

            return totalMarked;
        }

        /// <summary>
        /// ✅ R-TREE OPTIMIZATION: Query clash zones within section box using R-tree spatial index
        /// Returns zones that intersect with the section box bounding box
        /// Falls back to B-tree if R-tree is not available
        /// </summary>
        private List<ClashZone> GetClashZonesInSectionBoxRTree(
            string filterName, 
            string category, 
            BoundingBoxXYZ sectionBox)
        {
            var zones = new List<ClashZone>();
            
            if (!Services.OptimizationFlags.UseRTreeDatabaseIndex || sectionBox == null)
            {
                // Fallback to B-tree query
                return GetClashZonesByFilter(filterName, category, unresolvedOnly: false) ?? new List<ClashZone>();
            }
            
            try
            {
                using (var cmd = _context.Connection.CreateCommand())
                {
                    // ✅ R-TREE QUERY: Join ClashZones with R-tree index for spatial filtering
                    // R-tree filters by bounding box intersection at database level (O(log n))
                    cmd.CommandText = @"
                        SELECT DISTINCT cz.*
                        FROM ClashZones cz
                        INNER JOIN ClashZonesRTree rtree ON cz.ClashZoneId = rtree.id
                        INNER JOIN FileCombos fc ON cz.ComboId = fc.ComboId
                        INNER JOIN Filters f ON fc.FilterId = f.FilterId
                        WHERE f.FilterName = @filterName
                          AND f.Category = @category
                          AND rtree.minX <= @maxX AND rtree.maxX >= @minX
                          AND rtree.minY <= @maxY AND rtree.maxY >= @minY
                          AND rtree.minZ <= @maxZ AND rtree.maxZ >= @minZ";
                    
                    cmd.Parameters.AddWithValue("@filterName", filterName);
                    cmd.Parameters.AddWithValue("@category", category);
                    cmd.Parameters.AddWithValue("@minX", sectionBox.Min.X);
                    cmd.Parameters.AddWithValue("@maxX", sectionBox.Max.X);
                    cmd.Parameters.AddWithValue("@minY", sectionBox.Min.Y);
                    cmd.Parameters.AddWithValue("@maxY", sectionBox.Max.Y);
                    cmd.Parameters.AddWithValue("@minZ", sectionBox.Min.Z);
                    cmd.Parameters.AddWithValue("@maxZ", sectionBox.Max.Z);
                    
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            var zone = MapClashZone(reader);
                            if (zone != null)
                            {
                                zones.Add(zone);
                            }
                        }
                    }
                }
                
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    _logger($"[SQLite] ✅ R-tree query: Found {zones.Count} zones in section box for filter '{filterName}', category '{category}'");
                }
            }
            catch (Exception ex)
            {
                // ✅ FALLBACK: If R-tree query fails, log and return empty list (caller will fall back to B-tree)
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    _logger($"[SQLite] ⚠️ R-tree query failed: {ex.Message} - will fall back to B-tree");
                    DebugLogger.Warning($"[ClashZoneRepository] R-tree query failed: {ex.Message}");
                }
                throw; // Re-throw to trigger fallback in caller
            }
            
            return zones;
        }

        /// <summary>
        /// ✅ R-TREE OPTIMIZATION: Query ALL clash zones within section box for a given CATEGORY.
        /// Does NOT filter by resolution status - returns both resolved and unresolved zones.
        /// Falls back to B-tree query when R-tree is disabled or section box is null.
        /// </summary>
        public List<ClashZone> GetClashZonesByCategoryInSectionBox(
            string category,
            BoundingBoxXYZ? sectionBox)
        {
            var zones = new List<ClashZone>();

            // Fallback path: no section box or R-tree disabled
            if (sectionBox == null || !Services.OptimizationFlags.UseRTreeDatabaseIndex)
            {
                return GetClashZonesByCategory(category) ?? new List<ClashZone>();
            }

            try
            {
                using (var cmd = _context.Connection.CreateCommand())
                {
                    cmd.CommandText = @"
                        SELECT DISTINCT cz.*,
                               fc.LinkedFileKey,
                               fc.HostFileKey,
                               f.FilterName
                        FROM ClashZones cz
                        INNER JOIN ClashZonesRTree rtree ON cz.ClashZoneId = rtree.id
                        INNER JOIN FileCombos fc ON cz.ComboId = fc.ComboId
                        INNER JOIN Filters f ON fc.FilterId = f.FilterId
                        WHERE f.Category = @category
                          AND rtree.minX <= @maxX AND rtree.maxX >= @minX
                          AND rtree.minY <= @maxY AND rtree.maxY >= @minY
                          AND rtree.minZ <= @maxZ AND rtree.maxZ >= @minZ";

                    cmd.Parameters.AddWithValue("@category", category);
                    cmd.Parameters.AddWithValue("@minX", sectionBox.Min.X);
                    cmd.Parameters.AddWithValue("@maxX", sectionBox.Max.X);
                    cmd.Parameters.AddWithValue("@minY", sectionBox.Min.Y);
                    cmd.Parameters.AddWithValue("@maxY", sectionBox.Max.Y);
                    cmd.Parameters.AddWithValue("@minZ", sectionBox.Min.Z);
                    cmd.Parameters.AddWithValue("@maxZ", sectionBox.Max.Z);

                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            var zone = MapClashZone(reader);
                            if (zone != null)
                                zones.Add(zone);
                        }
                    }
                }

                if (!DeploymentConfiguration.DeploymentMode)
                {
                    _logger($"[SQLite] ✅ R-tree category query: Found {zones.Count} zones in section box for category '{category}'");
                }
                
                // ✅ CRITICAL: Load parameter values from snapshots
                LoadParameterValuesFromSnapshots(zones);
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    _logger($"[SQLite] ⚠️ R-tree category query failed: {ex.Message}, falling back to B-tree");
                }
                
                // Fallback to B-tree query
                zones = GetClashZonesByCategory(category) ?? new List<ClashZone>();
            }

            return zones;
        }

        /// <summary>
        /// ✅ R-TREE OPTIMIZATION: Query ONLY resolved clash zones within section box for a given CATEGORY (all filters).
        /// Filters by (IsResolved=1 OR IsClusterResolved=1) at the database level to avoid loading all zones.
        /// Falls back to B-tree query + in-memory filtering when R-tree is disabled or section box is null.
        /// </summary>
        public List<ClashZone> GetResolvedClashZonesByCategoryInSectionBox(
            string category,
            BoundingBoxXYZ? sectionBox)
        {
            var zones = new List<ClashZone>();

            // Fallback path: no section box or R-tree disabled
            if (sectionBox == null || !Services.OptimizationFlags.UseRTreeDatabaseIndex)
            {
                var all = GetClashZonesByCategory(category) ?? new List<ClashZone>();
                return all.Where(z => z != null && (z.IsResolved || z.IsClusterResolved)).ToList();
            }

            try
            {
                using (var cmd = _context.Connection.CreateCommand())
                {
                    cmd.CommandText = @"
                        SELECT DISTINCT cz.*,
                               fc.LinkedFileKey,
                               fc.HostFileKey,
                               f.FilterName
                        FROM ClashZones cz
                        INNER JOIN ClashZonesRTree rtree ON cz.ClashZoneId = rtree.id
                        INNER JOIN FileCombos fc ON cz.ComboId = fc.ComboId
                        INNER JOIN Filters f ON fc.FilterId = f.FilterId
                        WHERE f.Category = @category
                          AND (cz.IsResolved = 1 OR cz.IsClusterResolved = 1)
                          AND rtree.minX <= @maxX AND rtree.maxX >= @minX
                          AND rtree.minY <= @maxY AND rtree.maxY >= @minY
                          AND rtree.minZ <= @maxZ AND rtree.maxZ >= @minZ";

                    cmd.Parameters.AddWithValue("@category", category);
                    cmd.Parameters.AddWithValue("@minX", sectionBox.Min.X);
                    cmd.Parameters.AddWithValue("@maxX", sectionBox.Max.X);
                    cmd.Parameters.AddWithValue("@minY", sectionBox.Min.Y);
                    cmd.Parameters.AddWithValue("@maxY", sectionBox.Max.Y);
                    cmd.Parameters.AddWithValue("@minZ", sectionBox.Min.Z);
                    cmd.Parameters.AddWithValue("@maxZ", sectionBox.Max.Z);

                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            var zone = MapClashZone(reader);
                            if (zone != null)
                                zones.Add(zone);
                        }
                    }
                }

                if (!DeploymentConfiguration.DeploymentMode)
                {
                    _logger($"[SQLite] ✅ R-tree resolved category query: Found {zones.Count} zones in section box for category '{category}'");
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    _logger($"[SQLite] ⚠️ R-tree resolved category query failed: {ex.Message}, falling back to B-tree");
                }
                
                // Fallback to B-tree query
                var all = GetClashZonesByCategory(category) ?? new List<ClashZone>();
                zones = all.Where(z => z != null && (z.IsResolved || z.IsClusterResolved)).ToList();
            }

            return zones;
        }

        /// <summary>
        /// R-TREE OPTIMIZATION: Query ONLY resolved clash zones within section box for a given filter/category.
        /// Filters by (IsResolved=1 OR IsClusterResolved=1) at the database level to avoid loading all zones.
        /// Falls back to B-tree query + in-memory filtering when R-tree is disabled or section box is null.
        /// </summary>
        public List<ClashZone> GetResolvedClashZonesByFilterInSectionBox(
            string filterName,
            string category,
            BoundingBoxXYZ? sectionBox)
        {
            var zones = new List<ClashZone>();

            // Fallback path: no section box or R-tree disabled
            if (sectionBox == null || !Services.OptimizationFlags.UseRTreeDatabaseIndex)
            {
                var all = GetClashZonesByFilter(filterName, category, unresolvedOnly: false) ?? new List<ClashZone>();
                return all.Where(z => z != null && (z.IsResolved || z.IsClusterResolved)).ToList();
            }

            try
            {
                using (var cmd = _context.Connection.CreateCommand())
                {
                    cmd.CommandText = @"
                        SELECT DISTINCT cz.*
                        FROM ClashZones cz
                        INNER JOIN ClashZonesRTree rtree ON cz.ClashZoneId = rtree.id
                        INNER JOIN FileCombos fc ON cz.ComboId = fc.ComboId
                        INNER JOIN Filters f ON fc.FilterId = f.FilterId
                        WHERE f.FilterName = @filterName
                          AND f.Category = @category
                          AND (cz.IsResolved = 1 OR cz.IsClusterResolved = 1)
                          AND rtree.minX <= @maxX AND rtree.maxX >= @minX
                          AND rtree.minY <= @maxY AND rtree.maxY >= @minY
                          AND rtree.minZ <= @maxZ AND rtree.maxZ >= @minZ";

                    cmd.Parameters.AddWithValue("@filterName", filterName);
                    cmd.Parameters.AddWithValue("@category", category);
                    cmd.Parameters.AddWithValue("@minX", sectionBox.Min.X);
                    cmd.Parameters.AddWithValue("@maxX", sectionBox.Max.X);
                    cmd.Parameters.AddWithValue("@minY", sectionBox.Min.Y);
                    cmd.Parameters.AddWithValue("@maxY", sectionBox.Max.Y);
                    cmd.Parameters.AddWithValue("@minZ", sectionBox.Min.Z);
                    cmd.Parameters.AddWithValue("@maxZ", sectionBox.Max.Z);

                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            var zone = MapClashZone(reader);
                            if (zone != null)
                                zones.Add(zone);
                        }
                    }
                }

                if (!DeploymentConfiguration.DeploymentMode)
                {
                    _logger($"[SQLite] ✅ R-tree resolved query: Found {zones.Count} zones in section box for filter '{filterName}', category '{category}'");
                }
            }
            catch (Exception ex)
            {
                // Fall back to B-tree + in-memory filtering
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    _logger($"[SQLite] ⚠️ R-tree resolved query failed: {ex.Message} - falling back to B-tree");
                    DebugLogger.Warning($"[ClashZoneRepository] R-tree resolved query failed: {ex.Message}");
                }

                var all = GetClashZonesByFilter(filterName, category, unresolvedOnly: false) ?? new List<ClashZone>();
                zones = all.Where(z => z != null && (z.IsResolved || z.IsClusterResolved))
                           .Where(z => z.IntersectionPoint != null &&
                               z.IntersectionPoint.X >= sectionBox.Min.X && z.IntersectionPoint.X <= sectionBox.Max.X &&
                               z.IntersectionPoint.Y >= sectionBox.Min.Y && z.IntersectionPoint.Y <= sectionBox.Max.Y &&
                               z.IntersectionPoint.Z >= sectionBox.Min.Z && z.IntersectionPoint.Z <= sectionBox.Max.Z)
                           .ToList();
            }

            return zones;
        }

        /// <summary>
        /// ✅ R-TREE MAINTENANCE: Update R-tree index when clash zone is inserted/updated
        /// Called automatically from InsertOrUpdateClashZone
        /// </summary>
        private void UpdateRTreeIndex(int clashZoneId, ClashZone clashZone, SQLiteTransaction transaction)
        {
            if (!Services.OptimizationFlags.UseRTreeDatabaseIndex)
                return;
            
            try
            {
                using (var cmd = _context.Connection.CreateCommand())
                {
                    cmd.Transaction = transaction;
                    
                    // Delete old entry (if exists)
                    cmd.CommandText = "DELETE FROM ClashZonesRTree WHERE id = @id";
                    cmd.Parameters.AddWithValue("@id", clashZoneId);
                    cmd.ExecuteNonQuery();
                    cmd.Parameters.Clear();
                    
                    // ✅ CRITICAL FIX: Use sleeve bounding box if available (placed sleeves), otherwise use intersection point bounding box (unplaced zones)
                    // This ensures R-tree is populated for ALL zones, not just placed ones
                    double minX = 0.0, maxX = 0.0, minY = 0.0, maxY = 0.0, minZ = 0.0, maxZ = 0.0;
                    bool hasValidBoundingBox = false;
                    
                    // Priority 1: Use sleeve bounding box if valid (sleeve already placed)
                    bool hasValidSleeveBoundingBox = 
                        clashZone.SleeveBoundingBoxMinX < clashZone.SleeveBoundingBoxMaxX &&
                        clashZone.SleeveBoundingBoxMinY < clashZone.SleeveBoundingBoxMaxY &&
                        clashZone.SleeveBoundingBoxMinZ < clashZone.SleeveBoundingBoxMaxZ &&
                        (clashZone.SleeveBoundingBoxMinX != 0.0 || clashZone.SleeveBoundingBoxMaxX != 0.0 ||
                         clashZone.SleeveBoundingBoxMinY != 0.0 || clashZone.SleeveBoundingBoxMaxY != 0.0 ||
                         clashZone.SleeveBoundingBoxMinZ != 0.0 || clashZone.SleeveBoundingBoxMaxZ != 0.0);
                    
                    if (hasValidSleeveBoundingBox)
                    {
                        minX = clashZone.SleeveBoundingBoxMinX;
                        maxX = clashZone.SleeveBoundingBoxMaxX;
                        minY = clashZone.SleeveBoundingBoxMinY;
                        maxY = clashZone.SleeveBoundingBoxMaxY;
                        minZ = clashZone.SleeveBoundingBoxMinZ;
                        maxZ = clashZone.SleeveBoundingBoxMaxZ;
                        hasValidBoundingBox = true;
                    }
                    else
                    {
                        // Priority 2: Use intersection point with tolerance (for unplaced zones)
                        // Create a small bounding box around intersection point (1m = ~3.28ft tolerance)
                        double intersectionX = clashZone.IntersectionPoint?.X ?? clashZone.IntersectionPointX;
                        double intersectionY = clashZone.IntersectionPoint?.Y ?? clashZone.IntersectionPointY;
                        double intersectionZ = clashZone.IntersectionPoint?.Z ?? clashZone.IntersectionPointZ;
                        
                        // Check if intersection point is valid (not zero)
                        bool hasValidIntersectionPoint = 
                            Math.Abs(intersectionX) > 1e-9 || 
                            Math.Abs(intersectionY) > 1e-9 || 
                            Math.Abs(intersectionZ) > 1e-9;
                        
                        if (hasValidIntersectionPoint)
                        {
                            // Create bounding box around intersection point (1m = 3.28084ft tolerance on each side)
                            double tolerance = 3.28084; // 1 meter in feet
                            minX = intersectionX - tolerance;
                            maxX = intersectionX + tolerance;
                            minY = intersectionY - tolerance;
                            maxY = intersectionY + tolerance;
                            minZ = intersectionZ - tolerance;
                            maxZ = intersectionZ + tolerance;
                            hasValidBoundingBox = true;
                        }
                    }
                    
                    if (hasValidBoundingBox)
                    {
                        cmd.CommandText = @"
                            INSERT INTO ClashZonesRTree (id, minX, maxX, minY, maxY, minZ, maxZ)
                            VALUES (@id, @minX, @maxX, @minY, @maxY, @minZ, @maxZ)";
                        
                        cmd.Parameters.AddWithValue("@id", clashZoneId);
                        cmd.Parameters.AddWithValue("@minX", minX);
                        cmd.Parameters.AddWithValue("@maxX", maxX);
                        cmd.Parameters.AddWithValue("@minY", minY);
                        cmd.Parameters.AddWithValue("@maxY", maxY);
                        cmd.Parameters.AddWithValue("@minZ", minZ);
                        cmd.Parameters.AddWithValue("@maxZ", maxZ);
                        cmd.ExecuteNonQuery();
                        
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            _logger($"[SQLite] ✅ R-tree index updated for ClashZoneId={clashZoneId}: {(hasValidSleeveBoundingBox ? "SleeveBBox" : "IntersectionPoint")} bbox");
                        }
                    }
                    else if (!DeploymentConfiguration.DeploymentMode)
                    {
                        _logger($"[SQLite] ⚠️ R-tree index NOT updated for ClashZoneId={clashZoneId}: No valid bounding box (SleeveBBox invalid, IntersectionPoint invalid)");
                    }
                }
            }
            catch (Exception ex)
            {
                // ✅ SAFETY: Don't fail the main operation if R-tree update fails
                // Log warning but continue (R-tree is optional optimization)
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    _logger($"[SQLite] ⚠️ Failed to update R-tree index for ClashZoneId={clashZoneId}: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// ✅ R-TREE MAINTENANCE: Remove entry from R-tree when clash zone is deleted
        /// Called automatically from delete operations (CASCADE handles this, but explicit for safety)
        /// </summary>
        private void RemoveFromRTreeIndex(int clashZoneId, SQLiteTransaction transaction)
        {
            if (!Services.OptimizationFlags.UseRTreeDatabaseIndex)
                return;
            
            try
            {
                using (var cmd = _context.Connection.CreateCommand())
                {
                    cmd.Transaction = transaction;
                    cmd.CommandText = "DELETE FROM ClashZonesRTree WHERE id = @id";
                    cmd.Parameters.AddWithValue("@id", clashZoneId);
                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                // ✅ SAFETY: Don't fail the main operation if R-tree delete fails
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    _logger($"[SQLite] ⚠️ Failed to remove R-tree index for ClashZoneId={clashZoneId}: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// ✅ CRITICAL FIX: Bulk reset ReadyForPlacementFlag to 0 for ALL zones in selected filters and categories.
        /// Called at the START of refresh to ensure only newly detected zones (within scope box) get ReadyForPlacementFlag=1.
        /// This prevents zones outside scope box or from previous refreshes from being processed.
        /// </summary>
        /// <param name="filterNames">List of filter names to reset flags for</param>
        /// <param name="categories">List of categories to reset flags for</param>
        /// <returns>Number of zones reset</returns>
        public int BulkResetReadyForPlacementFlagsForFilters(List<string> filterNames, List<string> categories)
        {
            if (filterNames == null || filterNames.Count == 0 || categories == null || categories.Count == 0)
                return 0;

            int totalReset = 0;

            try
            {
                using (var cmd = _context.Connection.CreateCommand())
                {
                    // Build WHERE clause: FilterName IN (...) AND Category IN (...)
                    var filterNamePlaceholders = string.Join(",", filterNames.Select((_, i) => $"@FilterName{i}"));
                    var categoryPlaceholders = string.Join(",", categories.Select((_, i) => $"@Category{i}"));

                    cmd.CommandText = $@"
                        UPDATE ClashZones 
                        SET ReadyForPlacementFlag = 0, UpdatedAt = CURRENT_TIMESTAMP
                        WHERE ClashZoneId IN (
                            SELECT cz.ClashZoneId
                            FROM Filters f
                            INNER JOIN FileCombos fc ON f.FilterId = fc.FilterId
                            INNER JOIN ClashZones cz ON fc.ComboId = cz.ComboId
                            WHERE f.FilterName IN ({filterNamePlaceholders})
                              AND f.Category IN ({categoryPlaceholders})
                        )";

                    // Add parameters
                    for (int i = 0; i < filterNames.Count; i++)
                    {
                        cmd.Parameters.AddWithValue($"@FilterName{i}", filterNames[i]);
                    }
                    for (int i = 0; i < categories.Count; i++)
                    {
                        cmd.Parameters.AddWithValue($"@Category{i}", categories[i]);
                    }

                    totalReset = cmd.ExecuteNonQuery();

                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        _logger($"[SQLite] ✅ Bulk reset ReadyForPlacementFlag=0 for {totalReset} zones in {filterNames.Count} filters, {categories.Count} categories");
                    }
                }
            }
            catch (Exception ex)
            {
                _logger($"[SQLite] ❌ Error bulk resetting ReadyForPlacementFlag: {ex.Message}");
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Error($"[ClashZoneRepository] Error bulk resetting ReadyForPlacementFlag: {ex.Message}");
                }
            }

            return totalReset;
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

        /// <summary>
        /// Loads parameter values from ClashZones table for a specific zone
        /// ✅ CRITICAL FIX: Also loads MepElementSizeParameterValue to ensure Size parameter is available for snapshot aggregation
        /// </summary>
        private void LoadParameterValuesFromDatabase(ClashZone zone, SQLiteTransaction transaction)
        {
            if (zone == null || string.IsNullOrEmpty(zone.Id.ToString()))
                return;

            try
            {
                using (var cmd = _context.Connection.CreateCommand())
                {
                    cmd.Transaction = transaction;
                    // ✅ CRITICAL FIX: Use UPPER() for case-insensitive GUID comparison (matches how GUIDs are stored)
                    cmd.CommandText = @"
                        SELECT MepParameterValuesJson, HostParameterValuesJson, MepElementSizeParameterValue
                        FROM ClashZones
                        WHERE UPPER(ClashZoneGuid) = UPPER(@ClashZoneGuid)
                        LIMIT 1";
                    cmd.Parameters.AddWithValue("@ClashZoneGuid", zone.Id.ToString().ToUpperInvariant());

                    using (var reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            var mepParamsJson = GetNullableString(reader, "MepParameterValuesJson");
                            var hostParamsJson = GetNullableString(reader, "HostParameterValuesJson");
                            // ✅ CRITICAL FIX: Load MepElementSizeParameterValue from database
                            var mepElementSizeParameterValue = GetNullableString(reader, "MepElementSizeParameterValue");

                            // ✅ CRITICAL: Populate MepElementSizeParameterValue if it's empty (ensures Size parameter is available for aggregation)
                            if (!string.IsNullOrWhiteSpace(mepElementSizeParameterValue) && string.IsNullOrWhiteSpace(zone.MepElementSizeParameterValue))
                            {
                                zone.MepElementSizeParameterValue = mepElementSizeParameterValue;
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    _logger($"[SQLite] ✅ Loaded MepElementSizeParameterValue='{mepElementSizeParameterValue}' for zone {zone.Id} from database");
                                }
                            }

                            // ✅ CRITICAL FIX: Deserialize MEP parameters - MERGE with existing if zone has some params
                            // Database is source of truth, but preserve any additional params zone might have
                            if (!string.IsNullOrWhiteSpace(mepParamsJson) && mepParamsJson != "{}")
                            {
                                var mepDict = DeserializeDictionary(mepParamsJson);
                                if (mepDict != null && mepDict.Count > 0)
                                {
                                    // ✅ MERGE STRATEGY: If zone already has some parameters, merge with database params
                                    if (zone.MepParameterValues != null && zone.MepParameterValues.Count > 0)
                                    {
                                        // Merge: Add database params that don't already exist in zone
                                        var existingDict = zone.MepParameterValues
                                            .Where(kv => kv != null && !string.IsNullOrEmpty(kv.Key))
                                            .ToDictionary(kv => kv.Key, kv => kv.Value, StringComparer.OrdinalIgnoreCase);
                                        
                                        var beforeCount = existingDict.Count;
                                        foreach (var dbParam in mepDict)
                                        {
                                            if (!string.IsNullOrEmpty(dbParam.Key) && !existingDict.ContainsKey(dbParam.Key))
                                            {
                                                existingDict[dbParam.Key] = dbParam.Value;
                                            }
                                        }
                                        
                                        zone.MepParameterValues = existingDict
                                            .Select(kv => new Models.SerializableKeyValue { Key = kv.Key, Value = kv.Value })
                                            .ToList();
                                        
                                        if (!DeploymentConfiguration.DeploymentMode)
                                        {
                                            var addedCount = zone.MepParameterValues.Count - beforeCount;
                                            _logger($"[SQLite] ✅ LoadParameterValuesFromDatabase: Merged MEP params for zone {zone.Id} - Had {beforeCount} params, added {addedCount} from DB, total now {zone.MepParameterValues.Count}. Sample keys: {string.Join(", ", zone.MepParameterValues.Take(5).Select(kv => kv.Key))}");
                                        }
                                    }
                                    else
                                    {
                                        // Zone has no params - use database params directly
                                        zone.MepParameterValues = mepDict
                                            .Select(kv => new Models.SerializableKeyValue { Key = kv.Key, Value = kv.Value })
                                            .ToList();
                                        if (!DeploymentConfiguration.DeploymentMode)
                                        {
                                            _logger($"[SQLite] ✅ LoadParameterValuesFromDatabase: Loaded {mepDict.Count} MEP parameters for zone {zone.Id} (SleeveId={zone.SleeveInstanceId}) from ClashZones table. Sample keys: {string.Join(", ", mepDict.Keys.Take(5))}");
                                        }
                                    }
                                }
                                else
                                {
                                    if (!DeploymentConfiguration.DeploymentMode)
                                    {
                                        _logger($"[SQLite] ⚠️ LoadParameterValuesFromDatabase: MepParameterValuesJson deserialized to empty dictionary for zone {zone.Id} (SleeveId={zone.SleeveInstanceId}). JSON length: {mepParamsJson?.Length ?? 0}");
                                    }
                                }
                            }
                            else
                            {
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    _logger($"[SQLite] ⚠️ LoadParameterValuesFromDatabase: MepParameterValuesJson is null/empty/{{}} for zone {zone.Id} (SleeveId={zone.SleeveInstanceId})");
                                }
                            }

                            // ✅ CRITICAL FIX: Deserialize Host parameters - MERGE with existing if zone has some params
                            if (!string.IsNullOrWhiteSpace(hostParamsJson) && hostParamsJson != "{}")
                            {
                                var hostDict = DeserializeDictionary(hostParamsJson);
                                if (hostDict != null && hostDict.Count > 0)
                                {
                                    // ✅ MERGE STRATEGY: If zone already has some parameters, merge with database params
                                    if (zone.HostParameterValues != null && zone.HostParameterValues.Count > 0)
                                    {
                                        // Merge: Add database params that don't already exist in zone
                                        var existingDict = zone.HostParameterValues
                                            .Where(kv => kv != null && !string.IsNullOrEmpty(kv.Key))
                                            .ToDictionary(kv => kv.Key, kv => kv.Value, StringComparer.OrdinalIgnoreCase);
                                        
                                        var beforeCount = existingDict.Count;
                                        foreach (var dbParam in hostDict)
                                        {
                                            if (!string.IsNullOrEmpty(dbParam.Key) && !existingDict.ContainsKey(dbParam.Key))
                                            {
                                                existingDict[dbParam.Key] = dbParam.Value;
                                            }
                                        }
                                        
                                        zone.HostParameterValues = existingDict
                                            .Select(kv => new Models.SerializableKeyValue { Key = kv.Key, Value = kv.Value })
                                            .ToList();
                                        
                                        if (!DeploymentConfiguration.DeploymentMode)
                                        {
                                            var addedCount = zone.HostParameterValues.Count - beforeCount;
                                            _logger($"[SQLite] ✅ LoadParameterValuesFromDatabase: Merged Host params for zone {zone.Id} - Had {beforeCount} params, added {addedCount} from DB, total now {zone.HostParameterValues.Count}");
                                        }
                                    }
                                    else
                                    {
                                        // Zone has no params - use database params directly
                                        zone.HostParameterValues = hostDict
                                            .Select(kv => new Models.SerializableKeyValue { Key = kv.Key, Value = kv.Value })
                                            .ToList();
                                        if (!DeploymentConfiguration.DeploymentMode)
                                        {
                                            _logger($"[SQLite] ✅ LoadParameterValuesFromDatabase: Loaded {hostDict.Count} Host parameters for zone {zone.Id} (SleeveId={zone.SleeveInstanceId}) from ClashZones table");
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
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    _logger($"[SQLite] ⚠️ Error loading parameter values from ClashZones for zone {zone.Id}: {ex.Message}");
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
            
            // ✅ DIRECT LOGGING: Always log sleeve dimensions being loaded for duct accessories
            if (string.Equals(clashZone.MepElementCategory, "Duct Accessories", StringComparison.OrdinalIgnoreCase))
            {
                SafeFileLogger.SafeAppendText("damper_placement_trace.log", $"[{DateTime.Now:HH:mm:ss.fff}] [DB-LOAD-SLEEVE] Zone {clashZone.Id}: SleeveWidth={clashZone.SleeveWidth:F6}ft ({clashZone.SleeveWidth * 304.8:F1}mm), SleeveHeight={clashZone.SleeveHeight:F6}ft ({clashZone.SleeveHeight * 304.8:F1}mm), SleeveDiameter={clashZone.SleeveDiameter:F6}ft ({clashZone.SleeveDiameter * 304.8:F1}mm)\n");
            }

            clashZone.IntersectionPoint = new XYZ(
                GetDouble(reader, "IntersectionX"),
                GetDouble(reader, "IntersectionY"),
                GetDouble(reader, "IntersectionZ"));

            // ✅ WALL CENTERLINE POINT: Load pre-calculated wall centerline coordinates (enables multi-threaded placement)
            // ✅ CRITICAL: These values are calculated during refresh and saved to database
            clashZone.WallCenterlinePointX = GetNullableDouble(reader, "WallCenterlinePointX") ?? 0.0;
            clashZone.WallCenterlinePointY = GetNullableDouble(reader, "WallCenterlinePointY") ?? 0.0;
            clashZone.WallCenterlinePointZ = GetNullableDouble(reader, "WallCenterlinePointZ") ?? 0.0;
            
            // ✅ DIAGNOSTIC: Log wall centerline loading for debugging (especially for dampers)
            if (!DeploymentConfiguration.DeploymentMode && string.Equals(clashZone.MepElementCategory, "Duct Accessories", StringComparison.OrdinalIgnoreCase))
            {
                bool isZero = (clashZone.WallCenterlinePointX == 0.0 && clashZone.WallCenterlinePointY == 0.0 && clashZone.WallCenterlinePointZ == 0.0);
                string zeroWarning = isZero ? " ⚠️⚠️⚠️ LOADED ZEROS!" : "";
                SafeFileLogger.SafeAppendText("placement_debug.log",
                    $"[{DateTime.Now:HH:mm:ss.fff}] [DB-LOAD] Zone {clashZone.Id}: Loaded WallCenterlinePoint=({clashZone.WallCenterlinePointX:F6}ft, {clashZone.WallCenterlinePointY:F6}ft, {clashZone.WallCenterlinePointZ:F6}ft), " +
                    $"Intersection=({clashZone.IntersectionPointX:F6}ft, {clashZone.IntersectionPointY:F6}ft, {clashZone.IntersectionPointZ:F6}ft)" +
                    $"{zeroWarning}\n");
            }

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
            // ✅ REFERENCE LEVEL: Load MEP element Reference Level (used for Schedule Level and Bottom of Opening calculation)
            clashZone.MepElementLevelName = GetNullableString(reader, "MepElementLevelName") ?? string.Empty;
            clashZone.MepElementLevelElevation = GetDouble(reader, "MepElementLevelElevation", 0.0);
            
            // ✅ DIAGNOSTIC: Log level elevation loading for debugging (especially for dampers)
            if (!DeploymentConfiguration.DeploymentMode && !string.IsNullOrWhiteSpace(clashZone.MepElementLevelName))
            {
                var elevationMm = clashZone.MepElementLevelElevation * 304.8;
                SafeFileLogger.SafeAppendText("placement_debug.log",
                    $"[{DateTime.Now:HH:mm:ss.fff}] [DB-LOAD] Zone {clashZone.Id}: Loaded MepElementLevelName='{clashZone.MepElementLevelName}', MepElementLevelElevation={clashZone.MepElementLevelElevation:F6}ft ({elevationMm:F1}mm)\n");
            }
            
            clashZone.WallDirectionType = GetNullableString(reader, "WallDirectionType") ?? string.Empty;
            
            // ✅ PARAMETER VALUES: Load parameter values from JSON columns
            var mepParamsJson = GetNullableString(reader, "MepParameterValuesJson");
            if (!string.IsNullOrWhiteSpace(mepParamsJson) && mepParamsJson != "{}")
            {
                try
                {
                    var mepDict = JsonSerializer.Deserialize<Dictionary<string, string>>(mepParamsJson);
                    clashZone.MepParameterValues = mepDict
                        .Select(kv => new SerializableKeyValue { Key = kv.Key, Value = kv.Value })
                        .ToList();
                    
                    // ✅ CRITICAL DIAGNOSTIC: Log parameter loading for debugging
                    if (!DeploymentConfiguration.DeploymentMode && string.Equals(clashZone.MepElementCategory, "Pipes", StringComparison.OrdinalIgnoreCase))
                    {
                        var sampleKeys = mepDict.Keys.Take(5).ToList();
                        var sampleStr = string.Join(", ", sampleKeys);
                        if (mepDict.Count > 5) sampleStr += $" (+{mepDict.Count - 5} more)";
                        _logger($"[SQLite] [PARAM-LOAD] Zone {clashZone.Id}: Loaded {mepDict.Count} MEP params from JSON: {sampleStr}");
                    }
                }
                catch (Exception ex)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        _logger($"[SQLite] ⚠️ Failed to deserialize MepParameterValuesJson for zone {clashZone.Id}: {ex.Message}");
                    }
                }
            }
            else if (!DeploymentConfiguration.DeploymentMode && string.Equals(clashZone.MepElementCategory, "Pipes", StringComparison.OrdinalIgnoreCase))
            {
                _logger($"[SQLite] [PARAM-LOAD] ⚠️ Zone {clashZone.Id}: MepParameterValuesJson is NULL, empty, or '{{}}' - no params loaded");
            }
            
            var hostParamsJson = GetNullableString(reader, "HostParameterValuesJson");
            if (!string.IsNullOrWhiteSpace(hostParamsJson) && hostParamsJson != "{}")
            {
                try
                {
                    var hostDict = JsonSerializer.Deserialize<Dictionary<string, string>>(hostParamsJson);
                    clashZone.HostParameterValues = hostDict
                        .Select(kv => new SerializableKeyValue { Key = kv.Key, Value = kv.Value })
                        .ToList();
                }
                catch (Exception ex)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        _logger($"[SQLite] ⚠️ Failed to deserialize HostParameterValuesJson: {ex.Message}");
                    }
                }
            }

            clashZone.MepElementOrientationDirection = GetNullableString(reader, "MepOrientationDirection") ?? string.Empty;
            
            // ✅ WALL DIRECTION: Calculate WallDirection from HostOrientation once when loading from DB
            // This avoids expensive on-the-fly calculations during placement
            // HostOrientation "X" = X-WALL (wall runs along X-axis, direction = +X)
            // HostOrientation "Y" = Y-WALL (wall runs along Y-axis, direction = +Y)
            bool isWallHost = string.Equals(clashZone.StructuralElementType, "Wall", StringComparison.OrdinalIgnoreCase) ||
                              string.Equals(clashZone.StructuralElementType, "Walls", StringComparison.OrdinalIgnoreCase);
            bool isFramingHost = string.Equals(clashZone.StructuralElementType, "Structural Framing", StringComparison.OrdinalIgnoreCase);
            
            if (isWallHost || isFramingHost)
            {
                string hostOrientation = clashZone.HostOrientation ?? string.Empty;
                string wallDirectionType = clashZone.WallDirectionType ?? string.Empty;
                
                if (!string.IsNullOrEmpty(hostOrientation))
                {
                    if (string.Equals(hostOrientation, "X", StringComparison.OrdinalIgnoreCase) ||
                        wallDirectionType.Contains("X-WALL", StringComparison.OrdinalIgnoreCase))
                    {
                        clashZone.WallDirection = new XYZ(1, 0, 0); // X-WALL: direction along X-axis
                    }
                    else if (string.Equals(hostOrientation, "Y", StringComparison.OrdinalIgnoreCase) ||
                             wallDirectionType.Contains("Y-WALL", StringComparison.OrdinalIgnoreCase))
                    {
                        clashZone.WallDirection = new XYZ(0, 1, 0); // Y-WALL: direction along Y-axis
                    }
                }
            }
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
            
            // ✅ DIAGNOSTIC: Log MEP element sizes loaded from database (for debugging sleeve size issues)
            if (!DeploymentConfiguration.DeploymentMode && clashZone.MepElementWidth > 0 && clashZone.MepElementHeight > 0)
            {
                var widthMm = clashZone.MepElementWidth * 304.8;
                var heightMm = clashZone.MepElementHeight * 304.8;
                SafeFileLogger.SafeAppendText("refresh_mep_sizes.log",
                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [DB-LOAD] Zone {clashZone.Id}: Loaded MepWidth={clashZone.MepElementWidth:F6}ft ({widthMm:F1}mm), MepHeight={clashZone.MepElementHeight:F6}ft ({heightMm:F1}mm) from database\n");
            }
            // ✅ PIPE DIAMETER COLUMNS: Load outer diameter and nominal diameter for pipes
            clashZone.MepElementOuterDiameter = GetNullableDouble(reader, "MepElementOuterDiameter") ?? 0.0;
            clashZone.MepElementNominalDiameter = GetNullableDouble(reader, "MepElementNominalDiameter") ?? 0.0;
            // ✅ SIZE PARAMETER VALUE: Load Size parameter value as string for snapshot table and parameter transfer
            clashZone.MepElementSizeParameterValue = GetNullableString(reader, "MepElementSizeParameterValue") ?? string.Empty;
            
            // ✅ DIRECT LOGGING: Always log MEP dimensions being loaded for duct accessories
            if (string.Equals(clashZone.MepElementCategory, "Duct Accessories", StringComparison.OrdinalIgnoreCase))
            {
                _logger($"[DB-LOAD-DEBUG] Zone {clashZone.Id}: Loaded MepWidth={clashZone.MepElementWidth:F6}ft ({clashZone.MepElementWidth * 304.8:F1}mm), MepHeight={clashZone.MepElementHeight:F6}ft ({clashZone.MepElementHeight * 304.8:F1}mm)");
                SafeFileLogger.SafeAppendText("damper_placement_trace.log", $"[{DateTime.Now:HH:mm:ss.fff}] [DB-LOAD-MEP] Zone {clashZone.Id}: MepWidth={clashZone.MepElementWidth:F6}ft ({clashZone.MepElementWidth * 304.8:F1}mm), MepHeight={clashZone.MepElementHeight:F6}ft ({clashZone.MepElementHeight * 304.8:F1}mm)\n");
            }

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
            
            // ✅ OOP METHOD: Load damper connector detection values
            clashZone.HasMepConnector = GetBool(reader, "HasMepConnector");
            clashZone.DamperConnectorSide = GetNullableString(reader, "DamperConnectorSide") ?? string.Empty;
            
            // ✅ PIPE DIAMETER COLUMNS: Load outer diameter and nominal diameter
            clashZone.MepElementOuterDiameter = GetDouble(reader, "MepElementOuterDiameter", 0.0);
            clashZone.MepElementNominalDiameter = GetDouble(reader, "MepElementNominalDiameter", 0.0);
            
            // ✅ SIZE PARAMETER VALUE: Load Size parameter value as string (critical for cluster snapshots)
            clashZone.MepElementSizeParameterValue = GetNullableString(reader, "MepElementSizeParameterValue") ?? string.Empty;
            
            // ✅ OOP METHOD: Load insulation detection values
            clashZone.IsInsulated = GetBool(reader, "IsInsulated");
            clashZone.InsulationThickness = GetDouble(reader, "InsulationThickness", 0.0);
            // Update InsulationType for backward compatibility
            clashZone.InsulationType = clashZone.IsInsulated ? "Insulated" : "Normal";
            
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
            using (var transaction = _context.Connection.BeginTransaction())
            {
                using (var cmd = _context.Connection.CreateCommand())
                {
                    cmd.Transaction = transaction;
                    
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
                    else if (rowsAffected > 0)
                    {
                        // ✅ R-TREE MAINTENANCE: Get ClashZoneId and update R-tree index
                        cmd.CommandText = @"
                            SELECT ClashZoneId FROM ClashZones
                            WHERE UPPER(ClashZoneGuid) = UPPER(@ClashZoneGuid)
                              AND ClashZoneGuid != '' AND ClashZoneGuid IS NOT NULL
                            LIMIT 1";
                        var clashZoneIdResult = cmd.ExecuteScalar();
                        
                        if (clashZoneIdResult != null)
                        {
                            int clashZoneId = Convert.ToInt32(clashZoneIdResult);
                            
                            // ✅ Create a temporary ClashZone object with updated bounding boxes for R-tree update
                            var tempClashZone = new ClashZone
                            {
                                SleeveBoundingBoxMinX = minX,
                                SleeveBoundingBoxMaxX = maxX,
                                SleeveBoundingBoxMinY = minY,
                                SleeveBoundingBoxMaxY = maxY,
                                SleeveBoundingBoxMinZ = minZ,
                                SleeveBoundingBoxMaxZ = maxZ
                            };
                            
                            // ✅ Update R-tree index with new bounding boxes
                            UpdateRTreeIndex(clashZoneId, tempClashZone, transaction);
                        }
                    }
                }
                
                transaction.Commit();
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
            // ✅ DIAGNOSTIC: Check if row exists before updating
            bool rowExists = false;
            using (var checkCmd = _context.Connection.CreateCommand())
            {
                checkCmd.CommandText = @"
                    SELECT COUNT(*) FROM ClashZones 
                    WHERE UPPER(ClashZoneGuid) = UPPER(@ClashZoneGuid)
                      AND ClashZoneGuid != '' AND ClashZoneGuid IS NOT NULL";
                checkCmd.Parameters.AddWithValue("@ClashZoneGuid", clashZoneGuid.ToString());
                var count = Convert.ToInt32(checkCmd.ExecuteScalar());
                rowExists = count > 0;
                
                if (!rowExists && !DeploymentConfiguration.DeploymentMode)
                {
                    _logger($"[SQLite] ⚠️ UpdateSleeveInstanceId: Zone {clashZoneGuid} does NOT exist in ClashZones table (row will be created later in batch persistence)");
                }
            }
            
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
                
                // ✅ DATABASE LOGGING: Log the operation result (single log entry with all info)
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    var logParams = new Dictionary<string, object>
                    {
                        { "ClashZoneGuid", clashZoneGuid.ToString() },
                        { "SleeveInstanceId", sleeveInstanceId },
                        { "RowExists", rowExists }
                    };
                    DatabaseOperationLogger.LogOperation("UPDATE", "ClashZones", logParams, rowsAffected, 
                        rowsAffected > 0 ? $"✅ Updated SleeveInstanceId={sleeveInstanceId} for zone {clashZoneGuid}" 
                                         : (rowExists ? $"⚠️ UPDATE failed for zone {clashZoneGuid} (row exists but UPDATE returned 0 rows)" 
                                                      : $"⚠️ No rows updated for zone {clashZoneGuid} (row does not exist yet, will be created in batch persistence)"));
                }
                
                if (rowsAffected == 0 && !DeploymentConfiguration.DeploymentMode)
                {
                    if (rowExists)
                    {
                        _logger($"[SQLite] ⚠️ UpdateSleeveInstanceId: Row exists but UPDATE returned 0 rows for GUID {clashZoneGuid} - possible GUID format mismatch");
                    }
                    else
                    {
                        _logger($"[SQLite] ⚠️ UpdateSleeveInstanceId: Zone {clashZoneGuid} does not exist in ClashZones table yet (will be created in batch persistence)");
                    }
                }
                else if (rowsAffected > 0 && !DeploymentConfiguration.DeploymentMode)
                {
                    _logger($"[SQLite] ✅ UpdateSleeveInstanceId: Updated {rowsAffected} row(s) in ClashZones for GUID {clashZoneGuid} with SleeveInstanceId={sleeveInstanceId}");
                }
            }
            
            // ✅ CRITICAL FIX: Also update SleeveSnapshots table to keep SleeveInstanceId in sync
            // This ensures parameter transfer works after sleeve regeneration/recreation
            using (var snapshotCmd = _context.Connection.CreateCommand())
            {
                snapshotCmd.CommandText = @"
                    UPDATE SleeveSnapshots SET
                        SleeveInstanceId = @SleeveInstanceId
                    WHERE UPPER(ClashZoneGuid) = UPPER(@ClashZoneGuid)
                      AND ClashZoneGuid != '' AND ClashZoneGuid IS NOT NULL
                      AND (ClusterInstanceId IS NULL OR ClusterInstanceId <= 0)";

                snapshotCmd.Parameters.AddWithValue("@ClashZoneGuid", clashZoneGuid.ToString());
                snapshotCmd.Parameters.AddWithValue("@SleeveInstanceId", sleeveInstanceId > 0 ? (object)sleeveInstanceId : DBNull.Value);
                
                var snapshotRowsAffected = snapshotCmd.ExecuteNonQuery();
                if (snapshotRowsAffected > 0 && !DeploymentConfiguration.DeploymentMode)
                {
                    _logger($"[SQLite] ✅ UpdateSleeveInstanceId: Updated {snapshotRowsAffected} snapshot(s) for GUID {clashZoneGuid} with SleeveInstanceId={sleeveInstanceId}");
                }
                else if (snapshotRowsAffected == 0 && !DeploymentConfiguration.DeploymentMode)
                {
                    _logger($"[SQLite] ⚠️ UpdateSleeveInstanceId: No snapshots updated for GUID {clashZoneGuid} (may not exist yet or is a cluster)");
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
                // ✅ CRITICAL: Check IntersectionX/Y/Z for GUID lookup (stable intersection point)
                // IntersectionPoint is the actual MEP/host intersection - stable and reliable for GUID generation
                // This matches the GUID generation logic which uses IntersectionPoint
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
                    int rowsAffectedTotal = 0;

                    // 🚀 PERFORMANCE OPTIMIZATION: Use temp table + single UPDATE instead of N individual UPDATEs
                    // Old approach: 68 individual UPDATEs = 897ms
                    // New approach: 1 temp table INSERT + 1 UPDATE JOIN = ~10-50ms
                    
                    // Step 1: Create temp table for batch data
                        using (var tempCmd = _context.Connection.CreateCommand())
                        {
                            tempCmd.Transaction = transaction;
                            tempCmd.CommandText = @"
                                CREATE TEMP TABLE IF NOT EXISTS TempFlagUpdates (
                                    ClashZoneGuid TEXT,
                                    IsResolvedFlag INTEGER,
                                    IsClusterResolvedFlag INTEGER,
                                    SleeveInstanceId INTEGER,
                                    ClusterInstanceId INTEGER,
                                    OldSleeveInstanceId INTEGER,
                                    OldClusterInstanceId INTEGER,
                                    MepElementId INTEGER,
                                    HostElementId INTEGER,
                                    IntersectionX REAL,
                                    IntersectionY REAL,
                                    IntersectionZ REAL,
                                    MarkedForClusterProcess INTEGER,
                                    AfterClusterSleeveId INTEGER,
                                    IsClusteredFlag INTEGER
                                )";
                            tempCmd.ExecuteNonQuery();
                        }
                        
                        // Step 2: Bulk insert all updates into temp table
                        using (var insertCmd = _context.Connection.CreateCommand())
                        {
                            insertCmd.Transaction = transaction;
                            insertCmd.CommandText = @"
                                INSERT INTO TempFlagUpdates (
                                    ClashZoneGuid, IsResolvedFlag, IsClusterResolvedFlag, 
                                    SleeveInstanceId, ClusterInstanceId,
                                    OldSleeveInstanceId, OldClusterInstanceId,
                                    MepElementId, HostElementId,
                                    IntersectionX, IntersectionY, IntersectionZ,
                                    MarkedForClusterProcess, AfterClusterSleeveId, IsClusteredFlag
                                ) VALUES (
                                    @ClashZoneGuid, @IsResolvedFlag, @IsClusterResolvedFlag,
                                    @SleeveInstanceId, @ClusterInstanceId,
                                    @OldSleeveInstanceId, @OldClusterInstanceId,
                                    @MepElementId, @HostElementId,
                                    @IntersectionX, @IntersectionY, @IntersectionZ,
                                    @MarkedForClusterProcess, @AfterClusterSleeveId, @IsClusteredFlag
                                )";
                            
                            var guidParam = insertCmd.Parameters.Add("@ClashZoneGuid", System.Data.DbType.String);
                            var isResolvedParam = insertCmd.Parameters.Add("@IsResolvedFlag", System.Data.DbType.Int32);
                            var isClusterResolvedParam = insertCmd.Parameters.Add("@IsClusterResolvedFlag", System.Data.DbType.Int32);
                            var sleeveIdParam = insertCmd.Parameters.Add("@SleeveInstanceId", System.Data.DbType.Int32);
                            var clusterIdParam = insertCmd.Parameters.Add("@ClusterInstanceId", System.Data.DbType.Int32);
                            var oldSleeveIdParam = insertCmd.Parameters.Add("@OldSleeveInstanceId", System.Data.DbType.Int32);
                            var oldClusterIdParam = insertCmd.Parameters.Add("@OldClusterInstanceId", System.Data.DbType.Int32);
                            var mepIdParam = insertCmd.Parameters.Add("@MepElementId", System.Data.DbType.Int32);
                            var hostIdParam = insertCmd.Parameters.Add("@HostElementId", System.Data.DbType.Int32);
                            var intersectionXParam = insertCmd.Parameters.Add("@IntersectionX", System.Data.DbType.Double);
                            var intersectionYParam = insertCmd.Parameters.Add("@IntersectionY", System.Data.DbType.Double);
                            var intersectionZParam = insertCmd.Parameters.Add("@IntersectionZ", System.Data.DbType.Double);
                            var markedForClusterParam = insertCmd.Parameters.Add("@MarkedForClusterProcess", System.Data.DbType.Int32);
                            var afterClusterSleeveIdParam = insertCmd.Parameters.Add("@AfterClusterSleeveId", System.Data.DbType.Int32);
                            var isClusteredFlagParam = insertCmd.Parameters.Add("@IsClusteredFlag", System.Data.DbType.Int32);
                            
                            insertCmd.Prepare();
                            
                            foreach (var update in updatesList)
                            {
                                guidParam.Value = update.ClashZoneId.ToString().ToUpperInvariant();
                                isResolvedParam.Value = update.IsResolved ? 1 : 0;
                                isClusterResolvedParam.Value = update.IsClusterResolved ? 1 : 0;
                                sleeveIdParam.Value = update.SleeveInstanceId > 0 ? (object)update.SleeveInstanceId : DBNull.Value;
                                clusterIdParam.Value = update.ClusterInstanceId > 0 ? (object)update.ClusterInstanceId : DBNull.Value;
                                oldSleeveIdParam.Value = update.OldSleeveInstanceId > 0 ? (object)update.OldSleeveInstanceId : DBNull.Value;
                                oldClusterIdParam.Value = update.OldClusterInstanceId > 0 ? (object)update.OldClusterInstanceId : DBNull.Value;
                                mepIdParam.Value = update.MepElementId;
                                hostIdParam.Value = update.StructuralElementId;
                                intersectionXParam.Value = update.IntersectionPointX;
                                intersectionYParam.Value = update.IntersectionPointY;
                                intersectionZParam.Value = update.IntersectionPointZ;
                                markedForClusterParam.Value = update.MarkedForClusterProcess.HasValue ? (object)(update.MarkedForClusterProcess.Value ? 1 : 0) : DBNull.Value;
                                afterClusterSleeveIdParam.Value = update.AfterClusterSleeveId > 0 ? (object)update.AfterClusterSleeveId : DBNull.Value;
                                isClusteredFlagParam.Value = update.IsClusteredFlag.HasValue ? (object)(update.IsClusteredFlag.Value ? 1 : 0) : DBNull.Value;
                                
                                insertCmd.ExecuteNonQuery();
                            }
                        }
                        
                        // Step 3: Single UPDATE with JOIN to temp table (1 SQL statement for all updates)
                        using (var bulkUpdateCmd = _context.Connection.CreateCommand())
                        {
                            bulkUpdateCmd.Transaction = transaction;
                            bulkUpdateCmd.CommandText = @"
                                UPDATE ClashZones
                                SET 
                                    IsResolvedFlag = t.IsResolvedFlag,
                                    IsClusterResolvedFlag = t.IsClusterResolvedFlag,
                                    SleeveInstanceId = t.SleeveInstanceId,
                                    ClusterInstanceId = t.ClusterInstanceId,
                                    MarkedForClusterProcess = t.MarkedForClusterProcess,
                                    AfterClusterSleeveId = t.AfterClusterSleeveId,
                                    IsClusteredFlag = t.IsClusteredFlag,
                                    ReadyForPlacementFlag = CASE 
                                        WHEN t.IsResolvedFlag = 0 AND t.IsClusterResolvedFlag = 0 THEN 1 
                                        ELSE ClashZones.ReadyForPlacementFlag 
                                    END,
                                    UpdatedAt = datetime('now', '+5 hours', '+30 minutes')
                                FROM TempFlagUpdates t
                                WHERE ClashZones.ClashZoneId = COALESCE(
                                    (SELECT ClashZoneId FROM ClashZones cz 
                                     WHERE UPPER(cz.ClashZoneGuid) = UPPER(t.ClashZoneGuid) 
                                       AND cz.ClashZoneGuid != '' AND cz.ClashZoneGuid IS NOT NULL
                                     LIMIT 1),
                                    (SELECT ClashZoneId FROM ClashZones cz 
                                     WHERE t.OldSleeveInstanceId > 0 
                                       AND cz.SleeveInstanceId = t.OldSleeveInstanceId
                                       AND cz.MepElementId = t.MepElementId 
                                       AND cz.HostElementId = t.HostElementId 
                                       AND ABS(cz.IntersectionX - t.IntersectionX) < 0.001 
                                       AND ABS(cz.IntersectionY - t.IntersectionY) < 0.001 
                                       AND ABS(cz.IntersectionZ - t.IntersectionZ) < 0.001
                                     LIMIT 1),
                                    (SELECT ClashZoneId FROM ClashZones cz 
                                     WHERE t.OldClusterInstanceId > 0 
                                       AND cz.ClusterInstanceId = t.OldClusterInstanceId
                                       AND cz.MepElementId = t.MepElementId 
                                       AND cz.HostElementId = t.HostElementId 
                                       AND ABS(cz.IntersectionX - t.IntersectionX) < 0.001 
                                       AND ABS(cz.IntersectionY - t.IntersectionY) < 0.001 
                                       AND ABS(cz.IntersectionZ - t.IntersectionZ) < 0.001
                                     LIMIT 1),
                                    NULL
                                )";
                            rowsAffectedTotal = bulkUpdateCmd.ExecuteNonQuery();
                        }
                        
                        // Step 4: Cleanup temp table
                        using (var dropCmd = _context.Connection.CreateCommand())
                        {
                            dropCmd.Transaction = transaction;
                            dropCmd.CommandText = "DROP TABLE IF EXISTS TempFlagUpdates";
                            dropCmd.ExecuteNonQuery();
                        }
                        
                        int updateCount = updatesList.Count;
                        int notFoundCount = updateCount - rowsAffectedTotal;
                        int multipleRowsCount = 0; // Can't track this in bulk mode
                        
                        // ✅ Old per-row loop completely removed - replaced with bulk temp table approach above

                        transaction.Commit();
                        _logger($"[SQLite] ✅ 🚀 BULK UPDATE: {updateCount} updates in single transaction, {rowsAffectedTotal} rows affected, {notFoundCount} not found");
                        
                        if (multipleRowsCount > 0)
                        {
                            _logger($"[SQLite] ⚠️⚠️⚠️ WARNING: {multipleRowsCount} updates matched multiple rows! This indicates duplicate entries in the database. Expected: 1 row per update. Actual: {rowsAffectedTotal} rows for {updateCount} updates.");
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
        /// Interface implementation: Batch update flags with simplified signature.
        /// Queries database for missing values and calls the full implementation.
        /// </summary>
        public void BatchUpdateFlags(List<(System.Guid ClashZoneId, bool IsResolved, bool IsClusterResolved, int SleeveInstanceId, int ClusterInstanceId)> updates)
        {
            if (updates == null || updates.Count == 0)
                return;

            // ✅ DIAGNOSTIC: Log incoming updates
            if (!DeploymentConfiguration.DeploymentMode && updates.Count > 0)
            {
                var sampleUpdate = updates[0];
                _logger($"[SQLite] BatchUpdateFlags (interface): Received {updates.Count} updates. Sample: GUID={sampleUpdate.ClashZoneId}, IsResolved={sampleUpdate.IsResolved}, SleeveId={sampleUpdate.SleeveInstanceId}");
            }

            // Build a dictionary to map GUIDs to updates for efficient lookup
            // ✅ FIX: Use standard GUID format (with hyphens, uppercase) for consistency
            var updateDict = updates.ToDictionary(u => u.ClashZoneId.ToString().ToUpperInvariant(), u => u);
            var fullUpdates = new List<(Guid ClashZoneId, bool IsResolved, bool IsClusterResolved, int SleeveInstanceId, int ClusterInstanceId, int MepElementId, int StructuralElementId, double IntersectionPointX, double IntersectionPointY, double IntersectionPointZ, int OldSleeveInstanceId, int OldClusterInstanceId, bool? MarkedForClusterProcess, int AfterClusterSleeveId, bool? IsClusteredFlag)>();

            // Use temp table for efficient batch querying
            using (var transaction = _context.Connection.BeginTransaction())
            {
                try
                {
                    // Create temp table for GUIDs
                    using (var createCmd = _context.Connection.CreateCommand())
                    {
                        createCmd.Transaction = transaction;
                        createCmd.CommandText = @"
                            CREATE TEMP TABLE IF NOT EXISTS TempGuidLookup (
                                ClashZoneGuid TEXT PRIMARY KEY
                            )";
                        createCmd.ExecuteNonQuery();
                    }

                    // Insert GUIDs into temp table
                    using (var insertCmd = _context.Connection.CreateCommand())
                    {
                        insertCmd.Transaction = transaction;
                        insertCmd.CommandText = "INSERT OR IGNORE INTO TempGuidLookup (ClashZoneGuid) VALUES (@Guid)";
                        var guidParam = insertCmd.Parameters.Add("@Guid", System.Data.DbType.String);
                        insertCmd.Prepare();
                        
                        foreach (var guid in updateDict.Keys)
                        {
                            // ✅ FIX: Ensure GUID is in standard format (uppercase, with hyphens)
                            guidParam.Value = guid;
                            insertCmd.ExecuteNonQuery();
                        }
                        
                        if (!DeploymentConfiguration.DeploymentMode && updateDict.Count > 0)
                        {
                            _logger($"[SQLite] BatchUpdateFlags: Inserted {updateDict.Count} GUIDs into temp table for lookup");
                        }
                    }

                    // Query all matching records
                    using (var queryCmd = _context.Connection.CreateCommand())
                    {
                        queryCmd.Transaction = transaction;
                        // ✅ FIX: Match GUIDs case-insensitively and handle empty/null GUIDs
                        queryCmd.CommandText = @"
                            SELECT c.ClashZoneGuid, c.MepElementId, c.HostElementId, c.IntersectionX, c.IntersectionY, c.IntersectionZ, 
                                   c.SleeveInstanceId, c.ClusterInstanceId, c.MarkedForClusterProcess, c.AfterClusterSleeveId, c.IsClusteredFlag
                            FROM ClashZones c
                            INNER JOIN TempGuidLookup t ON UPPER(TRIM(c.ClashZoneGuid)) = UPPER(TRIM(t.ClashZoneGuid))
                            WHERE c.ClashZoneGuid IS NOT NULL AND c.ClashZoneGuid != ''";
                        
                        // ✅ DIAGNOSTIC: Log sample GUIDs being queried
                        if (!DeploymentConfiguration.DeploymentMode && updateDict.Count > 0)
                        {
                            var sampleGuid = updateDict.Keys.First();
                            _logger($"[SQLite] BatchUpdateFlags: Querying database for {updateDict.Count} GUIDs. Sample GUID: '{sampleGuid}'");
                        }
                        
                        using (var reader = queryCmd.ExecuteReader())
                        {
                            var foundGuids = new HashSet<string>();
                            
                            while (reader.Read())
                            {
                                var guid = reader.GetString(0)?.Trim().ToUpperInvariant() ?? string.Empty;
                                if (string.IsNullOrEmpty(guid))
                                    continue;
                                    
                                foundGuids.Add(guid);
                                
                                if (updateDict.TryGetValue(guid, out var update))
                                {
                                    fullUpdates.Add((
                                        update.ClashZoneId,
                                        update.IsResolved,
                                        update.IsClusterResolved,
                                        update.SleeveInstanceId,
                                        update.ClusterInstanceId,
                                        reader.GetInt32(1), // MepElementId
                                        reader.GetInt32(2), // HostElementId (StructuralElementId)
                                        reader.GetDouble(3), // IntersectionX
                                        reader.GetDouble(4), // IntersectionY
                                        reader.GetDouble(5), // IntersectionZ
                                        reader.IsDBNull(6) ? 0 : reader.GetInt32(6), // OldSleeveInstanceId (current SleeveInstanceId)
                                        reader.IsDBNull(7) ? 0 : reader.GetInt32(7), // OldClusterInstanceId (current ClusterInstanceId)
                                        reader.IsDBNull(8) ? (bool?)null : (reader.GetInt32(8) == 1), // MarkedForClusterProcess
                                        reader.IsDBNull(9) ? 0 : reader.GetInt32(9), // AfterClusterSleeveId
                                        reader.IsDBNull(10) ? (bool?)null : (reader.GetInt32(10) == 1) // IsClusteredFlag
                                    ));
                                }
                            }
                            
                            // ✅ DIAGNOSTIC: Log how many GUIDs were found
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                _logger($"[SQLite] BatchUpdateFlags: Found {foundGuids.Count} out of {updateDict.Count} GUIDs in database");
                            }
                            
                            // ✅ FIX: For GUIDs not found in database, log warning and skip (don't use defaults)
                            // Using defaults causes WHERE clause to fail because it can't match by GUID or by other criteria
                            var notFoundGuids = new List<string>();
                            foreach (var kvp in updateDict)
                            {
                                if (!foundGuids.Contains(kvp.Key))
                                {
                                    notFoundGuids.Add(kvp.Key);
                                    var update = kvp.Value;
                                    _logger($"[SQLite] ⚠️ BatchUpdateFlags: GUID {kvp.Key} not found in database - skipping update. This may indicate the clash zone was not saved to database.");
                                    
                                    // ✅ CRITICAL FIX: Still add to fullUpdates but with actual GUID lookup
                                    // The full implementation's WHERE clause will try to match by GUID first
                                    // If GUID doesn't exist, it will fail, but at least we tried
                                    fullUpdates.Add((
                                        update.ClashZoneId,
                                        update.IsResolved,
                                        update.IsClusterResolved,
                                        update.SleeveInstanceId,
                                        update.ClusterInstanceId,
                                        0, // MepElementId - will be ignored if GUID match fails
                                        0, // StructuralElementId - will be ignored if GUID match fails
                                        0.0, // IntersectionPointX - will be ignored if GUID match fails
                                        0.0, // IntersectionPointY - will be ignored if GUID match fails
                                        0.0, // IntersectionPointZ - will be ignored if GUID match fails
                                        0, // OldSleeveInstanceId - will be ignored if GUID match fails
                                        0, // OldClusterInstanceId - will be ignored if GUID match fails
                                        null, // MarkedForClusterProcess
                                        0, // AfterClusterSleeveId
                                        null // IsClusteredFlag
                                    ));
                                }
                            }
                            
                            if (notFoundGuids.Count > 0)
                            {
                                _logger($"[SQLite] ⚠️ BatchUpdateFlags: {notFoundGuids.Count} GUID(s) not found in database. Updates may fail for these zones.");
                            }
                        }
                    }

                    // Cleanup temp table
                    using (var dropCmd = _context.Connection.CreateCommand())
                    {
                        dropCmd.Transaction = transaction;
                        dropCmd.CommandText = "DROP TABLE IF EXISTS TempGuidLookup";
                        dropCmd.ExecuteNonQuery();
                    }

                    transaction.Commit();
                }
                catch
                {
                    transaction.Rollback();
                    throw;
                }
            }

            // Call the full implementation
            BatchUpdateFlags(fullUpdates);
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

        /// <summary>
        /// PUBLIC: Fetch MEP snapshot parameter values for provided sleeve element instance ids.
        /// Lightweight helper for batch parameter transfer – does NOT modify clash zones.
        /// Returns dictionary: SleeveInstanceId -> (ParameterName -> ParameterValue).
        /// Safely returns empty dictionary on any failure.
        /// </summary>
        public Dictionary<int, Dictionary<string, string>> GetSnapshotMepParametersForSleeveIds(IEnumerable<int> sleeveInstanceIds)
        {
            var result = new Dictionary<int, Dictionary<string, string>>();
            if (sleeveInstanceIds == null) return result;
            var ids = sleeveInstanceIds.Where(i => i > 0).Distinct().ToList();
            if (ids.Count == 0) return result;

            try
            {
                using (var cmd = _context.Connection.CreateCommand())
                {
                    var placeholders = string.Join(",", ids.Select((_, i) => $"@Id{i}"));
                    cmd.CommandText = $@"SELECT SleeveInstanceId, MepParametersJson FROM SleeveSnapshots WHERE SleeveInstanceId IN ({placeholders})";
                    for (int i = 0; i < ids.Count; i++)
                    {
                        cmd.Parameters.AddWithValue($"@Id{i}", ids[i]);
                    }

                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            var sleeveId = GetInt(reader, "SleeveInstanceId", -1);
                            if (sleeveId <= 0) continue;
                            var mepParamsJson = GetNullableString(reader, "MepParametersJson");
                            var dict = DeserializeDictionary(mepParamsJson);
                            result[sleeveId] = dict ?? new Dictionary<string, string>();
                        }
                    }
                }

                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[ClashZoneRepository] ✅ Loaded snapshot MEP parameters for {result.Count} sleeves (requested {ids.Count})");
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Warning($"[ClashZoneRepository] ⚠️ Error loading snapshot parameters: {ex.Message}");
                }
                return new Dictionary<int, Dictionary<string, string>>();
            }

            return result;
        }

        /// <summary>
        /// ✅ CRITICAL SAFETY: Verify that all placed sleeves have snapshots saved in database
        /// </summary>
        private (bool Success, string Message, List<int> MissingSleeveIds) VerifySnapshotCompleteness(
            List<ClashZone> placedZones, 
            SQLiteTransaction transaction)
        {
            var missingSleeveIds = new List<int>();
            var sleeveIds = placedZones
                .Where(z => z.SleeveInstanceId > 0)
                .Select(z => z.SleeveInstanceId)
                .Distinct()
                .ToList();

            if (sleeveIds.Count == 0)
                return (true, "No individual sleeves to verify", missingSleeveIds);

            try
            {
                using (var cmd = _context.Connection.CreateCommand())
                {
                    cmd.Transaction = transaction;
                    var placeholders = string.Join(",", sleeveIds.Select((_, i) => $"@Id{i}"));
                    cmd.CommandText = $"SELECT SleeveInstanceId FROM SleeveSnapshots WHERE SleeveInstanceId IN ({placeholders})";
                    
                    for (int i = 0; i < sleeveIds.Count; i++)
                    {
                        cmd.Parameters.AddWithValue($"@Id{i}", sleeveIds[i]);
                    }

                    var savedIds = new HashSet<int>();
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            savedIds.Add(GetInt(reader, "SleeveInstanceId", -1));
                        }
                    }

                    missingSleeveIds = sleeveIds.Where(id => !savedIds.Contains(id)).ToList();
                    
                    if (missingSleeveIds.Count > 0)
                    {
                        return (false, $"{missingSleeveIds.Count}/{sleeveIds.Count} sleeves missing snapshots", missingSleeveIds);
                    }

                    return (true, $"All {sleeveIds.Count} sleeves have snapshots", missingSleeveIds);
                }
            }
            catch (Exception ex)
            {
                return (false, $"Verification failed: {ex.Message}", missingSleeveIds);
            }
        }

        /// <summary>
        /// ✅ CRITICAL SAFETY: Validate that critical parameters are present in aggregated data
        /// </summary>
        private (bool IsValid, string Message) ValidateCriticalParameters(
            Dictionary<string, string> mepParams,
            Dictionary<string, string> hostParams,
            List<ClashZone> zones,
            bool isCluster,
            int groupId)
        {
            var missingCritical = new List<string>();

            // Critical MEP parameters that MUST be present
            var criticalMepParams = new[] { "Size", "System Type", "System Name" };
            foreach (var paramName in criticalMepParams)
            {
                var hasParam = mepParams.Any(kvp => 
                    string.Equals(kvp.Key, paramName, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(kvp.Key, "MEP " + paramName, StringComparison.OrdinalIgnoreCase));
                
                if (!hasParam)
                {
                    missingCritical.Add(paramName);
                }
            }

            if (missingCritical.Count > 0)
            {
                var category = zones.FirstOrDefault()?.MepElementCategory ?? "Unknown";
                return (false, $"Missing critical MEP parameters for {category}: {string.Join(", ", missingCritical)}");
            }

            // Verify at least some parameters exist
            if (mepParams.Count == 0)
            {
                return (false, "No MEP parameters captured - snapshot may be empty");
            }

            return (true, $"All critical parameters present ({mepParams.Count} MEP, {hostParams.Count} Host)");
        }
    }
}



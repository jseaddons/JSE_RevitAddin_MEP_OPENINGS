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
using JSE_RevitAddin_MEP_OPENINGS.Services.Refresh;

// Implements the missing interface member for BatchUpdateFlagsWithCurrentClash
// Must be inside the class

namespace JSE_RevitAddin_MEP_OPENINGS.Data.Repositories
{
    /// <summary>
    /// SQLite repository implementation for ClashZone persistence.
    /// Database is the sole source of truth for all clash zone data and flags.
    /// </summary>
    public class ClashZoneRepository : IClashZoneRepository, Services.Interfaces.ISleeveRepository
    {
        private readonly SleeveDbContext _context;
        private readonly Action<string> _logger;
        private readonly PerformanceMonitor _performanceMonitor;
        private static bool _dbVerifiedOnce = false; // session-level guard

        // SOLID Refactoring: Sub-repositories
        private readonly CombinedSleeveRepository _combinedRepo;
        private readonly ClusterSleeveRepository _clusterRepo;
        private readonly SleeveSnapshotRepository _snapshotRepo;

        /// <summary>
        /// Public property to expose the database path for services that need direct database access
        /// </summary>
        public string DatabasePath => _context.DatabasePath;

        public ClashZoneRepository(SleeveDbContext context, Action<string> logger = null, PerformanceMonitor performanceMonitor = null)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            _logger = logger ?? (msg => { });
            _performanceMonitor = performanceMonitor;

            // Initialize sub-repositories for delegation
            _combinedRepo = new CombinedSleeveRepository(_context);
            // ClusterSleeveRepository constructor requires context and optional logger
            _clusterRepo = new ClusterSleeveRepository(_context, _logger);
            // SleeveSnapshotRepository constructor requires context and optional logger
            _snapshotRepo = new SleeveSnapshotRepository(_context, _logger);
        }

        // Interface implementation for IClashZoneRepository
        public List<ClashZone> GetClashZonesByGuids(IEnumerable<Guid> clashZoneGuids)
        {
            var result = new List<ClashZone>();
            if (clashZoneGuids == null) return result;

            var guidList = clashZoneGuids.Distinct().ToList();
            if (guidList.Count == 0) return result;

            try
            {
                // Create comma-separated string of quoted GUIDs for IN clause
                var guidString = string.Join(",", guidList.Select(g => $"'{g.ToString().ToUpperInvariant()}'"));

                using (var cmd = _context.Connection.CreateCommand())
                {
                    cmd.CommandText = $@"
                        SELECT * 
                        FROM ClashZones 
                        WHERE UPPER(ClashZoneGuid) IN ({guidString})";

                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            result.Add(MapClashZone(reader));
                        }
                    }
                }

                if (!OptimizationFlags.DisableVerboseLogging)
                {
                    _logger($"[SQLite] Custom Fetch: Retrieved {result.Count} zones by {guidList.Count} GUIDs");
                }
            }
            catch (Exception ex)
            {
                _logger($"[SQLite] ❌ Error in GetClashZonesByGuids: {ex.Message}");
            }

            return result;
        }


        /// <summary>
        /// ✅ PUBLIC BATCH OPTIMIZATION: Insert or update clash zones in a single multi-category batch.
        /// Consolidates multiple transactions into one atomic operation.
        /// </summary>
        public void InsertOrUpdateClashZonesBulk(IEnumerable<ClashZone> clashZones, string filterName)
        {
            var zonesList = clashZones?.ToList() ?? new List<ClashZone>();
            if (zonesList.Count == 0) return;

            var sw = System.Diagnostics.Stopwatch.StartNew();
            _logger($"[SQLite][BULK-OPTIMIZED] ⚡ Starting multi-category bulk save for {zonesList.Count} zones");

            using (var transaction = _context.Connection.BeginTransaction())
            {
                try
                {
                    // Group by category to respect the database schema (Filters table)
                    var byCategory = zonesList
                        .Where(z => z != null && !string.IsNullOrWhiteSpace(z.MepElementCategory))
                        .GroupBy(z => z.MepElementCategory, StringComparer.OrdinalIgnoreCase)
                        .ToList();

                    foreach (var categoryGroup in byCategory)
                    {
                        string category = categoryGroup.Key;
                        var categoryZones = categoryGroup.ToList();

                        // Use the optimized single-category bulk method if flag is enabled
                        if (OptimizationFlags.UseBulkSqliteUpdates)
                        {
                            InsertOrUpdateClashZonesBulkInternal(categoryZones, filterName, category, transaction);
                        }
                        else
                        {
                            // Legacy per-zone fallback if needed (though not recommended for Phase 3)
                            InsertOrUpdateClashZones(categoryZones, filterName, category);
                        }
                    }

                    transaction.Commit();
                    sw.Stop();
                    _logger($"[SQLite][BULK-OPTIMIZED] ✅ Completed multi-category save in {sw.ElapsedMilliseconds}ms");
                }
                catch (Exception ex)
                {
                    _logger($"[SQLite][BULK-OPTIMIZED] ❌ Fatal error: {ex.Message}");
                    transaction.Rollback();
                    throw;
                }
            }
        }

        private void InsertOrUpdateClashZonesBulkInternal(List<ClashZone> zonesList, string filterName, string category, SQLiteTransaction transaction)
        {
            int filterId = GetOrCreateFilter(filterName, category, transaction);
            if (filterId <= 0) return;

            // 1. Batch lookup existing IDs by GUID
            var guidList = string.Join(",", zonesList.Select(z => $"'{z.Id.ToString().ToUpperInvariant()}'"));
            var existingMap = new Dictionary<Guid, int>();
            using (var cmd = _context.Connection.CreateCommand())
            {
                cmd.Transaction = transaction;
                cmd.CommandText = $@"SELECT ClashZoneGuid, ClashZoneId FROM ClashZones WHERE UPPER(ClashZoneGuid) IN ({guidList})";
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read()) existingMap[Guid.Parse(reader.GetString(0))] = reader.GetInt32(1);
                }
            }

            // 2. Batch File Combos
            var comboMap = BatchGetOrCreateFileCombos(zonesList, filterId, category, transaction);

            // 3. Update existing
            var toUpdate = zonesList.Where(z => existingMap.ContainsKey(z.Id)).ToList();
            if (toUpdate.Count > 0)
            {
                foreach (var zone in toUpdate) zone.ClashZoneId = existingMap[zone.Id];
                BulkUpdateClashZones(toUpdate, existingMap, comboMap, transaction);
                BulkUpdateRTreeIndex(toUpdate, transaction);
            }

            // 4. Handle Inserts (with UNIQUE constraint check)
            var toInsert = zonesList.Where(z => !existingMap.ContainsKey(z.Id) && comboMap.ContainsKey(z.Id)).ToList();
            if (toInsert.Count > 0)
            {
                // Unique constraint check... (simplified for now to keep diff clean, but ideally uses temp table)
                // For now, reuse the existing logic in the private method but inside this transaction
                var uniqueConstraintMap = GetUniqueConstraintMap(toInsert, comboMap, transaction);

                var actuallyNew = new List<ClashZone>();
                foreach (var zone in toInsert)
                {
                    var key = GetUniqueKey(zone, comboMap[zone.Id]);
                    if (uniqueConstraintMap.TryGetValue(key, out var existingId))
                    {
                        existingMap[zone.Id] = existingId;
                        zone.ClashZoneId = existingId;
                        BulkUpdateClashZones(new List<ClashZone> { zone }, existingMap, comboMap, transaction);
                    }
                    else actuallyNew.Add(zone);
                }

                if (actuallyNew.Count > 0)
                {
                    BulkInsertClashZones(actuallyNew, comboMap, transaction);
                    // Fetch new IDs for R-tree
                    var newGuids = string.Join(",", actuallyNew.Select(z => $"'{z.Id.ToString().ToUpperInvariant()}'"));
                    using (var fetchCmd = _context.Connection.CreateCommand())
                    {
                        fetchCmd.Transaction = transaction;
                        fetchCmd.CommandText = $"SELECT ClashZoneGuid, ClashZoneId FROM ClashZones WHERE UPPER(ClashZoneGuid) IN ({newGuids})";
                        using (var reader = fetchCmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                var guid = Guid.Parse(reader.GetString(0));
                                var zone = actuallyNew.FirstOrDefault(z => z.Id == guid);
                                if (zone != null) zone.ClashZoneId = reader.GetInt32(1);
                            }
                        }
                    }
                    BulkUpdateRTreeIndex(actuallyNew, transaction);
                }
            }

            // 5. Snapshots
            var processedZonesData = zonesList
                .Where(z => comboMap.ContainsKey(z.Id))
                .Select(z => (comboMap[z.Id], z))
                .ToList();
            if (processedZonesData.Count > 0) InsertOrUpdateSleeveSnapshots(filterId, processedZonesData, transaction);
        }

        private string GetUniqueKey(ClashZone zone, int comboId)
        {
            var mepId = zone.MepElementId?.IntegerValue ?? zone.MepElementIdValue;
            var hostId = zone.StructuralElementId?.IntegerValue ?? zone.StructuralElementIdValue;
            var interX = Math.Round(zone.IntersectionPoint?.X ?? zone.IntersectionPointX, 6);
            var interY = Math.Round(zone.IntersectionPoint?.Y ?? zone.IntersectionPointY, 6);
            var interZ = Math.Round(zone.IntersectionPoint?.Z ?? zone.IntersectionPointZ, 6);
            return $"{comboId}|{mepId}|{hostId}|{interX}|{interY}|{interZ}";
        }

        private Dictionary<string, int> GetUniqueConstraintMap(List<ClashZone> zones, Dictionary<Guid, int> comboMap, SQLiteTransaction transaction)
        {
            var map = new Dictionary<string, int>();
            if (zones.Count == 0) return map;

            using (var cmd = _context.Connection.CreateCommand())
            {
                cmd.Transaction = transaction;
                var conditions = new List<string>();
                int i = 0;
                foreach (var zone in zones)
                {
                    if (!comboMap.ContainsKey(zone.Id)) continue;

                    int cid = comboMap[zone.Id];
                    int mid = zone.MepElementId?.IntegerValue ?? zone.MepElementIdValue;
                    int hid = zone.StructuralElementId?.IntegerValue ?? zone.StructuralElementIdValue;
                    double x = zone.IntersectionPoint?.X ?? zone.IntersectionPointX;
                    double y = zone.IntersectionPoint?.Y ?? zone.IntersectionPointY;
                    double z = zone.IntersectionPoint?.Z ?? zone.IntersectionPointZ;

                    // Skip zones with invalid coordinates (NaN/Infinity cause SQL errors)
                    if (double.IsNaN(x) || double.IsInfinity(x) ||
                        double.IsNaN(y) || double.IsInfinity(y) ||
                        double.IsNaN(z) || double.IsInfinity(z))
                    {
                        _logger?.Invoke($"[SQLite] ⚠️ Skipping zone with invalid coordinates: {zone.Id}");
                        continue;
                    }

                    // Use parameterized query to avoid locale/formatting issues with doubles
                    conditions.Add($"(ComboId=@cid{i} AND MepElementId=@mid{i} AND HostElementId=@hid{i} AND ABS(IntersectionX-@x{i}) < 0.0001 AND ABS(IntersectionY-@y{i}) < 0.0001 AND ABS(IntersectionZ-@z{i}) < 0.0001)");
                    cmd.Parameters.AddWithValue($"@cid{i}", cid);
                    cmd.Parameters.AddWithValue($"@mid{i}", mid);
                    cmd.Parameters.AddWithValue($"@hid{i}", hid);
                    cmd.Parameters.AddWithValue($"@x{i}", x);
                    cmd.Parameters.AddWithValue($"@y{i}", y);
                    cmd.Parameters.AddWithValue($"@z{i}", z);

                    if (++i > 100) break; // Limit to 100 per check to avoid giant SQL
                }

                if (conditions.Count > 0)
                {
                    cmd.CommandText = $"SELECT ClashZoneId, ComboId, MepElementId, HostElementId, IntersectionX, IntersectionY, IntersectionZ FROM ClashZones WHERE {string.Join(" OR ", conditions)}";
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            var key = $"{reader.GetInt32(1)}|{reader.GetInt32(2)}|{reader.GetInt32(3)}|{Math.Round(reader.GetDouble(4), 6)}|{Math.Round(reader.GetDouble(5), 6)}|{Math.Round(reader.GetDouble(6), 6)}";
                            map[key] = reader.GetInt32(0);
                        }
                    }
                }
            }
            return map;
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

                    int filterId;
                    using (var filterOp = _performanceMonitor?.TrackOperation("9a2a. Get Filter ID"))
                    {
                        filterId = GetOrCreateFilter(filterName, category, transaction);
                    }
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
                    using (var fetchOp = _performanceMonitor?.TrackOperation("9a2b. Fetch Existing IDs"))
                    {
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
                    }

                    _logger($"[SQLite][BULK] Found {existingMap.Count}/{zonesList.Count} existing zones");

                    // ⚡ BATCH OPTIMIZATION: Get/create file combos in single batch operation
                    Dictionary<Guid, int> comboMap;
                    using (var comboOp = _performanceMonitor?.TrackOperation("9a5. Batch File Combos") as PerformanceMonitor.OperationTracker)
                    {
                        comboMap = BatchGetOrCreateFileCombos(zonesList, filterId, category, transaction);
                    }

                    // Batch UPDATE using CASE statements for all zones with existing ClashZoneIds
                    var toUpdate = zonesList.Where(z => existingMap.ContainsKey(z.Id)).ToList();
                    if (toUpdate.Count > 0)
                    {
                        using (var updateOp = _performanceMonitor?.TrackOperation("9a3. Bulk Update") as PerformanceMonitor.OperationTracker)
                        {
                            BulkUpdateClashZones(toUpdate, existingMap, comboMap, transaction);
                            updateOp?.SetItemCount(toUpdate.Count);
                        }
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
                                using (var constraintOp = _performanceMonitor?.TrackOperation("9a4. Constraint Lookup") as PerformanceMonitor.OperationTracker)
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
                                existingMap[zone.Id] = existingClashZoneId;
                                toUpdateByConstraint.Add((zone, existingClashZoneId));
                            }
                            else
                            {
                                actuallyNew.Add(zone);
                            }
                        }

                        // ✅ STEP 3: Execute updates and inserts with sub-timings
                        if (toUpdateByConstraint.Count > 0)
                        {
                            using (var updateOp = _performanceMonitor?.TrackOperation("9a1. Bulk Update (Constraint)") as PerformanceMonitor.OperationTracker)
                            {
                                var zonesToUpdate = toUpdateByConstraint.Select(t => t.zone).ToList();
                                BulkUpdateClashZones(zonesToUpdate, existingMap, comboMap, transaction);
                                _logger($"[SQLite][BULK] ✅ Updated {toUpdateByConstraint.Count} zones found by UNIQUE constraint");
                                updateOp?.SetItemCount(toUpdateByConstraint.Count);
                            }
                        }

                        // ✅ STEP 4: INSERT only truly new zones
                        if (actuallyNew.Count > 0)
                        {
                            using (var insertOp = _performanceMonitor?.TrackOperation("9a2. Bulk Insert") as PerformanceMonitor.OperationTracker)
                            {
                                BulkInsertClashZones(actuallyNew, comboMap, transaction);
                                _logger($"[SQLite][BULK] ✅ Inserted {actuallyNew.Count} new zones");
                                insertOp?.SetItemCount(actuallyNew.Count);
                            }
                        }
                    }

                    using (var commitOp = _performanceMonitor?.TrackOperation("9a2c. Commit Transaction"))
                    {
                        transaction.Commit();
                    }
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
        /// Verify if sleeves marked as resolved in the DB still exist in the current Revit model.
        /// If not found, reset their flags to false and IDs to -1.
        /// </summary>
        /// <summary>
        /// Verify if sleeves marked as resolved in the DB still exist in the current Revit model.
        /// If not found, reset their flags to false and IDs to -1 in a hierarchical manner.
        /// Hierarchy: Combined -> Cluster -> Individual
        /// </summary>
        public int VerifyExistingSleevesAndResetFlags(Autodesk.Revit.DB.Document doc, List<string> filterNames, List<string> categories)
        {
            try
            {
                _logger("[SQLite] 🔍 Starting Optimized Hierarchical Sleeve Verification...");
                var stopwatch = System.Diagnostics.Stopwatch.StartNew();

                // ⚡ OPTIMIZATION 1: High-performance batch collection from Revit
                var sleeveCollector = new JSE_RevitAddin_MEP_OPENINGS.Services.FlagManagement.RevitSleeveCollector();
                var revitSleeveIds = sleeveCollector.CollectAllSleeveIds(doc);
                _logger($"[SQLite] 🔍 Collected {revitSleeveIds.Count} sleeve elements from Revit in {stopwatch.ElapsedMilliseconds}ms");

                // ⚡ OPTIMIZATION 2: Single query to get all zones with ANY resolution flag set
                var zonesToCheck = new List<ZoneResolutionState>();
                using (var cmd = _context.Connection.CreateCommand())
                {
                    cmd.CommandText = @"
                        SELECT ClashZoneId, 
                               IsCombinedResolved, CombinedClusterSleeveInstanceId,
                               IsClusterResolvedFlag, ClusterInstanceId,
                               IsResolvedFlag, SleeveInstanceId,
                               ReadyForPlacementFlag
                        FROM ClashZones
                        WHERE IsCombinedResolved = 1 
                           OR IsClusterResolvedFlag = 1 
                           OR IsResolvedFlag = 1
                           OR (IsResolvedFlag = 0 AND IsClusterResolvedFlag = 0 AND IsCombinedResolved = 0 AND ReadyForPlacementFlag = 0)";

                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            zonesToCheck.Add(new ZoneResolutionState
                            {
                                ZoneId = reader.GetInt32(0),
                                IsCombined = reader.GetBoolean(1),
                                CombinedId = reader.IsDBNull(2) ? -1 : reader.GetInt32(2),
                                IsCluster = reader.GetBoolean(3),
                                ClusterId = reader.IsDBNull(4) ? -1 : reader.GetInt32(4),
                                IsIndividual = reader.GetBoolean(5) && !reader.GetBoolean(3) && !reader.GetBoolean(1),
                                IndividualId = reader.IsDBNull(6) ? -1 : reader.GetInt32(6),
                                ReadyForPlacement = reader.IsDBNull(7) ? false : reader.GetBoolean(7)
                            });
                        }
                    }
                }

                if (zonesToCheck.Count == 0)
                {
                    _logger("[SQLite] ✅ No resolved zones to verify");
                    return 0;
                }

                _logger($"[SQLite] 🔍 Checking {zonesToCheck.Count} zones in multi-threading...");

                // ⚡ OPTIMIZATION 3: Multi-threaded existence check using Parallel.ForEach
                var updates = new System.Collections.Concurrent.ConcurrentBag<ZoneResolutionUpdate>();

                System.Threading.Tasks.Parallel.ForEach(zonesToCheck, zone =>
                {
                    bool needsUpdate = false;
                    bool newIsCombined = zone.IsCombined;
                    bool newIsCluster = zone.IsCluster;
                    bool newIsIndividual = zone.IsIndividual;

                    // Hierarchical Check 1: Combined
                    if (zone.IsCombined)
                    {
                        if (zone.CombinedId <= 0 || !revitSleeveIds.Contains(zone.CombinedId))
                        {
                            newIsCombined = false;
                            needsUpdate = true;
                        }
                    }

                    // Hierarchical Check 2: Cluster
                    if (zone.IsCluster)
                    {
                        if (zone.ClusterId <= 0 || !revitSleeveIds.Contains(zone.ClusterId))
                        {
                            newIsCluster = false;
                            needsUpdate = true;
                        }
                    }

                    // Hierarchical Check 3: Individual
                    if (zone.IsIndividual)
                    {
                        if (zone.IndividualId <= 0 || !revitSleeveIds.Contains(zone.IndividualId))
                        {
                            newIsIndividual = false;
                            needsUpdate = true;
                        }
                    }

                    // ✅ NEW CASE: Unresolved zone with ReadyForPlacement=0 needs the flag set
                    if (!zone.IsCombined && !zone.IsCluster && !zone.IsIndividual && !zone.ReadyForPlacement)
                    {
                        needsUpdate = true;
                    }
                    if (needsUpdate)
                    {
                        updates.Add(new ZoneResolutionUpdate
                        {
                            ZoneId = zone.ZoneId,
                            IsCombined = newIsCombined,
                            IsCluster = newIsCluster,
                            IsIndividual = newIsIndividual
                        });
                    }
                });

                if (updates.Count == 0)
                {
                    _logger($"[SQLite] ✅ All {zonesToCheck.Count} zones verified. No resets needed. ({stopwatch.ElapsedMilliseconds}ms)");
                    return 0;
                }

                // ⚡ OPTIMIZATION 4: Batch DB Update
                _logger($"[SQLite] ⚠️ Resetting flags for {updates.Count} zones...");
                int totalReset = 0;

                using (var transaction = _context.Connection.BeginTransaction())
                {
                    using (var cmd = _context.Connection.CreateCommand())
                    {
                        cmd.Transaction = transaction;
                        cmd.CommandText = @"
                            UPDATE ClashZones 
                            SET IsCombinedResolved = @isCombined,
                                IsClusterResolvedFlag = @isCluster,
                                IsResolvedFlag = @isIndividual,
                                CombinedClusterSleeveInstanceId = CASE WHEN @isCombined = 0 THEN NULL ELSE CombinedClusterSleeveInstanceId END,
                                ClusterInstanceId = CASE WHEN @isCluster = 0 THEN -1 ELSE ClusterInstanceId END,
                                SleeveInstanceId = CASE WHEN @isIndividual = 0 THEN -1 ELSE SleeveInstanceId END,
                                MarkedForClusterProcess = CASE WHEN @isCluster = 0 THEN 1 ELSE MarkedForClusterProcess END,
                                IsClusteredFlag = CASE WHEN @isCluster = 0 THEN 0 ELSE IsClusteredFlag END,
                                AfterClusterSleeveId = CASE WHEN @isCluster = 0 THEN 0 ELSE AfterClusterSleeveId END,
                                IsCurrentClashFlag = 1,
                                ReadyForPlacementFlag = 1,
                                UpdatedAt = CURRENT_TIMESTAMP
                            WHERE ClashZoneId = @id";

                        var pId = cmd.CreateParameter(); pId.ParameterName = "@id"; cmd.Parameters.Add(pId);
                        var pCombined = cmd.CreateParameter(); pCombined.ParameterName = "@isCombined"; cmd.Parameters.Add(pCombined);
                        var pCluster = cmd.CreateParameter(); pCluster.ParameterName = "@isCluster"; cmd.Parameters.Add(pCluster);
                        var pIndividual = cmd.CreateParameter(); pIndividual.ParameterName = "@isIndividual"; cmd.Parameters.Add(pIndividual);

                        foreach (var update in updates)
                        {
                            pId.Value = update.ZoneId;
                            pCombined.Value = update.IsCombined ? 1 : 0;
                            pCluster.Value = update.IsCluster ? 1 : 0;
                            pIndividual.Value = update.IsIndividual ? 1 : 0;

                            totalReset += cmd.ExecuteNonQuery();
                        }
                    }
                    transaction.Commit();
                }

                stopwatch.Stop();
                _logger($"[SQLite] ✅ Optimized Verification complete: {totalReset} total zones updated in {stopwatch.ElapsedMilliseconds}ms");
                return totalReset;
            }
            catch (Exception ex)
            {
                _logger($"[SQLite] ❌ Error in VerifyExistingSleevesAndResetFlags: {ex.Message}");
                return 0;
            }
        }

        // Helper classes for flag reset logic
        private class ZoneResolutionState
        {
            public int ZoneId { get; set; }
            public bool IsCombined { get; set; }
            public int CombinedId { get; set; }
            public bool IsCluster { get; set; }
            public int ClusterId { get; set; }
            public bool IsIndividual { get; set; }
            public int IndividualId { get; set; }
            public bool ReadyForPlacement { get; set; }
        }

        private class ZoneResolutionUpdate
        {
            public int ZoneId { get; set; }
            public bool IsCombined { get; set; }
            public bool IsCluster { get; set; }
            public bool IsIndividual { get; set; }
        }

        /// <summary>
        /// Reset IsCurrentClashFlag to false for all zones in the specified filters/categories.
        /// This is called at the start of a refresh cycle to mark all existing zones as "stale" until re-detected.
        /// </summary>
        public int ResetIsCurrentClashFlag(List<string> filterNames, List<string> categories)
        {
            try
            {
                var sw = System.Diagnostics.Stopwatch.StartNew();

                // ✅ OPTIMIZATION: Only reset zones for selected filters/categories (not ALL zones)
                // This reduces UPDATE from thousands of rows to just the relevant ones

                using (var transaction = _context.Connection.BeginTransaction())
                {
                    using (var cmd = _context.Connection.CreateCommand())
                    {
                        cmd.Transaction = transaction;

                        // Build WHERE clause for selected filters and categories
                        var filterConditions = new List<string>();
                        int paramIndex = 0;

                        foreach (var filterName in filterNames ?? new List<string>())
                        {
                            foreach (var category in categories ?? new List<string>())
                            {
                                filterConditions.Add($"(f.FilterName = @fn{paramIndex} AND cz.MepCategory = @cat{paramIndex})");
                                cmd.Parameters.AddWithValue($"@fn{paramIndex}", filterName);
                                cmd.Parameters.AddWithValue($"@cat{paramIndex}", category);
                                paramIndex++;
                            }
                        }

                        if (filterConditions.Count == 0)
                        {
                            // No filters/categories selected - reset ALL
                            cmd.CommandText = "UPDATE ClashZones SET IsCurrentClashFlag = 0, ReadyForPlacementFlag = 0";
                        }
                        else
                        {
                            // ✅ OPTIMIZED: Only reset zones matching selected filters/categories
                            // Reset ALL zones in scope - SetReadyForPlacementForUnresolvedZonesInSectionBox 
                            // will set flag=1 for zones in section box
                            cmd.CommandText = $@"
                                UPDATE ClashZones 
                                SET IsCurrentClashFlag = 0, ReadyForPlacementFlag = 0
                                WHERE (IsResolvedFlag = 0 AND IsClusterResolvedFlag = 0 AND IsCombinedResolved = 0)
                                AND ClashZoneId IN (
                                    SELECT cz.ClashZoneId
                                    FROM ClashZones cz
                                    INNER JOIN FileCombos fc ON cz.ComboId = fc.ComboId
                                    INNER JOIN Filters f ON fc.FilterId = f.FilterId
                                    WHERE ({string.Join(" OR ", filterConditions)})
                                    AND cz.IsResolvedFlag = 0 AND cz.IsClusterResolvedFlag = 0
                                )";
                        }

                        int count = cmd.ExecuteNonQuery();
                        transaction.Commit();

                        sw.Stop();
                        _logger($"[IsCurrentClash] 🧹 Reset flag for {count} zones in {sw.ElapsedMilliseconds}ms (optimized)");
                        return count;
                    }
                }
            }
            catch (Exception ex)
            {
                _logger($"[SQLite] ❌ Error in ResetIsCurrentClashFlag: {ex.Message}");
                return 0;
            }
        }
        private void BulkUpdateClashZones(List<ClashZone> zones, Dictionary<Guid, int> existingMap, Dictionary<Guid, int> comboMap, SQLiteTransaction transaction)
        {
            if (zones.Count == 0) return;

            // ✅ PHASE 2 OPTIMIZATION: Use High-Performance Temp Table for bulk updates
            if (OptimizationFlags.UseTempTableForBulkUpdates)
            {
                BulkUpdateClashZonesWithTempTable(zones, existingMap, comboMap, transaction);
                return;
            }

            // Legacy CASE WHEN approach (deprecated)
            BulkUpdateClashZonesLegacy(zones, existingMap, comboMap, transaction);
        }

        /// <summary>
        /// ✅ PERFORMANCE OPTIMIZED: High-performance bulk update using a temporary table.
        /// Replaces the CASE WHEN approach which was causing a 21s bottleneck for 888 zones.
        /// Strategy: 
        /// 1. Create a TEMP TABLE for updates.
        /// 2. Bulk insert data into TEMP TABLE using prepared statement.
        /// 3. Execute ONE join-based UPDATE against the main table.
        /// </summary>
        private void BulkUpdateClashZonesWithTempTable(List<ClashZone> zones, Dictionary<Guid, int> existingMap, Dictionary<Guid, int> comboMap, SQLiteTransaction transaction)
        {
            var validZones = zones.Where(z => existingMap.ContainsKey(z.Id) && comboMap.ContainsKey(z.Id)).ToList();
            if (validZones.Count == 0) return;

            try
            {
                using (var cmd = _context.Connection.CreateCommand())
                {
                    cmd.Transaction = transaction;

                    // 1. Create Temp Table
                    cmd.CommandText = @"
                        CREATE TEMP TABLE IF NOT EXISTS BulkUpdateZones (
                            ClashZoneId INTEGER PRIMARY KEY,
                            IsResolvedFlag INTEGER,
                            IsClusterResolvedFlag INTEGER,
                            IsCombinedResolved INTEGER,
                            SleeveInstanceId INTEGER,
                            ClusterInstanceId INTEGER,
                            MepParameterValuesJson TEXT,
                            HostParameterValuesJson TEXT,
                            WallCenterlinePointX REAL,
                            WallCenterlinePointY REAL,
                            WallCenterlinePointZ REAL,
                            MepOrientationX REAL,
                            MepOrientationY REAL,
                            MepOrientationZ REAL,
                            MepRotationAngleRad REAL,
                            MepRotationAngleDeg REAL,
                            MepOrientationDirection TEXT,
                            StructuralThickness REAL,
                            MepElementTypeName TEXT,
                            MepElementFamilyName TEXT,
                            MepSystemType TEXT,
                            MepServiceType TEXT,
                            ElevationFromLevel REAL,
                            MarkedForClusterProcess INTEGER
                        )";
                    cmd.ExecuteNonQuery();
                    cmd.CommandText = "DELETE FROM BulkUpdateZones";
                    cmd.ExecuteNonQuery();

                    // 2. Fetch current flags to preserve IsCombinedResolved=1
                    var idList = string.Join(",", validZones.Select(z => existingMap[z.Id]));
                    var flagPreserveMap = new Dictionary<int, bool>();
                    cmd.CommandText = $"SELECT ClashZoneId, IsCombinedResolved FROM ClashZones WHERE ClashZoneId IN ({idList})";
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read()) flagPreserveMap[reader.GetInt32(0)] = reader.GetInt32(1) == 1;
                    }

                    // 3. Insert update data into Temp Table
                    cmd.CommandText = @"
                        INSERT INTO BulkUpdateZones (ClashZoneId, IsResolvedFlag, IsClusterResolvedFlag, IsCombinedResolved, SleeveInstanceId, ClusterInstanceId, MepParameterValuesJson, HostParameterValuesJson, WallCenterlinePointX, WallCenterlinePointY, WallCenterlinePointZ, MepOrientationX, MepOrientationY, MepOrientationZ, MepRotationAngleRad, MepRotationAngleDeg, MepOrientationDirection, StructuralThickness, MepElementTypeName, MepElementFamilyName, MepSystemType, MepServiceType, ElevationFromLevel, MarkedForClusterProcess) 
                        VALUES (@ClashZoneId, @IsResolvedFlag, @IsClusterResolvedFlag, @IsCombinedResolved, @SleeveInstanceId, @ClusterInstanceId, @MepParameterValuesJson, @HostParameterValuesJson, @WallCenterlinePointX, @WallCenterlinePointY, @WallCenterlinePointZ, @MepOrientationX, @MepOrientationY, @MepOrientationZ, @MepRotationAngleRad, @MepRotationAngleDeg, @MepOrientationDirection, @StructuralThickness, @MepElementTypeName, @MepElementFamilyName, @MepSystemType, @MepServiceType, @ElevationFromLevel, @MarkedForClusterProcess)";

                    var pId = cmd.Parameters.Add("@ClashZoneId", System.Data.DbType.Int32);
                    var pRes = cmd.Parameters.Add("@IsResolvedFlag", System.Data.DbType.Int32);
                    var pClust = cmd.Parameters.Add("@IsClusterResolvedFlag", System.Data.DbType.Int32);
                    var pCombo = cmd.Parameters.Add("@IsCombinedResolved", System.Data.DbType.Int32);
                    var pSleeve = cmd.Parameters.Add("@SleeveInstanceId", System.Data.DbType.Int32);
                    var pClustSleeve = cmd.Parameters.Add("@ClusterInstanceId", System.Data.DbType.Int32);
                    var pMepP = cmd.Parameters.Add("@MepParameterValuesJson", System.Data.DbType.String);
                    var pHostP = cmd.Parameters.Add("@HostParameterValuesJson", System.Data.DbType.String);
                    var pCX = cmd.Parameters.Add("@WallCenterlinePointX", System.Data.DbType.Double);
                    var pCY = cmd.Parameters.Add("@WallCenterlinePointY", System.Data.DbType.Double);
                    var pCZ = cmd.Parameters.Add("@WallCenterlinePointZ", System.Data.DbType.Double);
                    var pMX = cmd.Parameters.Add("@MepOrientationX", System.Data.DbType.Double);
                    var pMY = cmd.Parameters.Add("@MepOrientationY", System.Data.DbType.Double);
                    var pMZ = cmd.Parameters.Add("@MepOrientationZ", System.Data.DbType.Double);
                    var pRotRad = cmd.Parameters.Add("@MepRotationAngleRad", System.Data.DbType.Double);
                    var pRotDeg = cmd.Parameters.Add("@MepRotationAngleDeg", System.Data.DbType.Double);
                    var pDir = cmd.Parameters.Add("@MepOrientationDirection", System.Data.DbType.String);
                    var pStructThick = cmd.Parameters.Add("@StructuralThickness", System.Data.DbType.Double);
                    var pMepTypeName = cmd.Parameters.Add("@MepElementTypeName", System.Data.DbType.String);
                    var pMepFamilyName = cmd.Parameters.Add("@MepElementFamilyName", System.Data.DbType.String);
                    var pMepSysType = cmd.Parameters.Add("@MepSystemType", System.Data.DbType.String);
                    var pMepServType = cmd.Parameters.Add("@MepServiceType", System.Data.DbType.String);
                    var pElevLevel = cmd.Parameters.Add("@ElevationFromLevel", System.Data.DbType.Double);
                    var pMarked = cmd.Parameters.Add("@MarkedForClusterProcess", System.Data.DbType.Int32);

                    foreach (var zone in validZones)
                    {
                        int czId = existingMap[zone.Id];
                        bool preserveCombo = flagPreserveMap.TryGetValue(czId, out var pc) && pc;

                        pId.Value = czId;
                        pRes.Value = zone.IsResolved ? 1 : 0;
                        pClust.Value = zone.IsClusterResolved ? 1 : 0;
                        pCombo.Value = (zone.IsCombinedResolved || preserveCombo) ? 1 : 0;
                        pSleeve.Value = zone.SleeveInstanceId;
                        pClustSleeve.Value = zone.ClusterSleeveInstanceId;
                        pMepP.Value = JsonSerializer.Serialize(zone.MepParameterValues?.ToDictionary(kv => kv.Key, kv => kv.Value) ?? new Dictionary<string, string>());
                        pHostP.Value = JsonSerializer.Serialize(zone.HostParameterValues?.ToDictionary(kv => kv.Key, kv => kv.Value) ?? new Dictionary<string, string>());
                        pCX.Value = zone.WallCenterlinePointX;
                        pCY.Value = zone.WallCenterlinePointY;
                        pCZ.Value = zone.WallCenterlinePointZ;
                        pMX.Value = zone.MepOrientationX;
                        pMY.Value = zone.MepOrientationY;
                        pMZ.Value = zone.MepOrientationZ;
                        pRotRad.Value = zone.MepElementRotationAngle; // In radians
                        pRotDeg.Value = zone.MepElementRotationAngle * 180.0 / Math.PI; // Convert to degrees
                        pDir.Value = zone.MepElementOrientationDirection ?? string.Empty;
                        pStructThick.Value = zone.StructuralElementThickness; // ✅ CRITICAL FIX: Ensure thickness is updated
                        pMepTypeName.Value = zone.MepElementTypeName ?? (object)DBNull.Value;
                        pMepFamilyName.Value = zone.MepElementFamilyName ?? (object)DBNull.Value;
                        pMepSysType.Value = zone.MepSystemType ?? (object)DBNull.Value;
                        pMepServType.Value = zone.MepServiceType ?? (object)DBNull.Value;
                        pElevLevel.Value = zone.ElevationFromLevel;
                        pMarked.Value = (object)zone.MarkedForClusterProcess ?? DBNull.Value; // ✅ FIX: Use actual value or NULL (don't default to TRUE)
                        cmd.ExecuteNonQuery();
                    }

                    // 4. Perform Join-based UPDATE
                    cmd.CommandText = @"
                        UPDATE ClashZones 
                        SET 
                            IsResolvedFlag = (SELECT IsResolvedFlag FROM BulkUpdateZones WHERE BulkUpdateZones.ClashZoneId = ClashZones.ClashZoneId),
                            IsClusterResolvedFlag = (SELECT IsClusterResolvedFlag FROM BulkUpdateZones WHERE BulkUpdateZones.ClashZoneId = ClashZones.ClashZoneId),
                            IsCombinedResolved = (SELECT IsCombinedResolved FROM BulkUpdateZones WHERE BulkUpdateZones.ClashZoneId = ClashZones.ClashZoneId),
                            SleeveInstanceId = (SELECT SleeveInstanceId FROM BulkUpdateZones WHERE BulkUpdateZones.ClashZoneId = ClashZones.ClashZoneId),
                            ClusterInstanceId = (SELECT ClusterInstanceId FROM BulkUpdateZones WHERE BulkUpdateZones.ClashZoneId = ClashZones.ClashZoneId),
                            MepParameterValuesJson = (SELECT MepParameterValuesJson FROM BulkUpdateZones WHERE BulkUpdateZones.ClashZoneId = ClashZones.ClashZoneId),
                            HostParameterValuesJson = (SELECT HostParameterValuesJson FROM BulkUpdateZones WHERE BulkUpdateZones.ClashZoneId = ClashZones.ClashZoneId),
                            WallCenterlinePointX = (SELECT WallCenterlinePointX FROM BulkUpdateZones WHERE BulkUpdateZones.ClashZoneId = ClashZones.ClashZoneId),
                            WallCenterlinePointY = (SELECT WallCenterlinePointY FROM BulkUpdateZones WHERE BulkUpdateZones.ClashZoneId = ClashZones.ClashZoneId),
                            WallCenterlinePointZ = (SELECT WallCenterlinePointZ FROM BulkUpdateZones WHERE BulkUpdateZones.ClashZoneId = ClashZones.ClashZoneId),
                            MepOrientationX = (SELECT MepOrientationX FROM BulkUpdateZones WHERE BulkUpdateZones.ClashZoneId = ClashZones.ClashZoneId),
                            MepOrientationY = (SELECT MepOrientationY FROM BulkUpdateZones WHERE BulkUpdateZones.ClashZoneId = ClashZones.ClashZoneId),
                            MepOrientationZ = (SELECT MepOrientationZ FROM BulkUpdateZones WHERE BulkUpdateZones.ClashZoneId = ClashZones.ClashZoneId),
                            MepRotationAngleRad = (SELECT MepRotationAngleRad FROM BulkUpdateZones WHERE BulkUpdateZones.ClashZoneId = ClashZones.ClashZoneId),
                            MepRotationAngleDeg = (SELECT MepRotationAngleDeg FROM BulkUpdateZones WHERE BulkUpdateZones.ClashZoneId = ClashZones.ClashZoneId),
                            MepOrientationDirection = (SELECT MepOrientationDirection FROM BulkUpdateZones WHERE BulkUpdateZones.ClashZoneId = ClashZones.ClashZoneId),
                            StructuralThickness = (SELECT StructuralThickness FROM BulkUpdateZones WHERE BulkUpdateZones.ClashZoneId = ClashZones.ClashZoneId),
                            MepElementTypeName = (SELECT MepElementTypeName FROM BulkUpdateZones WHERE BulkUpdateZones.ClashZoneId = ClashZones.ClashZoneId),
                            MepElementFamilyName = (SELECT MepElementFamilyName FROM BulkUpdateZones WHERE BulkUpdateZones.ClashZoneId = ClashZones.ClashZoneId),
                            MepSystemType = (SELECT MepSystemType FROM BulkUpdateZones WHERE BulkUpdateZones.ClashZoneId = ClashZones.ClashZoneId),
                            MepServiceType = (SELECT MepServiceType FROM BulkUpdateZones WHERE BulkUpdateZones.ClashZoneId = ClashZones.ClashZoneId),
                            ElevationFromLevel = (SELECT ElevationFromLevel FROM BulkUpdateZones WHERE BulkUpdateZones.ClashZoneId = ClashZones.ClashZoneId),
                            MarkedForClusterProcess = (SELECT MarkedForClusterProcess FROM BulkUpdateZones WHERE BulkUpdateZones.ClashZoneId = ClashZones.ClashZoneId),
                            UpdatedAt = CURRENT_TIMESTAMP
                        WHERE ClashZoneId IN (SELECT ClashZoneId FROM BulkUpdateZones)";
                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                _logger($"[SQLite][BULK] ❌ Error in BulkUpdateClashZonesWithTempTable: {ex.Message}");
                // Fallback to individual updates if needed
            }
        }

        /// <summary>
        /// ✅ BATCH OPTIMIZATION: Update placement flags for multiple clash zones in a single transaction.
        /// Uses a temporary table approach for maximum performance.
        /// Updates IsResolved, IsClusterResolved, IsCombinedResolved, SleeveInstanceId, ClusterSleeveInstanceId, and IsCurrentClash flags.
        /// </summary>
        /// <summary>
        /// ✅ BATCH OPTIMIZATION: Update placement flags for multiple clash zones in a single transaction.
        /// Uses a temporary table approach for maximum performance.
        /// Updates IsResolved, IsClusterResolved, IsCombinedResolved, SleeveInstanceId, ClusterSleeveInstanceId, and IsCurrentClash flags.
        /// </summary>
        public void BatchUpdateFlagsWithCurrentClash(List<(System.Guid ClashZoneId, int ClashZoneIntId, bool IsResolvedFlag, bool IsClusterResolvedFlag, bool IsCombinedResolved, int SleeveInstanceId, int ClusterInstanceId, bool IsCurrentClashFlag, bool IsClusteredFlag, bool? MarkedForClusterProcess, int AfterClusterSleeveId, double SleeveWidth, double SleeveHeight, double SleeveDiameter)> updates)
        {
            _logger($"[SQLite][BATCH] 🚀 BatchUpdateFlagsWithCurrentClash called with {updates?.Count ?? 0} updates");

            if (updates == null || updates.Count == 0)
            {
                _logger($"[SQLite][BATCH] ⚠️ No updates to process, returning early");
                return;
            }

            try
            {
                using (var transaction = _context.Connection.BeginTransaction())
                {
                    using (var cmd = _context.Connection.CreateCommand())
                    {
                        cmd.Transaction = transaction;

                        // DROP temp table if exists
                        cmd.CommandText = "DROP TABLE IF EXISTS TempFlagUpdates";
                        cmd.ExecuteNonQuery();

                        // CREATE temp table with INTEGER PK (much faster and safer than TEXT GUID)
                        cmd.CommandText = @"
                            CREATE TEMP TABLE TempFlagUpdates (
                                ClashZoneId INTEGER PRIMARY KEY,
                                ClashZoneGuid TEXT,
                                IsResolvedFlag INTEGER,
                                IsClusterResolvedFlag INTEGER,
                                IsCombinedResolved INTEGER,
                                SleeveInstanceId INTEGER,
                                ClusterInstanceId INTEGER,
                                IsCurrentClashFlag INTEGER,
                                IsClusteredFlag INTEGER,
                                MarkedForClusterProcess INTEGER,
                                AfterClusterSleeveId INTEGER,
                                SleeveWidth REAL,
                                SleeveHeight REAL,
                                SleeveDiameter REAL
                            )";
                        cmd.ExecuteNonQuery();

                        // Bulk insert updates into temp table
                        cmd.CommandText = @"
                            INSERT INTO TempFlagUpdates 
                            (ClashZoneId, ClashZoneGuid, IsResolvedFlag, IsClusterResolvedFlag, IsCombinedResolved, 
                             SleeveInstanceId, ClusterInstanceId, IsCurrentClashFlag, IsClusteredFlag, MarkedForClusterProcess, AfterClusterSleeveId,
                             SleeveWidth, SleeveHeight, SleeveDiameter)
                            VALUES (@ClashZoneId, @ClashZoneGuid, @IsResolvedFlag, @IsClusterResolvedFlag, @IsCombinedResolved, 
                                    @SleeveInstanceId, @ClusterInstanceId, @IsCurrentClashFlag, @IsClusteredFlag, @MarkedForClusterProcess, @AfterClusterSleeveId,
                                    @SleeveWidth, @SleeveHeight, @SleeveDiameter)";

                        var pClashZoneIntId = cmd.CreateParameter();
                        pClashZoneIntId.ParameterName = "@ClashZoneId";
                        cmd.Parameters.Add(pClashZoneIntId);

                        var pClashZoneGuid = cmd.CreateParameter();
                        pClashZoneGuid.ParameterName = "@ClashZoneGuid";
                        cmd.Parameters.Add(pClashZoneGuid);

                        var pIsResolved = cmd.CreateParameter();
                        pIsResolved.ParameterName = "@IsResolvedFlag";
                        cmd.Parameters.Add(pIsResolved);

                        var pIsClusterResolved = cmd.CreateParameter();
                        pIsClusterResolved.ParameterName = "@IsClusterResolvedFlag";
                        cmd.Parameters.Add(pIsClusterResolved);

                        var pIsCombinedResolved = cmd.CreateParameter();
                        pIsCombinedResolved.ParameterName = "@IsCombinedResolved";
                        cmd.Parameters.Add(pIsCombinedResolved);

                        var pSleeveInstanceId = cmd.CreateParameter();
                        pSleeveInstanceId.ParameterName = "@SleeveInstanceId";
                        cmd.Parameters.Add(pSleeveInstanceId);

                        var pClusterInstanceId = cmd.CreateParameter();
                        pClusterInstanceId.ParameterName = "@ClusterInstanceId";
                        cmd.Parameters.Add(pClusterInstanceId);

                        var pIsCurrentClash = cmd.CreateParameter();
                        pIsCurrentClash.ParameterName = "@IsCurrentClashFlag";
                        cmd.Parameters.Add(pIsCurrentClash);

                        var pIsClusteredFlag = cmd.CreateParameter();
                        pIsClusteredFlag.ParameterName = "@IsClusteredFlag";
                        cmd.Parameters.Add(pIsClusteredFlag);

                        var pMarkedForClusterProcess = cmd.CreateParameter();
                        pMarkedForClusterProcess.ParameterName = "@MarkedForClusterProcess";
                        cmd.Parameters.Add(pMarkedForClusterProcess);

                        var pAfterClusterSleeveId = cmd.CreateParameter();
                        pAfterClusterSleeveId.ParameterName = "@AfterClusterSleeveId";
                        cmd.Parameters.Add(pAfterClusterSleeveId);

                        var pSleeveWidth = cmd.CreateParameter();
                        pSleeveWidth.ParameterName = "@SleeveWidth";
                        cmd.Parameters.Add(pSleeveWidth);

                        var pSleeveHeight = cmd.CreateParameter();
                        pSleeveHeight.ParameterName = "@SleeveHeight";
                        cmd.Parameters.Add(pSleeveHeight);

                        var pSleeveDiameter = cmd.CreateParameter();
                        pSleeveDiameter.ParameterName = "@SleeveDiameter";
                        cmd.Parameters.Add(pSleeveDiameter);

                        foreach (var update in updates)
                        {
                            pClashZoneIntId.Value = update.ClashZoneIntId;
                            pClashZoneGuid.Value = update.ClashZoneId.ToString();
                            pIsResolved.Value = update.IsResolvedFlag ? 1 : 0;
                            pIsClusterResolved.Value = update.IsClusterResolvedFlag ? 1 : 0;
                            pIsCombinedResolved.Value = update.IsCombinedResolved ? 1 : 0;
                            pSleeveInstanceId.Value = update.SleeveInstanceId;
                            pClusterInstanceId.Value = update.ClusterInstanceId;
                            pIsCurrentClash.Value = update.IsCurrentClashFlag ? 1 : 0;
                            pIsClusteredFlag.Value = update.IsClusteredFlag ? 1 : 0;
                            pMarkedForClusterProcess.Value = (object)(update.MarkedForClusterProcess.HasValue ? (update.MarkedForClusterProcess.Value ? 1 : 0) : DBNull.Value);
                            pAfterClusterSleeveId.Value = update.AfterClusterSleeveId;
                            pSleeveWidth.Value = update.SleeveWidth;
                            pSleeveHeight.Value = update.SleeveHeight;
                            pSleeveDiameter.Value = update.SleeveDiameter;

                            cmd.ExecuteNonQuery();
                        }
                    }

                    // ✅ CRITICAL: Execute UPDATE using JOIN from temp table to main table via INTEGER KEY (Fast!)
                    using (var updateCmd = _context.Connection.CreateCommand())
                    {
                        updateCmd.Transaction = transaction;
                        updateCmd.CommandText = @"
                            UPDATE ClashZones
                            SET 
                                IsResolvedFlag = CASE WHEN (SELECT IsResolvedFlag FROM TempFlagUpdates WHERE TempFlagUpdates.ClashZoneId = ClashZones.ClashZoneId) = 1 THEN 1 ELSE IsResolvedFlag END,
                                IsClusterResolvedFlag = CASE WHEN (SELECT IsClusterResolvedFlag FROM TempFlagUpdates WHERE TempFlagUpdates.ClashZoneId = ClashZones.ClashZoneId) = 1 THEN 1 ELSE IsClusterResolvedFlag END,
                                IsCombinedResolved = CASE WHEN (SELECT IsCombinedResolved FROM TempFlagUpdates WHERE TempFlagUpdates.ClashZoneId = ClashZones.ClashZoneId) = 1 THEN 1 ELSE IsCombinedResolved END,
                                SleeveInstanceId = CASE WHEN (SELECT SleeveInstanceId FROM TempFlagUpdates WHERE TempFlagUpdates.ClashZoneId = ClashZones.ClashZoneId) > 0 THEN (SELECT SleeveInstanceId FROM TempFlagUpdates WHERE TempFlagUpdates.ClashZoneId = ClashZones.ClashZoneId) ELSE SleeveInstanceId END,
                                ClusterInstanceId = CASE WHEN (SELECT ClusterInstanceId FROM TempFlagUpdates WHERE TempFlagUpdates.ClashZoneId = ClashZones.ClashZoneId) > 0 THEN (SELECT ClusterInstanceId FROM TempFlagUpdates WHERE TempFlagUpdates.ClashZoneId = ClashZones.ClashZoneId) ELSE ClusterInstanceId END,
                                IsCurrentClashFlag = (SELECT IsCurrentClashFlag FROM TempFlagUpdates WHERE TempFlagUpdates.ClashZoneId = ClashZones.ClashZoneId),
                                IsClusteredFlag = (SELECT IsClusteredFlag FROM TempFlagUpdates WHERE TempFlagUpdates.ClashZoneId = ClashZones.ClashZoneId),
                                ReadyForPlacementFlag = (SELECT IsCurrentClashFlag FROM TempFlagUpdates WHERE TempFlagUpdates.ClashZoneId = ClashZones.ClashZoneId),
                                MarkedForClusterProcess = (SELECT MarkedForClusterProcess FROM TempFlagUpdates WHERE TempFlagUpdates.ClashZoneId = ClashZones.ClashZoneId),
                                AfterClusterSleeveId = (SELECT AfterClusterSleeveId FROM TempFlagUpdates WHERE TempFlagUpdates.ClashZoneId = ClashZones.ClashZoneId),
                                SleeveWidth = CASE WHEN (SELECT SleeveWidth FROM TempFlagUpdates WHERE TempFlagUpdates.ClashZoneId = ClashZones.ClashZoneId) > 0 THEN (SELECT SleeveWidth FROM TempFlagUpdates WHERE TempFlagUpdates.ClashZoneId = ClashZones.ClashZoneId) ELSE SleeveWidth END,
                                SleeveHeight = CASE WHEN (SELECT SleeveHeight FROM TempFlagUpdates WHERE TempFlagUpdates.ClashZoneId = ClashZones.ClashZoneId) > 0 THEN (SELECT SleeveHeight FROM TempFlagUpdates WHERE TempFlagUpdates.ClashZoneId = ClashZones.ClashZoneId) ELSE SleeveHeight END,
                                SleeveDiameter = CASE WHEN (SELECT SleeveDiameter FROM TempFlagUpdates WHERE TempFlagUpdates.ClashZoneId = ClashZones.ClashZoneId) > 0 THEN (SELECT SleeveDiameter FROM TempFlagUpdates WHERE TempFlagUpdates.ClashZoneId = ClashZones.ClashZoneId) ELSE SleeveDiameter END,
                                UpdatedAt = CURRENT_TIMESTAMP
                            WHERE ClashZoneId IN (SELECT ClashZoneId FROM TempFlagUpdates)";
                        int rowsAffected = updateCmd.ExecuteNonQuery();
                        _logger($"[SQLite][BATCH] ✅ UPDATE executed: {rowsAffected} rows affected");
                    }

                    // ✅ DIAGNOSTIC: Check if the flags were actually set
                    if (updates.Count > 0)
                    {
                        var sampleId = updates[0].ClashZoneIntId;
                        using (var verifyCmd = _context.Connection.CreateCommand())
                        {
                            verifyCmd.Transaction = transaction;
                            verifyCmd.CommandText = $"SELECT IsResolvedFlag, IsClusterResolvedFlag, SleeveInstanceId, ClusterInstanceId FROM ClashZones WHERE ClashZoneId = {sampleId}";
                            using (var reader = verifyCmd.ExecuteReader())
                            {
                                if (reader.Read())
                                {
                                    int isResolved = GetInt(reader, "IsResolvedFlag", 0);
                                    int isClusterResolved = GetInt(reader, "IsClusterResolvedFlag", 0);
                                    int sleeveId = GetInt(reader, "SleeveInstanceId", 0);
                                    int clusterId = GetInt(reader, "ClusterInstanceId", 0);
                                    _logger($"[SQLite][BATCH] 🔍 VERIFY: ClashZone ID {sampleId} after update: IsResolvedFlag={isResolved}, IsClusterResolvedFlag={isClusterResolved}, SleeveId={sleeveId}, ClusterId={clusterId}");
                                }
                                else
                                {
                                    _logger($"[SQLite][BATCH] ❌ ERROR: ClashZone ID {sampleId} not found in database!");
                                }
                            }
                        }
                    }

                    // Clean up temp table
                    using (var dropCmd = _context.Connection.CreateCommand())
                    {
                        dropCmd.Transaction = transaction;
                        dropCmd.CommandText = "DROP TABLE IF EXISTS TempFlagUpdates";
                        dropCmd.ExecuteNonQuery();
                    }

                    // ✅ CRITICAL: Commit the transaction to persist changes
                    transaction.Commit();
                    _logger($"[SQLite][BATCH] ✅ Transaction committed successfully for {updates.Count} flag updates");
                }
            }
            catch (Exception ex)
            {
                _logger($"[SQLite][BATCH] ❌ Error in BatchUpdateFlagsWithCurrentClash: {ex.Message}");
                _logger($"[SQLite][BATCH] ❌ Stack trace: {ex.StackTrace}");
                throw;
            }
        }

        private void BulkUpdateClashZonesLegacy(List<ClashZone> zones, Dictionary<Guid, int> existingMap, Dictionary<Guid, int> comboMap, SQLiteTransaction transaction)
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
                var flagMap = new Dictionary<int, (bool IsResolved, bool IsClusterResolved, bool IsCombinedResolved, int SleeveId, int ClusterId)>();
                var idList = string.Join(",", validZonesForQuery.Select(z => existingMap[z.Id]));

                cmd.CommandText = $@"
                    SELECT ClashZoneId, IsResolvedFlag, IsClusterResolvedFlag, IsCombinedResolved, SleeveInstanceId, ClusterInstanceId
                    FROM ClashZones WHERE ClashZoneId IN ({idList})";

                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        var id = reader.GetInt32(0);
                        // SQLite stores booleans as integers (0/1). Avoid GetBoolean cast errors.
                        var isResolved = !reader.IsDBNull(1) && reader.GetInt32(1) == 1;
                        var isClusterResolved = !reader.IsDBNull(2) && reader.GetInt32(2) == 1;
                        var isCombinedResolved = !reader.IsDBNull(3) && reader.GetInt32(3) == 1;
                        var sleeveId = reader.IsDBNull(4) ? -1 : reader.GetInt32(4);
                        var clusterId = reader.IsDBNull(5) ? -1 : reader.GetInt32(5);
                        flagMap[id] = (isResolved, isClusterResolved, isCombinedResolved, sleeveId, clusterId);
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
                    bool finalIsCombinedResolved = zone.IsCombinedResolved;
                    int finalSleeveId = zone.SleeveInstanceId;
                    int finalClusterId = zone.ClusterSleeveInstanceId;

                    if (flagMap.TryGetValue(clashZoneId, out var dbFlags))
                    {
                        // ✅ CRITICAL: ALWAYS preserve IsCombinedResolved=1 from database
                        // This prevents zones in combined sleeves from being overwritten
                        if (dbFlags.IsCombinedResolved)
                        {
                            finalIsCombinedResolved = true;
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                _logger($"[SQLite][BULK] ✅ PRESERVING IsCombinedResolved=1 for ClashZoneId={clashZoneId}, GUID={zone.Id}");
                            }
                        }

                        // If database has reset flags (false, false, -1, -1), preserve them
                        if (!dbFlags.IsResolved && !dbFlags.IsClusterResolved && !dbFlags.IsCombinedResolved && dbFlags.SleeveId == -1 && dbFlags.ClusterId == -1)
                        {
                            finalIsResolved = false;
                            finalIsClusterResolved = false;
                            finalIsCombinedResolved = false;
                            finalSleeveId = -1;
                            finalClusterId = -1;

                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                _logger($"[SQLite][BULK] ✅ PRESERVING reset flags for ClashZoneId={clashZoneId}, GUID={zone.Id}: IsResolved=false, IsClusterResolved=false, IsCombinedResolved=false");
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
                    cmd.Parameters.AddWithValue($"@IsCombinedResolved{i}", finalIsCombinedResolved ? 1 : 0);
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
                    cmd.Parameters.AddWithValue($"@WallCenterlinePointX{i}", zone.WallCenterlinePointX);
                    cmd.Parameters.AddWithValue($"@WallCenterlinePointY{i}", zone.WallCenterlinePointY);
                    cmd.Parameters.AddWithValue($"@WallCenterlinePointZ{i}", zone.WallCenterlinePointZ);

                    // ✅ CRITICAL FIX: Add MEP System Name to bulk update parameters
                    cmd.Parameters.AddWithValue($"@MepSystemName{i}", (object)zone.MepSystemName ?? DBNull.Value);

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
                sql.AppendLine("  WallCenterlinePointX = CASE ClashZoneId");
                for (int i = 0; i < validZones.Count; i++)
                {
                    sql.AppendLine($"    WHEN @ZoneId{i} THEN @WallCenterlinePointX{i}");
                }
                sql.AppendLine("    ELSE WallCenterlinePointX END,");

                sql.AppendLine("  WallCenterlinePointY = CASE ClashZoneId");
                for (int i = 0; i < validZones.Count; i++)
                {
                    sql.AppendLine($"    WHEN @ZoneId{i} THEN @WallCenterlinePointY{i}");
                }
                sql.AppendLine("    ELSE WallCenterlinePointY END,");

                sql.AppendLine("  WallCenterlinePointZ = CASE ClashZoneId");
                for (int i = 0; i < validZones.Count; i++)
                {
                    sql.AppendLine($"    WHEN @ZoneId{i} THEN @WallCenterlinePointZ{i}");
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

                sql.AppendLine("  IsCombinedResolved = CASE ClashZoneId");
                for (int i = 0; i < validZones.Count; i++)
                {
                    sql.AppendLine($"    WHEN @ZoneId{i} THEN @IsCombinedResolved{i}");
                }
                sql.AppendLine("    ELSE IsCombinedResolved END,");

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

                // ✅ CRITICAL FIX: Add MEP System Name to bulk update parameters
                sql.AppendLine("  MepSystemName = CASE ClashZoneId");
                for (int i = 0; i < validZones.Count; i++)
                {
                    sql.AppendLine($"    WHEN @ZoneId{i} THEN @MepSystemName{i}");
                }
                sql.AppendLine("    ELSE MepSystemName END,");

                sql.AppendLine("  MepServiceType = CASE ClashZoneId");
                for (int i = 0; i < validZones.Count; i++)
                {
                    sql.AppendLine($"    WHEN @ZoneId{i} THEN @MepServiceType{i}");
                }
                sql.AppendLine("    ELSE MepServiceType END,");

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

        /// <summary>
        /// ✅ PERFORMANCE OPTIMIZED: True multi-row INSERT for clash zones.
        /// Replaces individual INSERTs which were causing a 19.2s bottleneck.
        /// Expected: < 100ms for 1000 zones.
        /// </summary>
        private void BulkInsertClashZones(List<ClashZone> zones, Dictionary<Guid, int> comboMap, SQLiteTransaction transaction)
        {
            if (zones == null || zones.Count == 0) return;

            const int batchSize = 100; // SQLite limit around 500 parameters per query
            for (int i = 0; i < zones.Count; i += batchSize)
            {
                var currentBatch = zones.Skip(i).Take(batchSize).ToList();
                using (var cmd = _context.Connection.CreateCommand())
                {
                    cmd.Transaction = transaction;
                    var sql = new System.Text.StringBuilder();
                    sql.Append(@"INSERT INTO ClashZones (
                        ClashZoneGuid, ComboId, MepElementId, HostElementId, 
                        IntersectionX, IntersectionY, IntersectionZ, 
                        WallCenterlinePointX, WallCenterlinePointY, WallCenterlinePointZ,
                        MepCategory, StructuralType, 
                        MepWidth, MepHeight,
                        MepElementOuterDiameter, MepElementNominalDiameter,
                        MepElementTypeName, MepElementFamilyName, MepElementSizeParameterValue,
                        MepElementLevelName, MepElementLevelElevation,
                        StructuralThickness, WallThickness, FramingThickness,
                        IsInsulated, InsulationThickness,
                        HasMepConnector, DamperConnectorSide,
                        SourceDocKey, HostDocKey, MepElementUniqueId,
                        HostOrientation, MepOrientationDirection,
                        MepOrientationX, MepOrientationY, MepOrientationZ,
                        MepRotationAngleRad, MepRotationAngleDeg,
                        MepAngleToXRad, MepAngleToXDeg, MepAngleToYRad, MepAngleToYDeg,
                        MepParameterValuesJson, HostParameterValuesJson,
                        MepElementSystemAbbreviation, MepElementFormattedSize, IsStandardDamper,
                        MepSystemType, MepSystemName, MepServiceType, ElevationFromLevel,
                        IsCurrentClashFlag, ReadyForPlacementFlag, UpdatedAt, MarkedForClusterProcess
                    ) VALUES ");

                    for (int j = 0; j < currentBatch.Count; j++)
                    {
                        var zone = currentBatch[j];
                        var comboId = comboMap[zone.Id];
                        var mepId = zone.MepElementId?.IntegerValue ?? zone.MepElementIdValue;
                        var hostId = zone.StructuralElementId?.IntegerValue ?? zone.StructuralElementIdValue;
                        var interX = zone.IntersectionPoint?.X ?? zone.IntersectionPointX;
                        var interY = zone.IntersectionPoint?.Y ?? zone.IntersectionPointY;
                        var interZ = zone.IntersectionPoint?.Z ?? zone.IntersectionPointZ;
                        var wallX = zone.WallCenterlinePoint?.X ?? interX;
                        var wallY = zone.WallCenterlinePoint?.Y ?? interY;
                        var wallZ = zone.WallCenterlinePoint?.Z ?? interZ;

                        sql.Append($"(@G{j}, @C{j}, @M{j}, @H{j}, @IX{j}, @IY{j}, @IZ{j}, @WX{j}, @WY{j}, @WZ{j}, " +
                                   $"@CAT{j}, @ST{j}, @MW{j}, @MH{j}, @OD{j}, @ND{j}, @TN{j}, @FN{j}, @SP{j}, " +
                                   $"@LN{j}, @LE{j}, @STH{j}, @WTH{j}, @FTH{j}, @INS{j}, @INSTH{j}, @CONN{j}, @CONNSIDE{j}, " +
                                   $"@SDK{j}, @HDK{j}, @UID{j}, @HO{j}, @MOD{j}, @MOX{j}, @MOY{j}, @MOZ{j}, " +
                                    $"@MRAR{j}, @MRAD{j}, @MAXR{j}, @MAXD{j}, @MAYR{j}, @MAYD{j}, " +
                                   $"@MPJ{j}, @HPJ{j}, @MSA{j}, @MFS{j}, @ISD{j}, " +
                                   $"@MST{j}, @MSN{j}, @MSVT{j}, @EFL{j}, 1, 1, @T{j}, @MFCP{j})");
                        if (j < currentBatch.Count - 1) sql.Append(",");

                        cmd.Parameters.AddWithValue($"@G{j}", zone.Id.ToString());
                        cmd.Parameters.AddWithValue($"@C{j}", comboId);
                        cmd.Parameters.AddWithValue($"@M{j}", mepId);
                        cmd.Parameters.AddWithValue($"@H{j}", hostId);
                        cmd.Parameters.AddWithValue($"@IX{j}", interX);
                        cmd.Parameters.AddWithValue($"@IY{j}", interY);
                        cmd.Parameters.AddWithValue($"@IZ{j}", interZ);
                        cmd.Parameters.AddWithValue($"@WX{j}", wallX);
                        cmd.Parameters.AddWithValue($"@WY{j}", wallY);
                        cmd.Parameters.AddWithValue($"@WZ{j}", wallZ);
                        cmd.Parameters.AddWithValue($"@CAT{j}", zone.MepElementCategory ?? string.Empty);
                        cmd.Parameters.AddWithValue($"@ST{j}", zone.StructuralElementType ?? "Walls");
                        cmd.Parameters.AddWithValue($"@MW{j}", zone.MepElementWidth);
                        cmd.Parameters.AddWithValue($"@MH{j}", zone.MepElementHeight);
                        cmd.Parameters.AddWithValue($"@OD{j}", zone.MepElementOuterDiameter);
                        cmd.Parameters.AddWithValue($"@ND{j}", zone.MepElementNominalDiameter);
                        cmd.Parameters.AddWithValue($"@TN{j}", zone.MepElementTypeName ?? string.Empty);
                        cmd.Parameters.AddWithValue($"@FN{j}", zone.MepElementFamilyName ?? string.Empty);
                        cmd.Parameters.AddWithValue($"@SP{j}", zone.MepElementSizeParameterValue ?? string.Empty);
                        cmd.Parameters.AddWithValue($"@LN{j}", zone.MepElementLevelName ?? string.Empty);
                        cmd.Parameters.AddWithValue($"@LE{j}", zone.MepElementLevelElevation);
                        // ✅ STRUCTURAL THICKNESS FIX: Use actual thickness (required for cluster depth calculation on floors)
                        cmd.Parameters.AddWithValue($"@STH{j}", zone.StructuralElementThickness);
                        cmd.Parameters.AddWithValue($"@WTH{j}", zone.WallThickness);
                        cmd.Parameters.AddWithValue($"@FTH{j}", zone.FramingThickness);
                        cmd.Parameters.AddWithValue($"@INS{j}", zone.IsInsulated ? 1 : 0);
                        cmd.Parameters.AddWithValue($"@INSTH{j}", zone.InsulationThickness);
                        cmd.Parameters.AddWithValue($"@CONN{j}", zone.HasMepConnector ? 1 : 0);
                        cmd.Parameters.AddWithValue($"@CONNSIDE{j}", zone.DamperConnectorSide ?? string.Empty);
                        cmd.Parameters.AddWithValue($"@SDK{j}", zone.SourceDocKey ?? string.Empty);
                        cmd.Parameters.AddWithValue($"@HDK{j}", zone.HostDocKey ?? string.Empty);
                        cmd.Parameters.AddWithValue($"@UID{j}", zone.MepElementUniqueId ?? string.Empty);
                        cmd.Parameters.AddWithValue($"@HO{j}", zone.HostOrientation ?? string.Empty);

                        // ✅ MEP ORIENTATION DIRECTION: Populated in enrichment loop
                        cmd.Parameters.AddWithValue($"@MOD{j}", zone.MepElementOrientationDirection ?? string.Empty);

                        // ✅ MEP ORIENTATION FIX: Use actual orientation vector from ClashZone (calculated during refresh)
                        cmd.Parameters.AddWithValue($"@MOX{j}", zone.MepOrientationX);
                        cmd.Parameters.AddWithValue($"@MOY{j}", zone.MepOrientationY);
                        cmd.Parameters.AddWithValue($"@MOZ{j}", zone.MepOrientationZ);

                        // ✅ FLOOR ROTATION FIX: Use actual rotation angle from ClashZone (calculated during refresh)
                        double rotationAngleRad = zone.MepElementRotationAngle; // In radians
                        double rotationAngleDeg = rotationAngleRad * 180.0 / Math.PI; // Convert to degrees
                        cmd.Parameters.AddWithValue($"@MRAR{j}", rotationAngleRad);
                        cmd.Parameters.AddWithValue($"@MRAD{j}", rotationAngleDeg);

                        cmd.Parameters.AddWithValue($"@MAXR{j}", 0.0); // MepAngleToXRad - populated later
                        cmd.Parameters.AddWithValue($"@MAXD{j}", 0.0); // MepAngleToXDeg - populated later
                        cmd.Parameters.AddWithValue($"@MAYR{j}", 0.0); // MepAngleToYRad - populated later
                        cmd.Parameters.AddWithValue($"@MAYD{j}", 0.0); // MepAngleToYDeg - populated later

                        // ✅ CRITICAL: Serialize MEP and Host parameter values to JSON as Dictionary format
                        // Uses Dictionary format {"key":"value"} for consistency with AddClashZoneParameters
                        string mepParamJson = "{}";
                        if (zone.MepParameterValues != null && zone.MepParameterValues.Count > 0)
                        {
                            try
                            {
                                var mepDict = zone.MepParameterValues
                                    .Where(kv => kv != null && !string.IsNullOrEmpty(kv.Key))
                                    .ToDictionary(kv => kv.Key, kv => kv.Value ?? string.Empty, StringComparer.OrdinalIgnoreCase);
                                mepParamJson = System.Text.Json.JsonSerializer.Serialize(mepDict);
                            }
                            catch { mepParamJson = "{}"; }
                        }
                        cmd.Parameters.AddWithValue($"@MPJ{j}", mepParamJson);

                        string hostParamJson = "{}";
                        if (zone.HostParameterValues != null && zone.HostParameterValues.Count > 0)
                        {
                            try
                            {
                                var hostDict = zone.HostParameterValues
                                    .Where(kv => kv != null && !string.IsNullOrEmpty(kv.Key))
                                    .ToDictionary(kv => kv.Key, kv => kv.Value ?? string.Empty, StringComparer.OrdinalIgnoreCase);
                                hostParamJson = System.Text.Json.JsonSerializer.Serialize(hostDict);
                            }
                            catch { hostParamJson = "{}"; }
                        }
                        cmd.Parameters.AddWithValue($"@HPJ{j}", hostParamJson);


                        cmd.Parameters.AddWithValue($"@MSA{j}", zone.MepElementSystemAbbreviation ?? string.Empty);
                        cmd.Parameters.AddWithValue($"@MFS{j}", zone.MepElementFormattedSize ?? string.Empty);
                        cmd.Parameters.AddWithValue($"@ISD{j}", zone.IsStandardDamper ? 1 : 0);
                        cmd.Parameters.AddWithValue($"@ISD{j}", zone.IsStandardDamper ? 1 : 0);
                        cmd.Parameters.AddWithValue($"@MST{j}", zone.MepSystemType ?? string.Empty);
                        // ✅ CRITICAL FIX: Add "MSN" (MepSystemName) parameter
                        cmd.Parameters.AddWithValue($"@MSN{j}", zone.MepSystemName ?? string.Empty);
                        cmd.Parameters.AddWithValue($"@MSVT{j}", zone.MepServiceType ?? string.Empty);
                        cmd.Parameters.AddWithValue($"@EFL{j}", zone.ElevationFromLevel);

                        // ✅ CRITICAL FIX: Populate MarkedForClusterProcess for new zones (Default to NULL for discovery)
                        cmd.Parameters.AddWithValue($"@MFCP{j}", (object)zone.MarkedForClusterProcess ?? DBNull.Value);

                        cmd.Parameters.AddWithValue($"@T{j}", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                    }

                    cmd.CommandText = sql.ToString();
                    cmd.ExecuteNonQuery();
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

                            var comboId = GetOrCreateFileCombo(filterId, mepCategory, hostCategories ?? new List<string>(), clashZone, transaction);

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
            // ✅ NORMALIZE: Remove category suffixes from filter name to prevent duplicates
            // This ensures "Plumbing_pipes" and "Plumbing" both resolve to "Plumbing"
            filterName = FilterNameHelper.NormalizeBaseName(filterName, filterName, category);

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

            // ✅ STEP 2: Filter doesn't exist - CREATE it with SelectedHostCategories
            // This ensures clash zones can be saved even if filter wasn't pre-created
            using (var insertCmd = _context.Connection.CreateCommand())
            {
                insertCmd.Transaction = transaction;
                insertCmd.CommandText = @"
                    INSERT INTO Filters (FilterName, Category, SelectedHostCategories, IsFilterComboNew, CreatedAt, UpdatedAt)
                    VALUES (@FilterName, @Category, @SelectedHostCategories, 0, CURRENT_TIMESTAMP, CURRENT_TIMESTAMP);
                    SELECT last_insert_rowid();";
                insertCmd.Parameters.AddWithValue("@FilterName", filterName);
                insertCmd.Parameters.AddWithValue("@Category", category);
                insertCmd.Parameters.AddWithValue("@SelectedHostCategories", "Walls"); // Default to Walls for all filters

                var newId = insertCmd.ExecuteScalar();
                if (newId != null && int.TryParse(newId.ToString(), out int insertedId))
                {
                    _logger($"[SQLite] ✅ Created filter '{filterName}' (Category='{category}', SelectedHostCategories='Walls') in database (FilterId={insertedId})");
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

        /// <summary>
        /// ⚡ BATCH OPTIMIZATION: Get or create file combos for all zones in a single batch operation
        /// Replaces N individual GetOrCreateFileCombo calls with 1 SELECT + 1 multi-row INSERT
        /// Expected improvement: 10-15s → <1s for 888 zones
        /// </summary>
        private Dictionary<Guid, int> BatchGetOrCreateFileCombos(
            List<ClashZone> zones,
            int filterId,
            string category,
            SQLiteTransaction transaction)
        {
            var sw = System.Diagnostics.Stopwatch.StartNew();
            var result = new Dictionary<Guid, int>();

            // STEP 1: Extract unique file combo tuples
            var uniqueCombos = new Dictionary<string, (string LinkedFile, string HostFile, string Category, List<string> HostCategories)>();

            foreach (var zone in zones)
            {
                var linkedFileKey = NormalizeDocumentKey(
                    !string.IsNullOrWhiteSpace(zone.SourceDocKey)
                        ? zone.SourceDocKey
                        : (!string.IsNullOrWhiteSpace(zone.DocumentPath) ? zone.DocumentPath : "unknown-linked"));

                var hostFileKey = NormalizeDocumentKey(
                    !string.IsNullOrWhiteSpace(zone.HostDocKey)
                        ? zone.HostDocKey
                        : (!string.IsNullOrWhiteSpace(zone.StructuralElementDocumentTitle) ? zone.StructuralElementDocumentTitle : "unknown-host"));

                var normalizedCategory = MepCategoryConstants.Normalize(zone.MepElementCategory ?? category ?? "Unknown");

                var key = $"{linkedFileKey}|{hostFileKey}|{normalizedCategory}";
                if (!uniqueCombos.ContainsKey(key))
                {
                    List<string> hostCategories = null;
                    if (FilterUiStateProvider.GetSelectedHostCategories != null)
                    {
                        hostCategories = FilterUiStateProvider.GetSelectedHostCategories.Invoke();
                    }
                    uniqueCombos[key] = (linkedFileKey, hostFileKey, normalizedCategory, hostCategories);
                }
            }

            _logger($"[SQLite][BATCH] Extracted {uniqueCombos.Count} unique file combos from {zones.Count} zones");

            // STEP 2: Single SELECT to find existing combos
            var existingCombos = new Dictionary<string, int>(); // Key: "LinkedFile|HostFile|Category", Value: ComboId

            if (uniqueCombos.Count > 0)
            {
                var whereConditions = new List<string>();
                using (var cmd = _context.Connection.CreateCommand())
                {
                    cmd.Transaction = transaction;

                    int paramIndex = 0;
                    foreach (var kvp in uniqueCombos)
                    {
                        whereConditions.Add($"(LinkedFileKey = @Linked{paramIndex} AND HostFileKey = @Host{paramIndex} AND Category = @Cat{paramIndex})");
                        cmd.Parameters.AddWithValue($"@Linked{paramIndex}", kvp.Value.LinkedFile);
                        cmd.Parameters.AddWithValue($"@Host{paramIndex}", kvp.Value.HostFile);
                        cmd.Parameters.AddWithValue($"@Cat{paramIndex}", kvp.Value.Category);
                        paramIndex++;
                    }

                    cmd.CommandText = $@"
                        SELECT ComboId, LinkedFileKey, HostFileKey, Category 
                        FROM FileCombos 
                        WHERE FilterId = @FilterId AND ({string.Join(" OR ", whereConditions)})";
                    cmd.Parameters.AddWithValue("@FilterId", filterId);

                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            var comboId = reader.GetInt32(0);
                            var linked = reader.GetString(1);
                            var host = reader.GetString(2);
                            var cat = reader.GetString(3);
                            var key = $"{linked}|{host}|{cat}";
                            existingCombos[key] = comboId;
                        }
                    }
                }

                _logger($"[SQLite][BATCH] Found {existingCombos.Count}/{uniqueCombos.Count} existing file combos");
                if (existingCombos.Count > 0)
                    _logger($"[SQLite][BATCH-DEBUG] Sample existing combo key: {existingCombos.Keys.First()}");
            }

            // STEP 3: Single multi-row INSERT for new combos
            var newCombos = uniqueCombos.Where(kvp => !existingCombos.ContainsKey(kvp.Key)).ToList();
            _logger($"[SQLite][BATCH-DEBUG] New combos to insert: {newCombos.Count}");
            if (newCombos.Count > 0)
                _logger($"[SQLite][BATCH-DEBUG] Sample new combo: Linked='{newCombos[0].Value.LinkedFile}', Host='{newCombos[0].Value.HostFile}', Cat='{newCombos[0].Value.Category}'");

            if (newCombos.Count > 0)
            {
                using (var cmd = _context.Connection.CreateCommand())
                {
                    cmd.Transaction = transaction;

                    var valuesClauses = new List<string>();
                    for (int i = 0; i < newCombos.Count; i++)
                    {
                        valuesClauses.Add($"(@FilterId, @Cat{i}, @HostCats{i}, @Linked{i}, @Host{i}, 1, CURRENT_TIMESTAMP, CURRENT_TIMESTAMP, CURRENT_TIMESTAMP)");

                        var combo = newCombos[i].Value;
                        cmd.Parameters.AddWithValue($"@Cat{i}", combo.Category);

                        string hostCategoriesJson = null;
                        if (combo.HostCategories != null && combo.HostCategories.Count > 0)
                        {
                            var normalizedCategories = combo.HostCategories
                                .Where(c => !string.IsNullOrWhiteSpace(c))
                                .Select(c => c.Trim())
                                .ToList();
                            if (normalizedCategories.Count > 0)
                            {
                                hostCategoriesJson = string.Join(", ", normalizedCategories);
                            }
                        }
                        cmd.Parameters.AddWithValue($"@HostCats{i}", hostCategoriesJson ?? (object)DBNull.Value);
                        cmd.Parameters.AddWithValue($"@Linked{i}", combo.LinkedFile);
                        cmd.Parameters.AddWithValue($"@Host{i}", combo.HostFile);
                    }

                    cmd.Parameters.AddWithValue("@FilterId", filterId);

                    cmd.CommandText = $@"
                        INSERT INTO FileCombos (FilterId, Category, SelectedHostCategories, LinkedFileKey, HostFileKey, IsFilterComboNew, ProcessedAt, CreatedAt, UpdatedAt)
                        VALUES {string.Join(", ", valuesClauses)}";

                    var insertedRows = cmd.ExecuteNonQuery();
                    _logger($"[SQLite][BATCH-DEBUG] INSERT executed, rows affected: {insertedRows}");

                    // Retrieve the new ComboIds
                    cmd.Parameters.Clear();
                    var whereConditions = new List<string>();
                    int paramIndex = 0;
                    foreach (var combo in newCombos)
                    {
                        whereConditions.Add($"(LinkedFileKey = @Linked{paramIndex} AND HostFileKey = @Host{paramIndex} AND Category = @Cat{paramIndex})");
                        cmd.Parameters.AddWithValue($"@Linked{paramIndex}", combo.Value.LinkedFile);
                        cmd.Parameters.AddWithValue($"@Host{paramIndex}", combo.Value.HostFile);
                        cmd.Parameters.AddWithValue($"@Cat{paramIndex}", combo.Value.Category);
                        paramIndex++;
                    }

                    cmd.CommandText = $@"
                        SELECT ComboId, LinkedFileKey, HostFileKey, Category 
                        FROM FileCombos 
                        WHERE FilterId = @FilterId AND ({string.Join(" OR ", whereConditions)})";
                    cmd.Parameters.AddWithValue("@FilterId", filterId);

                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            var comboId = reader.GetInt32(0);
                            var linked = reader.GetString(1);
                            var host = reader.GetString(2);
                            var cat = reader.GetString(3);
                            var key = $"{linked}|{host}|{cat}";
                            existingCombos[key] = comboId;
                        }
                    }
                }

                _logger($"[SQLite][BATCH] Created {newCombos.Count} new file combos");
            }

            // STEP 4: Map zones to comboIds
            foreach (var zone in zones)
            {
                var linkedFileKey = NormalizeDocumentKey(
                    !string.IsNullOrWhiteSpace(zone.SourceDocKey)
                        ? zone.SourceDocKey
                        : (!string.IsNullOrWhiteSpace(zone.DocumentPath) ? zone.DocumentPath : "unknown-linked"));

                var hostFileKey = NormalizeDocumentKey(
                    !string.IsNullOrWhiteSpace(zone.HostDocKey)
                        ? zone.HostDocKey
                        : (!string.IsNullOrWhiteSpace(zone.StructuralElementDocumentTitle) ? zone.StructuralElementDocumentTitle : "unknown-host"));

                var normalizedCategory = MepCategoryConstants.Normalize(zone.MepElementCategory ?? category ?? "Unknown");

                var key = $"{linkedFileKey}|{hostFileKey}|{normalizedCategory}";
                if (existingCombos.TryGetValue(key, out int comboId))
                {
                    result[zone.Id] = comboId;
                }
            }

            sw.Stop();
            _logger($"[SQLite][BATCH-DEBUG] Mapped {result.Count}/{zones.Count} zones to ComboIds");
            if (result.Count < zones.Count)
            {
                _logger($"[SQLite][BATCH-WARNING] ❌ {zones.Count - result.Count} zones NOT mapped to ComboIds!");
            }
            _logger($"[SQLite][BATCH] ⚡ BatchGetOrCreateFileCombos completed in {sw.ElapsedMilliseconds}ms for {zones.Count} zones ({sw.ElapsedMilliseconds / (double)zones.Count:F1}ms per zone)");

            return result;
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
                        MepElementTypeName, MepElementFamilyName,
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
                        IsResolvedFlag, IsClusterResolvedFlag, IsCombinedResolved, IsClusteredFlag,
                        MarkedForClusterProcess, AfterClusterSleeveId,
                        HasDamperNearbyFlag, IsCurrentClashFlag, ReadyForPlacementFlag,
                        StructuralThickness, WallThickness, FramingThickness,
                        MepParameterValuesJson, HostParameterValuesJson,
                        HasMepConnector, DamperConnectorSide,
                        IsInsulated, InsulationThickness,
                        MepSystemType, MepServiceType, ElevationFromLevel
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
                        @MepElementTypeName, @MepElementFamilyName,
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
                        @IsResolvedFlag, @IsClusterResolvedFlag, @IsCombinedResolved, @IsClusteredFlag,
                        @MarkedForClusterProcess, @AfterClusterSleeveId,
                        @HasDamperNearbyFlag, @IsCurrentClashFlag, @ReadyForPlacementFlag,
                        @StructuralThickness, @WallThickness, @FramingThickness,
                        @MepParameterValuesJson, @HostParameterValuesJson,
                        @HasMepConnector, @DamperConnectorSide,
                        @IsInsulated, @InsulationThickness,
                        @MepSystemType, @MepServiceType, @ElevationFromLevel
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
                        MepElementTypeName = @MepElementTypeName,
                        MepElementFamilyName = @MepElementFamilyName,
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
                        IsCombinedResolved = @IsCombinedResolved,
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
                        MepSystemType = @MepSystemType,
                        MepServiceType = @MepServiceType,
                        ElevationFromLevel = @ElevationFromLevel,
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
            if (clashZone.IsClusterResolvedFlag && clashZone.ClusterSleeveInstanceId > 0)
                sleeveState = 2; // ClusterPlaced
            else if (clashZone.IsResolvedFlag && clashZone.SleeveInstanceId > 0)
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
            var markedForCluster = clashZone.MarkedForClusterProcess.HasValue
                ? (object)(clashZone.MarkedForClusterProcess.Value ? 1 : 0)
                : DBNull.Value;
            var isClustered = clashZone.MarkedForClusterProcess.HasValue
                ? (object)(clashZone.MarkedForClusterProcess.Value ? 1 : 0)
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

            // ✅ DIAGNOSTIC: Log WallCenterlinePoint values being saved (ALWAYS log for dampers, even in deployment mode)
            if (string.Equals(clashZone.MepElementCategory, "Duct Accessories", StringComparison.OrdinalIgnoreCase))
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
            // ✅ COMBINED SLEEVES: Add IsCombinedResolved flag parameter
            cmd.Parameters.AddWithValue("@IsCombinedResolved", clashZone.IsCombinedResolved ? 1 : 0);
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
            // ✅ MEP METADATA: Add pre-calculated type and family names for persistence
            cmd.Parameters.AddWithValue("@MepElementTypeName", (object)clashZone.MepElementTypeName ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@MepElementFamilyName", (object)clashZone.MepElementFamilyName ?? DBNull.Value);
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
                cmd.Parameters.AddWithValue("@IsResolvedFlag", clashZone.IsResolvedFlag ? 1 : 0);
                cmd.Parameters.AddWithValue("@IsClusterResolvedFlag", clashZone.IsClusterResolvedFlag ? 1 : 0);
            }

            cmd.Parameters.AddWithValue("@IsClusteredFlag", isClustered);
            cmd.Parameters.AddWithValue("@MarkedForClusterProcess", markedForCluster);
            cmd.Parameters.AddWithValue("@AfterClusterSleeveId", clashZone.AfterClusterSleevePlacedSleeveInstanceId);
            cmd.Parameters.AddWithValue("@HasDamperNearbyFlag", clashZone.HasDamperNearby ? 1 : 0);
            cmd.Parameters.AddWithValue("@IsCurrentClashFlag", clashZone.IsCurrentClashFlag ? 1 : 0);
            cmd.Parameters.AddWithValue("@ReadyForPlacementFlag", clashZone.ReadyForPlacementFlag ? 1 : 0);

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
                            // ✅ CRITICAL FIX: Allow empty values for essential parameters to ensure they are saved
                            var k = kv.Key;
                            if (k.Equals("System Type", StringComparison.OrdinalIgnoreCase) ||
                                k.Equals("Service Type", StringComparison.OrdinalIgnoreCase) ||
                                k.Equals("System Name", StringComparison.OrdinalIgnoreCase) ||
                                k.Equals("System Abbreviation", StringComparison.OrdinalIgnoreCase) ||
                                k.Equals("Level", StringComparison.OrdinalIgnoreCase) ||
                                k.Equals("Reference Level", StringComparison.OrdinalIgnoreCase) ||
                                k.Equals("Schedule Level", StringComparison.OrdinalIgnoreCase) ||
                                k.Equals("Schedule of Level", StringComparison.OrdinalIgnoreCase))
                            {
                                // Allow empty value - proceed to return true
                            }
                            else
                            {
                                filteredOut.Add($"{kv.Key} (empty value)");
                                return false;
                            }
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

            // ✅ MEP METADATA: Add system type, service type, and elevation from level
            cmd.Parameters.AddWithValue("@MepSystemType", (object)clashZone.MepSystemType ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@MepServiceType", (object)clashZone.MepServiceType ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@ElevationFromLevel", clashZone.ElevationFromLevel);
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
                        ClashZoneGuid = clashZoneGuidString ?? string.Empty, // ✅ DETERMINISTIC: Save ClashZoneGuid from zone.Id
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
        private int? GetExistingSnapshotId(string clashZoneGuid, int? sleeveInstanceId, int? clusterInstanceId, SQLiteTransaction transaction)
        {
            using (var cmd = _context.Connection.CreateCommand())
            {
                cmd.Transaction = transaction;

                // ✅ PRIORITY 1: Check by ClashZoneGuid (deterministic GUID) - most reliable
                if (!string.IsNullOrWhiteSpace(clashZoneGuid))
                {
                    // ✅ CRITICAL FIX: Normalize GUID to uppercase for consistent comparison
                    var normalizedGuid = clashZoneGuid != null ? clashZoneGuid.ToUpperInvariant().Trim() : string.Empty;
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
            cmd.Parameters.AddWithValue("@SleeveInstanceId", snapshot.SleeveInstanceId != null ? (object)snapshot.SleeveInstanceId : DBNull.Value);
            cmd.Parameters.AddWithValue("@ClusterInstanceId", snapshot.ClusterInstanceId != null ? (object)snapshot.ClusterInstanceId : DBNull.Value);
            cmd.Parameters.AddWithValue("@SourceType", snapshot.SourceType ?? "Individual");
            cmd.Parameters.AddWithValue("@FilterId", snapshot.FilterId != null ? (object)snapshot.FilterId : DBNull.Value);
            cmd.Parameters.AddWithValue("@ComboId", snapshot.ComboId != null ? (object)snapshot.ComboId : DBNull.Value);
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

                    // ✅ RESTORED: Add "MEP Size" alias to snapshot for redundancy
                    // Although individual sleeves read from Revit, strict aggregators might look for this key.
                    if (!string.IsNullOrWhiteSpace(zone.MepElementSizeParameterValue))
                    {
                        // Check if "MEP Size" is already in bag to avoid duplicates
                        bool hasMepSize = bag.Any(kv => kv != null &&
                           string.Equals(kv.Key?.Trim(), "MEP Size", StringComparison.OrdinalIgnoreCase));

                        if (!hasMepSize)
                        {
                            bag.Add(new Models.SerializableKeyValue { Key = "MEP Size", Value = zone.MepElementSizeParameterValue });
                        }
                    }
                }

                if (bag == null) continue;

                foreach (var kv in bag)
                {
                    if (kv == null) continue;
                    var key = kv.Key?.Trim();
                    var value = kv.Value?.Trim();

                    // ✅ CRITICAL FIX: Allow empty values for essential parameters
                    // This ensures keys like "System Name" are preserved even if empty,
                    // allowing downstream logic to handle them correctly (e.g. valid empty string vs missing key)
                    bool keepEmpty = false;
                    if (!string.IsNullOrEmpty(key))
                    {
                        if (key.Equals("System Type", StringComparison.OrdinalIgnoreCase) ||
                            key.Equals("Service Type", StringComparison.OrdinalIgnoreCase) ||
                            key.Equals("System Name", StringComparison.OrdinalIgnoreCase) ||
                            key.Equals("System Abbreviation", StringComparison.OrdinalIgnoreCase) ||
                            key.Equals("Level", StringComparison.OrdinalIgnoreCase) ||
                            key.Equals("Reference Level", StringComparison.OrdinalIgnoreCase) ||
                            key.Equals("Schedule Level", StringComparison.OrdinalIgnoreCase) ||
                            key.Equals("Schedule of Level", StringComparison.OrdinalIgnoreCase))
                        {
                            keepEmpty = true;
                        }
                    }

                    if (string.IsNullOrEmpty(key) || (string.IsNullOrEmpty(value) && !keepEmpty))
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
            // ✅ NORMALIZE: Remove category suffixes from filter name before querying
            // This ensures "Plumbing_pipes" and "Plumbing" both resolve to "Plumbing"
            filterName = FilterNameHelper.NormalizeBaseName(filterName, filterName, category);

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

                // ✅ CRITICAL: ALWAYS exclude zones already in combined sleeves
                // This prevents placing individual/cluster sleeves on top of combined sleeves
                whereConditions.Add("cz.IsCombinedResolved = 0");

                if (unresolvedOnly)
                    whereConditions.Add("(cz.IsResolvedFlag = 0 AND cz.IsClusterResolvedFlag = 0)");

                if (readyForPlacementOnly)
                    whereConditions.Add("cz.ReadyForPlacementFlag = 1");

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
                    while (reader.Read())
                    {
                        var clashZone = MapClashZone(reader);
                        SetMetadataFromReader(clashZone, reader);

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
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Returns all ClashZones matching the given filter names and categories.
        /// </summary>
        public List<ClashZone> GetClashZonesByFilterAndCategory(List<string> filterNames, List<string> categories)
        {
            if (filterNames == null || filterNames.Count == 0 || categories == null || categories.Count == 0)
                return new List<ClashZone>();

            var result = new List<ClashZone>();
            try
            {
                foreach (var filterName in filterNames)
                {
                    foreach (var category in categories)
                    {
                        var zones = GetClashZonesByFilter(filterName, category, unresolvedOnly: false);
                        if (zones != null) result.AddRange(zones);
                    }
                }
            }
            catch (Exception ex)
            {
                _logger($"[SQLite] ❌ GetClashZonesByFilterAndCategory error: {ex.Message}");
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

        public void ResetIsCurrentClashFlag()
        {
            try
            {
                using (var cmd = _context.Connection.CreateCommand())
                {
                    cmd.CommandText = "UPDATE ClashZones SET IsCurrentClashFlag = 0 WHERE IsCurrentClashFlag = 1";
                    int rows = cmd.ExecuteNonQuery();
                    if (!DeploymentConfiguration.DeploymentMode && rows > 0)
                        _logger($"[SQLite] 🔄 ResetIsCurrentClashFlag: Cleared flag for {rows} zones.");
                }
            }
            catch (Exception ex)
            {
                _logger($"[SQLite] ❌ Error in ResetIsCurrentClashFlag: {ex.Message}");
            }
        }

        public void BulkSetIsCurrentClashFlag(List<Guid> guids, bool value)
        {
            if (guids == null || guids.Count == 0) return;

            var unique = guids.Distinct().ToList();
            const int batchSize = 100;
            int totalUpdated = 0;

            for (int i = 0; i < unique.Count; i += batchSize)
            {
                var batch = unique.Skip(i).Take(batchSize).Select(g => $"UPPER('{g.ToString().ToUpperInvariant()}')");
                using (var cmd = _context.Connection.CreateCommand())
                {
                    cmd.CommandText = $"UPDATE ClashZones SET IsCurrentClashFlag = {(value ? 1 : 0)}, UpdatedAt = CURRENT_TIMESTAMP WHERE UPPER(ClashZoneGuid) IN ({string.Join(",", batch)})";
                    totalUpdated += cmd.ExecuteNonQuery();
                }
            }
            if (!DeploymentConfiguration.DeploymentMode && totalUpdated > 0)
                _logger($"[SQLite] ✅ BulkSetIsCurrentClashFlag: Set {value} for {totalUpdated} zones.");
        }

        // ====================================================================================
        // ✅ 2-STEP FLAG LOGIC (USER REQUESTED)
        // 1. ResetIsCurrentClashFlag: Reset ALL zones to 0 initially.
        // 2. SetIsCurrentClashFlagForSectionBox: Set to 1 based on Filters + SectionBox.
        // 3. SetReadyForPlacementForUnresolvedZonesInSectionBox: Set R4P=1 based on Step 2.
        // ====================================================================================


        public int SetIsCurrentClashFlagForSectionBox(
            List<string> filterNames,
            List<string> categories,
            BoundingBoxXYZ sectionBox)
        {
            if (filterNames == null || filterNames.Count == 0 || categories == null || categories.Count == 0)
                return 0;

            int markedCount = 0;

            try
            {
                // We'll iterate through filters and categories and use GetClashZonesByFilter
                var zonesToCheck = new List<ClashZone>();
                foreach (var filterName in filterNames)
                {
                    foreach (var category in categories)
                    {
                        var zones = GetClashZonesByFilter(filterName, category, unresolvedOnly: false);
                        if (zones != null) zonesToCheck.AddRange(zones);
                    }
                }

                if (zonesToCheck.Count == 0) return 0;

                // Apply Section Box Check
                var zonesToMark = new List<Guid>();
                foreach (var zone in zonesToCheck)
                {
                    // Spatial Check
                    if (sectionBox != null)
                    {
                        var pt = zone.IntersectionPoint;
                        // Reconstruct if needed
                        if (pt == null && (Math.Abs(zone.IntersectionPointX) > 1e-9 || Math.Abs(zone.IntersectionPointY) > 1e-9))
                            pt = new XYZ(zone.IntersectionPointX, zone.IntersectionPointY, zone.IntersectionPointZ);

                        if (pt == null || !JSE_RevitAddin_MEP_OPENINGS.Helpers.SectionBoxHelper.IsPointInBoundingBox(pt, sectionBox))
                            continue;
                    }

                    zonesToMark.Add(zone.Id);
                }

                if (zonesToMark.Count > 0)
                {
                    // Bulk Update to 1
                    BulkSetIsCurrentClashFlag(zonesToMark, true);
                    markedCount = zonesToMark.Count;
                }
            }
            catch (Exception ex)
            {
                _logger($"[SQLite] ❌ Error in SetIsCurrentClashFlagForSectionBox: {ex.Message}");
            }

            return markedCount;
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
                    // ✅ CRITICAL FIX: Handle IsCurrentClashFlag differently based on context:
                    // - When value=TRUE (during REFRESH): Set BOTH flags to 1 (zone is active in session)
                    // - When value=FALSE (after PLACEMENT): ONLY reset ReadyForPlacementFlag, leave IsCurrentClashFlag=1
                    // This ensures IsCurrentClashFlag remains 1 after placement until next refresh cycle
                    if (value)
                    {
                        // REFRESH: Set both flags to 1
                        cmd.CommandText = $"UPDATE ClashZones SET ReadyForPlacementFlag = 1, IsCurrentClashFlag = 1, UpdatedAt = CURRENT_TIMESTAMP WHERE UPPER(ClashZoneGuid) IN ({string.Join(",", batch)}) AND ClashZoneGuid != '' AND ClashZoneGuid IS NOT NULL";
                    }
                    else
                    {
                        // PLACEMENT RESET: Only reset ReadyForPlacementFlag, leave IsCurrentClashFlag unchanged
                        cmd.CommandText = $"UPDATE ClashZones SET ReadyForPlacementFlag = 0, UpdatedAt = CURRENT_TIMESTAMP WHERE UPPER(ClashZoneGuid) IN ({string.Join(",", batch)}) AND ClashZoneGuid != '' AND ClashZoneGuid IS NOT NULL";
                    }
                    var rowsAffected = cmd.ExecuteNonQuery();
                    totalUpdated += rowsAffected;
                }
            }

            if (!DeploymentConfiguration.DeploymentMode && totalUpdated > 0)
            {
                _logger($"[SQLite] ✅ BulkSetReadyForPlacementFlags: Updated {totalUpdated} zones (ReadyForPlacement={value}, IsCurrentClash={(value ? "SET TO 1" : "UNCHANGED")}, batches={(list.Count + batchSize - 1) / batchSize})");
            }
        }

        /// <summary>
        /// ✅ OPTIMIZED: Single-query batch UPDATE for ReadyForPlacementFlag
        /// Uses SQL WHERE clause for section box + unresolved check instead of memory filtering
        /// Expected: 18x faster than legacy path (367ms → 20ms)
        /// </summary>
        private int SetReadyForPlacementBatchOptimized(
            List<string> filterNames,
            List<string> categories,
            BoundingBoxXYZ sectionBox)
        {
            if (filterNames == null || filterNames.Count == 0 || categories == null || categories.Count == 0)
            {
                return 0;
            }

            try
            {
                using (var transaction = _context.Connection.BeginTransaction())
                {
                    using (var cmd = _context.Connection.CreateCommand())
                    {
                        cmd.Transaction = transaction;

                        // Build filter/category conditions
                        var filterConditions = new List<string>();
                        int paramIndex = 0;

                        foreach (var filterName in filterNames)
                        {
                            foreach (var category in categories)
                            {
                                if (string.IsNullOrWhiteSpace(filterName) || string.IsNullOrWhiteSpace(category))
                                    continue;

                                filterConditions.Add($"(f.FilterName = @fn{paramIndex} AND cz.MepCategory = @cat{paramIndex})");
                                cmd.Parameters.AddWithValue($"@fn{paramIndex}", filterName);
                                cmd.Parameters.AddWithValue($"@cat{paramIndex}", category);
                                paramIndex++;
                            }
                        }

                        if (filterConditions.Count == 0)
                        {
                            return 0;
                        }

                        // Build section box clause
                        var sectionBoxClause = "";
                        if (sectionBox != null)
                        {
                            if (OptimizationFlags.UseRTreeDatabaseIndex)
                            {
                                // ✅ R-TREE OPTIMIZATION: Use spatial index join for O(log n) filtering
                                // This is significantly faster for large databases with small section boxes
                                sectionBoxClause = @"
                                    AND cz.ClashZoneId IN (
                                        SELECT id FROM ClashZonesRTree 
                                        WHERE minX <= @MaxX AND maxX >= @MinX
                                          AND minY <= @MaxY AND maxY >= @MinY
                                          AND minZ <= @MaxZ AND maxZ >= @MinZ
                                    )";
                            }
                            else
                            {
                                // B-tree range query (fallback)
                                sectionBoxClause = @"
                                    AND cz.IntersectionPointX >= @MinX AND cz.IntersectionPointX <= @MaxX
                                    AND cz.IntersectionPointY >= @MinY AND cz.IntersectionPointY <= @MaxY
                                    AND cz.IntersectionPointZ >= @MinZ AND cz.IntersectionPointZ <= @MaxZ";
                            }
                            cmd.Parameters.AddWithValue("@MinX", sectionBox.Min.X);
                            cmd.Parameters.AddWithValue("@MaxX", sectionBox.Max.X);
                            cmd.Parameters.AddWithValue("@MinY", sectionBox.Min.Y);
                            cmd.Parameters.AddWithValue("@MaxY", sectionBox.Max.Y);
                            cmd.Parameters.AddWithValue("@MinZ", sectionBox.Min.Z);
                            cmd.Parameters.AddWithValue("@MaxZ", sectionBox.Max.Z);
                        }

                        // ✅ ATOMIC UPDATE: Sets flag=1 for IN-SCOPE zones, flag=0 for OUT-OF-SCOPE zones
                        // This eliminates the need for a separate reset step and avoids timing issues
                        // In-scope = matches filter/category + unresolved + in section box
                        // Out-of-scope = matches filter/category but resolved OR outside section box
                        cmd.CommandText = $@"
                            UPDATE ClashZones
                            SET ReadyForPlacementFlag = CASE 
                                    WHEN ClashZoneId IN (
                                        SELECT cz2.ClashZoneId
                                        FROM ClashZones cz2
                                        INNER JOIN FileCombos ffc2 ON cz2.ComboId = ffc2.ComboId
                                        INNER JOIN Filters f2 ON ffc2.FilterId = f2.FilterId
                                        WHERE 
                                            ({string.Join(" OR ", filterConditions)})
                                            AND cz2.IsResolvedFlag = 0 
                                            AND cz2.IsClusterResolvedFlag = 0
                                            AND cz2.IsCombinedResolved = 0
                                            {sectionBoxClause.Replace("cz.", "cz2.")}
                                    ) THEN 1 ELSE 0 END,
                                IsCurrentClashFlag = CASE 
                                    WHEN ClashZoneId IN (
                                        SELECT cz3.ClashZoneId
                                        FROM ClashZones cz3
                                        INNER JOIN FileCombos ffc3 ON cz3.ComboId = ffc3.ComboId
                                        INNER JOIN Filters f3 ON ffc3.FilterId = f3.FilterId
                                        WHERE 
                                            ({string.Join(" OR ", filterConditions)})
                                            AND cz3.IsResolvedFlag = 0 
                                            AND cz3.IsClusterResolvedFlag = 0
                                            AND cz3.IsCombinedResolved = 0
                                            {sectionBoxClause.Replace("cz.", "cz3.")}
                                    ) THEN 1 ELSE 0 END
                            WHERE ClashZoneId IN (
                                SELECT cz.ClashZoneId
                                FROM ClashZones cz
                                INNER JOIN FileCombos ffc ON cz.ComboId = ffc.ComboId
                                INNER JOIN Filters f ON ffc.FilterId = f.FilterId
                                WHERE ({string.Join(" OR ", filterConditions)})
                            )";

                        int count = cmd.ExecuteNonQuery();
                        transaction.Commit();

                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            _logger($"[SQLite] [BATCH-OPTIMIZED] ✅ Set ReadyForPlacementFlag=1 for {count} zones (single query, filters={filterNames.Count}, categories={categories.Count}, sectionBox={(sectionBox != null ? "Active" : "NULL")})");
                        }

                        return count;
                    }
                }
            }
            catch (Exception ex)
            {
                _logger($"[SQLite] [BATCH-OPTIMIZED] ❌ Failed: {ex.Message}");
                return 0;
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
        /// Update ReadyForPlacementFlag based *solely* on IsCurrentClashFlag (set by SessionContextService).
        /// This method no longer performs spatial or category filtering itself.
        /// </summary>
        public int SetReadyForPlacementForUnresolvedZonesInSectionBox(
            List<string> ignoredData1 = null,
            List<string> ignoredData2 = null,
            BoundingBoxXYZ ignoredData3 = null)
        {
            // ✅ SIMPLIFIED: Pure SQL update based on IsCurrentClashFlag.
            // Logic moved to SessionContextService to adhere to SOLID / User Request.

            int updatedCount = 0;
            try
            {
                using (var cmd = _context.Connection.CreateCommand())
                {
                    cmd.CommandText = @"
                        UPDATE ClashZones 
                        SET ReadyForPlacementFlag = 1, 
                            UpdatedAt = CURRENT_TIMESTAMP
                        WHERE IsCurrentClashFlag = 1 
                          AND IsResolvedFlag = 0 
                          AND IsClusterResolvedFlag = 0 
                          AND IsCombinedResolved = 0";

                    updatedCount = cmd.ExecuteNonQuery();

                    if (!DeploymentConfiguration.DeploymentMode && updatedCount > 0)
                    {
                        _logger($"[SQLite] ✅ SetReadyForPlacement (SQL): Marked {updatedCount} zones based on IsCurrentClashFlag=1.");
                    }
                }
            }
            catch (Exception ex)
            {
                _logger($"[SQLite] ❌ Error in SetReadyForPlacementForUnresolvedZonesInSectionBox: {ex.Message}");
            }
            return updatedCount;
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
            BoundingBoxXYZ sectionBox)
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
            BoundingBoxXYZ sectionBox)
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
            BoundingBoxXYZ sectionBox)
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
        /// <summary>
        /// ✅ BATCH OPTIMIZATION: Update R-tree index for multiple clash zones in one pass.
        /// Strategy: 
        /// 1. Delete existing entries for the IDs (batch DELETE).
        /// 2. Batch INSERT new entries.
        /// </summary>
        public void BulkUpdateRTreeIndex(IEnumerable<ClashZone> zones, SQLiteTransaction? transaction = null)
        {
            if (!Services.OptimizationFlags.UseRTreeDatabaseIndex)
                return;

            var zonesList = zones?.ToList() ?? new List<ClashZone>();
            if (zonesList.Count == 0) return;

            try
            {
                // We need IDs for R-tree. If IDs are missing, we can't update.
                // Note: zonesList must have ClashZoneId populated or we must have another way to get them.
                var zonesWithId = zonesList.Where(z => z.ClashZoneId > 0).ToList();
                if (zonesWithId.Count == 0) return;

                bool ownTransaction = transaction == null;
                var currentTransaction = transaction ?? _context.Connection.BeginTransaction();

                try
                {
                    using (var cmd = _context.Connection.CreateCommand())
                    {
                        cmd.Transaction = currentTransaction;

                        // 1. Batch DELETE
                        var idList = string.Join(",", zonesWithId.Select(z => z.ClashZoneId));
                        cmd.CommandText = $"DELETE FROM ClashZonesRTree WHERE id IN ({idList})";
                        cmd.ExecuteNonQuery();

                        // 2. Batch INSERT
                        const int batchSize = 100;
                        for (int i = 0; i < zonesWithId.Count; i += batchSize)
                        {
                            var batch = zonesWithId.Skip(i).Take(batchSize).ToList();
                            var sql = new System.Text.StringBuilder();
                            sql.Append("INSERT INTO ClashZonesRTree (id, minX, maxX, minY, maxY, minZ, maxZ) VALUES ");

                            for (int j = 0; j < batch.Count; j++)
                            {
                                var zone = batch[j];
                                double minX = 0, maxX = 0, minY = 0, maxY = 0, minZ = 0, maxZ = 0;
                                bool hasValidBBox = false;

                                // Tolerance check (same as UpdateRTreeIndex)
                                bool hasValidSleeveBBox =
                                    zone.SleeveBoundingBoxMinX < zone.SleeveBoundingBoxMaxX &&
                                    zone.SleeveBoundingBoxMinY < zone.SleeveBoundingBoxMaxY &&
                                    zone.SleeveBoundingBoxMinZ < zone.SleeveBoundingBoxMaxZ &&
                                    (zone.SleeveBoundingBoxMinX != 0.0 || zone.SleeveBoundingBoxMaxX != 0.0);

                                if (hasValidSleeveBBox)
                                {
                                    minX = zone.SleeveBoundingBoxMinX;
                                    maxX = zone.SleeveBoundingBoxMaxX;
                                    minY = zone.SleeveBoundingBoxMinY;
                                    maxY = zone.SleeveBoundingBoxMaxY;
                                    minZ = zone.SleeveBoundingBoxMinZ;
                                    maxZ = zone.SleeveBoundingBoxMaxZ;
                                    hasValidBBox = true;
                                }
                                else
                                {
                                    double interX = zone.IntersectionPoint?.X ?? zone.IntersectionPointX;
                                    double interY = zone.IntersectionPoint?.Y ?? zone.IntersectionPointY;
                                    double interZ = zone.IntersectionPoint?.Z ?? zone.IntersectionPointZ;

                                    if (Math.Abs(interX) > 1e-9 || Math.Abs(interY) > 1e-9)
                                    {
                                        double tol = 3.28; // ~1m
                                        minX = interX - tol; maxX = interX + tol;
                                        minY = interY - tol; maxY = interY + tol;
                                        minZ = interZ - tol; maxZ = interZ + tol;
                                        hasValidBBox = true;
                                    }
                                }

                                if (hasValidBBox)
                                {
                                    sql.Append($"({zone.ClashZoneId}, {minX:F6}, {maxX:F6}, {minY:F6}, {maxY:F6}, {minZ:F6}, {maxZ:F6})");
                                    if (j < batch.Count - 1) sql.Append(",");
                                }
                            }

                            // Clean trailing comma if last entries were skipped
                            string finalSql = sql.ToString().TrimEnd(',');
                            if (finalSql.EndsWith("VALUES ")) continue;

                            cmd.CommandText = finalSql;
                            cmd.ExecuteNonQuery();
                        }
                    }

                    if (ownTransaction) currentTransaction.Commit();

                    if (!DeploymentConfiguration.DeploymentMode)
                        _logger($"[SQLite] ✅ Bulk R-tree index updated for {zonesWithId.Count} zones");
                }
                catch (Exception)
                {
                    if (ownTransaction) currentTransaction.Rollback();
                    throw;
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    _logger($"[SQLite] ⚠️ Bulk R-tree update failed: {ex.Message}");
            }
        }

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
                        WHERE f.Category = @Category
                          AND cz.IsCombinedResolved = 0";

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
            if (string.IsNullOrWhiteSpace(json) || json == "{}" || json == "NULL" || json == "[]")
                return new Dictionary<string, string>();

            try
            {
                // ✅ CRITICAL FIX: Handle TWO possible JSON formats:
                // 1. Dictionary format: {"key1":"value1","key2":"value2"} (from AddClashZoneParameters)
                // 2. Array of Key/Value objects: [{"Key":"k1","Value":"v1"},{"Key":"k2","Value":"v2"}] (from BulkInsertClashZones)

                json = json.Trim();

                // Try Dictionary format first (starts with '{')
                if (json.StartsWith("{"))
                {
                    return System.Text.Json.JsonSerializer.Deserialize<Dictionary<string, string>>(json)
                        ?? new Dictionary<string, string>();
                }

                // Try Array of Key/Value objects format (starts with '[')
                if (json.StartsWith("["))
                {
                    var list = System.Text.Json.JsonSerializer.Deserialize<List<Models.SerializableKeyValue>>(json);
                    if (list != null && list.Count > 0)
                    {
                        var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                        foreach (var kv in list)
                        {
                            if (kv != null && !string.IsNullOrEmpty(kv.Key))
                            {
                                dict[kv.Key] = kv.Value ?? string.Empty;
                            }
                        }
                        return dict;
                    }
                }

                return new Dictionary<string, string>();
            }
            catch (Exception ex)
            {
                // Log the error if in debug mode
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    SafeFileLogger.SafeAppendText("cluster_debug.log",
                        $"[{DateTime.Now:HH:mm:ss}] ❌ DeserializeDictionary failed: {ex.Message}, JSON prefix: {json?.Substring(0, Math.Min(50, json?.Length ?? 0))}\n");
                }
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
                clashZone.ClashZoneGuid = guidString; // ✅ CRITICAL FIX: Also set ClashZoneGuid property for clustering
            }

            // ✅ R-TREE SYNC: Load the database auto-increment ID
            clashZone.ClashZoneId = GetInt(reader, "ClashZoneId", -1);
            
            // ✅ CLUSTER SLEEVES: Load ComboId for legacy ClusterSleeves table population
            clashZone.ComboId = GetInt(reader, "ComboId", -1);

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
            
            // ✅ NEW: MEP System Type, Service Type, and Elevation from Level
            clashZone.MepSystemType = GetNullableString(reader, "MepSystemType") ?? string.Empty;
            clashZone.MepServiceType = GetNullableString(reader, "MepServiceType") ?? string.Empty;
            clashZone.ElevationFromLevel = GetDouble(reader, "ElevationFromLevel", 0.0);

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
                        ((string)wallDirectionType).IndexOf("X-WALL", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        clashZone.WallDirection = new XYZ(1, 0, 0); // X-WALL: direction along X-axis
                    }
                    else if (string.Equals(hostOrientation, "Y", StringComparison.OrdinalIgnoreCase) ||
                             ((string)wallDirectionType).IndexOf("Y-WALL", StringComparison.OrdinalIgnoreCase) >= 0)
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
                // Reduced logging noise
                // SafeFileLogger.SafeAppendText("damper_placement_trace.log", $"[{DateTime.Now:HH:mm:ss.fff}] [DB-LOAD-MEP] Zone {clashZone.Id}: MepWidth={clashZone.MepElementWidth:F6}ft ({clashZone.MepElementWidth * 304.8:F1}mm), MepHeight={clashZone.MepElementHeight:F6}ft ({clashZone.MepElementHeight * 304.8:F1}mm)\n");
            }

            clashZone.SleeveFamilyName = GetNullableString(reader, "SleeveFamilyName") ?? string.Empty;
            clashZone.SourceDocKey = GetNullableString(reader, "SourceDocKey") ?? string.Empty;
            clashZone.HostDocKey = GetNullableString(reader, "HostDocKey") ?? string.Empty;
            clashZone.MepElementUniqueId = GetNullableString(reader, "MepElementUniqueId") ?? string.Empty;

            clashZone.IsResolved = GetBool(reader, "IsResolvedFlag");
            clashZone.IsClusterResolved = GetBool(reader, "IsClusterResolvedFlag");
            clashZone.IsCombinedResolved = GetBool(reader, "IsCombinedResolved");
            clashZone.MarkedForClusterProcess = GetNullableBool(reader, "MarkedForClusterProcess");
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
            double rotationAngleRad, string sleeveFamilyName = null, bool markedForClusterProcess = true)
        {
            using (var command = _context.Connection.CreateCommand())
            {
                command.CommandText = @"
                    UPDATE ClashZones 
                    SET SleeveInstanceId = @SleeveInstanceId,
                        SleeveWidth = @Width,
                        SleeveHeight = @Height,
                        SleeveDiameter = @Diameter,
                        IntersectionX = @PlacementX,
                        IntersectionY = @PlacementY,
                        IntersectionZ = @PlacementZ,
                        WallCenterlinePointX = @PlacementActiveX,
                        WallCenterlinePointY = @PlacementActiveY,
                        WallCenterlinePointZ = @PlacementActiveZ,
                        MepElementRotationAngle = @RotationAngleRad,
                        SleeveFamilyName = @SleeveFamilyName,
                        MarkedForClusterProcess = @MarkedForClusterProcess,
                        UpdatedAt = CURRENT_TIMESTAMP
                    WHERE ClashZoneGuid = @ClashZoneGuid";

                command.Parameters.AddWithValue("@SleeveInstanceId", sleeveInstanceId);
                command.Parameters.AddWithValue("@Width", width);
                command.Parameters.AddWithValue("@Height", height);
                command.Parameters.AddWithValue("@Diameter", diameter);
                command.Parameters.AddWithValue("@PlacementX", placementX);
                command.Parameters.AddWithValue("@PlacementY", placementY);
                command.Parameters.AddWithValue("@PlacementZ", placementZ);
                command.Parameters.AddWithValue("@PlacementActiveX", placementActiveX);
                command.Parameters.AddWithValue("@PlacementActiveY", placementActiveY);
                command.Parameters.AddWithValue("@PlacementActiveZ", placementActiveZ);
                command.Parameters.AddWithValue("@RotationAngleRad", rotationAngleRad);
                command.Parameters.AddWithValue("@SleeveFamilyName", (object)sleeveFamilyName ?? DBNull.Value);
                command.Parameters.AddWithValue("@MarkedForClusterProcess", markedForClusterProcess); // ✅ FIX: Allow resetting this flag
                command.Parameters.AddWithValue("@ClashZoneGuid", clashZoneGuid.ToString());

                command.ExecuteNonQuery();
            }
        }

        public void UpdateClusterPlacement(System.Guid clashZoneId, int clusterInstanceId, double minX, double minY, double minZ,
            double maxX, double maxY, double maxZ, double? placementX = null, double? placementY = null, double? placementZ = null,
            double? rotatedMinX = null, double? rotatedMinY = null, double? rotatedMinZ = null,
            double? rotatedMaxX = null, double? rotatedMaxY = null, double? rotatedMaxZ = null,
            bool? isClustered = null, bool? markedForCluster = null, string? sleeveFamilyName = null,
            double sleeveWidth = 0, double sleeveHeight = 0, double sleeveDiameter = 0)
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

                // ✅ CLUSTER DIMENSIONS: Save cluster dimensions to ClashZones table
                if (sleeveWidth > 0)
                {
                    updateFields.Add("SleeveWidth = @SleeveWidth");
                    cmd.Parameters.AddWithValue("@SleeveWidth", sleeveWidth);
                }
                if (sleeveHeight > 0)
                {
                    updateFields.Add("SleeveHeight = @SleeveHeight");
                    cmd.Parameters.AddWithValue("@SleeveHeight", sleeveHeight);
                }
                if (sleeveDiameter > 0)
                {
                    updateFields.Add("SleeveDiameter = @SleeveDiameter");
                    cmd.Parameters.AddWithValue("@SleeveDiameter", sleeveDiameter);
                }

                // ✅ PERSISTENCE FIX: Save Family Name for validation in ClashZones table
                if (!string.IsNullOrEmpty(sleeveFamilyName))
                {
                    updateFields.Add("SleeveFamilyName = @SleeveFamilyName");
                    cmd.Parameters.AddWithValue("@SleeveFamilyName", sleeveFamilyName);
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

                    cmd.Parameters.AddWithValue("@RotatedBoundingBoxMinX", rotatedMinX.Value);
                    cmd.Parameters.AddWithValue("@RotatedBoundingBoxMinY", rotatedMinY.Value);
                    cmd.Parameters.AddWithValue("@RotatedBoundingBoxMinZ", rotatedMinZ.Value);
                    cmd.Parameters.AddWithValue("@RotatedBoundingBoxMaxX", rotatedMaxX.Value);
                    cmd.Parameters.AddWithValue("@RotatedBoundingBoxMaxY", rotatedMaxY.Value);
                    cmd.Parameters.AddWithValue("@RotatedBoundingBoxMaxZ", rotatedMaxZ.Value);
                }

                if (isClustered.HasValue)
                {
                    updateFields.Add("IsClusteredFlag = @IsClusteredFlag");
                    cmd.Parameters.AddWithValue("@IsClusteredFlag", isClustered.Value ? 1 : 0);
                }

                // ✅ CRITICAL FIX: Always resolve the zone when cluster ID is provided
                updateFields.Add("IsResolvedFlag = CASE WHEN @ClusterInstanceId > 0 THEN 1 ELSE IsResolvedFlag END");
                updateFields.Add("IsClusterResolvedFlag = CASE WHEN @ClusterInstanceId > 0 THEN 1 ELSE IsClusterResolvedFlag END");

                if (markedForCluster.HasValue)
                {
                    updateFields.Add("MarkedForClusterProcess = @MarkedForClusterProcess");
                    cmd.Parameters.AddWithValue("@MarkedForClusterProcess", markedForCluster.Value ? 1 : 0);
                }

                updateFields.Add("UpdatedAt = CURRENT_TIMESTAMP");

                cmd.CommandText = $@"
                    UPDATE ClashZones SET
                        {string.Join(",\n                        ", updateFields)}
                    WHERE UPPER(ClashZoneGuid) = UPPER(@ClashZoneId)
                      AND ClashZoneGuid != '' AND ClashZoneGuid IS NOT NULL";

                cmd.Parameters.AddWithValue("@ClashZoneId", clashZoneId.ToString()); // ✅ Pass Guid as string
                cmd.Parameters.AddWithValue("@ClusterInstanceId", clusterInstanceId);
                cmd.Parameters.AddWithValue("@BoundingBoxMinX", minX);
                cmd.Parameters.AddWithValue("@BoundingBoxMinY", minY);
                cmd.Parameters.AddWithValue("@BoundingBoxMinZ", minZ);
                cmd.Parameters.AddWithValue("@BoundingBoxMaxX", maxX);
                cmd.Parameters.AddWithValue("@BoundingBoxMaxY", maxY);
                cmd.Parameters.AddWithValue("@BoundingBoxMaxZ", maxZ);

                // ✅ Add optional parameters
                /*
                if (placementX.HasValue && placementY.HasValue && placementZ.HasValue)
                {
                    cmd.Parameters.AddWithValue("@SleevePlacementX", placementX.Value);
                    cmd.Parameters.AddWithValue("@SleevePlacementY", placementY.Value);
                    cmd.Parameters.AddWithValue("@SleevePlacementZ", placementZ.Value);
                }
                */

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

        public void UpdateMepCategory(Guid clashZoneGuid, string category)
        {
            using (var cmd = _context.Connection.CreateCommand())
            {
                cmd.CommandText = "UPDATE ClashZones SET MepCategory = @Category, UpdatedAt = CURRENT_TIMESTAMP WHERE UPPER(ClashZoneGuid) = UPPER(@ClashZoneGuid)";
                cmd.Parameters.AddWithValue("@Category", category);
                cmd.Parameters.AddWithValue("@ClashZoneGuid", clashZoneGuid.ToString());
                cmd.ExecuteNonQuery();
            }
        }

        public void UpdateSleeveFamilyName(Guid clashZoneGuid, string familyName)
        {
            using (var cmd = _context.Connection.CreateCommand())
            {
                cmd.CommandText = "UPDATE ClashZones SET SleeveFamilyName = @FamilyName, UpdatedAt = CURRENT_TIMESTAMP WHERE UPPER(ClashZoneGuid) = UPPER(@ClashZoneGuid)";
                cmd.Parameters.AddWithValue("@FamilyName", familyName ?? string.Empty);
                cmd.Parameters.AddWithValue("@ClashZoneGuid", clashZoneGuid.ToString());
                cmd.ExecuteNonQuery();
            }
        }

        public void UpdateSleeveFamilyNameBulk(IEnumerable<Guid> clashZoneGuids, string familyName)
        {
            using (var transaction = _context.Connection.BeginTransaction())
            {
                using (var cmd = _context.Connection.CreateCommand())
                {
                    cmd.Transaction = transaction;
                    cmd.CommandText = "UPDATE ClashZones SET SleeveFamilyName = @FamilyName, UpdatedAt = CURRENT_TIMESTAMP WHERE UPPER(ClashZoneGuid) = UPPER(@ClashZoneGuid)";

                    cmd.Parameters.Add("@FamilyName", System.Data.DbType.String);
                    cmd.Parameters.Add("@ClashZoneGuid", System.Data.DbType.String);

                    cmd.Parameters["@FamilyName"].Value = familyName ?? string.Empty;

                    foreach (var guid in clashZoneGuids)
                    {
                        cmd.Parameters["@ClashZoneGuid"].Value = guid.ToString();
                        cmd.ExecuteNonQuery();
                    }
                }
                transaction.Commit();
            }
        }

        /// <summary>
        /// ✅ CLUSTER SLEEVE CORNERS: Update pre-calculated 4 corner coordinates for Cluster Sleeves (Phase 3)
        /// Calculated after cluster placement, stored for downstream processes.
        /// </summary>
        public void UpdateClusterSleeveCorners(int clusterInstanceId,
            double corner1X, double corner1Y, double corner1Z,
            double corner2X, double corner2Y, double corner2Z,
            double corner3X, double corner3Y, double corner3Z,
            double corner4X, double corner4Y, double corner4Z)
        {
            try
            {
                using (var transaction = _context.Connection.BeginTransaction())
                {
                    using (var cmd = _context.Connection.CreateCommand())
                    {
                        cmd.Transaction = transaction;
                        cmd.CommandText = @"
                            UPDATE ClusterSleeves SET
                                Corner1X = @Corner1X,
                                Corner1Y = @Corner1Y,
                                Corner1Z = @Corner1Z,
                                Corner2X = @Corner2X,
                                Corner2Y = @Corner2Y,
                                Corner2Z = @Corner2Z,
                                Corner3X = @Corner3X,
                                Corner3Y = @Corner3Y,
                                Corner3Z = @Corner3Z,
                                Corner4X = @Corner4X,
                                Corner4Y = @Corner4Y,
                                Corner4Z = @Corner4Z,
                                UpdatedAt = CURRENT_TIMESTAMP
                            WHERE ClusterInstanceId = @ClusterInstanceId";

                        cmd.Parameters.AddWithValue("@ClusterInstanceId", clusterInstanceId);
                        cmd.Parameters.AddWithValue("@Corner1X", corner1X);
                        cmd.Parameters.AddWithValue("@Corner1Y", corner1Y);
                        cmd.Parameters.AddWithValue("@Corner1Z", corner1Z);
                        cmd.Parameters.AddWithValue("@Corner2X", corner2X);
                        cmd.Parameters.AddWithValue("@Corner2Y", corner2Y);
                        cmd.Parameters.AddWithValue("@Corner2Z", corner2Z);
                        cmd.Parameters.AddWithValue("@Corner3X", corner3X);
                        cmd.Parameters.AddWithValue("@Corner3Y", corner3Y);
                        cmd.Parameters.AddWithValue("@Corner3Z", corner3Z);
                        cmd.Parameters.AddWithValue("@Corner4X", corner4X);
                        cmd.Parameters.AddWithValue("@Corner4Y", corner4Y);
                        cmd.Parameters.AddWithValue("@Corner4Z", corner4Z);

                        var rowsAffected = cmd.ExecuteNonQuery();

                        // ✅ DIAGNOSTIC LOGGING (User Request)
                        var logParams = new Dictionary<string, object>
                        {
                            { "ClusterInstanceId", clusterInstanceId },
                            { "Corner1", $"({corner1X:F2},{corner1Y:F2},{corner1Z:F2})" },
                            { "RowsAffected", rowsAffected }
                        };

                        if (rowsAffected == 0)
                        {
                            DatabaseOperationLogger.LogOperation("UPDATE", "ClusterSleeves", logParams, rowsAffected,
                                $"⚠️ Failed to update corners for Cluster {clusterInstanceId} (Row not found?)");
                            _logger($"[SQLite] ⚠️ UpdateClusterSleeveCorners: No rows updated for ClusterInstanceId {clusterInstanceId}");
                        }
                        else
                        {
                            DatabaseOperationLogger.LogOperation("UPDATE", "ClusterSleeves", logParams, rowsAffected,
                                $"✅ Saved corners for Cluster {clusterInstanceId}");
                            // Only log to file if not in strict deployment mode, or if crucial
                            if (!DeploymentConfiguration.DeploymentMode)
                                _logger($"[SQLite] ✅ Saved 4 corners for ClusterInstanceId {clusterInstanceId}");
                        }
                    }
                    transaction.Commit();
                }
            }
            catch (Exception ex)
            {
                _logger($"[SQLite] ❌ Error updating corners for ClusterSleeve {clusterInstanceId}: {ex.Message}");
                DatabaseOperationLogger.LogOperation("UPDATE", "ClusterSleeves", null, 0, $"❌ Error: {ex.Message}");
            }
        }

        /// <summary>
        /// ✅ PLACEMENT OPTIMIZATION: Batch update sleeve placement data in a single transaction.
        /// Replaces multiple UpdateSleevePlacement calls with one batch operation (50x faster).
        /// </summary>
        public void BatchUpdateSleevePlacement(
            IEnumerable<(Guid ClashZoneGuid, int SleeveInstanceId, double Width, double Height, double Diameter,
                double PlacementX, double PlacementY, double PlacementZ,
                double PlacementActiveX, double PlacementActiveY, double PlacementActiveZ,
                double RotationAngleRad, string SleeveFamilyName)> updates)
        {
            var updateList = updates?.ToList();
            if (updateList == null || updateList.Count == 0) return;

            var sw = System.Diagnostics.Stopwatch.StartNew();
            int totalRows = 0;

            try
            {
                using (var transaction = _context.Connection.BeginTransaction())
                {
                    using (var cmd = _context.Connection.CreateCommand())
                    {
                        cmd.Transaction = transaction;
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
                                SleeveFamilyName = @SleeveFamilyName,
                                UpdatedAt = CURRENT_TIMESTAMP
                            WHERE UPPER(ClashZoneGuid) = UPPER(@ClashZoneGuid)
                               AND ClashZoneGuid != '' AND ClashZoneGuid IS NOT NULL";

                        // ✅ PREPARED STATEMENT: Create parameters once, reuse for all updates
                        var pGuid = cmd.Parameters.Add("@ClashZoneGuid", System.Data.DbType.String);
                        var pSleeveId = cmd.Parameters.Add("@SleeveInstanceId", System.Data.DbType.Int32);
                        var pWidth = cmd.Parameters.Add("@SleeveWidth", System.Data.DbType.Double);
                        var pHeight = cmd.Parameters.Add("@SleeveHeight", System.Data.DbType.Double);
                        var pDiameter = cmd.Parameters.Add("@SleeveDiameter", System.Data.DbType.Double);
                        var pX = cmd.Parameters.Add("@SleevePlacementX", System.Data.DbType.Double);
                        var pY = cmd.Parameters.Add("@SleevePlacementY", System.Data.DbType.Double);
                        var pZ = cmd.Parameters.Add("@SleevePlacementZ", System.Data.DbType.Double);
                        var pActiveX = cmd.Parameters.Add("@SleevePlacementActiveX", System.Data.DbType.Double);
                        var pActiveY = cmd.Parameters.Add("@SleevePlacementActiveY", System.Data.DbType.Double);
                        var pActiveZ = cmd.Parameters.Add("@SleevePlacementActiveZ", System.Data.DbType.Double);
                        var pRotRad = cmd.Parameters.Add("@MepRotationAngleRad", System.Data.DbType.Double);
                        var pRotDeg = cmd.Parameters.Add("@MepRotationAngleDeg", System.Data.DbType.Double);
                        var pRotCos = cmd.Parameters.Add("@MepRotationCos", System.Data.DbType.Double);
                        var pRotSin = cmd.Parameters.Add("@MepRotationSin", System.Data.DbType.Double);
                        var pFamily = cmd.Parameters.Add("@SleeveFamilyName", System.Data.DbType.String);

                        foreach (var u in updateList)
                        {
                            pGuid.Value = u.ClashZoneGuid.ToString();
                            pSleeveId.Value = u.SleeveInstanceId;
                            pWidth.Value = u.Width;
                            pHeight.Value = u.Height;
                            pDiameter.Value = u.Diameter;
                            pX.Value = u.PlacementX;
                            pY.Value = u.PlacementY;
                            pZ.Value = u.PlacementZ;
                            pActiveX.Value = u.PlacementActiveX;
                            pActiveY.Value = u.PlacementActiveY;
                            pActiveZ.Value = u.PlacementActiveZ;
                            pRotRad.Value = u.RotationAngleRad;
                            pRotDeg.Value = u.RotationAngleRad * 180.0 / Math.PI;
                            pRotCos.Value = Math.Cos(u.RotationAngleRad);
                            pRotSin.Value = Math.Sin(u.RotationAngleRad);
                            pFamily.Value = u.SleeveFamilyName ?? string.Empty;

                            totalRows += cmd.ExecuteNonQuery();
                        }
                    }
                    transaction.Commit();
                }

                if (!DeploymentConfiguration.DeploymentMode)
                {
                    _logger($"[SQLite] ✅ BatchUpdateSleevePlacement: Updated {totalRows} rows in {sw.ElapsedMilliseconds}ms ({updateList.Count} sleeves)");
                }
            }
            catch (Exception ex)
            {
                _logger($"[SQLite] ❌ BatchUpdateSleevePlacement failed: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// ✅ PLACEMENT OPTIMIZATION: Batch update cluster placement data in a single transaction.
        /// </summary>
        public void BatchUpdateClusterPlacement(
            IEnumerable<(System.Guid ClashZoneGuid, int ClusterInstanceId,
                double Width, double Height, double Diameter,
                double BoundingBoxMinX, double BoundingBoxMinY, double BoundingBoxMinZ,
                double BoundingBoxMaxX, double BoundingBoxMaxY, double BoundingBoxMaxZ,
                string SleeveFamilyName)> updates)
        {
            var updateList = updates?.ToList();
            if (updateList == null || updateList.Count == 0) return;

            var sw = System.Diagnostics.Stopwatch.StartNew();
            int totalRows = 0;

            try
            {
                using (var transaction = _context.Connection.BeginTransaction())
                {
                    using (var cmd = _context.Connection.CreateCommand())
                    {
                        cmd.Transaction = transaction;
                        cmd.CommandText = @"
                            UPDATE ClashZones SET
                                SleeveState = 2,
                                ClusterInstanceId = @ClusterInstanceId,
                                SleeveWidth = @SleeveWidth,
                                SleeveHeight = @SleeveHeight,
                                SleeveDiameter = @SleeveDiameter,
                                BoundingBoxMinX = @BoundingBoxMinX,
                                BoundingBoxMinY = @BoundingBoxMinY,
                                BoundingBoxMinZ = @BoundingBoxMinZ,
                                BoundingBoxMaxX = @BoundingBoxMaxX,
                                BoundingBoxMaxY = @BoundingBoxMaxY,
                                BoundingBoxMaxZ = @BoundingBoxMaxZ,
                                SleeveFamilyName = @SleeveFamilyName,
                                IsClustered = 1,
                                IsClusterResolved = 1,
                                MarkedForClusterProcess = 1,
                                UpdatedAt = CURRENT_TIMESTAMP
                            WHERE UPPER(ClashZoneGuid) = UPPER(@ClashZoneGuid)
                               AND ClashZoneGuid != '' AND ClashZoneGuid IS NOT NULL";

                        var pGuid = cmd.Parameters.Add("@ClashZoneGuid", System.Data.DbType.String);
                        var pClusterId = cmd.Parameters.Add("@ClusterInstanceId", System.Data.DbType.Int32);
                        var pWidth = cmd.Parameters.Add("@SleeveWidth", System.Data.DbType.Double);
                        var pHeight = cmd.Parameters.Add("@SleeveHeight", System.Data.DbType.Double);
                        var pDiameter = cmd.Parameters.Add("@SleeveDiameter", System.Data.DbType.Double);
                        var pMinX = cmd.Parameters.Add("@BoundingBoxMinX", System.Data.DbType.Double);
                        var pMinY = cmd.Parameters.Add("@BoundingBoxMinY", System.Data.DbType.Double);
                        var pMinZ = cmd.Parameters.Add("@BoundingBoxMinZ", System.Data.DbType.Double);
                        var pMaxX = cmd.Parameters.Add("@BoundingBoxMaxX", System.Data.DbType.Double);
                        var pMaxY = cmd.Parameters.Add("@BoundingBoxMaxY", System.Data.DbType.Double);
                        var pMaxZ = cmd.Parameters.Add("@BoundingBoxMaxZ", System.Data.DbType.Double);
                        var pFamily = cmd.Parameters.Add("@SleeveFamilyName", System.Data.DbType.String);

                        foreach (var u in updateList)
                        {
                            pGuid.Value = u.ClashZoneGuid.ToString();
                            pClusterId.Value = u.ClusterInstanceId;
                            pWidth.Value = u.Width;
                            pHeight.Value = u.Height;
                            pDiameter.Value = u.Diameter;
                            pMinX.Value = u.BoundingBoxMinX;
                            pMinY.Value = u.BoundingBoxMinY;
                            pMinZ.Value = u.BoundingBoxMinZ;
                            pMaxX.Value = u.BoundingBoxMaxX;
                            pMaxY.Value = u.BoundingBoxMaxY;
                            pMaxZ.Value = u.BoundingBoxMaxZ;
                            pFamily.Value = u.SleeveFamilyName ?? string.Empty;

                            totalRows += cmd.ExecuteNonQuery();
                        }
                    }
                    transaction.Commit();
                }

                sw.Stop();
                DatabaseOperationLogger.LogOperation("UPDATE", "ClashZones (ClusterBatch)", null, totalRows,
                    $"Done. {totalRows} rows in {sw.ElapsedMilliseconds}ms");
            }
            catch (Exception ex)
            {
                _logger($"[SQLite] ❌ Error in BatchUpdateClusterPlacement: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// ✅ PLACEMENT OPTIMIZATION: Batch update sleeve corners in a single transaction.
        /// Replaces multiple UpdateSleeveCorners calls with one batch operation (50x faster).
        /// </summary>
        public void BatchUpdateSleeveCorners(IEnumerable<(Guid Guid, double c1x, double c1y, double c1z, double c2x, double c2y, double c2z, double c3x, double c3y, double c3z, double c4x, double c4y, double c4z)> updates)
        {
            if (updates == null || !updates.Any()) return;

            try
            {
                using (var transaction = _context.Connection.BeginTransaction())
                using (var cmd = _context.Connection.CreateCommand())
                {
                    cmd.Transaction = transaction;
                    cmd.CommandText = @"
                        UPDATE ClashZones SET
                            SleeveCorner1X = @c1x, SleeveCorner1Y = @c1y, SleeveCorner1Z = @c1z,
                            SleeveCorner2X = @c2x, SleeveCorner2Y = @c2y, SleeveCorner2Z = @c2z,
                            SleeveCorner3X = @c3x, SleeveCorner3Y = @c3y, SleeveCorner3Z = @c3z,
                            SleeveCorner4X = @c4x, SleeveCorner4Y = @c4y, SleeveCorner4Z = @c4z,
                            CalculatedAt = CURRENT_TIMESTAMP
                        WHERE ClashZoneGuid = @Guid";

                    var pGuid = cmd.Parameters.Add("@Guid", System.Data.DbType.String);
                    var pC1X = cmd.Parameters.Add("@c1x", System.Data.DbType.Double);
                    var pC1Y = cmd.Parameters.Add("@c1y", System.Data.DbType.Double);
                    var pC1Z = cmd.Parameters.Add("@c1z", System.Data.DbType.Double);
                    var pC2X = cmd.Parameters.Add("@c2x", System.Data.DbType.Double);
                    var pC2Y = cmd.Parameters.Add("@c2y", System.Data.DbType.Double);
                    var pC2Z = cmd.Parameters.Add("@c2z", System.Data.DbType.Double);
                    var pC3X = cmd.Parameters.Add("@c3x", System.Data.DbType.Double);
                    var pC3Y = cmd.Parameters.Add("@c3y", System.Data.DbType.Double);
                    var pC3Z = cmd.Parameters.Add("@c3z", System.Data.DbType.Double);
                    var pC4X = cmd.Parameters.Add("@c4x", System.Data.DbType.Double);
                    var pC4Y = cmd.Parameters.Add("@c4y", System.Data.DbType.Double);
                    var pC4Z = cmd.Parameters.Add("@c4z", System.Data.DbType.Double);

                    foreach (var item in updates)
                    {
                        pGuid.Value = item.Guid.ToString();
                        pC1X.Value = item.c1x; pC1Y.Value = item.c1y; pC1Z.Value = item.c1z;
                        pC2X.Value = item.c2x; pC2Y.Value = item.c2y; pC2Z.Value = item.c2z;
                        pC3X.Value = item.c3x; pC3Y.Value = item.c3y; pC3Z.Value = item.c3z;
                        pC4X.Value = item.c4x; pC4Y.Value = item.c4y; pC4Z.Value = item.c4z;
                        cmd.ExecuteNonQuery();
                    }
                    transaction.Commit();
                }
            }
            catch (Exception ex)
            {
                _logger($"[SQLite] ❌ BatchUpdateSleeveCorners failed: {ex.Message}");
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
                        IsResolvedFlag = CASE WHEN @SleeveInstanceId > 0 THEN 1 ELSE IsResolvedFlag END,
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
        /// ✅ BATCH OPTIMIZATION: Finds existing GUIDs for a collection of MEP+Host+Point triples.
        /// Returns a dictionary mapping (MepId, HostId, PointKey) -> Guid.
        /// </summary>
        public Dictionary<(int MepId, int HostId, string PointKey), Guid> FindGuidsByMepHostAndPointsBulk(
            IEnumerable<(int MepId, int HostId, double X, double Y, double Z)> targets,
            double tolerance = 0.001)
        {
            var results = new Dictionary<(int MepId, int HostId, string PointKey), Guid>();
            var targetsList = targets.ToList();
            if (targetsList.Count == 0) return results;

            try
            {
                using (var transaction = _context.Connection.BeginTransaction())
                {
                    // Create temporary table to store targets
                    using (var cmd = _context.Connection.CreateCommand())
                    {
                        cmd.Transaction = transaction;
                        cmd.CommandText = "CREATE TEMP TABLE IF NOT EXISTS TargetPoints (MepId INTEGER, HostId INTEGER, X REAL, Y REAL, Z REAL)";
                        cmd.ExecuteNonQuery();

                        cmd.CommandText = "DELETE FROM TargetPoints";
                        cmd.ExecuteNonQuery();

                        // Bulk insert targets
                        cmd.CommandText = "INSERT INTO TargetPoints (MepId, HostId, X, Y, Z) VALUES (@MepId, @HostId, @X, @Y, @Z)";
                        var pMep = cmd.Parameters.Add("@MepId", System.Data.DbType.Int32);
                        var pHost = cmd.Parameters.Add("@HostId", System.Data.DbType.Int32);
                        var pX = cmd.Parameters.Add("@X", System.Data.DbType.Double);
                        var pY = cmd.Parameters.Add("@Y", System.Data.DbType.Double);
                        var pZ = cmd.Parameters.Add("@Z", System.Data.DbType.Double);

                        foreach (var target in targetsList)
                        {
                            pMep.Value = target.MepId;
                            pHost.Value = target.HostId;
                            pX.Value = target.X;
                            pY.Value = target.Y;
                            pZ.Value = target.Z;
                            cmd.ExecuteNonQuery();
                        }

                        // Join with ClashZones using tolerance
                        cmd.CommandText = @"
                            SELECT t.MepId, t.HostId, t.X, t.Y, t.Z, cz.ClashZoneGuid
                            FROM TargetPoints t
                            JOIN ClashZones cz ON t.MepId = cz.MepElementId AND t.HostId = cz.HostElementId
                            WHERE ABS(t.X - cz.IntersectionX) < @Tolerance 
                              AND ABS(t.Y - cz.IntersectionY) < @Tolerance 
                              AND ABS(t.Z - cz.IntersectionZ) < @Tolerance
                              AND cz.ClashZoneGuid IS NOT NULL AND cz.ClashZoneGuid != ''";

                        cmd.Parameters.Clear();
                        cmd.Parameters.AddWithValue("@Tolerance", tolerance);

                        using (var reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                int mepId = reader.GetInt32(0);
                                int hostId = reader.GetInt32(1);
                                double x = reader.GetDouble(2);
                                double y = reader.GetDouble(3);
                                double z = reader.GetDouble(4);
                                string guidStr = reader.GetString(reader.GetOrdinal("ClashZoneGuid"));

                                if (Guid.TryParse(guidStr, out var guid))
                                {
                                    string pointKey = $"{Math.Round(x, 4)}_{Math.Round(y, 4)}_{Math.Round(z, 4)}";
                                    var key = (mepId, hostId, pointKey);
                                    if (!results.ContainsKey(key))
                                    {
                                        results[key] = guid;
                                    }
                                }
                            }
                        }
                    }
                    transaction.Commit();
                }
            }
            catch (Exception ex)
            {
                _logger($"[SQLite] ❌ Error in FindGuidsByMepHostAndPointsBulk: {ex.Message}");
            }

            return results;
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
        /// ✅ INTERFACE IMPLEMENTATION: Simplified batch update for interface compliance
        /// Now includes IsClusteredFlag, MarkedForClusterProcess, and AfterClusterSleeveId.
        /// </summary>
        public void BatchUpdateFlags(List<(Guid ClashZoneId, bool IsResolved, bool IsClusterResolved, bool IsCombinedResolved, int SleeveInstanceId, int ClusterInstanceId, bool IsClusteredFlag, bool MarkedForClusterProcess, int AfterClusterSleeveId)> updates)
        {
            if (updates == null || updates.Count == 0)
                return;

            // Convert to full signature with default values for optional parameters
            var fullUpdates = updates.Select(u => (
                ClashZoneId: u.ClashZoneId,
                IsResolved: u.IsResolved,
                IsClusterResolved: u.IsClusterResolved,
                IsCombinedResolved: u.IsCombinedResolved,
                SleeveInstanceId: u.SleeveInstanceId,
                ClusterInstanceId: u.ClusterInstanceId,
                MepElementId: 0,
                StructuralElementId: 0,
                IntersectionPointX: 0.0,
                IntersectionPointY: 0.0,
                IntersectionPointZ: 0.0,
                OldSleeveInstanceId: 0,
                OldClusterInstanceId: 0,
                MarkedForClusterProcess: (bool?)u.MarkedForClusterProcess,
                AfterClusterSleeveId: u.AfterClusterSleeveId,
                IsClusteredFlag: (bool?)u.IsClusteredFlag
            ));

            // Delegate to full implementation
            BatchUpdateFlags(fullUpdates);
        }

        /// <summary>
        /// ✅ DATABASE FLAG MANAGEMENT: Batch update flags for multiple clash zones (full signature)
        /// Updates IsResolvedFlag, IsClusterResolvedFlag, SleeveInstanceId, and ClusterInstanceId
        /// Uses GUID, OLD SleeveInstanceId/ClusterInstanceId, or MEP+Host+Point for matching
        /// </summary>
        public void BatchUpdateFlags(IEnumerable<(Guid ClashZoneId, bool IsResolved, bool IsClusterResolved, bool IsCombinedResolved, int SleeveInstanceId, int ClusterInstanceId, int MepElementId, int StructuralElementId, double IntersectionPointX, double IntersectionPointY, double IntersectionPointZ, int OldSleeveInstanceId, int OldClusterInstanceId, bool? MarkedForClusterProcess, int AfterClusterSleeveId, bool? IsClusteredFlag)> updates)
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
                                    IsCombinedResolved INTEGER,
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
                                    IsClusteredFlag INTEGER,
                                    IsCurrentClashFlag INTEGER
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
                                    MarkedForClusterProcess, AfterClusterSleeveId, IsClusteredFlag,
                                    IsCombinedResolved, IsCurrentClashFlag
                                ) VALUES (
                                    @ClashZoneGuid, @IsResolvedFlag, @IsClusterResolvedFlag, 
                                    @SleeveInstanceId, @ClusterInstanceId,
                                    @OldSleeveInstanceId, @OldClusterInstanceId,
                                    @MepElementId, @HostElementId,
                                    @IntersectionX, @IntersectionY, @IntersectionZ,
                                    @MarkedForClusterProcess, @AfterClusterSleeveId, @IsClusteredFlag,
                                    @IsCombinedResolved, @IsCurrentClashFlag
                                )";

                        var guidParam = insertCmd.Parameters.Add("@ClashZoneGuid", System.Data.DbType.String);
                        var isResolvedParam = insertCmd.Parameters.Add("@IsResolvedFlag", System.Data.DbType.Int32);
                        var isClusterResolvedParam = insertCmd.Parameters.Add("@IsClusterResolvedFlag", System.Data.DbType.Int32);
                        var isCombinedResolvedParam = insertCmd.Parameters.Add("@IsCombinedResolved", System.Data.DbType.Int32);
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
                        var isCurrentClashFlagParam = insertCmd.Parameters.Add("@IsCurrentClashFlag", System.Data.DbType.Int32);

                        insertCmd.Prepare();

                        foreach (var update in updatesList)
                        {
                            guidParam.Value = update.ClashZoneId.ToString().ToUpperInvariant();
                            isResolvedParam.Value = update.IsResolved ? 1 : 0;
                            isClusterResolvedParam.Value = update.IsClusterResolved ? 1 : 0;
                            isCombinedResolvedParam.Value = update.IsCombinedResolved ? 1 : 0;
                            
                            // ✅ USER REQUEST: Preserve -1 for instance IDs (indicates "Processed but deleted/none")
                            // 0 usually means "Unprocessed/Null" in Revit terms, but database uses NULL or -1
                            sleeveIdParam.Value = update.SleeveInstanceId != 0 ? (object)update.SleeveInstanceId : DBNull.Value;
                            clusterIdParam.Value = update.ClusterInstanceId != 0 ? (object)update.ClusterInstanceId : DBNull.Value;
                            oldSleeveIdParam.Value = update.OldSleeveInstanceId != 0 ? (object)update.OldSleeveInstanceId : DBNull.Value;
                            oldClusterIdParam.Value = update.OldClusterInstanceId != 0 ? (object)update.OldClusterInstanceId : DBNull.Value;
                            
                            mepIdParam.Value = update.MepElementId;
                            hostIdParam.Value = update.StructuralElementId;
                            intersectionXParam.Value = update.IntersectionPointX;
                            intersectionYParam.Value = update.IntersectionPointY;
                            intersectionZParam.Value = update.IntersectionPointZ;
                            markedForClusterParam.Value = update.MarkedForClusterProcess.HasValue ? (object)(update.MarkedForClusterProcess.Value ? 1 : 0) : DBNull.Value;
                            
                            // ✅ USER REQUEST: AfterClusterSleeveId should also preserve -1 if provided
                            afterClusterSleeveIdParam.Value = update.AfterClusterSleeveId != 0 ? (object)update.AfterClusterSleeveId : DBNull.Value;
                            isClusteredFlagParam.Value = update.IsClusteredFlag.HasValue ? (object)(update.IsClusteredFlag.Value ? 1 : 0) : DBNull.Value;
                            isCurrentClashFlagParam.Value = DBNull.Value; // Standard BatchUpdateFlags doesn't set this from the input tuple, but we use the column to preserve state

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
                                    IsCombinedResolved = t.IsCombinedResolved,
                                    SleeveInstanceId = t.SleeveInstanceId,
                                    ClusterInstanceId = t.ClusterInstanceId,
                                    MarkedForClusterProcess = t.MarkedForClusterProcess,
                                    AfterClusterSleeveId = t.AfterClusterSleeveId,
                                    IsClusteredFlag = t.IsClusteredFlag,
                                    IsCurrentClashFlag = COALESCE(t.IsCurrentClashFlag, ClashZones.IsCurrentClashFlag), -- Update only if provided, else keep existing
                                    ReadyForPlacementFlag = CASE 
                                        WHEN t.IsResolvedFlag = 0 AND t.IsClusterResolvedFlag = 0 AND t.IsCombinedResolved = 0 THEN 1 
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

                    // ✅ FALLBACK: If bulk update missed rows, try individual updates for reliability
                    if (rowsAffectedTotal < updateCount)
                    {
                        _logger($"[SQLite] ⚠️ Bulk update affected {rowsAffectedTotal} rows but {updateCount} were expected. Attempting {updateCount} individual updates as fallback.");

                        foreach (var update in updatesList)
                        {
                            using (var indUpdateCmd = _context.Connection.CreateCommand())
                            {
                                indUpdateCmd.Transaction = transaction;
                                indUpdateCmd.CommandText = @"
                                        UPDATE ClashZones
                                        SET 
                                            IsResolvedFlag = @IsResolvedFlag,
                                            IsClusterResolvedFlag = @IsClusterResolvedFlag,
                                            IsCombinedResolved = @IsCombinedResolved,
                                            SleeveInstanceId = @SleeveInstanceId,
                                            ClusterInstanceId = @ClusterInstanceId,
                                            MarkedForClusterProcess = @MarkedForClusterProcess,
                                            AfterClusterSleeveId = @AfterClusterSleeveId,
                                            IsClusteredFlag = @IsClusteredFlag,
                                            ReadyForPlacementFlag = CASE 
                                                WHEN @IsResolvedFlag = 0 AND @IsClusterResolvedFlag = 0 AND @IsCombinedResolved = 0 THEN 1 
                                                ELSE ReadyForPlacementFlag 
                                            END,
                                            UpdatedAt = datetime('now', '+5 hours', '+30 minutes')
                                        WHERE UPPER(ClashZoneGuid) = @ClashZoneGuid";

                                indUpdateCmd.Parameters.AddWithValue("@ClashZoneGuid", update.ClashZoneId.ToString().ToUpperInvariant());
                                indUpdateCmd.Parameters.AddWithValue("@IsResolvedFlag", update.IsResolved ? 1 : 0);
                                indUpdateCmd.Parameters.AddWithValue("@IsClusterResolvedFlag", update.IsClusterResolved ? 1 : 0);
                                indUpdateCmd.Parameters.AddWithValue("@IsCombinedResolved", update.IsCombinedResolved ? 1 : 0);
                                indUpdateCmd.Parameters.AddWithValue("@SleeveInstanceId", update.SleeveInstanceId > 0 ? (object)update.SleeveInstanceId : DBNull.Value);
                                indUpdateCmd.Parameters.AddWithValue("@ClusterInstanceId", update.ClusterInstanceId > 0 ? (object)update.ClusterInstanceId : DBNull.Value);
                                indUpdateCmd.Parameters.AddWithValue("@MarkedForClusterProcess", update.MarkedForClusterProcess.HasValue ? (object)(update.MarkedForClusterProcess.Value ? 1 : 0) : DBNull.Value);
                                indUpdateCmd.Parameters.AddWithValue("@AfterClusterSleeveId", update.AfterClusterSleeveId > 0 ? (object)update.AfterClusterSleeveId : DBNull.Value);
                                indUpdateCmd.Parameters.AddWithValue("@IsClusteredFlag", update.IsClusteredFlag.HasValue ? (object)(update.IsClusteredFlag.Value ? 1 : 0) : DBNull.Value);

                                indUpdateCmd.ExecuteNonQuery();
                            }
                        }
                        _logger($"[SQLite] ✅ Fallback individual updates completed.");
                    }

                    transaction.Commit();
                    _logger($"[SQLite] ✅ 🚀 BATCH UPDATE COMPLETED: {updateCount} items processed (with fallback capability).");

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
        /// ✅ BATCH UPDATE: Optimized version that also updates IsCurrentClashFlag
        /// Used by FlagManagerService for placement updates to clear the "Current Clash" status
        /// </summary>


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

        /// <summary>
        /// Force Detection Mode: Reset all flags (IsResolved, IsClusterResolved) to false and clear sleeve IDs
        /// for all zones in the specified filters and categories, while preserving GUIDs.
        /// Used when ForceDetectionMode is enabled in global settings.
        /// </summary>
        public int ResetAllFlagsForForceDetectionMode(List<string> filterNames, List<string> categories)
        {
            if (filterNames == null || filterNames.Count == 0 || categories == null || categories.Count == 0)
                return 0;

            int totalReset = 0;

            try
            {
                using (var transaction = _context.Connection.BeginTransaction())
                {
                    try
                    {
                        foreach (var filterName in filterNames)
                        {
                            foreach (var category in categories)
                            {
                                using (var cmd = _context.Connection.CreateCommand())
                                {
                                    cmd.Transaction = transaction;
                                    cmd.CommandText = @"
                                        UPDATE ClashZones
                                        SET IsResolvedFlag = 0,
                                            IsClusterResolvedFlag = 0,
                                            SleeveInstanceId = NULL,
                                            ClusterInstanceId = NULL,
                                            AfterClusterSleeveId = NULL,
                                            MarkedForClusterProcess = 0
                                        WHERE FilterName = @FilterName
                                          AND MepElementCategory = @Category";

                                    cmd.Parameters.AddWithValue("@FilterName", filterName);
                                    cmd.Parameters.AddWithValue("@Category", category);

                                    int rowsAffected = cmd.ExecuteNonQuery();
                                    totalReset += rowsAffected;
                                }
                            }
                        }

                        transaction.Commit();

                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[ClashZoneRepository] [FORCE-DETECTION] Reset flags for {totalReset} zones (filters={filterNames.Count}, categories={categories.Count})");
                        }

                        return totalReset;
                    }
                    catch (Exception ex)
                    {
                        transaction.Rollback();
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Error($"[ClashZoneRepository] [FORCE-DETECTION] Failed to reset flags: {ex.Message}");
                        }
                        throw;
                    }
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Error($"[ClashZoneRepository] [FORCE-DETECTION] Transaction error: {ex.Message}");
                }
                return 0;
            }
        }


        /// <summary>
        /// Retrieve ClashZone objects associated with the given Revit Sleeve Instance IDs.
        /// This checks both individual SleeveInstanceId and ClusterSleeveInstanceId.
        /// </summary>
        public List<ClashZone> GetClashZonesBySleeveIds(IEnumerable<int> sleeveInstanceIds)
        {
            var ids = sleeveInstanceIds?.Where(id => id > 0).Distinct().ToList();
            if (ids == null || ids.Count == 0) return new List<ClashZone>();

            var zones = new List<ClashZone>();

            try
            {
                var idString = string.Join(",", ids);
                using (var cmd = _context.Connection.CreateCommand())
                {
                    // Query for both Individual IDs and Cluster IDs (ClusterSleeveInstanceId stores the ID of the cluster family instance)
                    // Also check 'AfterClusterSleevePlacedSleeveInstanceId' which is used for single-zone clusters
                    cmd.CommandText = $@"
                        SELECT * 
                        FROM ClashZones 
                        WHERE SleeveInstanceId IN ({idString}) 
                           OR ClusterInstanceId IN ({idString})";

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

                // IMPORTANT: Populate parameters if needed? 
                // Currently CombinedClusterFormationService works on geometry fields (ClusterSleeveBoundingBoxMinX etc)
                // which MapClashZone should populate.
                // We do NOT need full Parameter Maps for the geometry calculation, 
                // but might need them if we want to preserve parameters?
                // For now, let's assume geometry is the priority.
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    _logger($"[SQLite] Error querying by sleeve IDs: {ex.Message}");
                }
            }

            return zones;
        }


        public Dictionary<int, string> GetMepCategoriesForSleeveIds(IEnumerable<int> sleeveInstanceIds)
        {
            var result = new Dictionary<int, string>();
            if (sleeveInstanceIds == null || !sleeveInstanceIds.Any()) return result;

            try
            {
                var uniqueIds = sleeveInstanceIds.Distinct().ToList();
                const int batchSize = 500;

                for (int i = 0; i < uniqueIds.Count; i += batchSize)
                {
                    var batch = uniqueIds.Skip(i).Take(batchSize);
                    var idString = string.Join(",", batch);

                    using (var cmd = _context.Connection.CreateCommand())
                    {
                        cmd.CommandText = $@"
                            SELECT SleeveInstanceId, MepCategory 
                            FROM ClashZones 
                            WHERE SleeveInstanceId IN ({idString})
                              AND MepCategory IS NOT NULL 
                              AND MepCategory != ''";

                        using (var reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                int id = Convert.ToInt32(reader["SleeveInstanceId"]);
                                string category = reader["MepCategory"]?.ToString();

                                if (!result.ContainsKey(id) && !string.IsNullOrEmpty(category))
                                {
                                    result[id] = category;
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger($"[SQLite] ❌ Error getting MEP Categories: {ex.Message}");
            }

            return result;
        }



        /// <summary>
        /// Retrieves cluster sleeves associated with the given instance IDs.
        /// </summary>
        public List<ClusterSleeve> GetClusterSleevesByInstanceIds(IEnumerable<int> instanceIds)
        {
            // SOLID Refactoring: Delegate to specialized repository if enabled
            if (OptimizationFlags.UseSolidRefactoredRepositories)
            {
                // Use ClusterSleeveData from _clusterRepo and map to ClusterSleeve
                var dataList = _clusterRepo.GetClusterSleevesByInstanceIds(instanceIds);
                return dataList?.Select(MapClusterSleeve).ToList() ?? new List<ClusterSleeve>();
            }

            var results = new List<ClusterSleeve>();
            var ids = instanceIds?.ToList();
            if (ids == null || ids.Count == 0) return results;

            try
            {
                var idString = string.Join(",", ids);
                using (var cmd = _context.Connection.CreateCommand())
                {
                    cmd.CommandText = $@"
                        SELECT * 
                        FROM ClusterSleeves 
                        WHERE ClusterInstanceId IN ({idString})";

                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            results.Add(MapClusterSleeve(reader));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger($"[SQLite] ❌ Error in GetClusterSleevesByInstanceIds: {ex.Message}");
            }
            return results;
        }

        /// <summary>
        /// Retrieves all cluster sleeves from the database.
        /// </summary>
        public List<ClusterSleeve> GetAllClusterSleeves()
        {
            // SOLID Refactoring: Delegate to specialized repository if enabled
            if (OptimizationFlags.UseSolidRefactoredRepositories)
            {
                // Use ClusterSleeveData from _clusterRepo and map to ClusterSleeve
                var dataList = _clusterRepo.GetAllClusterSleeves();
                return dataList?.Select(MapClusterSleeve).ToList() ?? new List<ClusterSleeve>();
            }

            var results = new List<ClusterSleeve>();

            try
            {
                using (var cmd = _context.Connection.CreateCommand())
                {
                    cmd.CommandText = "SELECT * FROM ClusterSleeves";

                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            results.Add(MapClusterSleeve(reader));
                        }
                    }
                }

                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[ClashZoneRepository] Retrieved {results.Count} cluster sleeves from database");
                }
            }
            catch (Exception ex)
            {
                _logger($"[SQLite] ❌ Error in GetAllClusterSleeves: {ex.Message}");
            }

            return results;
        }

        private ClusterSleeve MapClusterSleeve(ClusterSleeveData data)
        {
            if (data == null) return null;

            return new ClusterSleeve
            {
                // ClusterSleeveId not present in ClusterSleeveData, defaulting to 0 or -1 appropriate for non-persisted ID
                ClusterSleeveId = 0,
                ClusterInstanceId = data.ClusterInstanceId,
                Category = data.Category,
                // ClusterSleeveData has double, ClusterSleeve likely has double?. Mapping directly.
                Corner1X = data.Corner1X,
                Corner1Y = data.Corner1Y,
                Corner1Z = data.Corner1Z,
                Corner2X = data.Corner2X,
                Corner2Y = data.Corner2Y,
                Corner2Z = data.Corner2Z,
                Corner3X = data.Corner3X,
                Corner3Y = data.Corner3Y,
                Corner3Z = data.Corner3Z,
                Corner4X = data.Corner4X,
                Corner4Y = data.Corner4Y,
                Corner4Z = data.Corner4Z,
                RotationAngleDeg = data.RotationAngleDeg,
                HostType = data.HostType,
                HostOrientation = data.HostOrientation
            };
        }

        private ClusterSleeve MapClusterSleeve(System.Data.SQLite.SQLiteDataReader reader)
        {
            return new ClusterSleeve
            {
                ClusterSleeveId = reader.GetInt32(reader.GetOrdinal("ClusterSleeveId")),
                ClusterInstanceId = reader.GetInt32(reader.GetOrdinal("ClusterInstanceId")),
                Category = reader.IsDBNull(reader.GetOrdinal("Category")) ? string.Empty : reader.GetString(reader.GetOrdinal("Category")),
                Corner1X = reader.IsDBNull(reader.GetOrdinal("Corner1X")) ? (double?)null : reader.GetDouble(reader.GetOrdinal("Corner1X")),
                Corner1Y = reader.IsDBNull(reader.GetOrdinal("Corner1Y")) ? (double?)null : reader.GetDouble(reader.GetOrdinal("Corner1Y")),
                Corner1Z = reader.IsDBNull(reader.GetOrdinal("Corner1Z")) ? (double?)null : reader.GetDouble(reader.GetOrdinal("Corner1Z")),
                Corner2X = reader.IsDBNull(reader.GetOrdinal("Corner2X")) ? (double?)null : reader.GetDouble(reader.GetOrdinal("Corner2X")),
                Corner2Y = reader.IsDBNull(reader.GetOrdinal("Corner2Y")) ? (double?)null : reader.GetDouble(reader.GetOrdinal("Corner2Y")),
                Corner2Z = reader.IsDBNull(reader.GetOrdinal("Corner2Z")) ? (double?)null : reader.GetDouble(reader.GetOrdinal("Corner2Z")),
                Corner3X = reader.IsDBNull(reader.GetOrdinal("Corner3X")) ? (double?)null : reader.GetDouble(reader.GetOrdinal("Corner3X")),
                Corner3Y = reader.IsDBNull(reader.GetOrdinal("Corner3Y")) ? (double?)null : reader.GetDouble(reader.GetOrdinal("Corner3Y")),
                Corner3Z = reader.IsDBNull(reader.GetOrdinal("Corner3Z")) ? (double?)null : reader.GetDouble(reader.GetOrdinal("Corner3Z")),
                Corner4X = reader.IsDBNull(reader.GetOrdinal("Corner4X")) ? (double?)null : reader.GetDouble(reader.GetOrdinal("Corner4X")),
                Corner4Y = reader.IsDBNull(reader.GetOrdinal("Corner4Y")) ? (double?)null : reader.GetDouble(reader.GetOrdinal("Corner4Y")),
                Corner4Z = reader.IsDBNull(reader.GetOrdinal("Corner4Z")) ? (double?)null : reader.GetDouble(reader.GetOrdinal("Corner4Z")),
                RotationAngleDeg = reader.IsDBNull(reader.GetOrdinal("RotationAngleDeg")) ? (double?)null : reader.GetDouble(reader.GetOrdinal("RotationAngleDeg")),

                // Map Host Info
                HostType = reader.IsDBNull(reader.GetOrdinal("HostType")) ? string.Empty : reader.GetString(reader.GetOrdinal("HostType")),
                HostOrientation = reader.IsDBNull(reader.GetOrdinal("HostOrientation")) ? string.Empty : reader.GetString(reader.GetOrdinal("HostOrientation"))
            };
        }
        public (int ComboId, int FilterId) GetComboAndFilterId(Guid clashZoneId)
        {
            using (var cmd = _context.Connection.CreateCommand())
            {
                cmd.CommandText = @"
                    SELECT cz.ComboId, fc.FilterId 
                    FROM ClashZones cz
                    INNER JOIN FileCombos fc ON cz.ComboId = fc.ComboId
                    WHERE cz.ClashZoneId = @Id";

                cmd.Parameters.AddWithValue("@Id", clashZoneId.ToString());

                using (var reader = cmd.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        return (reader.GetInt32(0), reader.GetInt32(1));
                    }
                }
            }
            return (0, 0);
        }

        public List<string> GetDistinctCategories()
        {
            var categories = new List<string>();
            try
            {
                using (var cmd = _context.Connection.CreateCommand())
                {
                    cmd.CommandText = "SELECT DISTINCT MepElementCategory FROM ClashZones WHERE MepElementCategory IS NOT NULL AND MepElementCategory != ''";
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            categories.Add(reader.GetString(0));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    _logger($"[SQLite] ❌ Error in GetDistinctCategories: {ex.Message}");
            }
            return categories;
        }

        public List<ClashZone> GetAllClashZones()
        {
            var list = new List<ClashZone>();
            try
            {
                 using (var cmd = _context.Connection.CreateCommand())
                {
                    cmd.CommandText = "SELECT * FROM ClashZones WHERE ClashZoneGuid IS NOT NULL AND ClashZoneGuid != ''";
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            list.Add(MapClashZone(reader));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                 if (!DeploymentConfiguration.DeploymentMode)
                    _logger($"[SQLite] Error Get All Clash Zones: {ex.Message}");
            }
            return list;
        }

        public void Update(ClashZone zone)
        {
             if (zone == null) return;
             // Use existing bulk update method for single item
             InsertOrUpdateClashZones(new[] { zone }, "SingleUpdate", zone.MepElementCategory);
        }

        /// <summary>
        /// ✅ PLACEMENT OPTIMIZATION: Batch update SleeveInstanceId for multiple zones in a single transaction.
        /// Consolidates individual Updates into a single transaction (50x faster).
        /// </summary>
        public void BatchUpdateSleeveInstanceIds(IEnumerable<(Guid ClashZoneId, int SleeveInstanceId)> updates)
        {
            if (updates == null || !updates.Any()) return;
            var updatesList = updates.ToList();

            if (!DeploymentConfiguration.DeploymentMode)
            {
                DebugLogger.Info($"[ClashZoneRepository] [BATCH-UPDATE] Updating SleeveInstanceId for {updatesList.Count} zones...");
            }

            try
            {
                using (var transaction = _context.Connection.BeginTransaction())
                {
                    // 1. Update ClashZones (Optimized Reusable Command)
                    using (var cmd = _context.Connection.CreateCommand())
                    {
                        cmd.Transaction = transaction;
                        cmd.CommandText = "UPDATE ClashZones SET SleeveInstanceId = @SleeveInstanceId, UpdatedAt = datetime('now', '+5 hours', '+30 minutes') WHERE UPPER(ClashZoneGuid) = UPPER(@ClashZoneGuid)";
                        var pGuid = cmd.Parameters.AddWithValue("@ClashZoneGuid", "");
                        var pId = cmd.Parameters.AddWithValue("@SleeveInstanceId", DBNull.Value);
                        
                        int totalRowsAffected = 0;
                        foreach (var update in updatesList)
                        {
                            pGuid.Value = update.ClashZoneId.ToString();
                            pId.Value = update.SleeveInstanceId > 0 ? (object)update.SleeveInstanceId : DBNull.Value;
                            int rowsAffected = cmd.ExecuteNonQuery();
                            totalRowsAffected += rowsAffected;
                            
                            // Log if no rows were affected (GUID mismatch)
                            if (rowsAffected == 0)
                            {
                                SafeFileLogger.SafeAppendText("db_update_debug.log", 
                                    $"[{DateTime.Now:HH:mm:ss}] ⚠️ UPDATE 0 ROWS: ClashZoneGuid={update.ClashZoneId}, SleeveInstanceId={update.SleeveInstanceId}\n");
                            }
                        }
                        
                        SafeFileLogger.SafeAppendText("db_update_debug.log", 
                            $"[{DateTime.Now:HH:mm:ss}] ✅ BatchUpdateSleeveInstanceIds: {totalRowsAffected}/{updatesList.Count} rows updated in ClashZones\n");
                    }

                    // 2. Update SleeveSnapshots (Mirrors logic in UpdateSleeveInstanceId)
                    using (var snapshotCmd = _context.Connection.CreateCommand())
                    {
                        snapshotCmd.Transaction = transaction;
                        snapshotCmd.CommandText = @"
                            UPDATE SleeveSnapshots SET
                                SleeveInstanceId = @SleeveInstanceId
                            WHERE UPPER(ClashZoneGuid) = UPPER(@ClashZoneGuid)
                                AND ClashZoneGuid != '' AND ClashZoneGuid IS NOT NULL
                                AND (ClusterInstanceId IS NULL OR ClusterInstanceId <= 0)";

                        var pGuid = snapshotCmd.Parameters.AddWithValue("@ClashZoneGuid", "");
                        var pId = snapshotCmd.Parameters.AddWithValue("@SleeveInstanceId", DBNull.Value);

                        foreach (var update in updatesList)
                        {
                            pGuid.Value = update.ClashZoneId.ToString();
                            pId.Value = update.SleeveInstanceId > 0 ? (object)update.SleeveInstanceId : DBNull.Value;
                            snapshotCmd.ExecuteNonQuery();
                        }
                    }

                    transaction.Commit();
                }

                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[ClashZoneRepository] [BATCH-UPDATE] Successfully updated {updatesList.Count} items.");
                }
            }
            catch (Exception ex)
            {
                 if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Error($"[ClashZoneRepository] [BATCH-UPDATE] FAILED: {ex.Message}");
                }
                throw;
            }
        }

/// <summary>
/// ✅ PLACEMENT FIX: Update sleeve placement data (dimensions, coordinates, and InstanceId).
/// Persists the calculated geometry used for placement and the resulting instance ID.
/// </summary>
public void BatchUpdateSleevePlacementData(IEnumerable<ClashZone> placedZones)
{
    if (placedZones == null || !placedZones.Any()) return;
    var zonesList = placedZones.Where(z => z.SleeveInstanceId > 0).ToList();
    
    if (!zonesList.Any()) return;

    if (!DeploymentConfiguration.DeploymentMode)
    {
        DebugLogger.Info($"[ClashZoneRepository] [BATCH-UPDATE-DATA] Updating placement data for {zonesList.Count} zones...");
    }

    try
    {
        using (var transaction = _context.Connection.BeginTransaction())
        {
            // 1. Update ClashZones with FULL data
            using (var cmd = _context.Connection.CreateCommand())
            {
                cmd.Transaction = transaction;
                cmd.CommandText = @"
                    UPDATE ClashZones SET 
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
                        UpdatedAt = datetime('now', '+5 hours', '+30 minutes') 
                    WHERE UPPER(ClashZoneGuid) = UPPER(@ClashZoneGuid)";

                var pGuid = cmd.Parameters.AddWithValue("@ClashZoneGuid", "");
                var pId = cmd.Parameters.AddWithValue("@SleeveInstanceId", DBNull.Value);
                var pWidth = cmd.Parameters.AddWithValue("@SleeveWidth", DBNull.Value);
                var pHeight = cmd.Parameters.AddWithValue("@SleeveHeight", DBNull.Value);
                var pDiameter = cmd.Parameters.AddWithValue("@SleeveDiameter", DBNull.Value);
                var pX = cmd.Parameters.AddWithValue("@SleevePlacementX", DBNull.Value);
                var pY = cmd.Parameters.AddWithValue("@SleevePlacementY", DBNull.Value);
                var pZ = cmd.Parameters.AddWithValue("@SleevePlacementZ", DBNull.Value);
                var pActiveX = cmd.Parameters.AddWithValue("@SleevePlacementActiveX", DBNull.Value);
                var pActiveY = cmd.Parameters.AddWithValue("@SleevePlacementActiveY", DBNull.Value);
                var pActiveZ = cmd.Parameters.AddWithValue("@SleevePlacementActiveZ", DBNull.Value);
                
                int totalRowsAffected = 0;
                foreach (var zone in zonesList)
                {
                    pGuid.Value = zone.ClashZoneGuid.ToString();
                    pId.Value = zone.SleeveInstanceId;
                    pWidth.Value = zone.SleeveWidth;
                    pHeight.Value = zone.SleeveHeight;
                    pDiameter.Value = zone.SleeveDiameter;
                    pX.Value = zone.SleevePlacementPointX;
                    pY.Value = zone.SleevePlacementPointY;
                    pZ.Value = zone.SleevePlacementPointZ;
                    pActiveX.Value = zone.SleevePlacementPointActiveDocumentX;
                    pActiveY.Value = zone.SleevePlacementPointActiveDocumentY;
                    pActiveZ.Value = zone.SleevePlacementPointActiveDocumentZ;

                    totalRowsAffected += cmd.ExecuteNonQuery();
                }
                
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[ClashZoneRepository] [BATCH-UPDATE-FULL] Updated {totalRowsAffected}/{zonesList.Count} rows in ClashZones.");
                }
            }

            // 2. Update SleeveSnapshots (InstanceId only is sufficient here)
            using (var snapshotCmd = _context.Connection.CreateCommand())
            {
                snapshotCmd.Transaction = transaction;
                snapshotCmd.CommandText = @"
                    UPDATE SleeveSnapshots SET
                        SleeveInstanceId = @SleeveInstanceId
                    WHERE UPPER(ClashZoneGuid) = UPPER(@ClashZoneGuid)
                        AND ClashZoneGuid != '' AND ClashZoneGuid IS NOT NULL
                        AND (ClusterInstanceId IS NULL OR ClusterInstanceId <= 0)";

                var pGuid = snapshotCmd.Parameters.AddWithValue("@ClashZoneGuid", "");
                var pId = snapshotCmd.Parameters.AddWithValue("@SleeveInstanceId", DBNull.Value);

                foreach (var zone in zonesList)
                {
                    pGuid.Value = zone.ClashZoneGuid.ToString();
                    pId.Value = zone.SleeveInstanceId;
                    snapshotCmd.ExecuteNonQuery();
                }
            }

            transaction.Commit();
        }
    }
    catch (Exception ex)
    {
         if (!DeploymentConfiguration.DeploymentMode)
        {
            DebugLogger.Error($"[ClashZoneRepository] [BATCH-UPDATE-FULL] FAILED: {ex.Message}");
        }
        throw;
    }
}

        /// Sets IsCombinedResolved=true, IsResolved=false, IsClusterResolved=false, 
        /// and links to the combined sleeve ID.
        /// </summary>
        public void UpdateCombinedResolutionFlags(IEnumerable<Guid> zoneGuids, int combinedSleeveId)
        {
            // SOLID Refactoring: Delegate to specialized repository if enabled
            if (OptimizationFlags.UseSolidRefactoredRepositories)
            {
                var constituents = zoneGuids.Select(g => new SleeveConstituent
                {
                    Type = ConstituentType.Individual,
                    ClashZoneGuid = g
                }).ToList();

                _combinedRepo.MarkConstituentsAsResolved(constituents, combinedSleeveId);
                return;
            }

            var guidsList = zoneGuids.ToList();
            if (guidsList.Count == 0) return;

            try
            {
                var guidString = string.Join(",", guidsList.Select(g => $"'{g.ToString().ToUpperInvariant()}'"));

                using (var transaction = _context.Connection.BeginTransaction())
                {
                    using (var cmd = _context.Connection.CreateCommand())
                    {
                        cmd.Transaction = transaction;
                        cmd.CommandText = $@"
                            UPDATE ClashZones 
                            SET IsCombinedResolved = 1,
                                IsResolvedFlag = 0,
                                IsClusterResolvedFlag = 0,
                                CombinedClusterSleeveInstanceId = @CombinedSleeveId,
                                UpdatedAt = CURRENT_TIMESTAMP
                            WHERE UPPER(ClashZoneGuid) IN ({guidString})";

                        cmd.Parameters.AddWithValue("@CombinedSleeveId", combinedSleeveId);

                        int rows = cmd.ExecuteNonQuery();
                        if (!OptimizationFlags.DisableVerboseLogging)
                        {
                            _logger($"[SQLite] Updated {rows} zones for combined sleeve ID {combinedSleeveId}");
                        }
                    }
                    transaction.Commit();
                }
            }
            catch (Exception ex)
            {
                _logger($"[SQLite] ❌ Error in UpdateCombinedResolutionFlags: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Updates resolution flags for all zones belonging to the specified cluster instance IDs.
        /// This ensures that even if individual zone GUIDs are not known, all members of the cluster are updated.
        /// </summary>
        /// <param name="clusterInstanceIds">The list of ClusterInstanceIds (from ClusterSleeves table) whose member zones should be updated.</param>
        /// <param name="combinedSleeveId">The ID of the new combined sleeve.</param>
        public void UpdateCombinedResolutionFlagsByClusterIds(IEnumerable<int> clusterInstanceIds, int combinedSleeveId)
        {
            var ids = clusterInstanceIds?.ToList();
            if (ids == null || ids.Count == 0) return;

            _logger($"[SQLite] 🔍 UpdateCombinedResolutionFlagsByClusterIds called with {ids.Count} cluster IDs: [{string.Join(", ", ids)}], CombinedSleeveId={combinedSleeveId}");

            try
            {
                var idString = string.Join(",", ids);

                using (var transaction = _context.Connection.BeginTransaction())
                {
                    using (var cmd = _context.Connection.CreateCommand())
                    {
                        cmd.Transaction = transaction;
                        cmd.CommandText = $@"
                            UPDATE ClashZones 
                            SET IsCombinedResolved = 1,
                                IsResolvedFlag = 0,
                                IsClusterResolvedFlag = 0,
                                CombinedClusterSleeveInstanceId = @CombinedSleeveId,
                                UpdatedAt = CURRENT_TIMESTAMP
                            WHERE ClusterInstanceId IN ({idString})";

                        cmd.Parameters.AddWithValue("@CombinedSleeveId", combinedSleeveId);

                        _logger($"[SQLite] 🔍 Executing SQL: UPDATE ClashZones SET IsCombinedResolved=1 WHERE ClusterInstanceId IN ({idString})");

                        int rows = cmd.ExecuteNonQuery();

                        _logger($"[SQLite] ✅ Updated {rows} zones for {ids.Count} clusters (CombinedId={combinedSleeveId})");
                    }
                    transaction.Commit();
                }
            }
            catch (Exception ex)
            {
                _logger($"[SQLite] ❌ Error in UpdateCombinedResolutionFlagsByClusterIds: {ex.Message}");
                throw;
            }
        }

        // ✅ NEW HELPER: Get single zone by sleeve ID
        public ClashZone GetClashZoneBySleeveId(int sleeveInstanceId)
        {
            if (sleeveInstanceId <= 0) return null;

            const string sql = @"
                SELECT * FROM ClashZones 
                WHERE SleeveInstanceId = @id OR ClusterInstanceId = @id 
                LIMIT 1";

            using (var cmd = _context.Connection.CreateCommand())
            {
                cmd.CommandText = sql;
                cmd.Parameters.AddWithValue("@id", sleeveInstanceId);

                using (var reader = cmd.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        return MapClashZone(reader);
                    }
                }
            }
            return null;
        }

        // ✅ NEW HELPER: Get single zone by MEP Element ID
        public ClashZone GetClashZoneByMepElementId(int mepElementId)
        {
            if (mepElementId <= 0) return null;

            const string sql = @"
                SELECT * FROM ClashZones 
                WHERE MepElementId = @id 
                LIMIT 1";

            using (var cmd = _context.Connection.CreateCommand())
            {
                cmd.CommandText = sql;
                cmd.Parameters.AddWithValue("@id", mepElementId);

                using (var reader = cmd.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        return MapClashZone(reader);
                    }
                }
            }
            return null;
        }


        // ✅ NEW HELPER: Get single zone by ClashZoneGuid
        public ClashZone GetClashZoneByGuid(Guid clashZoneGuid)
        {
            if (clashZoneGuid == Guid.Empty) return null;

            const string sql = @"
                SELECT * FROM ClashZones 
                WHERE UPPER(ClashZoneGuid) = UPPER(@guid) 
                  AND ClashZoneGuid IS NOT NULL 
                  AND ClashZoneGuid != ''
                LIMIT 1";

            using (var cmd = _context.Connection.CreateCommand())
            {
                cmd.CommandText = sql;
                cmd.Parameters.AddWithValue("@guid", clashZoneGuid.ToString().ToUpperInvariant());

                using (var reader = cmd.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        return MapClashZone(reader);
                    }
                }
            }
            return null;
        }



        /// <summary>
        /// ✅ DEBUG: Get flag statistics for all zones in DB
        /// </summary>
        public (int Total, int IsCurrentClashSet, int ReadyForPlacementSet, int IsResolvedSet) GetFlagStatistics()
        {
            try
            {
                using (var cmd = _context.Connection.CreateCommand())
                {
                    cmd.CommandText = @"
                        SELECT 
                            COUNT(*) as Total,
                            SUM(CASE WHEN IsCurrentClashFlag = 1 THEN 1 ELSE 0 END) as IsCurrentClashSet,
                            SUM(CASE WHEN ReadyForPlacementFlag = 1 THEN 1 ELSE 0 END) as ReadyForPlacementSet,
                            SUM(CASE WHEN IsResolved = 1 THEN 1 ELSE 0 END) as IsResolvedSet
                        FROM ClashZones";

                    using (var reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            return (
                                Convert.ToInt32(reader["Total"]),
                                Convert.ToInt32(reader["IsCurrentClashSet"]),
                                Convert.ToInt32(reader["ReadyForPlacementSet"]),
                                Convert.ToInt32(reader["IsResolvedSet"])
                            );
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger($"[SQLite] ❌ GetFlagStatistics error: {ex.Message}");
            }
            return (0, 0, 0, 0);
        }

        public void UpdateCalculatedSleeveData(System.Guid clashZoneGuid,
            double width, double height, double depth,
            double rotation, string familyName,
            string status, string batchId)
        {
            try
            {
                using (var cmd = _context.Connection.CreateCommand())
                {
                    cmd.CommandText = @"
                        UPDATE ClashZones SET
                            CalculatedSleeveWidth = @Width,
                            CalculatedSleeveHeight = @Height,
                            CalculatedSleeveDepth = @Depth,
                            CalculatedRotation = @Rotation,
                            CalculatedFamilyName = @FamilyName,
                            PlacementStatus = @Status,
                            CalculationBatchId = @BatchId,
                            CalculatedAt = CURRENT_TIMESTAMP
                        WHERE UPPER(ClashZoneGuid) = UPPER(@Guid)
                          AND ClashZoneGuid != '' AND ClashZoneGuid IS NOT NULL";

                    cmd.Parameters.AddWithValue("@Guid", clashZoneGuid.ToString());
                    // Persist dimensions in Revit internal units (feet)
                    cmd.Parameters.AddWithValue("@Width", width);
                    cmd.Parameters.AddWithValue("@Height", height);
                    cmd.Parameters.AddWithValue("@Depth", depth);
                    cmd.Parameters.AddWithValue("@Rotation", rotation);
                    cmd.Parameters.AddWithValue("@FamilyName", familyName ?? string.Empty);
                    cmd.Parameters.AddWithValue("@Status", status ?? "Pending");
                    cmd.Parameters.AddWithValue("@BatchId", batchId ?? string.Empty);

                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    _logger($"[SQLite] ❌ Error in UpdateCalculatedSleeveData: {ex.Message}");
                }
                throw;
            }
        }

        /// <summary>
        /// ✅ DEBUGTOOL: Reset IsResolved and IsClusterResolved flags to false for all zones 
        /// whose Center point falls within the given Section Box (World Coordinates).
        /// </summary>
        public int ResetResolvedFlagsInSectionBox(BoundingBoxXYZ sectionBox)
        {
            if (sectionBox == null) return 0;

            try
            {
                var min = sectionBox.Min;
                var max = sectionBox.Max;

                // Ensure Min is actually smaller than Max (Revit BBox can be flipped in some transforms, but Min/Max properties usually corrected? 
                // Actually BBox.Min/Max are just points. Standardize values for query.)
                double minX = Math.Min(min.X, max.X);
                double maxX = Math.Max(min.X, max.X);
                double minY = Math.Min(min.Y, max.Y);
                double maxY = Math.Max(min.Y, max.Y);
                double minZ = Math.Min(min.Z, max.Z);
                double maxZ = Math.Max(min.Z, max.Z);

                using (var cmd = _context.Connection.CreateCommand())
                {
                    cmd.CommandText = @"
                        UPDATE ClashZones 
                        SET IsResolved = 0, 
                            IsClusterResolved = 0, 
                            IsCombinedResolved = 0
                        WHERE CenterX >= @minX AND CenterX <= @maxX
                          AND CenterY >= @minY AND CenterY <= @maxY
                          AND CenterZ >= @minZ AND CenterZ <= @maxZ;
                    ";

                    cmd.Parameters.AddWithValue("@minX", minX);
                    cmd.Parameters.AddWithValue("@maxX", maxX);
                    cmd.Parameters.AddWithValue("@minY", minY);
                    cmd.Parameters.AddWithValue("@maxY", maxY);
                    cmd.Parameters.AddWithValue("@minZ", minZ);
                    cmd.Parameters.AddWithValue("@maxZ", maxZ);

                    int affected = cmd.ExecuteNonQuery();
                    _logger($"[SQLite] ResetResolvedFlagsInSectionBox: Reset {affected} zones in provided bounds.");
                    return affected;
                }
            }
            catch (Exception ex)
            {
                _logger($"[SQLite] ❌ Error in ResetResolvedFlagsInSectionBox: {ex.Message}");
                return 0;
            }
        }



        public List<ClashZone> GetPlacedClashZones()
        {
            var list = new List<ClashZone>();
            try
            {
                using (var cmd = _context.Connection.CreateCommand())
                {
                    cmd.CommandText = "SELECT * FROM ClashZones WHERE SleeveInstanceId > 0 AND SleeveInstanceId IS NOT NULL";
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            list.Add(MapClashZone(reader));
                        }
                    }
                }
                
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[ClashZoneRepository] [DB-LOAD] Loaded {list.Count} PLACED clash zones (SleeveInstanceId > 0)");
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    _logger($"[SQLite] Error Get Placed Clash Zones: {ex.Message}");
            }
            return list;
        }
        /// <summary>
        /// ✅ PRE-PLACEMENT PERSISTENCE: Save calculated sleeve data (dimensions, coordinates) to database BEFORE placement.
        /// This ensures data is saved even if placement fails or crashes. Uses ClashZoneGuid for lookup.
        /// </summary>
        public void UpdateSleeveCalculatedData(ClashZone zone)
        {
            if (zone == null || string.IsNullOrEmpty(zone.ClashZoneGuid)) return;

            // Log that we are attempting to save
            if (!DeploymentConfiguration.DeploymentMode)
            {
                SafeFileLogger.SafeAppendText("db_update_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss.fff}] [ClashZoneRepository] ? PRE-SAVE: Saving calculated data for Zone {zone.Id} (GUID: {zone.ClashZoneGuid})\n");
            }

            try
            {
                using (var cmd = _context.Connection.CreateCommand()) // Changed _connection to _context.Connection
                {
                    cmd.CommandText = @"
                        UPDATE ClashZones 
                        SET 
                            SleeveWidth = @Width,
                            SleeveHeight = @Height,
                            SleeveDiameter = @Diameter,
                            SleevePlacementX = @X,
                            SleevePlacementY = @Y,
                            SleevePlacementZ = @Z,
                            SleevePlacementActiveX = @ActiveX,
                            SleevePlacementActiveY = @ActiveY,
                            SleevePlacementActiveZ = @ActiveZ,
                            
                            -- ? CRITICAL FIX: Persist Calculated properties
                            CalculatedSleeveWidth = @CalcWidth,
                            CalculatedSleeveHeight = @CalcHeight,
                            CalculatedSleeveDiameter = @CalcDiameter,
                            CalculatedSleeveDepth = @CalcDepth,
                            CalculatedRotation = @CalcRotation,
                            CalculatedPlacementX = @CalcX,
                            CalculatedPlacementY = @CalcY,
                            CalculatedPlacementZ = @CalcZ,
                            CalculatedFamilyName = @CalcFam
                        WHERE ClashZoneGuid = @Guid";

                    cmd.Parameters.AddWithValue("@Width", zone.SleeveWidth);
                    cmd.Parameters.AddWithValue("@Height", zone.SleeveHeight);
                    cmd.Parameters.AddWithValue("@Diameter", zone.SleeveDiameter);
                    cmd.Parameters.AddWithValue("@X", zone.SleevePlacementPointX);
                    cmd.Parameters.AddWithValue("@Y", zone.SleevePlacementPointY);
                    cmd.Parameters.AddWithValue("@Z", zone.SleevePlacementPointZ);
                    cmd.Parameters.AddWithValue("@ActiveX", zone.SleevePlacementPointActiveDocumentX);
                    cmd.Parameters.AddWithValue("@ActiveY", zone.SleevePlacementPointActiveDocumentY);
                    cmd.Parameters.AddWithValue("@ActiveZ", zone.SleevePlacementPointActiveDocumentZ);
                    
                    // Calculated Parameters
                    cmd.Parameters.AddWithValue("@CalcWidth", zone.CalculatedSleeveWidth);
                    cmd.Parameters.AddWithValue("@CalcHeight", zone.CalculatedSleeveHeight);
                    cmd.Parameters.AddWithValue("@CalcDiameter", zone.CalculatedSleeveDiameter);
                    cmd.Parameters.AddWithValue("@CalcDepth", zone.CalculatedSleeveDepth);
                    cmd.Parameters.AddWithValue("@CalcRotation", zone.CalculatedRotation);
                    cmd.Parameters.AddWithValue("@CalcX", zone.CalculatedPlacementX);
                    cmd.Parameters.AddWithValue("@CalcY", zone.CalculatedPlacementY);
                    cmd.Parameters.AddWithValue("@CalcZ", zone.CalculatedPlacementZ);
                    cmd.Parameters.AddWithValue("@CalcFam", zone.CalculatedFamilyName ?? string.Empty);
                    
                    cmd.Parameters.AddWithValue("@Guid", zone.ClashZoneGuid);

                    int rows = cmd.ExecuteNonQuery();
                    
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                         SafeFileLogger.SafeAppendText("db_update_debug.log", 
                            $"[{DateTime.Now:HH:mm:ss.fff}] [ClashZoneRepository] ✅ UPDATE COMPLETE: Zone {zone.Id}, GUID={zone.ClashZoneGuid}, Rows Affected: {rows}\n");
                             
                         if (rows == 0)
                         {
                             SafeFileLogger.SafeAppendText("db_update_debug.log", 
                                $"[{DateTime.Now:HH:mm:ss.fff}] [ClashZoneRepository] ❌ PERSISTENCE FAILURE: No row found for GUID {zone.ClashZoneGuid}. Check for casing or whitespace issues.\n");
                         }
                    }
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[ClashZoneRepository] Error updating sleeve calculated data: {ex.Message}");
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    SafeFileLogger.SafeAppendText("db_update_debug.log", 
                        $"[{DateTime.Now:HH:mm:ss.fff}] [ClashZoneRepository] ?? ERROR: Failed to save calculated data for Zone {zone.Id}: {ex.Message}\n");
                }
            }
        }
        /// <summary>
        /// Stub implementation for ISleeveRepository compatibility.
        /// Feature deprecated in favor of SQLite database.
        /// </summary>
        public void SaveSleeveDataToClusterXml(string xmlFilePath, List<ClashZone> clashZones, Document doc)
        {
            // No-op: XML storage deprecated
        }

        /// <summary>
        /// Stub implementation for ISleeveRepository compatibility.
        /// Feature deprecated in favor of SQLite database.
        /// </summary>
        public void RegenerateClusterXmlFiles(Document doc, List<ClashZone> placedZones)
        {
            // No-op: XML storage deprecated
        }

        /// <summary>
        /// Stub implementation for ISleeveRepository compatibility.
        /// Feature deprecated in favor of SQLite database.
        /// </summary>
        public List<ClashZone> LoadDuctAccessoriesClashZones(string xmlFilePath)
        {
            return new List<ClashZone>(); // Return empty list as we rely on DB
        }
    }
}

using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using System.Text.Json;
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Data.Repositories
{
    /// <summary>
    /// 🚀 BATCH SAVE: Data transfer object for batch cluster save operations
    /// </summary>
    public class ClusterSaveData
    {
        public int ClusterInstanceId { get; set; }
        public int ComboId { get; set; }
        public int FilterId { get; set; }
        public string Category { get; set; }
        public double BoundingBoxMinX { get; set; }
        public double BoundingBoxMinY { get; set; }
        public double BoundingBoxMinZ { get; set; }
        public double BoundingBoxMaxX { get; set; }
        public double BoundingBoxMaxY { get; set; }
        public double BoundingBoxMaxZ { get; set; }
        public double ClusterWidth { get; set; }
        public double ClusterHeight { get; set; }
        public double ClusterDepth { get; set; }
        public double RotationAngleDeg { get; set; }
        public bool IsRotated { get; set; }
        public double PlacementX { get; set; }
        public double PlacementY { get; set; }
        public double PlacementZ { get; set; }
        public string HostType { get; set; }
        public string HostOrientation { get; set; }
        public List<Guid> ClashZoneIds { get; set; }
        
        // ✅ PERSISTENCE FIX: Save Family Name for validation
        public string SleeveFamilyName { get; set; }

        // ✅ CORNER PERISISTENCE (Added for proper cluster sizing)
        public double Corner1X { get; set; }
        public double Corner1Y { get; set; }
        public double Corner1Z { get; set; }
        public double Corner2X { get; set; }
        public double Corner2Y { get; set; }
        public double Corner2Z { get; set; }
        public double Corner3X { get; set; }
        public double Corner3Y { get; set; }
        public double Corner3Z { get; set; }
        public double Corner4X { get; set; }
        public double Corner4Y { get; set; }
        public double Corner4Z { get; set; }
    }

    /// <summary>
    /// ✅ CLUSTER SLEEVE STORAGE: Repository for storing and retrieving cluster sleeve calculation results
    /// Enables PATH 1 (Replay) to use pre-calculated cluster data without recalculating
    /// </summary>
    public class ClusterSleeveRepository
    {
        private readonly SleeveDbContext _context;
        private readonly Action<string> _logger;

        public ClusterSleeveRepository(SleeveDbContext context, Action<string> logger = null)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            _logger = logger ?? (_ => { });
        }

        /// <summary>
        /// Retrieves all cluster sleeves from the database.
        /// </summary>
        public List<ClusterSleeveData> GetAllClusterSleeves()
        {
            var clusters = new List<ClusterSleeveData>();
            using (var cmd = _context.Connection.CreateCommand())
            {
                cmd.CommandText = @"SELECT ClusterInstanceId, ComboId, FilterId, Category,
                                       BoundingBoxMinX, BoundingBoxMinY, BoundingBoxMinZ,
                                       BoundingBoxMaxX, BoundingBoxMaxY, BoundingBoxMaxZ,
                                       ClusterWidth, ClusterHeight, ClusterDepth,
                                       RotationAngleDeg, IsRotated,
                                       PlacementX, PlacementY, PlacementZ,
                                       HostType, HostOrientation, ClashZoneIdsJson,
                                       SleeveFamilyName,
                                       Corner1X, Corner1Y, Corner1Z,
                                       Corner2X, Corner2Y, Corner2Z,
                                       Corner3X, Corner3Y, Corner3Z,
                                       Corner4X, Corner4Y, Corner4Z
                                FROM ClusterSleeves";

                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        var cluster = new ClusterSleeveData
                        {
                            ClusterInstanceId = GetInt(reader, "ClusterInstanceId", -1),
                            ComboId = GetInt(reader, "ComboId", -1),
                            FilterId = GetInt(reader, "FilterId", -1),
                            Category = GetString(reader, "Category"),
                            BoundingBoxMinX = GetDouble(reader, "BoundingBoxMinX", 0.0),
                            BoundingBoxMinY = GetDouble(reader, "BoundingBoxMinY", 0.0),
                            BoundingBoxMinZ = GetDouble(reader, "BoundingBoxMinZ", 0.0),
                            BoundingBoxMaxX = GetDouble(reader, "BoundingBoxMaxX", 0.0),
                            BoundingBoxMaxY = GetDouble(reader, "BoundingBoxMaxY", 0.0),
                            BoundingBoxMaxZ = GetDouble(reader, "BoundingBoxMaxZ", 0.0),
                            ClusterWidth = GetDouble(reader, "ClusterWidth", 0.0),
                            ClusterHeight = GetDouble(reader, "ClusterHeight", 0.0),
                            ClusterDepth = GetDouble(reader, "ClusterDepth", 0.0),
                            RotationAngleDeg = GetDouble(reader, "RotationAngleDeg", 0.0),
                            IsRotated = GetBool(reader, "IsRotated"),
                            PlacementX = GetDouble(reader, "PlacementX", 0.0),
                            PlacementY = GetDouble(reader, "PlacementY", 0.0),
                            PlacementZ = GetDouble(reader, "PlacementZ", 0.0),
                            HostType = GetString(reader, "HostType"),
                            HostOrientation = GetString(reader, "HostOrientation"),
                            Corner1X = GetDouble(reader, "Corner1X", 0.0),
                            Corner1Y = GetDouble(reader, "Corner1Y", 0.0),
                            Corner1Z = GetDouble(reader, "Corner1Z", 0.0),
                            Corner2X = GetDouble(reader, "Corner2X", 0.0),
                            Corner2Y = GetDouble(reader, "Corner2Y", 0.0),
                            Corner2Z = GetDouble(reader, "Corner2Z", 0.0),
                            Corner3X = GetDouble(reader, "Corner3X", 0.0),
                            Corner3Y = GetDouble(reader, "Corner3Y", 0.0),
                            Corner3Z = GetDouble(reader, "Corner3Z", 0.0),
                            Corner4X = GetDouble(reader, "Corner4X", 0.0),
                            Corner4Y = GetDouble(reader, "Corner4Y", 0.0),
                            Corner4Z = GetDouble(reader, "Corner4Z", 0.0)
                        };

                        // Deserialize ClashZoneIds from JSON
                        var clashZoneIdsJson = GetString(reader, "ClashZoneIdsJson");
                        if (!string.IsNullOrWhiteSpace(clashZoneIdsJson))
                        {
                            try
                            {
                                var guidStrings = JsonSerializer.Deserialize<List<string>>(clashZoneIdsJson);
                                cluster.ClashZoneIds = guidStrings?.Select(g => Guid.Parse(g)).ToList() ?? new List<Guid>();
                            }
                            catch
                            {
                                cluster.ClashZoneIds = new List<Guid>();
                            }
                        }
                        else
                        {
                            cluster.ClashZoneIds = new List<Guid>();
                        }

                        clusters.Add(cluster);
                    }
                }
            }
            return clusters;
        }

        /// <summary>
        /// Retrieves cluster sleeves associated with the given instance IDs.
        /// </summary>
        public List<ClusterSleeveData> GetClusterSleevesByInstanceIds(IEnumerable<int> instanceIds)
        {
            var clusters = new List<ClusterSleeveData>();
            var ids = instanceIds?.ToList();
            if (ids == null || ids.Count == 0) return clusters;

            var idString = string.Join(",", ids);
            using (var cmd = _context.Connection.CreateCommand())
            {
                cmd.CommandText = $@"SELECT ClusterInstanceId, ComboId, FilterId, Category,
                                       BoundingBoxMinX, BoundingBoxMinY, BoundingBoxMinZ,
                                       BoundingBoxMaxX, BoundingBoxMaxY, BoundingBoxMaxZ,
                                       ClusterWidth, ClusterHeight, ClusterDepth,
                                       RotationAngleDeg, IsRotated,
                                       PlacementX, PlacementY, PlacementZ,
                                       HostType, HostOrientation, ClashZoneIdsJson,
                                       SleeveFamilyName,
                                       Corner1X, Corner1Y, Corner1Z,
                                       Corner2X, Corner2Y, Corner2Z,
                                       Corner3X, Corner3Y, Corner3Z,
                                       Corner4X, Corner4Y, Corner4Z
                                FROM ClusterSleeves
                                WHERE ClusterInstanceId IN ({idString})";

                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        var cluster = new ClusterSleeveData
                        {
                            ClusterInstanceId = GetInt(reader, "ClusterInstanceId", -1),
                            ComboId = GetInt(reader, "ComboId", -1),
                            FilterId = GetInt(reader, "FilterId", -1),
                            Category = GetString(reader, "Category"),
                            BoundingBoxMinX = GetDouble(reader, "BoundingBoxMinX", 0.0),
                            BoundingBoxMinY = GetDouble(reader, "BoundingBoxMinY", 0.0),
                            BoundingBoxMinZ = GetDouble(reader, "BoundingBoxMinZ", 0.0),
                            BoundingBoxMaxX = GetDouble(reader, "BoundingBoxMaxX", 0.0),
                            BoundingBoxMaxY = GetDouble(reader, "BoundingBoxMaxY", 0.0),
                            BoundingBoxMaxZ = GetDouble(reader, "BoundingBoxMaxZ", 0.0),
                            ClusterWidth = GetDouble(reader, "ClusterWidth", 0.0),
                            ClusterHeight = GetDouble(reader, "ClusterHeight", 0.0),
                            ClusterDepth = GetDouble(reader, "ClusterDepth", 0.0),
                            RotationAngleDeg = GetDouble(reader, "RotationAngleDeg", 0.0),
                            IsRotated = GetBool(reader, "IsRotated"),
                            PlacementX = GetDouble(reader, "PlacementX", 0.0),
                            PlacementY = GetDouble(reader, "PlacementY", 0.0),
                            PlacementZ = GetDouble(reader, "PlacementZ", 0.0),
                            HostType = GetString(reader, "HostType"),
                            HostOrientation = GetString(reader, "HostOrientation"),
                            Corner1X = GetDouble(reader, "Corner1X", 0.0),
                            Corner1Y = GetDouble(reader, "Corner1Y", 0.0),
                            Corner1Z = GetDouble(reader, "Corner1Z", 0.0),
                            Corner2X = GetDouble(reader, "Corner2X", 0.0),
                            Corner2Y = GetDouble(reader, "Corner2Y", 0.0),
                            Corner2Z = GetDouble(reader, "Corner2Z", 0.0),
                            Corner3X = GetDouble(reader, "Corner3X", 0.0),
                            Corner3Y = GetDouble(reader, "Corner3Y", 0.0),
                            Corner3Z = GetDouble(reader, "Corner3Z", 0.0),
                            Corner4X = GetDouble(reader, "Corner4X", 0.0),
                            Corner4Y = GetDouble(reader, "Corner4Y", 0.0),
                            Corner4Z = GetDouble(reader, "Corner4Z", 0.0)
                        };

                        // Deserialize ClashZoneIds from JSON
                        var clashZoneIdsJson = GetString(reader, "ClashZoneIdsJson");
                        if (!string.IsNullOrWhiteSpace(clashZoneIdsJson))
                        {
                            try
                            {
                                var guidStrings = JsonSerializer.Deserialize<List<string>>(clashZoneIdsJson);
                                cluster.ClashZoneIds = guidStrings?.Select(g => Guid.Parse(g)).ToList() ?? new List<Guid>();
                            }
                            catch
                            {
                                cluster.ClashZoneIds = new List<Guid>();
                            }
                        }
                        else
                        {
                            cluster.ClashZoneIds = new List<Guid>();
                        }

                        clusters.Add(cluster);
                    }
                }
            }
            return clusters;
        }
            // ...existing code...
        


        /// <summary>
        /// Save cluster sleeve calculation results to database
        /// Called after cluster calculation completes in PATH 2/3
        /// </summary>
        public void SaveClusterSleeve(
            int clusterInstanceId,
            int comboId,
            int filterId,
            string category,
            double boundingBoxMinX, double boundingBoxMinY, double boundingBoxMinZ,
            double boundingBoxMaxX, double boundingBoxMaxY, double boundingBoxMaxZ,
            double clusterWidth, double clusterHeight, double clusterDepth,
            double rotationAngleDeg,
            bool isRotated,
            double placementX, double placementY, double placementZ,
            string hostType,
            string hostOrientation,
            List<Guid> clashZoneIds,
            // ✅ PERSISTENCE FIX: Added SleeveFamilyName
            string sleeveFamilyName,
            // ✅ CORNER PERISISTENCE (Added for proper cluster sizing)
            double corner1X, double corner1Y, double corner1Z,
            double corner2X, double corner2Y, double corner2Z,
            double corner3X, double corner3Y, double corner3Z,
            double corner4X, double corner4Y, double corner4Z)
        {
            if (clusterInstanceId <= 0)
                throw new ArgumentException("ClusterInstanceId must be greater than 0", nameof(clusterInstanceId));

            // ✅ CRITICAL FIX: Generate deterministic ClusterGuid from sorted ClashZoneIds (like ClashZoneGuid for snapshots)
            // This allows proper upserting even if ClusterInstanceId changes
            string clusterGuid = GenerateDeterministicClusterGuid(clashZoneIds);

            using (var transaction = _context.Connection.BeginTransaction())
            {
                try
                {
                    // ✅ CRITICAL FIX: Check by ClusterGuid first (deterministic), then fallback to ClusterInstanceId
                    using (var checkCmd = _context.Connection.CreateCommand())
                    {
                        checkCmd.Transaction = transaction;
                        if (!string.IsNullOrWhiteSpace(clusterGuid))
                        {
                            // Priority 1: Check by ClusterGuid (deterministic)
                            checkCmd.CommandText = @"
                                SELECT ClusterSleeveId FROM ClusterSleeves 
                                WHERE ClusterGuid = @ClusterGuid";
                            checkCmd.Parameters.AddWithValue("@ClusterGuid", clusterGuid);
                        }
                        else
                        {
                            // Fallback: Check by ClusterInstanceId (may change)
                            checkCmd.CommandText = @"
                            SELECT ClusterSleeveId FROM ClusterSleeves 
                            WHERE ClusterInstanceId = @ClusterInstanceId";
                            checkCmd.Parameters.AddWithValue("@ClusterInstanceId", clusterInstanceId);
                        }

                        var existingId = checkCmd.ExecuteScalar();

                        if (existingId != null)
                        {
                            // Update existing cluster
                            using (var updateCmd = _context.Connection.CreateCommand())
                            {
                                updateCmd.Transaction = transaction;
                                // ✅ CRITICAL FIX: Update by ClusterGuid (deterministic) if available, otherwise by ClusterInstanceId
                                if (!string.IsNullOrWhiteSpace(clusterGuid))
                                {
                                    updateCmd.CommandText = @"
                                        UPDATE ClusterSleeves SET
                                            ClusterInstanceId = @ClusterInstanceId,
                                            ComboId = @ComboId,
                                            FilterId = @FilterId,
                                            Category = @Category,
                                            BoundingBoxMinX = @BoundingBoxMinX,
                                            BoundingBoxMinY = @BoundingBoxMinY,
                                            BoundingBoxMinZ = @BoundingBoxMinZ,
                                            BoundingBoxMaxX = @BoundingBoxMaxX,
                                            BoundingBoxMaxY = @BoundingBoxMaxY,
                                            BoundingBoxMaxZ = @BoundingBoxMaxZ,
                                            ClusterWidth = @ClusterWidth,
                                            ClusterHeight = @ClusterHeight,
                                            ClusterDepth = @ClusterDepth,
                                            RotationAngleDeg = @RotationAngleDeg,
                                            IsRotated = @IsRotated,
                                            PlacementX = @PlacementX,
                                            PlacementY = @PlacementY,
                                            PlacementZ = @PlacementZ,
                                            HostType = @HostType,
                                            HostOrientation = @HostOrientation,
                                            ClashZoneIdsJson = @ClashZoneIdsJson,
                                            ClashZoneGuids = @ClashZoneGuids,
                                            MepSizes = @MepSizes,
                                            MepSystemNames = @MepSystemNames,
                                            MepElementIds = @MepElementIds,
                                            UpdatedAt = CURRENT_TIMESTAMP
                                        WHERE ClusterGuid = @ClusterGuid";
                                }
                                else
                                {
                                    updateCmd.CommandText = @"
                                    UPDATE ClusterSleeves SET
                                        ComboId = @ComboId,
                                        FilterId = @FilterId,
                                        Category = @Category,
                                        BoundingBoxMinX = @BoundingBoxMinX,
                                        BoundingBoxMinY = @BoundingBoxMinY,
                                        BoundingBoxMinZ = @BoundingBoxMinZ,
                                        BoundingBoxMaxX = @BoundingBoxMaxX,
                                        BoundingBoxMaxY = @BoundingBoxMaxY,
                                        BoundingBoxMaxZ = @BoundingBoxMaxZ,
                                        ClusterWidth = @ClusterWidth,
                                        ClusterHeight = @ClusterHeight,
                                        ClusterDepth = @ClusterDepth,
                                        RotationAngleDeg = @RotationAngleDeg,
                                        IsRotated = @IsRotated,
                                        PlacementX = @PlacementX,
                                        PlacementY = @PlacementY,
                                        PlacementZ = @PlacementZ,
                                        HostType = @HostType,
                                        HostOrientation = @HostOrientation,
                                        ClashZoneIdsJson = @ClashZoneIdsJson,
                                        SleeveFamilyName = @SleeveFamilyName,
                                        ClashZoneGuids = @ClashZoneGuids,
                                        MepSizes = @MepSizes,
                                        MepSystemNames = @MepSystemNames,
                                        MepElementIds = @MepElementIds,
                                        UpdatedAt = CURRENT_TIMESTAMP
                                    WHERE ClusterInstanceId = @ClusterInstanceId";
                                }

                                AddClusterSleeveParameters(updateCmd, clusterInstanceId, comboId, filterId, category,
                                    boundingBoxMinX, boundingBoxMinY, boundingBoxMinZ,
                                    boundingBoxMaxX, boundingBoxMaxY, boundingBoxMaxZ,
                                    clusterWidth, clusterHeight, clusterDepth,
                                    rotationAngleDeg, isRotated,
                                    placementX, placementY, placementZ,
                                    hostType, hostOrientation, clashZoneIds,
                                    corner1X, corner1Y, corner1Z,
                                    corner2X, corner2Y, corner2Z,
                                    corner3X, corner3Y, corner3Z,
                                    corner4X, corner4Y, corner4Z,
                                    clusterGuid);

                                // Add ClusterGuid parameter for WHERE clause
                                if (!string.IsNullOrWhiteSpace(clusterGuid))
                                {
                                    updateCmd.Parameters.AddWithValue("@ClusterGuid", clusterGuid);
                                }

                                updateCmd.ExecuteNonQuery();
                                _logger($"[SQLite] ✅ Updated cluster sleeve {clusterInstanceId} in database");
                            }
                        }
                        else
                        {
                            // Insert new cluster
                            using (var insertCmd = _context.Connection.CreateCommand())
                            {
                                insertCmd.Transaction = transaction;
                                insertCmd.CommandText = @"
                                    INSERT INTO ClusterSleeves (
                                        ClusterInstanceId, ClusterGuid, ComboId, FilterId, Category,
                                        BoundingBoxMinX, BoundingBoxMinY, BoundingBoxMinZ,
                                        BoundingBoxMaxX, BoundingBoxMaxY, BoundingBoxMaxZ,
                                        ClusterWidth, ClusterHeight, ClusterDepth,
                                        RotationAngleDeg, IsRotated,
                                        PlacementX, PlacementY, PlacementZ,
                                        HostType, HostOrientation, ClashZoneIdsJson,
                                        SleeveFamilyName,
                                        ClashZoneGuids, MepSizes, MepSystemNames, MepElementIds,
                                        CreatedAt, UpdatedAt
                                    ) VALUES (
                                        @ClusterInstanceId, @ClusterGuid, @ComboId, @FilterId, @Category,
                                        @BoundingBoxMinX, @BoundingBoxMinY, @BoundingBoxMinZ,
                                        @BoundingBoxMaxX, @BoundingBoxMaxY, @BoundingBoxMaxZ,
                                        @ClusterWidth, @ClusterHeight, @ClusterDepth,
                                        @RotationAngleDeg, @IsRotated,
                                        @PlacementX, @PlacementY, @PlacementZ,
                                        @HostType, @HostOrientation, @ClashZoneIdsJson,
                                        @SleeveFamilyName,
                                        @ClashZoneGuids, @MepSizes, @MepSystemNames, @MepElementIds,
                                        CURRENT_TIMESTAMP, CURRENT_TIMESTAMP
                                    )";

                                AddClusterSleeveParameters(insertCmd, clusterInstanceId, comboId, filterId, category,
                                    boundingBoxMinX, boundingBoxMinY, boundingBoxMinZ,
                                    boundingBoxMaxX, boundingBoxMaxY, boundingBoxMaxZ,
                                    clusterWidth, clusterHeight, clusterDepth,
                                    rotationAngleDeg, isRotated,
                                    placementX, placementY, placementZ,
                                    hostType, hostOrientation, clashZoneIds,
                                    corner1X, corner1Y, corner1Z,
                                    corner2X, corner2Y, corner2Z,
                                    corner3X, corner3Y, corner3Z,
                                    corner4X, corner4Y, corner4Z,
                                    clusterGuid);

                                var rowsAffected = insertCmd.ExecuteNonQuery();

                                // ✅ LOG: INSERT operation - ALL COLUMNS (one sample row)
                                var clashZoneIdsJson = clashZoneIds != null && clashZoneIds.Count > 0
                                    ? JsonSerializer.Serialize(clashZoneIds.Select(g => g.ToString()).ToList())
                                    : "[]";
                                var (clashZoneGuids, mepSizes, mepSystemNames, mepElementIds) = GetCommaSeparatedMepData(clashZoneIds);

                                var insertParams = new Dictionary<string, object>
                                {
                                    { "ClusterInstanceId", clusterInstanceId },
                                    { "ComboId", comboId },
                                    { "FilterId", filterId },
                                    { "Category", category ?? "NULL" },
                                    { "BoundingBoxMinX", boundingBoxMinX },
                                    { "BoundingBoxMinY", boundingBoxMinY },
                                    { "BoundingBoxMinZ", boundingBoxMinZ },
                                    { "BoundingBoxMaxX", boundingBoxMaxX },
                                    { "BoundingBoxMaxY", boundingBoxMaxY },
                                    { "BoundingBoxMaxZ", boundingBoxMaxZ },
                                    { "ClusterWidth", clusterWidth },
                                    { "ClusterHeight", clusterHeight },
                                    { "ClusterDepth", clusterDepth },
                                    { "RotationAngleDeg", rotationAngleDeg },
                                    { "IsRotated", isRotated ? 1 : 0 },
                                    { "PlacementX", placementX },
                                    { "PlacementY", placementY },
                                    { "PlacementZ", placementZ },
                                    { "HostType", hostType ?? "NULL" },
                                    { "HostOrientation", hostOrientation ?? "NULL" },
                                    { "ClashZoneIdsJson", clashZoneIdsJson },
                                    { "ClashZoneGuids", clashZoneGuids ?? "NULL" },
                                    { "MepSizes", mepSizes ?? "NULL" },
                                    { "MepSystemNames", mepSystemNames ?? "NULL" },
                                    { "MepElementIds", mepElementIds ?? "NULL" }
                                };

                                DatabaseOperationLogger.LogOperation(
                                    "INSERT",
                                    "ClusterSleeves",
                                    insertParams,
                                    rowsAffected,
                                    $"✅ Sample row: All columns logged (ClusterInstanceId={clusterInstanceId})");

                                _logger($"[SQLite] ✅ Saved cluster sleeve {clusterInstanceId} to database");
                            }
                        }
                    }

                    DatabaseOperationLogger.LogTransaction("COMMIT", "SUCCESS", null);
                    transaction.Commit();
                }
                catch (Exception ex)
                {
                    DatabaseOperationLogger.LogTransaction("ROLLBACK", "FAILED", ex.Message);
                    transaction.Rollback();
                    _logger($"[SQLite] ❌ Error saving cluster sleeve {clusterInstanceId}: {ex.Message}");
                    throw;
                }
            }
        }

        /// <summary>
        /// 🚀 BATCH SAVE: Save multiple cluster sleeves in a single transaction
        /// Significantly faster than calling SaveClusterSleeve in a loop (113ms → ~10ms)
        /// </summary>
        public void BatchSaveClusterSleeves(List<ClusterSaveData> clusters)
        {
            if (clusters == null || clusters.Count == 0)
                return;

            // ✅ BULK OPTIMIZATION: Use bulk operations if enabled
            if (OptimizationFlags.UseBulkClusterSave)
            {
                BatchSaveClusterSleevesBulk(clusters);
                return;
            }

            using (var transaction = _context.Connection.BeginTransaction())
            {
                try
                {
                    foreach (var cluster in clusters)
                    {
                        // ✅ CRITICAL FIX: Generate deterministic ClusterGuid from sorted ClashZoneIds
                        string clusterGuid = GenerateDeterministicClusterGuid(cluster.ClashZoneIds);

                        // Check if cluster already exists
                        using (var checkCmd = _context.Connection.CreateCommand())
                        {
                            checkCmd.Transaction = transaction;
                            // ✅ CRITICAL FIX: Check by ClusterGuid first (deterministic), then fallback to ClusterInstanceId
                            if (!string.IsNullOrWhiteSpace(clusterGuid))
                            {
                                checkCmd.CommandText = @"
                                    SELECT ClusterSleeveId FROM ClusterSleeves 
                                    WHERE ClusterGuid = @ClusterGuid";
                                checkCmd.Parameters.AddWithValue("@ClusterGuid", clusterGuid);
                            }
                            else
                            {
                                checkCmd.CommandText = @"
                                SELECT ClusterSleeveId FROM ClusterSleeves 
                                WHERE ClusterInstanceId = @ClusterInstanceId";
                                checkCmd.Parameters.AddWithValue("@ClusterInstanceId", cluster.ClusterInstanceId);
                            }

                            var existingId = checkCmd.ExecuteScalar();

                            if (existingId != null)
                            {
                                // Update existing cluster
                                using (var updateCmd = _context.Connection.CreateCommand())
                                {
                                    updateCmd.Transaction = transaction;
                                    // ✅ CRITICAL FIX: Update by ClusterGuid (deterministic) if available, otherwise by ClusterInstanceId
                                    if (!string.IsNullOrWhiteSpace(clusterGuid))
                                    {
                                        updateCmd.CommandText = @"
                                            UPDATE ClusterSleeves SET
                                                ClusterInstanceId = @ClusterInstanceId,
                                                ComboId = @ComboId,
                                                FilterId = @FilterId,
                                                Category = @Category,
                                                BoundingBoxMinX = @BoundingBoxMinX,
                                                BoundingBoxMinY = @BoundingBoxMinY,
                                                BoundingBoxMinZ = @BoundingBoxMinZ,
                                                BoundingBoxMaxX = @BoundingBoxMaxX,
                                                BoundingBoxMaxY = @BoundingBoxMaxY,
                                                BoundingBoxMaxZ = @BoundingBoxMaxZ,
                                                ClusterWidth = @ClusterWidth,
                                                ClusterHeight = @ClusterHeight,
                                                ClusterDepth = @ClusterDepth,
                                                RotationAngleDeg = @RotationAngleDeg,
                                                IsRotated = @IsRotated,
                                                PlacementX = @PlacementX,
                                                PlacementY = @PlacementY,
                                                PlacementZ = @PlacementZ,
                                                HostType = @HostType,
                                                HostOrientation = @HostOrientation,
                                                ClashZoneIdsJson = @ClashZoneIdsJson,
                                                SleeveFamilyName = @SleeveFamilyName,
                                                ClashZoneGuids = @ClashZoneGuids,
                                                MepSizes = @MepSizes,
                                                MepSystemNames = @MepSystemNames,
                                                MepElementIds = @MepElementIds,
                                                UpdatedAt = CURRENT_TIMESTAMP
                                            WHERE ClusterGuid = @ClusterGuid";
                                    }
                                    else
                                    {
                                        updateCmd.CommandText = @"
                                        UPDATE ClusterSleeves SET
                                            ComboId = @ComboId,
                                            FilterId = @FilterId,
                                            Category = @Category,
                                            BoundingBoxMinX = @BoundingBoxMinX,
                                            BoundingBoxMinY = @BoundingBoxMinY,
                                            BoundingBoxMinZ = @BoundingBoxMinZ,
                                            BoundingBoxMaxX = @BoundingBoxMaxX,
                                            BoundingBoxMaxY = @BoundingBoxMaxY,
                                            BoundingBoxMaxZ = @BoundingBoxMaxZ,
                                            ClusterWidth = @ClusterWidth,
                                            ClusterHeight = @ClusterHeight,
                                            ClusterDepth = @ClusterDepth,
                                            RotationAngleDeg = @RotationAngleDeg,
                                            IsRotated = @IsRotated,
                                            PlacementX = @PlacementX,
                                            PlacementY = @PlacementY,
                                            PlacementZ = @PlacementZ,
                                            HostType = @HostType,
                                            HostOrientation = @HostOrientation,
                                            ClashZoneIdsJson = @ClashZoneIdsJson,
                                            ClashZoneGuids = @ClashZoneGuids,
                                            MepSizes = @MepSizes,
                                            MepSystemNames = @MepSystemNames,
                                            MepElementIds = @MepElementIds,
                                            Corner1X = @Corner1X, Corner1Y = @Corner1Y, Corner1Z = @Corner1Z,
                                            Corner2X = @Corner2X, Corner2Y = @Corner2Y, Corner2Z = @Corner2Z,
                                            Corner3X = @Corner3X, Corner3Y = @Corner3Y, Corner3Z = @Corner3Z,
                                            Corner4X = @Corner4X, Corner4Y = @Corner4Y, Corner4Z = @Corner4Z,
                                            UpdatedAt = CURRENT_TIMESTAMP
                                        WHERE ClusterInstanceId = @ClusterInstanceId";
                                    }

                                    AddClusterSleeveParameters(updateCmd, cluster);

                                    // Add ClusterGuid parameter for WHERE clause
                                    if (!string.IsNullOrWhiteSpace(clusterGuid))
                                    {
                                        updateCmd.Parameters.AddWithValue("@ClusterGuid", clusterGuid);
                                    }

                                    updateCmd.ExecuteNonQuery();
                                }
                            }
                            else
                            {
                                // Insert new cluster
                                using (var insertCmd = _context.Connection.CreateCommand())
                                {
                                    insertCmd.Transaction = transaction;
                                    insertCmd.CommandText = @"
                                        INSERT INTO ClusterSleeves (
                                            ClusterInstanceId, ClusterGuid, ComboId, FilterId, Category,
                                            BoundingBoxMinX, BoundingBoxMinY, BoundingBoxMinZ,
                                            BoundingBoxMaxX, BoundingBoxMaxY, BoundingBoxMaxZ,
                                            ClusterWidth, ClusterHeight, ClusterDepth,
                                            RotationAngleDeg, IsRotated,
                                            PlacementX, PlacementY, PlacementZ,
                                            HostType, HostOrientation, ClashZoneIdsJson,
                                            SleeveFamilyName,
                                            ClashZoneGuids, MepSizes, MepSystemNames, MepElementIds,
                                            Corner1X, Corner1Y, Corner1Z,
                                            Corner2X, Corner2Y, Corner2Z,
                                            Corner3X, Corner3Y, Corner3Z,
                                            Corner4X, Corner4Y, Corner4Z,
                                            CreatedAt, UpdatedAt
                                        ) VALUES (
                                            @ClusterInstanceId, @ClusterGuid, @ComboId, @FilterId, @Category,
                                            @BoundingBoxMinX, @BoundingBoxMinY, @BoundingBoxMinZ,
                                            @BoundingBoxMaxX, @BoundingBoxMaxY, @BoundingBoxMaxZ,
                                            @ClusterWidth, @ClusterHeight, @ClusterDepth,
                                            @RotationAngleDeg, @IsRotated,
                                            @PlacementX, @PlacementY, @PlacementZ,
                                            @HostType, @HostOrientation, @ClashZoneIdsJson,
                                            @SleeveFamilyName,
                                            @ClashZoneGuids, @MepSizes, @MepSystemNames, @MepElementIds,
                                            @Corner1X, @Corner1Y, @Corner1Z,
                                            @Corner2X, @Corner2Y, @Corner2Z,
                                            @Corner3X, @Corner3Y, @Corner3Z,
                                            @Corner4X, @Corner4Y, @Corner4Z,
                                            CURRENT_TIMESTAMP, CURRENT_TIMESTAMP
                                        )";

                                    AddClusterSleeveParameters(insertCmd, cluster);
                                    insertCmd.ExecuteNonQuery();
                                }
                            }
                        }
                    }

                    transaction.Commit();
                    DatabaseOperationLogger.LogTransaction("COMMIT", "SUCCESS", $"Batch saved {clusters.Count} clusters");
                    _logger($"[SQLite] ✅ Batch saved {clusters.Count} cluster sleeves in single transaction");
                }
                catch (Exception ex)
                {
                    DatabaseOperationLogger.LogTransaction("ROLLBACK", "FAILED", ex.Message);
                    transaction.Rollback();
                    _logger($"[SQLite] ❌ Error batch saving cluster sleeves: {ex.Message}");
                    throw;
                }
            }
        }

        /// <summary>
        /// ✅ SIMPLIFIED: Save clusters using INSERT OR REPLACE based on ClusterInstanceId (PRIMARY KEY).
        /// This eliminates complex bulk insert/update logic and parameter misalignment issues.
        /// </summary>
        private void BatchSaveClusterSleevesBulk(List<ClusterSaveData> clusters)
        {
            if (clusters == null || clusters.Count == 0)
                return;

            var sw = System.Diagnostics.Stopwatch.StartNew();

            using (var transaction = _context.Connection.BeginTransaction())
            {
                try
                {
                    int savedCount = 0;

                    foreach (var cluster in clusters)
                    {
                        using (var cmd = _context.Connection.CreateCommand())
                        {
                            cmd.Transaction = transaction;

                            // Generate deterministic ClusterGuid
                            var clusterGuid = GenerateDeterministicClusterGuid(cluster.ClashZoneIds);

                            // Get comma-separated MEP data
                            var (clashZoneGuids, mepSizes, mepSystemNames, mepElementIds) = GetCommaSeparatedMepData(cluster.ClashZoneIds);

                            // Serialize ClashZoneIds to JSON
                            var clashZoneIdsJson = cluster.ClashZoneIds != null && cluster.ClashZoneIds.Count > 0
                                ? JsonSerializer.Serialize(cluster.ClashZoneIds.Select(g => g.ToString()).ToList())
                                : "[]";

                            // ✅ SIMPLE: INSERT OR REPLACE based on ClusterInstanceId (PRIMARY KEY)
                            cmd.CommandText = @"
                                INSERT OR REPLACE INTO ClusterSleeves (
                                    ClusterInstanceId, ClusterGuid, ComboId, FilterId, Category,
                                    BoundingBoxMinX, BoundingBoxMinY, BoundingBoxMinZ,
                                    BoundingBoxMaxX, BoundingBoxMaxY, BoundingBoxMaxZ,
                                    ClusterWidth, ClusterHeight, ClusterDepth,
                                    RotationAngleDeg, IsRotated,
                                    PlacementX, PlacementY, PlacementZ,
                                    HostType, HostOrientation, ClashZoneIdsJson,
                                    SleeveFamilyName,
                                    ClashZoneGuids, MepSizes, MepSystemNames, MepElementIds,
                                    Corner1X, Corner1Y, Corner1Z,
                                    Corner2X, Corner2Y, Corner2Z,
                                    Corner3X, Corner3Y, Corner3Z,
                                    Corner4X, Corner4Y, Corner4Z,
                                    CreatedAt, UpdatedAt
                                ) VALUES (
                                    @ClusterInstanceId, @ClusterGuid, @ComboId, @FilterId, @Category,
                                    @BoundingBoxMinX, @BoundingBoxMinY, @BoundingBoxMinZ,
                                    @BoundingBoxMaxX, @BoundingBoxMaxY, @BoundingBoxMaxZ,
                                    @ClusterWidth, @ClusterHeight, @ClusterDepth,
                                    @RotationAngleDeg, @IsRotated,
                                    @PlacementX, @PlacementY, @PlacementZ,
                                    @HostType, @HostOrientation, @ClashZoneIdsJson,
                                    @SleeveFamilyName,
                                    @ClashZoneGuids, @MepSizes, @MepSystemNames, @MepElementIds,
                                    @Corner1X, @Corner1Y, @Corner1Z,
                                    @Corner2X, @Corner2Y, @Corner2Z,
                                    @Corner3X, @Corner3Y, @Corner3Z,
                                    @Corner4X, @Corner4Y, @Corner4Z,
                                    COALESCE((SELECT CreatedAt FROM ClusterSleeves WHERE ClusterInstanceId = @ClusterInstanceId), CURRENT_TIMESTAMP),
                                    CURRENT_TIMESTAMP
                                )";

                            // Add parameters in exact order
                            cmd.Parameters.AddWithValue("@ClusterInstanceId", cluster.ClusterInstanceId);
                            cmd.Parameters.AddWithValue("@ClusterGuid", string.IsNullOrWhiteSpace(clusterGuid) ? (object)DBNull.Value : clusterGuid);
                            cmd.Parameters.AddWithValue("@ComboId", cluster.ComboId);
                            cmd.Parameters.AddWithValue("@FilterId", cluster.FilterId);
                            cmd.Parameters.AddWithValue("@Category", cluster.Category ?? string.Empty);
                            cmd.Parameters.AddWithValue("@BoundingBoxMinX", cluster.BoundingBoxMinX);
                            cmd.Parameters.AddWithValue("@BoundingBoxMinY", cluster.BoundingBoxMinY);
                            cmd.Parameters.AddWithValue("@BoundingBoxMinZ", cluster.BoundingBoxMinZ);
                            cmd.Parameters.AddWithValue("@BoundingBoxMaxX", cluster.BoundingBoxMaxX);
                            cmd.Parameters.AddWithValue("@BoundingBoxMaxY", cluster.BoundingBoxMaxY);
                            cmd.Parameters.AddWithValue("@BoundingBoxMaxZ", cluster.BoundingBoxMaxZ);
                            cmd.Parameters.AddWithValue("@ClusterWidth", cluster.ClusterWidth);
                            cmd.Parameters.AddWithValue("@ClusterHeight", cluster.ClusterHeight);
                            cmd.Parameters.AddWithValue("@ClusterDepth", cluster.ClusterDepth);
                            cmd.Parameters.AddWithValue("@RotationAngleDeg", cluster.RotationAngleDeg);
                            cmd.Parameters.AddWithValue("@IsRotated", cluster.IsRotated ? 1 : 0);
                            cmd.Parameters.AddWithValue("@PlacementX", cluster.PlacementX);
                            cmd.Parameters.AddWithValue("@PlacementY", cluster.PlacementY);
                            cmd.Parameters.AddWithValue("@PlacementZ", cluster.PlacementZ);
                            cmd.Parameters.AddWithValue("@HostType", cluster.HostType ?? (object)DBNull.Value);
                            cmd.Parameters.AddWithValue("@HostOrientation", cluster.HostOrientation ?? (object)DBNull.Value);
                            cmd.Parameters.AddWithValue("@ClashZoneIdsJson", clashZoneIdsJson);
                            cmd.Parameters.AddWithValue("@SleeveFamilyName", cluster.SleeveFamilyName ?? (object)DBNull.Value);
                            cmd.Parameters.AddWithValue("@ClashZoneGuids", clashZoneGuids ?? (object)DBNull.Value);
                            cmd.Parameters.AddWithValue("@MepSizes", mepSizes ?? (object)DBNull.Value);
                            cmd.Parameters.AddWithValue("@MepSystemNames", mepSystemNames ?? (object)DBNull.Value);
                            cmd.Parameters.AddWithValue("@MepElementIds", mepElementIds ?? (object)DBNull.Value);

                            // ✅ CORNERS: Add in exact order
                            cmd.Parameters.AddWithValue("@Corner1X", cluster.Corner1X);
                            cmd.Parameters.AddWithValue("@Corner1Y", cluster.Corner1Y);
                            cmd.Parameters.AddWithValue("@Corner1Z", cluster.Corner1Z);
                            cmd.Parameters.AddWithValue("@Corner2X", cluster.Corner2X);
                            cmd.Parameters.AddWithValue("@Corner2Y", cluster.Corner2Y);
                            cmd.Parameters.AddWithValue("@Corner2Z", cluster.Corner2Z);
                            cmd.Parameters.AddWithValue("@Corner3X", cluster.Corner3X);
                            cmd.Parameters.AddWithValue("@Corner3Y", cluster.Corner3Y);
                            cmd.Parameters.AddWithValue("@Corner3Z", cluster.Corner3Z);
                            cmd.Parameters.AddWithValue("@Corner4X", cluster.Corner4X);
                            cmd.Parameters.AddWithValue("@Corner4Y", cluster.Corner4Y);
                            cmd.Parameters.AddWithValue("@Corner4Z", cluster.Corner4Z);

                            // 🔥 DEBUG: Log before execution
                            try
                            {
                                var versionTag = "R2023"; // Hardcoded for simplicity
                                var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                                var logDir = Path.Combine(appData, "JSE_MEP_Openings", "Logs", versionTag);
                                var logPath = Path.Combine(logDir, "cluster_debug.log");
                                File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] 💾 EXECUTING INSERT OR REPLACE for ClusterInstanceId={cluster.ClusterInstanceId}, Corner1X={cluster.Corner1X:F2}, Corner1Y={cluster.Corner1Y:F2}\n");
                            }
                            catch { }

                            int rowsAffected = cmd.ExecuteNonQuery();

                            // 🔥 DEBUG: Log after execution
                            try
                            {
                                var versionTag = "R2023";
                                var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                                var logDir = Path.Combine(appData, "JSE_MEP_Openings", "Logs", versionTag);
                                var logPath = Path.Combine(logDir, "cluster_debug.log");
                                File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] ✅ INSERT OR REPLACE COMPLETED: rowsAffected={rowsAffected}\n");
                            }
                            catch { }

                            DatabaseOperationLogger.LogOperation("INSERT OR REPLACE", "ClusterSleeves",
                                new Dictionary<string, object>
                                {
                                    { "ClusterInstanceId", cluster.ClusterInstanceId },
                                    { "Corner1X", cluster.Corner1X },
                                    { "Corner1Y", cluster.Corner1Y },
                                    { "Corner1Z", cluster.Corner1Z }
                                });

                            savedCount++;
                        }
                    }

                    transaction.Commit();
                    sw.Stop();

                    DatabaseOperationLogger.LogTransaction("COMMIT", "SUCCESS", $"Saved {savedCount} clusters using INSERT OR REPLACE");
                    _logger($"[SQLite] ✅ Saved {savedCount} cluster sleeves in {sw.ElapsedMilliseconds}ms using INSERT OR REPLACE");
                }
                catch (Exception ex)
                {
                    DatabaseOperationLogger.LogTransaction("ROLLBACK", "FAILED", ex.Message);
                    transaction.Rollback();
                    _logger($"[SQLite] ❌ Error saving cluster sleeves: {ex.Message}");
                    throw;
                }
            }
        }

        /// <summary>
        /// ✅ BULK INSERT: Insert multiple clusters in a single operation using parameterized VALUES.
        /// </summary>
        private void BulkInsertClusters(List<ClusterSaveData> clusters, SQLiteTransaction transaction)
        {
            if (clusters == null || clusters.Count == 0)
                return;

            using (var insertCmd = _context.Connection.CreateCommand())
            {
                insertCmd.Transaction = transaction;

                // Build bulk INSERT with VALUES clause
                var valuesClauses = new List<string>();
                for (int i = 0; i < clusters.Count; i++)
                {
                    var cluster = clusters[i];
                    var clusterGuid = GenerateDeterministicClusterGuid(cluster.ClashZoneIds);

                    valuesClauses.Add($@"
                        (@ClusterInstanceId{i}, @ClusterGuid{i}, @ComboId{i}, @FilterId{i}, @Category{i},
                         @BoundingBoxMinX{i}, @BoundingBoxMinY{i}, @BoundingBoxMinZ{i},
                         @BoundingBoxMaxX{i}, @BoundingBoxMaxY{i}, @BoundingBoxMaxZ{i},
                         @ClusterWidth{i}, @ClusterHeight{i}, @ClusterDepth{i},
                         @RotationAngleDeg{i}, @IsRotated{i},
                         @PlacementX{i}, @PlacementY{i}, @PlacementZ{i},
                         @HostType{i}, @HostOrientation{i}, @ClashZoneIdsJson{i},
                         @ClashZoneGuids{i}, @MepSizes{i}, @MepSystemNames{i}, @MepElementIds{i},
                         @Corner1X{i}, @Corner1Y{i}, @Corner1Z{i},
                         @Corner2X{i}, @Corner2Y{i}, @Corner2Z{i},
                         @Corner3X{i}, @Corner3Y{i}, @Corner3Z{i},
                         @Corner4X{i}, @Corner4Y{i}, @Corner4Z{i},
                         CURRENT_TIMESTAMP, CURRENT_TIMESTAMP)");

                    // Add parameters for this cluster
                    AddClusterSleeveParametersBulk(insertCmd, cluster, i, clusterGuid);
                }

                insertCmd.CommandText = $@"
                    INSERT INTO ClusterSleeves (
                        ClusterInstanceId, ClusterGuid, ComboId, FilterId, Category,
                        BoundingBoxMinX, BoundingBoxMinY, BoundingBoxMinZ,
                        BoundingBoxMaxX, BoundingBoxMaxY, BoundingBoxMaxZ,
                        ClusterWidth, ClusterHeight, ClusterDepth,
                        RotationAngleDeg, IsRotated,
                        PlacementX, PlacementY, PlacementZ,
                        HostType, HostOrientation, ClashZoneIdsJson,
                        ClashZoneGuids, MepSizes, MepSystemNames, MepElementIds,
                        Corner1X, Corner1Y, Corner1Z,
                        Corner2X, Corner2Y, Corner2Z,
                        Corner3X, Corner3Y, Corner3Z,
                        Corner4X, Corner4Y, Corner4Z,
                        CreatedAt, UpdatedAt
                    ) VALUES {string.Join(",", valuesClauses)}";

                int rowsAffected = insertCmd.ExecuteNonQuery();
                _logger($"[SQLite] ✅ Bulk INSERT: {rowsAffected} rows affected for {clusters.Count} clusters");
            }
        }

        /// <summary>
        /// ✅ BULK UPDATE: Update multiple clusters using CASE statements (similar to BulkUpdateClashZones).
        /// </summary>
        private void BulkUpdateClusters(List<ClusterSaveData> clusters, Dictionary<ClusterSaveData, string> clusterGuidMap, SQLiteTransaction transaction)
        {
            if (clusters == null || clusters.Count == 0)
                return;

            using (var updateCmd = _context.Connection.CreateCommand())
            {
                updateCmd.Transaction = transaction;

                // Build bulk UPDATE with CASE statements
                var sql = new System.Text.StringBuilder();
                sql.AppendLine("UPDATE ClusterSleeves SET");
                sql.AppendLine("  UpdatedAt = CURRENT_TIMESTAMP,");

                // Build CASE statements for each field
                var fields = new[]
                {
                    ("ClusterInstanceId", "int"),
                    ("ComboId", "int"),
                    ("FilterId", "int"),
                    ("Category", "string"),
                    ("BoundingBoxMinX", "double"),
                    ("BoundingBoxMinY", "double"),
                    ("BoundingBoxMinZ", "double"),
                    ("BoundingBoxMaxX", "double"),
                    ("BoundingBoxMaxY", "double"),
                    ("BoundingBoxMaxZ", "double"),
                    ("ClusterWidth", "double"),
                    ("ClusterHeight", "double"),
                    ("ClusterDepth", "double"),
                    ("RotationAngleDeg", "double"),
                    ("IsRotated", "int"),
                    ("PlacementX", "double"),
                    ("PlacementY", "double"),
                    ("PlacementZ", "double"),
                    ("HostType", "string"),
                    ("HostOrientation", "string"),
                    ("ClashZoneIdsJson", "string"),
                    ("ClashZoneGuids", "string"),
                    ("MepSizes", "string"),
                    ("MepSystemNames", "string"),
                    ("MepElementIds", "string"),
                    ("Corner1X", "double"), ("Corner1Y", "double"), ("Corner1Z", "double"),
                    ("Corner2X", "double"), ("Corner2Y", "double"), ("Corner2Z", "double"),
                    ("Corner3X", "double"), ("Corner3Y", "double"), ("Corner3Z", "double"),
                    ("Corner4X", "double"), ("Corner4Y", "double"), ("Corner4Z", "double")
                };

                for (int f = 0; f < fields.Length; f++)
                {
                    var (fieldName, fieldType) = fields[f];
                    sql.Append($"  {fieldName} = CASE ClusterGuid");

                    for (int i = 0; i < clusters.Count; i++)
                    {
                        var cluster = clusters[i];
                        var clusterGuid = clusterGuidMap.ContainsKey(cluster)
                            ? clusterGuidMap[cluster]
                            : GenerateDeterministicClusterGuid(cluster.ClashZoneIds);

                        if (string.IsNullOrWhiteSpace(clusterGuid))
                            continue; // Skip clusters without GUID

                        sql.AppendLine();
                        sql.Append($"    WHEN @ClusterGuid{i} THEN @{fieldName}{i}");

                        // Add parameter value
                        object paramValue = GetFieldValue(cluster, fieldName, fieldType);
                        updateCmd.Parameters.AddWithValue($"@{fieldName}{i}", paramValue ?? DBNull.Value);
                        updateCmd.Parameters.AddWithValue($"@ClusterGuid{i}", clusterGuid);
                    }

                    sql.AppendLine();
                    sql.Append("    ELSE ").Append(fieldName); // Keep existing value if not matched
                    sql.AppendLine("  END");

                    if (f < fields.Length - 1)
                        sql.AppendLine(",");
                }

                // WHERE clause: Update only clusters that match our GUIDs
                var guidParams = new List<string>();
                for (int i = 0; i < clusters.Count; i++)
                {
                    var cluster = clusters[i];
                    var clusterGuid = clusterGuidMap.ContainsKey(cluster)
                        ? clusterGuidMap[cluster]
                        : GenerateDeterministicClusterGuid(cluster.ClashZoneIds);

                    if (!string.IsNullOrWhiteSpace(clusterGuid))
                    {
                        guidParams.Add($"@WhereGuid{i}");
                        updateCmd.Parameters.AddWithValue($"@WhereGuid{i}", clusterGuid);
                    }
                }

                if (guidParams.Count > 0)
                {
                    sql.AppendLine($"WHERE ClusterGuid IN ({string.Join(", ", guidParams)})");
                }
                else
                {
                    // Fallback: Use ClusterInstanceId if no GUIDs
                    var instanceIdParams = new List<string>();
                    for (int i = 0; i < clusters.Count; i++)
                    {
                        instanceIdParams.Add($"@WhereInstanceId{i}");
                        updateCmd.Parameters.AddWithValue($"@WhereInstanceId{i}", clusters[i].ClusterInstanceId);
                    }
                    sql.AppendLine($"WHERE ClusterInstanceId IN ({string.Join(", ", instanceIdParams)})");
                }

                updateCmd.CommandText = sql.ToString();
                int rowsAffected = updateCmd.ExecuteNonQuery();
                _logger($"[SQLite] ✅ Bulk UPDATE: {rowsAffected} rows affected for {clusters.Count} clusters");
            }
        }

        /// <summary>
        /// Helper: Get field value from ClusterSaveData by field name.
        /// </summary>
        private object GetFieldValue(ClusterSaveData cluster, string fieldName, string fieldType)
        {
            switch (fieldName)
            {
                case "ClusterInstanceId": return cluster.ClusterInstanceId;
                case "ComboId": return cluster.ComboId;
                case "FilterId": return cluster.FilterId;
                case "Category": return cluster.Category ?? string.Empty;
                case "BoundingBoxMinX": return cluster.BoundingBoxMinX;
                case "BoundingBoxMinY": return cluster.BoundingBoxMinY;
                case "BoundingBoxMinZ": return cluster.BoundingBoxMinZ;
                case "BoundingBoxMaxX": return cluster.BoundingBoxMaxX;
                case "BoundingBoxMaxY": return cluster.BoundingBoxMaxY;
                case "BoundingBoxMaxZ": return cluster.BoundingBoxMaxZ;
                case "ClusterWidth": return cluster.ClusterWidth;
                case "ClusterHeight": return cluster.ClusterHeight;
                case "ClusterDepth": return cluster.ClusterDepth;
                case "RotationAngleDeg": return cluster.RotationAngleDeg;
                case "IsRotated": return cluster.IsRotated ? 1 : 0;
                case "PlacementX": return cluster.PlacementX;
                case "PlacementY": return cluster.PlacementY;
                case "PlacementZ": return cluster.PlacementZ;
                case "HostType": return cluster.HostType ?? (object)DBNull.Value;
                case "HostOrientation": return cluster.HostOrientation ?? (object)DBNull.Value;
                case "ClashZoneIdsJson":
                    return cluster.ClashZoneIds != null && cluster.ClashZoneIds.Count > 0
                        ? JsonSerializer.Serialize(cluster.ClashZoneIds.Select(g => g.ToString()).ToList())
                        : "[]";
                case "ClashZoneGuids":
                case "MepSizes":
                case "MepSystemNames":
                case "MepElementIds":
                    var (guids, sizes, names, ids) = GetCommaSeparatedMepData(cluster.ClashZoneIds);
                    switch (fieldName)
                    {
                        case "ClashZoneGuids": return guids ?? (object)DBNull.Value;
                        case "MepSizes": return sizes ?? (object)DBNull.Value;
                        case "MepSystemNames": return names ?? (object)DBNull.Value;
                        case "MepElementIds": return ids ?? (object)DBNull.Value;
                    }
                    break;
                case "Corner1X": return cluster.Corner1X;
                case "Corner1Y": return cluster.Corner1Y;
                case "Corner1Z": return cluster.Corner1Z;
                case "Corner2X": return cluster.Corner2X;
                case "Corner2Y": return cluster.Corner2Y;
                case "Corner2Z": return cluster.Corner2Z;
                case "Corner3X": return cluster.Corner3X;
                case "Corner3Y": return cluster.Corner3Y;
                case "Corner3Z": return cluster.Corner3Z;
                case "Corner4X": return cluster.Corner4X;
                case "Corner4Y": return cluster.Corner4Y;
                case "Corner4Z": return cluster.Corner4Z;
            }
            return DBNull.Value;
        }

        /// <summary>
        /// Helper: Add parameters for bulk INSERT operation.
        /// </summary>
        private void AddClusterSleeveParametersBulk(SQLiteCommand cmd, ClusterSaveData cluster, int index, string clusterGuid)
        {
            cmd.Parameters.AddWithValue($"@ClusterInstanceId{index}", cluster.ClusterInstanceId);
            cmd.Parameters.AddWithValue($"@ClusterGuid{index}", string.IsNullOrWhiteSpace(clusterGuid) ? (object)DBNull.Value : clusterGuid);
            cmd.Parameters.AddWithValue($"@ComboId{index}", cluster.ComboId);
            cmd.Parameters.AddWithValue($"@FilterId{index}", cluster.FilterId);
            cmd.Parameters.AddWithValue($"@Category{index}", cluster.Category ?? string.Empty);
            cmd.Parameters.AddWithValue($"@BoundingBoxMinX{index}", cluster.BoundingBoxMinX);
            cmd.Parameters.AddWithValue($"@BoundingBoxMinY{index}", cluster.BoundingBoxMinY);
            cmd.Parameters.AddWithValue($"@BoundingBoxMinZ{index}", cluster.BoundingBoxMinZ);
            cmd.Parameters.AddWithValue($"@BoundingBoxMaxX{index}", cluster.BoundingBoxMaxX);
            cmd.Parameters.AddWithValue($"@BoundingBoxMaxY{index}", cluster.BoundingBoxMaxY);
            cmd.Parameters.AddWithValue($"@BoundingBoxMaxZ{index}", cluster.BoundingBoxMaxZ);
            cmd.Parameters.AddWithValue($"@ClusterWidth{index}", cluster.ClusterWidth);
            cmd.Parameters.AddWithValue($"@ClusterHeight{index}", cluster.ClusterHeight);
            cmd.Parameters.AddWithValue($"@ClusterDepth{index}", cluster.ClusterDepth);
            cmd.Parameters.AddWithValue($"@RotationAngleDeg{index}", cluster.RotationAngleDeg);
            cmd.Parameters.AddWithValue($"@IsRotated{index}", cluster.IsRotated ? 1 : 0);
            cmd.Parameters.AddWithValue($"@PlacementX{index}", cluster.PlacementX);
            cmd.Parameters.AddWithValue($"@PlacementY{index}", cluster.PlacementY);
            cmd.Parameters.AddWithValue($"@PlacementZ{index}", cluster.PlacementZ);
            cmd.Parameters.AddWithValue($"@HostType{index}", cluster.HostType ?? (object)DBNull.Value);
            cmd.Parameters.AddWithValue($"@HostOrientation{index}", cluster.HostOrientation ?? (object)DBNull.Value);

            // Serialize ClashZoneIds to JSON
            var clashZoneIdsJson = cluster.ClashZoneIds != null && cluster.ClashZoneIds.Count > 0
                ? JsonSerializer.Serialize(cluster.ClashZoneIds.Select(g => g.ToString()).ToList())
                : "[]";
            cmd.Parameters.AddWithValue($"@ClashZoneIdsJson{index}", clashZoneIdsJson);

            // ✅ COMMA-SEPARATED VALUES: Load MEP data from SleeveSnapshots table
            var (clashZoneGuids, mepSizes, mepSystemNames, mepElementIds) = GetCommaSeparatedMepData(cluster.ClashZoneIds);
            cmd.Parameters.AddWithValue($"@ClashZoneGuids{index}", clashZoneGuids ?? (object)DBNull.Value);
            cmd.Parameters.AddWithValue($"@MepSizes{index}", mepSizes ?? (object)DBNull.Value);
            cmd.Parameters.AddWithValue($"@MepSystemNames{index}", mepSystemNames ?? (object)DBNull.Value);
            cmd.Parameters.AddWithValue($"@MepElementIds{index}", mepElementIds ?? (object)DBNull.Value);

            // ✅ Corners
            cmd.Parameters.AddWithValue($"@Corner1X{index}", cluster.Corner1X);
            cmd.Parameters.AddWithValue($"@Corner1Y{index}", cluster.Corner1Y);
            cmd.Parameters.AddWithValue($"@Corner1Z{index}", cluster.Corner1Z);
            cmd.Parameters.AddWithValue($"@Corner2X{index}", cluster.Corner2X);
            cmd.Parameters.AddWithValue($"@Corner2Y{index}", cluster.Corner2Y);
            cmd.Parameters.AddWithValue($"@Corner2Z{index}", cluster.Corner2Z);
            cmd.Parameters.AddWithValue($"@Corner3X{index}", cluster.Corner3X);
            cmd.Parameters.AddWithValue($"@Corner3Y{index}", cluster.Corner3Y);
            cmd.Parameters.AddWithValue($"@Corner3Z{index}", cluster.Corner3Z);
            cmd.Parameters.AddWithValue($"@Corner4X{index}", cluster.Corner4X);
            cmd.Parameters.AddWithValue($"@Corner4Y{index}", cluster.Corner4Y);
            cmd.Parameters.AddWithValue($"@Corner4Z{index}", cluster.Corner4Z);
        }

        private void AddClusterSleeveParameters(SQLiteCommand cmd, ClusterSaveData cluster)
        {
            // Generate deterministic ClusterGuid from ClashZoneIds
            string clusterGuid = GenerateDeterministicClusterGuid(cluster.ClashZoneIds);

            AddClusterSleeveParameters(
                cmd,
                cluster.ClusterInstanceId,
                cluster.ComboId,
                cluster.FilterId,
                cluster.Category,
                cluster.BoundingBoxMinX, cluster.BoundingBoxMinY, cluster.BoundingBoxMinZ,
                cluster.BoundingBoxMaxX, cluster.BoundingBoxMaxY, cluster.BoundingBoxMaxZ,
                cluster.ClusterWidth, cluster.ClusterHeight, cluster.ClusterDepth,
                cluster.RotationAngleDeg,
                cluster.IsRotated,
                cluster.PlacementX, cluster.PlacementY, cluster.PlacementZ,
                cluster.HostType,
                cluster.HostOrientation,
                cluster.ClashZoneIds,
                cluster.Corner1X, cluster.Corner1Y, cluster.Corner1Z,
                cluster.Corner2X, cluster.Corner2Y, cluster.Corner2Z,
                cluster.Corner3X, cluster.Corner3Y, cluster.Corner3Z,
                cluster.Corner4X, cluster.Corner4Y, cluster.Corner4Z,
                clusterGuid);
        }

        /// <summary>
        /// ✅ CRITICAL FIX: Generate deterministic ClusterGuid from sorted ClashZoneIds
        /// This allows proper upserting even if ClusterInstanceId changes (like ClashZoneGuid for snapshots)
        /// </summary>
        private string GenerateDeterministicClusterGuid(List<Guid> clashZoneIds)
        {
            if (clashZoneIds == null || clashZoneIds.Count == 0)
                return null;

            // Sort GUIDs to ensure deterministic result
            var sortedGuids = clashZoneIds.OrderBy(g => g.ToString()).ToList();

            // Create a deterministic string from sorted GUIDs
            var guidString = string.Join("|", sortedGuids.Select(g => g.ToString().ToUpperInvariant()));

            // Generate a deterministic GUID from the string using MD5 (like ClashZoneGuid)
            using (var md5 = System.Security.Cryptography.MD5.Create())
            {
                var hash = md5.ComputeHash(System.Text.Encoding.UTF8.GetBytes(guidString));
                var guid = new Guid(hash);
                return guid.ToString().ToUpperInvariant();
            }
        }

        private void AddClusterSleeveParameters(
            SQLiteCommand cmd,
            int clusterInstanceId,
            int comboId,
            int filterId,
            string category,
            double boundingBoxMinX, double boundingBoxMinY, double boundingBoxMinZ,
            double boundingBoxMaxX, double boundingBoxMaxY, double boundingBoxMaxZ,
            double clusterWidth, double clusterHeight, double clusterDepth,
            double rotationAngleDeg,
            bool isRotated,
            double placementX, double placementY, double placementZ,
            string hostType,
            string hostOrientation,
            List<Guid> clashZoneIds,
            double corner1X, double corner1Y, double corner1Z,
            double corner2X, double corner2Y, double corner2Z,
            double corner3X, double corner3Y, double corner3Z,
            double corner4X, double corner4Y, double corner4Z,
            string sleeveFamilyName = null,
            string clusterGuid = null)
        {
            cmd.Parameters.AddWithValue("@ClusterInstanceId", clusterInstanceId);
            cmd.Parameters.AddWithValue("@ClusterGuid", string.IsNullOrWhiteSpace(clusterGuid) ? (object)DBNull.Value : clusterGuid);
            cmd.Parameters.AddWithValue("@ComboId", comboId);
            cmd.Parameters.AddWithValue("@FilterId", filterId);
            cmd.Parameters.AddWithValue("@Category", category ?? string.Empty);
            cmd.Parameters.AddWithValue("@BoundingBoxMinX", boundingBoxMinX);
            cmd.Parameters.AddWithValue("@BoundingBoxMinY", boundingBoxMinY);
            cmd.Parameters.AddWithValue("@BoundingBoxMinZ", boundingBoxMinZ);
            cmd.Parameters.AddWithValue("@BoundingBoxMaxX", boundingBoxMaxX);
            cmd.Parameters.AddWithValue("@BoundingBoxMaxY", boundingBoxMaxY);
            cmd.Parameters.AddWithValue("@BoundingBoxMaxZ", boundingBoxMaxZ);
            cmd.Parameters.AddWithValue("@ClusterWidth", clusterWidth);
            cmd.Parameters.AddWithValue("@ClusterHeight", clusterHeight);
            cmd.Parameters.AddWithValue("@ClusterDepth", clusterDepth);
            cmd.Parameters.AddWithValue("@RotationAngleDeg", rotationAngleDeg);
            cmd.Parameters.AddWithValue("@IsRotated", isRotated ? 1 : 0);
            cmd.Parameters.AddWithValue("@PlacementX", placementX);
            cmd.Parameters.AddWithValue("@PlacementY", placementY);
            cmd.Parameters.AddWithValue("@PlacementZ", placementZ);
            cmd.Parameters.AddWithValue("@HostType", hostType ?? (object)DBNull.Value);
            cmd.Parameters.AddWithValue("@HostOrientation", hostOrientation ?? (object)DBNull.Value);
            cmd.Parameters.AddWithValue("@SleeveFamilyName", sleeveFamilyName ?? (object)DBNull.Value);
            
            // Serialize ClashZoneIds to JSON
            var clashZoneIdsJson = clashZoneIds != null && clashZoneIds.Count > 0
                ? JsonSerializer.Serialize(clashZoneIds.Select(g => g.ToString()).ToList())
                : "[]";
            cmd.Parameters.AddWithValue("@ClashZoneIdsJson", clashZoneIdsJson);

            // ✅ COMMA-SEPARATED VALUES: Load MEP data from SleeveSnapshots table
            var (clashZoneGuids, mepSizes, mepSystemNames, mepElementIds) = GetCommaSeparatedMepData(clashZoneIds);
            cmd.Parameters.AddWithValue("@ClashZoneGuids", clashZoneGuids ?? (object)DBNull.Value);
            cmd.Parameters.AddWithValue("@MepSizes", mepSizes ?? (object)DBNull.Value);
            cmd.Parameters.AddWithValue("@MepSystemNames", mepSystemNames ?? (object)DBNull.Value);
            cmd.Parameters.AddWithValue("@MepElementIds", mepElementIds ?? (object)DBNull.Value);

            // ✅ Corners
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
        }

        /// <summary>
        /// ✅ COMMA-SEPARATED VALUES: Load MEP data from SleeveSnapshots table for cluster sleeves
        /// Returns comma-separated strings for GUIDs, sizes, system names, and element IDs
        /// ✅ CRITICAL FIX: Queries SleeveSnapshots.MepParametersJson to get Size parameter (aggregated from snapshot)
        /// Falls back to ClashZones.MepElementSizeParameterValue if snapshot not found
        /// </summary>
        private (string clashZoneGuids, string mepSizes, string mepSystemNames, string mepElementIds) GetCommaSeparatedMepData(List<Guid> clashZoneIds)
        {
            if (clashZoneIds == null || clashZoneIds.Count == 0)
                return (null, null, null, null);

            try
            {
                // ✅ CRITICAL FIX: Query SleeveSnapshots first to get Size from MepParametersJson (aggregated Size parameter)
                // This ensures cluster sleeves get the same Size parameter format as individual sleeves (from snapshot table)
                using (var cmd = _context.Connection.CreateCommand())
                {
                    // Build WHERE clause for GUIDs
                    var guidPlaceholders = string.Join(", ", clashZoneIds.Select((_, i) => $"@Guid{i}"));

                    // ✅ PRIORITY 1: Query SleeveSnapshots to get Size from MepParametersJson
                    cmd.CommandText = $@"
                        SELECT DISTINCT
                            ss.ClashZoneGuid,
                            ss.MepParametersJson,
                            ss.MepElementIdsJson,
                            cz.MepElementId,
                            cz.MepElementSizeParameterValue,
                            cz.MepElementFormattedSize,
                            cz.MepElementSizeData,
                            cz.MepElementSystemName,
                            cz.MepElementSystemAbbreviation
                        FROM ClashZones cz
                        LEFT JOIN SleeveSnapshots ss ON UPPER(ss.ClashZoneGuid) = UPPER(cz.ClashZoneGuid)
                        WHERE UPPER(cz.ClashZoneGuid) IN ({guidPlaceholders})
                          AND cz.ClashZoneGuid != '' AND cz.ClashZoneGuid IS NOT NULL
                        ORDER BY cz.ClashZoneGuid";

                    // Add GUID parameters
                    for (int i = 0; i < clashZoneIds.Count; i++)
                    {
                        cmd.Parameters.AddWithValue($"@Guid{i}", clashZoneIds[i].ToString().ToUpperInvariant());
                    }

                    var guidList = new List<string>();
                    var sizeList = new List<string>();
                    var systemNameList = new List<string>();
                    var elementIdList = new List<string>();

                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            var guid = reader.IsDBNull(0) ? null : reader.GetString(0);
                            var mepParamsJson = reader.IsDBNull(1) ? null : reader.GetString(1); // ✅ PRIORITY 1: MepParametersJson from SleeveSnapshots
                            var mepElementIdsJson = reader.IsDBNull(2) ? null : reader.GetString(2);
                            var mepElementId = reader.IsDBNull(3) ? (long?)null : reader.GetInt64(3);
                            var mepSizeParameterValue = reader.IsDBNull(4) ? null : reader.GetString(4); // ✅ FALLBACK: MepElementSizeParameterValue from ClashZones
                            var mepFormattedSize = reader.IsDBNull(5) ? null : reader.GetString(5); // ✅ FALLBACK: MepElementFormattedSize
                            var mepSizeDataJson = reader.IsDBNull(6) ? null : reader.GetString(6);
                            var systemName = reader.IsDBNull(7) ? null : reader.GetString(7);
                            var systemAbbr = reader.IsDBNull(8) ? null : reader.GetString(8);

                            if (!string.IsNullOrWhiteSpace(guid))
                                guidList.Add(guid);

                            // ✅ CRITICAL FIX: Prioritize Size from SleeveSnapshots.MepParametersJson (aggregated Size parameter)
                            // This ensures cluster sleeves get the same Size parameter format as individual sleeves
                            string sizeValue = null;

                            // Priority 1: Extract Size from SleeveSnapshots.MepParametersJson (aggregated Size parameter from snapshot)
                            if (!string.IsNullOrWhiteSpace(mepParamsJson) && mepParamsJson != "{}")
                            {
                                try
                                {
                                    var mepParams = System.Text.Json.JsonSerializer.Deserialize<Dictionary<string, string>>(mepParamsJson);
                                    if (mepParams != null && mepParams.TryGetValue("Size", out var sizeFromSnapshot))
                                    {
                                        if (!string.IsNullOrWhiteSpace(sizeFromSnapshot))
                                        {
                                            sizeValue = sizeFromSnapshot.Trim();
                                            if (!DeploymentConfiguration.DeploymentMode)
                                            {
                                                _logger($"[ClusterSleeve] ✅ Got Size='{sizeValue}' from SleeveSnapshots.MepParametersJson for ClashZoneGuid={guid}");
                                            }
                                        }
                                    }
                                    else if (!DeploymentConfiguration.DeploymentMode)
                                    {
                                        var keys = mepParams != null ? mepParams.Keys.ToList() : new List<string>();
                                        _logger($"[ClusterSleeve] ⚠️ Size parameter not found in SleeveSnapshots.MepParametersJson for ClashZoneGuid={guid}, JSON keys: {string.Join(", ", keys)}");
                                    }
                                }
                                catch (Exception jsonEx)
                                {
                                    if (!DeploymentConfiguration.DeploymentMode)
                                    {
                                        _logger($"[ClusterSleeve] ⚠️ Error parsing MepParametersJson for ClashZoneGuid={guid}: {jsonEx.Message}");
                                    }
                                }
                            }
                            else if (!DeploymentConfiguration.DeploymentMode)
                            {
                                _logger($"[ClusterSleeve] ⚠️ No SleeveSnapshots.MepParametersJson found for ClashZoneGuid={guid}, falling back to ClashZones");
                            }

                            // Priority 2: Use MepElementSizeParameterValue from ClashZones (raw Size parameter value, e.g., "20 mmø", "200 mm dia symbol")
                            if (string.IsNullOrWhiteSpace(sizeValue) && !string.IsNullOrWhiteSpace(mepSizeParameterValue))
                            {
                                sizeValue = mepSizeParameterValue.Trim();
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    _logger($"[ClusterSleeve] ✅ Got Size='{sizeValue}' from ClashZones.MepElementSizeParameterValue for ClashZoneGuid={guid}");
                                }
                            }
                            else if (string.IsNullOrWhiteSpace(sizeValue) && !DeploymentConfiguration.DeploymentMode)
                            {
                                _logger($"[ClusterSleeve] ⚠️ ClashZones.MepElementSizeParameterValue is empty for ClashZoneGuid={guid}");
                            }
                            // Priority 3: Fall back to MepElementFormattedSize (formatted size, e.g., "Ø20")
                            else if (string.IsNullOrWhiteSpace(sizeValue) && !string.IsNullOrWhiteSpace(mepFormattedSize))
                            {
                                sizeValue = mepFormattedSize.Trim();
                            }
                            // Priority 4: Fall back to JSON parsing (legacy support)
                            else if (string.IsNullOrWhiteSpace(sizeValue) && !string.IsNullOrWhiteSpace(mepSizeDataJson))
                            {
                                try
                                {
                                    var sizeData = JsonSerializer.Deserialize<Dictionary<string, object>>(mepSizeDataJson);
                                    if (sizeData != null)
                                    {
                                        if (sizeData.TryGetValue("FormattedSize", out var formattedSizeObj) && formattedSizeObj != null)
                                        {
                                            sizeValue = formattedSizeObj.ToString();
                                        }
                                        else if (sizeData.TryGetValue("Shape", out var shapeObj) && shapeObj?.ToString() == "Round")
                                        {
                                            if (sizeData.TryGetValue("Diameter", out var diamObj))
                                            {
                                                var diamMm = Convert.ToDouble(diamObj) * 304.8; // Convert feet to mm
                                                sizeValue = $"Ø{diamMm:F0}";
                                            }
                                        }
                                        else if (sizeData.TryGetValue("Width", out var widthObj) && sizeData.TryGetValue("Height", out var heightObj))
                                        {
                                            var widthMm = Convert.ToDouble(widthObj) * 304.8;
                                            var heightMm = Convert.ToDouble(heightObj) * 304.8;
                                            sizeValue = $"{widthMm:F0}×{heightMm:F0}";
                                        }
                                    }
                                }
                                catch { /* Ignore JSON parse errors */ }
                            }

                            if (!string.IsNullOrWhiteSpace(sizeValue))
                            {
                                sizeList.Add(sizeValue);
                            }

                            // Extract system name
                            if (!string.IsNullOrWhiteSpace(systemName))
                            {
                                systemNameList.Add(systemName);
                            }
                            else if (!string.IsNullOrWhiteSpace(systemAbbr))
                            {
                                systemNameList.Add(systemAbbr);
                            }

                            // Add MEP element ID
                            if (mepElementId.HasValue && mepElementId.Value > 0)
                            {
                                elementIdList.Add(mepElementId.Value.ToString());
                            }
                        }
                    }

                    return (
                        guidList.Count > 0 ? string.Join(", ", guidList) : null,
                        sizeList.Count > 0 ? string.Join(", ", sizeList) : null,
                        systemNameList.Count > 0 ? string.Join(", ", systemNameList) : null,
                        elementIdList.Count > 0 ? string.Join(", ", elementIdList) : null
                    );
                }
            }
            catch (Exception ex)
            {
                _logger($"[SQLite] ⚠️ Error loading MEP data for cluster: {ex.Message}");
                return (null, null, null, null);
            }
        }

        /// <summary>
        /// Load cluster sleeve data for a specific ComboId and Category
        /// Used by PATH 1 (Replay) to check if cluster data exists
        /// </summary>
        public List<ClusterSleeveData> LoadClusterSleevesForCombo(int comboId, string category = null)
        {
            var clusters = new List<ClusterSleeveData>();

            // ✅ LOG: SELECT operation
            var whereClause = string.IsNullOrWhiteSpace(category)
                ? $"ComboId={comboId}"
                : $"ComboId={comboId} AND Category='{category}'";

            DatabaseOperationLogger.LogSelect(
                "ClusterSleeves",
                whereClause,
                additionalInfo: "Loading cluster data for PATH 1 replay");

            using (var cmd = _context.Connection.CreateCommand())
            {
                if (string.IsNullOrWhiteSpace(category))
                {
                    cmd.CommandText = @"
                        SELECT ClusterInstanceId, ComboId, FilterId, Category,
                               BoundingBoxMinX, BoundingBoxMinY, BoundingBoxMinZ,
                               BoundingBoxMaxX, BoundingBoxMaxY, BoundingBoxMaxZ,
                               ClusterWidth, ClusterHeight, ClusterDepth,
                               RotationAngleDeg, IsRotated,
                               PlacementX, PlacementY, PlacementZ,
                               HostType, HostOrientation, ClashZoneIdsJson
                        FROM ClusterSleeves
                        WHERE ComboId = @ComboId";
                    cmd.Parameters.AddWithValue("@ComboId", comboId);
                }
                else
                {
                    cmd.CommandText = @"
                        SELECT ClusterInstanceId, ComboId, FilterId, Category,
                               BoundingBoxMinX, BoundingBoxMinY, BoundingBoxMinZ,
                               BoundingBoxMaxX, BoundingBoxMaxY, BoundingBoxMaxZ,
                               ClusterWidth, ClusterHeight, ClusterDepth,
                               RotationAngleDeg, IsRotated,
                               PlacementX, PlacementY, PlacementZ,
                               HostType, HostOrientation, ClashZoneIdsJson
                        FROM ClusterSleeves
                        WHERE ComboId = @ComboId AND Category = @Category";
                    cmd.Parameters.AddWithValue("@ComboId", comboId);
                    cmd.Parameters.AddWithValue("@Category", category);
                }

                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        var cluster = new ClusterSleeveData
                        {
                            ClusterInstanceId = GetInt(reader, "ClusterInstanceId", -1),
                            ComboId = GetInt(reader, "ComboId", -1),
                            FilterId = GetInt(reader, "FilterId", -1),
                            Category = GetString(reader, "Category"),
                            BoundingBoxMinX = GetDouble(reader, "BoundingBoxMinX", 0.0),
                            BoundingBoxMinY = GetDouble(reader, "BoundingBoxMinY", 0.0),
                            BoundingBoxMinZ = GetDouble(reader, "BoundingBoxMinZ", 0.0),
                            BoundingBoxMaxX = GetDouble(reader, "BoundingBoxMaxX", 0.0),
                            BoundingBoxMaxY = GetDouble(reader, "BoundingBoxMaxY", 0.0),
                            BoundingBoxMaxZ = GetDouble(reader, "BoundingBoxMaxZ", 0.0),
                            ClusterWidth = GetDouble(reader, "ClusterWidth", 0.0),
                            ClusterHeight = GetDouble(reader, "ClusterHeight", 0.0),
                            ClusterDepth = GetDouble(reader, "ClusterDepth", 0.0),
                            RotationAngleDeg = GetDouble(reader, "RotationAngleDeg", 0.0),
                            IsRotated = GetBool(reader, "IsRotated"),
                            PlacementX = GetDouble(reader, "PlacementX", 0.0),
                            PlacementY = GetDouble(reader, "PlacementY", 0.0),
                            PlacementZ = GetDouble(reader, "PlacementZ", 0.0),
                            HostType = GetString(reader, "HostType"),
                            HostOrientation = GetString(reader, "HostOrientation")
                        };

                        // Deserialize ClashZoneIds from JSON
                        var clashZoneIdsJson = GetString(reader, "ClashZoneIdsJson");
                        if (!string.IsNullOrWhiteSpace(clashZoneIdsJson))
                        {
                            try
                            {
                                var guidStrings = JsonSerializer.Deserialize<List<string>>(clashZoneIdsJson);
                                cluster.ClashZoneIds = guidStrings?.Select(g => Guid.Parse(g)).ToList() ?? new List<Guid>();
                            }
                            catch
                            {
                                cluster.ClashZoneIds = new List<Guid>();
                            }
                        }
                        else
                        {
                            cluster.ClashZoneIds = new List<Guid>();
                        }

                        clusters.Add(cluster);
                    }
                }
            }

            // ✅ LOG: SELECT results
            DatabaseOperationLogger.LogSelect(
                "ClusterSleeves",
                whereClause,
                resultCount: clusters.Count,
                sampleRow: clusters.Count > 0 ? new Dictionary<string, object>
                {
                    { "ClusterInstanceId", clusters[0].ClusterInstanceId },
                    { "ComboId", clusters[0].ComboId },
                    { "Category", clusters[0].Category },
                    { "ClashZoneIdsCount", clusters[0].ClashZoneIds?.Count ?? 0 }
                } : null);

            return clusters;
        }

        /// <summary>
        /// Load cluster sleeve data by FilterId and Category (for PATH 1 check when comboId is unknown).
        /// This checks if any clusters exist for a given filter+category combination.
        /// </summary>
        public List<ClusterSleeveData> LoadClusterSleevesByFilter(int filterId, string category)
        {
            var clusters = new List<ClusterSleeveData>();

            if (filterId <= 0 || string.IsNullOrWhiteSpace(category))
                return clusters;

            var whereClause = $"FilterId={filterId} AND Category='{category}'";

            DatabaseOperationLogger.LogSelect(
                "ClusterSleeves",
                whereClause,
                additionalInfo: "Checking for existing cluster data by FilterId+Category");

            using (var cmd = _context.Connection.CreateCommand())
            {
                cmd.CommandText = @"
                    SELECT ClusterInstanceId, ComboId, FilterId, Category,
                           BoundingBoxMinX, BoundingBoxMinY, BoundingBoxMinZ,
                           BoundingBoxMaxX, BoundingBoxMaxY, BoundingBoxMaxZ,
                           ClusterWidth, ClusterHeight, ClusterDepth,
                           RotationAngleDeg, IsRotated,
                           PlacementX, PlacementY, PlacementZ,
                           HostType, HostOrientation, ClashZoneIdsJson
                    FROM ClusterSleeves
                    WHERE FilterId = @FilterId AND Category = @Category";
                cmd.Parameters.AddWithValue("@FilterId", filterId);
                cmd.Parameters.AddWithValue("@Category", category);

                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        var cluster = new ClusterSleeveData
                        {
                            ClusterInstanceId = GetInt(reader, "ClusterInstanceId", -1),
                            ComboId = GetInt(reader, "ComboId", -1),
                            FilterId = GetInt(reader, "FilterId", -1),
                            Category = GetString(reader, "Category"),
                            BoundingBoxMinX = GetDouble(reader, "BoundingBoxMinX", 0.0),
                            BoundingBoxMinY = GetDouble(reader, "BoundingBoxMinY", 0.0),
                            BoundingBoxMinZ = GetDouble(reader, "BoundingBoxMinZ", 0.0),
                            BoundingBoxMaxX = GetDouble(reader, "BoundingBoxMaxX", 0.0),
                            BoundingBoxMaxY = GetDouble(reader, "BoundingBoxMaxY", 0.0),
                            BoundingBoxMaxZ = GetDouble(reader, "BoundingBoxMaxZ", 0.0),
                            ClusterWidth = GetDouble(reader, "ClusterWidth", 0.0),
                            ClusterHeight = GetDouble(reader, "ClusterHeight", 0.0),
                            ClusterDepth = GetDouble(reader, "ClusterDepth", 0.0),
                            RotationAngleDeg = GetDouble(reader, "RotationAngleDeg", 0.0),
                            IsRotated = GetBool(reader, "IsRotated"),
                            PlacementX = GetDouble(reader, "PlacementX", 0.0),
                            PlacementY = GetDouble(reader, "PlacementY", 0.0),
                            PlacementZ = GetDouble(reader, "PlacementZ", 0.0),
                            HostType = GetString(reader, "HostType"),
                            HostOrientation = GetString(reader, "HostOrientation")
                        };

                        // Deserialize ClashZoneIds from JSON
                        var clashZoneIdsJson = GetString(reader, "ClashZoneIdsJson");
                        if (!string.IsNullOrWhiteSpace(clashZoneIdsJson))
                        {
                            try
                            {
                                cluster.ClashZoneIds = JsonSerializer.Deserialize<List<Guid>>(clashZoneIdsJson) ?? new List<Guid>();
                            }
                            catch
                            {
                                cluster.ClashZoneIds = new List<Guid>();
                            }
                        }
                        else
                        {
                            cluster.ClashZoneIds = new List<Guid>();
                        }

                        clusters.Add(cluster);
                    }
                }
            }

            DatabaseOperationLogger.LogSelect(
                "ClusterSleeves",
                whereClause,
                resultCount: clusters.Count);

            return clusters;
        }

        /// <summary>
        /// Check if cluster data exists for a ComboId (used by PATH 1 to decide if recalculation is needed)
        /// </summary>
        public bool HasClusterDataForCombo(int comboId, string category = null)
        {
            using (var cmd = _context.Connection.CreateCommand())
            {
                if (string.IsNullOrWhiteSpace(category))
                {
                    cmd.CommandText = @"
                        SELECT COUNT(*) FROM ClusterSleeves 
                        WHERE ComboId = @ComboId";
                    cmd.Parameters.AddWithValue("@ComboId", comboId);
                }
                else
                {
                    cmd.CommandText = @"
                        SELECT COUNT(*) FROM ClusterSleeves 
                        WHERE ComboId = @ComboId AND Category = @Category";
                    cmd.Parameters.AddWithValue("@ComboId", comboId);
                    cmd.Parameters.AddWithValue("@Category", category);
                }

                var count = Convert.ToInt32(cmd.ExecuteScalar());
                return count > 0;
            }
        }

        /// <summary>
        /// Delete cluster sleeve data (called when cluster sleeve is deleted from Revit)
        /// </summary>
        public void DeleteClusterSleeve(int clusterInstanceId)
        {
            using (var cmd = _context.Connection.CreateCommand())
            {
                cmd.CommandText = "DELETE FROM ClusterSleeves WHERE ClusterInstanceId = @ClusterInstanceId";
                cmd.Parameters.AddWithValue("@ClusterInstanceId", clusterInstanceId);
                cmd.ExecuteNonQuery();
                _logger($"[SQLite] ✅ Deleted cluster sleeve {clusterInstanceId} from database");
            }
        }

        // Helper methods
        private int GetInt(SQLiteDataReader reader, string columnName, int defaultValue = 0)
        {
            var ordinal = reader.GetOrdinal(columnName);
            return reader.IsDBNull(ordinal) ? defaultValue : reader.GetInt32(ordinal);
        }

        private double GetDouble(SQLiteDataReader reader, string columnName, double defaultValue = 0.0)
        {
            var ordinal = reader.GetOrdinal(columnName);
            return reader.IsDBNull(ordinal) ? defaultValue : reader.GetDouble(ordinal);
        }

        private string GetString(SQLiteDataReader reader, string columnName)
        {
            var ordinal = reader.GetOrdinal(columnName);
            return reader.IsDBNull(ordinal) ? string.Empty : reader.GetString(ordinal);
        }

        private bool GetBool(SQLiteDataReader reader, string columnName)
        {
            var ordinal = reader.GetOrdinal(columnName);
            return !reader.IsDBNull(ordinal) && reader.GetInt32(ordinal) != 0;
        }
    } 


    /// <summary>
    /// Data transfer object for cluster sleeve data
    /// </summary>
    public class ClusterSleeveData
    {
        public int ClusterInstanceId { get; set; }
        public int ComboId { get; set; }
        public int FilterId { get; set; }
        public string Category { get; set; } = string.Empty;
        public double BoundingBoxMinX { get; set; }
        public double BoundingBoxMinY { get; set; }
        public double BoundingBoxMinZ { get; set; }
        public double BoundingBoxMaxX { get; set; }
        public double BoundingBoxMaxY { get; set; }
        public double BoundingBoxMaxZ { get; set; }
        public double ClusterWidth { get; set; }
        public double ClusterHeight { get; set; }
        public double ClusterDepth { get; set; }
        public double RotationAngleDeg { get; set; }
        public bool IsRotated { get; set; }
        public double PlacementX { get; set; }
        public double PlacementY { get; set; }
        public double PlacementZ { get; set; }
        public string HostType { get; set; } = string.Empty;
        public string HostOrientation { get; set; } = string.Empty;
        public List<Guid> ClashZoneIds { get; set; } = new List<Guid>();
        
        // Corner coordinates for precise geometric proximity checks
        public double Corner1X { get; set; }
        public double Corner1Y { get; set; }
        public double Corner1Z { get; set; }
        public double Corner2X { get; set; }
        public double Corner2Y { get; set; }
        public double Corner2Z { get; set; }
        public double Corner3X { get; set; }
        public double Corner3Y { get; set; }
        public double Corner3Z { get; set; }
        public double Corner4X { get; set; }
        public double Corner4Y { get; set; }
        public double Corner4Z { get; set; }
    }
}



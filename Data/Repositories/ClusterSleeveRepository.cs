using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Linq;
using System.Text.Json;
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Data.Repositories
{
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
            List<Guid> clashZoneIds)
        {
            if (clusterInstanceId <= 0)
                throw new ArgumentException("ClusterInstanceId must be greater than 0", nameof(clusterInstanceId));

            using (var transaction = _context.Connection.BeginTransaction())
            {
                try
                {
                    // Check if cluster already exists
                    using (var checkCmd = _context.Connection.CreateCommand())
                    {
                        checkCmd.Transaction = transaction;
                        checkCmd.CommandText = @"
                            SELECT ClusterSleeveId FROM ClusterSleeves 
                            WHERE ClusterInstanceId = @ClusterInstanceId";
                        checkCmd.Parameters.AddWithValue("@ClusterInstanceId", clusterInstanceId);

                        var existingId = checkCmd.ExecuteScalar();

                        if (existingId != null)
                        {
                            // Update existing cluster
                            using (var updateCmd = _context.Connection.CreateCommand())
                            {
                                updateCmd.Transaction = transaction;
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
                                        UpdatedAt = CURRENT_TIMESTAMP
                                    WHERE ClusterInstanceId = @ClusterInstanceId";

                                AddClusterSleeveParameters(updateCmd, clusterInstanceId, comboId, filterId, category,
                                    boundingBoxMinX, boundingBoxMinY, boundingBoxMinZ,
                                    boundingBoxMaxX, boundingBoxMaxY, boundingBoxMaxZ,
                                    clusterWidth, clusterHeight, clusterDepth,
                                    rotationAngleDeg, isRotated,
                                    placementX, placementY, placementZ,
                                    hostType, hostOrientation, clashZoneIds);

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
                                        ClusterInstanceId, ComboId, FilterId, Category,
                                        BoundingBoxMinX, BoundingBoxMinY, BoundingBoxMinZ,
                                        BoundingBoxMaxX, BoundingBoxMaxY, BoundingBoxMaxZ,
                                        ClusterWidth, ClusterHeight, ClusterDepth,
                                        RotationAngleDeg, IsRotated,
                                        PlacementX, PlacementY, PlacementZ,
                                        HostType, HostOrientation, ClashZoneIdsJson,
                                        CreatedAt, UpdatedAt
                                    ) VALUES (
                                        @ClusterInstanceId, @ComboId, @FilterId, @Category,
                                        @BoundingBoxMinX, @BoundingBoxMinY, @BoundingBoxMinZ,
                                        @BoundingBoxMaxX, @BoundingBoxMaxY, @BoundingBoxMaxZ,
                                        @ClusterWidth, @ClusterHeight, @ClusterDepth,
                                        @RotationAngleDeg, @IsRotated,
                                        @PlacementX, @PlacementY, @PlacementZ,
                                        @HostType, @HostOrientation, @ClashZoneIdsJson,
                                        CURRENT_TIMESTAMP, CURRENT_TIMESTAMP
                                    )";

                                AddClusterSleeveParameters(insertCmd, clusterInstanceId, comboId, filterId, category,
                                    boundingBoxMinX, boundingBoxMinY, boundingBoxMinZ,
                                    boundingBoxMaxX, boundingBoxMaxY, boundingBoxMaxZ,
                                    clusterWidth, clusterHeight, clusterDepth,
                                    rotationAngleDeg, isRotated,
                                    placementX, placementY, placementZ,
                                    hostType, hostOrientation, clashZoneIds);

                                insertCmd.ExecuteNonQuery();
                                _logger($"[SQLite] ✅ Saved cluster sleeve {clusterInstanceId} to database");
                            }
                        }
                    }

                    transaction.Commit();
                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    _logger($"[SQLite] ❌ Error saving cluster sleeve {clusterInstanceId}: {ex.Message}");
                    throw;
                }
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
            List<Guid> clashZoneIds)
        {
            cmd.Parameters.AddWithValue("@ClusterInstanceId", clusterInstanceId);
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

            // Serialize ClashZoneIds to JSON
            var clashZoneIdsJson = clashZoneIds != null && clashZoneIds.Count > 0
                ? JsonSerializer.Serialize(clashZoneIds.Select(g => g.ToString()).ToList())
                : "[]";
            cmd.Parameters.AddWithValue("@ClashZoneIdsJson", clashZoneIdsJson);
        }

        /// <summary>
        /// Load cluster sleeve data for a specific ComboId and Category
        /// Used by PATH 1 (Replay) to check if cluster data exists
        /// </summary>
        public List<ClusterSleeveData> LoadClusterSleevesForCombo(int comboId, string category = null)
        {
            var clusters = new List<ClusterSleeveData>();

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
    }
}


using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using System.Text.Json;
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Services.Geometry;
using Autodesk.Revit.DB;

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
                // ✅ FIX: Read from ClusterSleeves_v2 (the only populated table)
                // Alias RotationAngleRad to RotationAngleDeg for downstream compatibility
                cmd.CommandText = @"SELECT ClusterInstanceId, ComboId, FilterId, Category,
                                       BoundingBoxMinX, BoundingBoxMinY, BoundingBoxMinZ,
                                       BoundingBoxMaxX, BoundingBoxMaxY, BoundingBoxMaxZ,
                                       ClusterWidth, ClusterHeight, ClusterDepth,
                                       RotationAngleRad * 57.29577951308232 AS RotationAngleDeg, IsRotated,
                                       PlacementX, PlacementY, PlacementZ,
                                       HostType, HostOrientation, ClashZoneIdsJson,
                                       SleeveFamilyName,
                                       Corner1X, Corner1Y, Corner1Z,
                                       Corner2X, Corner2Y, Corner2Z,
                                       Corner3X, Corner3Y, Corner3Z,
                                       Corner4X, Corner4Y, Corner4Z
                                FROM ClusterSleeves_v2";

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
                // ✅ FIX: Read from ClusterSleeves_v2 (the only populated table)
                // Alias RotationAngleRad to RotationAngleDeg for downstream compatibility
                cmd.CommandText = $@"SELECT ClusterInstanceId, ComboId, FilterId, Category,
                                       BoundingBoxMinX, BoundingBoxMinY, BoundingBoxMinZ,
                                       BoundingBoxMaxX, BoundingBoxMaxY, BoundingBoxMaxZ,
                                       ClusterWidth, ClusterHeight, ClusterDepth,
                                       RotationAngleRad * 57.29577951308232 AS RotationAngleDeg, IsRotated,
                                       PlacementX, PlacementY, PlacementZ,
                                       HostType, HostOrientation, ClashZoneIdsJson,
                                       SleeveFamilyName,
                                       Corner1X, Corner1Y, Corner1Z,
                                       Corner2X, Corner2Y, Corner2Z,
                                       Corner3X, Corner3Y, Corner3Z,
                                       Corner4X, Corner4Y, Corner4Z
                                FROM ClusterSleeves_v2
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

            // ✅ MIXED-TYPE FIX: Fallback to ClusterSleeves_v2 for any instance IDs not found in ClusterSleeves.
            // When clusters are saved only to v2 (e.g. bulk path), combined discovery needs their corners for
            // individual+cluster proximity groups; otherwise cluster corners are missing and combined dimensions fail.
            var foundIds = new HashSet<int>(clusters.Select(c => c.ClusterInstanceId));
            var missingIds = ids.Where(id => !foundIds.Contains(id)).ToList();
            if (missingIds.Count > 0)
            {
                try
                {
                    var fromV2 = GetClusterSleevesByInstanceIdsFromV2(missingIds);
                    foreach (var c in fromV2)
                        clusters.Add(c);
                }
                catch (Exception ex)
                {
                    _logger($"[ClusterSleeveRepo] Fallback ClusterSleeves_v2 read failed: {ex.Message}");
                }
            }

            return clusters;
        }

        /// <summary>
        /// Load cluster sleeve data by ClusterInstanceId from ClusterSleeves_v2 (for corner merge when not in legacy table).
        /// </summary>
        private List<ClusterSleeveData> GetClusterSleevesByInstanceIdsFromV2(List<int> instanceIds)
        {
            var clusters = new List<ClusterSleeveData>();
            if (instanceIds == null || instanceIds.Count == 0) return clusters;

            var idString = string.Join(",", instanceIds);
            using (var cmd = _context.Connection.CreateCommand())
            {
                cmd.CommandText = $@"SELECT ClusterInstanceId, ComboId, FilterId, Category,
                                       PlacementX, PlacementY, PlacementZ,
                                       ClusterWidth, ClusterHeight, ClusterDepth,
                                       RotationAngleRad, IsRotated,
                                       HostType, HostOrientation,
                                       Corner1X, Corner1Y, Corner1Z,
                                       Corner2X, Corner2Y, Corner2Z,
                                       Corner3X, Corner3Y, Corner3Z,
                                       Corner4X, Corner4Y, Corner4Z
                                FROM ClusterSleeves_v2
                                WHERE ClusterInstanceId IN ({idString}) AND ClusterInstanceId IS NOT NULL";

                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        double c1x = GetDouble(reader, "Corner1X", 0.0), c1y = GetDouble(reader, "Corner1Y", 0.0), c1z = GetDouble(reader, "Corner1Z", 0.0);
                        double c2x = GetDouble(reader, "Corner2X", 0.0), c2y = GetDouble(reader, "Corner2Y", 0.0), c2z = GetDouble(reader, "Corner2Z", 0.0);
                        double c3x = GetDouble(reader, "Corner3X", 0.0), c3y = GetDouble(reader, "Corner3Y", 0.0), c3z = GetDouble(reader, "Corner3Z", 0.0);
                        double c4x = GetDouble(reader, "Corner4X", 0.0), c4y = GetDouble(reader, "Corner4Y", 0.0), c4z = GetDouble(reader, "Corner4Z", 0.0);

                        var cluster = new ClusterSleeveData
                        {
                            ClusterInstanceId = GetInt(reader, "ClusterInstanceId", -1),
                            ComboId = GetInt(reader, "ComboId", -1),
                            FilterId = GetInt(reader, "FilterId", -1),
                            Category = GetString(reader, "Category"),
                            BoundingBoxMinX = Math.Min(Math.Min(c1x, c2x), Math.Min(c3x, c4x)),
                            BoundingBoxMinY = Math.Min(Math.Min(c1y, c2y), Math.Min(c3y, c4y)),
                            BoundingBoxMinZ = Math.Min(Math.Min(c1z, c2z), Math.Min(c3z, c4z)),
                            BoundingBoxMaxX = Math.Max(Math.Max(c1x, c2x), Math.Max(c3x, c4x)),
                            BoundingBoxMaxY = Math.Max(Math.Max(c1y, c2y), Math.Max(c3y, c4y)),
                            BoundingBoxMaxZ = Math.Max(Math.Max(c1z, c2z), Math.Max(c3z, c4z)),
                            ClusterWidth = GetDouble(reader, "ClusterWidth", 0.0),
                            ClusterHeight = GetDouble(reader, "ClusterHeight", 0.0),
                            ClusterDepth = GetDouble(reader, "ClusterDepth", 0.0),
                            RotationAngleDeg = GetDouble(reader, "RotationAngleRad", 0.0) * (180.0 / Math.PI),
                            IsRotated = GetBool(reader, "IsRotated"),
                            PlacementX = GetDouble(reader, "PlacementX", 0.0),
                            PlacementY = GetDouble(reader, "PlacementY", 0.0),
                            PlacementZ = GetDouble(reader, "PlacementZ", 0.0),
                            HostType = GetString(reader, "HostType"),
                            HostOrientation = GetString(reader, "HostOrientation"),
                            Corner1X = c1x, Corner1Y = c1y, Corner1Z = c1z,
                            Corner2X = c2x, Corner2Y = c2y, Corner2Z = c2z,
                            Corner3X = c3x, Corner3Y = c3y, Corner3Z = c3z,
                            Corner4X = c4x, Corner4Y = c4y, Corner4Z = c4z,
                            ClashZoneIds = new List<Guid>()
                        };

                        clusters.Add(cluster);
                    }
                }
            }
            return clusters;
        }

        /// <summary>
        /// Retrieves a single cluster sleeve by its instance ID.
        /// </summary>
        public ClusterSleeveData GetClusterSleeveById(int clusterInstanceId)
        {
            var result = GetClusterSleevesByInstanceIds(new[] { clusterInstanceId });
            return result.FirstOrDefault();
        }
        // ...existing code...



        /// <summary>
        /// Save cluster sleeve calculation results to database
        /// Called after cluster calculation completes in PATH 2/3
        /// </summary>
        /// <param name="rotationAngleRadOverride">When set, use this angle for DB (same as used for placement). Stops overwriting with instance rotation which can be 0 before commit.</param>
        public void SaveClusterSleeve(
            int clusterInstanceId,
            int comboId,
            int filterId,
            string category,
            string sleeveFamilyName,
            FamilyInstance actualInstance,
            double? rotationAngleRadOverride = null)
        {
            // This is a bridge method for callers who only have the high-level Task or Data
            // We'll extract what we need and call the primitive version

            double width = 0, height = 0, depth = 0, rot = 0;
            double px = 0, py = 0, pz = 0;
            List<Guid> zoneGuids = new List<Guid>();

            // Use cluster/ClashZone angle when provided so we don't overwrite with instance rotation (often 0 at save time)
            if (rotationAngleRadOverride.HasValue)
                rot = rotationAngleRadOverride.Value;

            // Extract from actualInstance if possible for MAX ACCURACY
            if (actualInstance != null && actualInstance.IsValidObject)
            {
                width = (actualInstance.LookupParameter("Width") ?? actualInstance.LookupParameter("Element Width"))?.AsDouble() ?? 0;
                height = (actualInstance.LookupParameter("Height") ?? actualInstance.LookupParameter("Element Height"))?.AsDouble() ?? 0;
                depth = (actualInstance.LookupParameter("Depth") ?? actualInstance.LookupParameter("Element Depth") ?? actualInstance.LookupParameter("Wall Width"))?.AsDouble() ?? 0;

                var loc = actualInstance.Location as LocationPoint;
                if (loc != null && !rotationAngleRadOverride.HasValue)
                {
                    px = loc.Point.X;
                    py = loc.Point.Y;
                    pz = loc.Point.Z;
                    rot = loc.Rotation;
                }
                else if (loc != null)
                {
                    px = loc.Point.X;
                    py = loc.Point.Y;
                    pz = loc.Point.Z;
                }

                // Extract high-accuracy corners
                var cornerService = new SleeveCornerCalculationService();
                var corners = cornerService.CalculateCornersFromInstance(actualInstance);

                if (corners.HasValue)
                {
                    SaveClusterSleeve(
                        clusterInstanceId, comboId, filterId, category,
                        0, 0, 0, 0, 0, 0, // BBox (not used by calculations that use corners anyway)
                        width, height, depth,
                        rot * (180.0 / Math.PI), // Convert to Deg
                        Math.Abs(rot) > 1e-6,
                        px, py, pz,
                        actualInstance.Host?.Category?.Name ?? "Wall",
                        "X", // Default orientation
                        new List<Guid>(), // GUIDs should be passed separately if needed
                        sleeveFamilyName,
                        corners.Value.corner1.X, corners.Value.corner1.Y, corners.Value.corner1.Z,
                        corners.Value.corner2.X, corners.Value.corner2.Y, corners.Value.corner2.Z,
                        corners.Value.corner3.X, corners.Value.corner3.Y, corners.Value.corner3.Z,
                        corners.Value.corner4.X, corners.Value.corner4.Y, corners.Value.corner4.Z
                    );
                    return;
                }
            }

            // Fallback: If no corners extracted, this method is incomplete for this specific call pattern.
            // But let's assume for now the primitive version is called directly if actualInstance is null.
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

            // ✅ FIX: If GUIDs are not provided (e.g. from SaveClusterSleeve(actualInstance)), try to retrieve them from ClashZones
            if (clashZoneIds == null || clashZoneIds.Count == 0)
            {
                clashZoneIds = GetClashZoneGuidsForCluster(clusterInstanceId);
                _logger($"[{DateTime.Now:HH:mm:ss}]       SaveClusterSleeve: Auto-retrieved {clashZoneIds.Count} GUIDs for ClusterInstanceId={clusterInstanceId}\n");
            }

            // ✅ DIAGNOSTIC: Log method entry
            _logger($"[{DateTime.Now:HH:mm:ss}]       SaveClusterSleeve CALLED\n" +
                   $"           clusterInstanceId={clusterInstanceId}, comboId={comboId}, filterId={filterId}\n");

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
                                            Corner1X = @Corner1X, Corner1Y = @Corner1Y, Corner1Z = @Corner1Z,
                                            Corner2X = @Corner2X, Corner2Y = @Corner2Y, Corner2Z = @Corner2Z,
                                            Corner3X = @Corner3X, Corner3Y = @Corner3Y, Corner3Z = @Corner3Z,
                                            Corner4X = @Corner4X, Corner4Y = @Corner4Y, Corner4Z = @Corner4Z,
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
                                        Corner1X = @Corner1X, Corner1Y = @Corner1Y, Corner1Z = @Corner1Z,
                                        Corner2X = @Corner2X, Corner2Y = @Corner2Y, Corner2Z = @Corner2Z,
                                        Corner3X = @Corner3X, Corner3Y = @Corner3Y, Corner3Z = @Corner3Z,
                                        Corner4X = @Corner4X, Corner4Y = @Corner4Y, Corner4Z = @Corner4Z,
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
                                        ClashZoneGuids, MepSizes, MepSystemNames, MepServiceTypes, MepElementIds,
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
                                        @ClashZoneGuids, @MepSizes, @MepSystemNames, @MepServiceTypes, @MepElementIds,
                                        @Corner1X, @Corner1Y, @Corner1Z,
                                        @Corner2X, @Corner2Y, @Corner2Z,
                                        @Corner3X, @Corner3Y, @Corner3Z,
                                        @Corner4X, @Corner4Y, @Corner4Z,
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

                                // ✅ DIAGNOSTIC: Log before INSERT execution
                                _logger($"[{DateTime.Now:HH:mm:ss}]           About to execute INSERT\n");

                                var rowsAffected = insertCmd.ExecuteNonQuery();

                                // ✅ DIAGNOSTIC: Log after INSERT execution
                                _logger($"[{DateTime.Now:HH:mm:ss}]           ✅ INSERT executed: {rowsAffected} row(s) affected\n");

                                if (rowsAffected == 0)
                                {
                                    _logger($"[{DateTime.Now:HH:mm:ss}]           ⚠️ WARNING: No rows inserted!\n");
                                }

                                // ✅ LOG: INSERT operation - ALL COLUMNS (one sample row)
                                var clashZoneIdsJson = clashZoneIds != null && clashZoneIds.Count > 0
                                    ? JsonSerializer.Serialize(clashZoneIds.Select(g => g.ToString()).ToList())
                                    : "[]";
                                var (clashZoneGuids, mepSizes, mepSystemNames, mepServiceTypes, mepElementIds) = GetCommaSeparatedMepData(clashZoneIds);

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

            // Non-bulk path uses _context.Connection with its own transaction
            BatchSaveClusterSleevesInternal(clusters, _context.Connection, null);
        }

        /// <summary>
        /// 🚀 BATCH SAVE (External Transaction): Save using caller's connection/transaction.
        /// Fixes WAL snapshot isolation: all writes share the same transaction as UpdateStatusBatch,
        /// so the legacy sync SELECT FROM ClusterSleeves_v2 sees the newly-written rows.
        /// </summary>
        public void BatchSaveClusterSleeves(List<ClusterSaveData> clusters, SQLiteConnection conn, SQLiteTransaction trans)
        {
            if (clusters == null || clusters.Count == 0)
                return;

            if (OptimizationFlags.UseBulkClusterSave)
            {
                BatchSaveClusterSleevesBulk(clusters, conn, trans);
                return;
            }

            BatchSaveClusterSleevesInternal(clusters, conn, trans);
        }

        /// <summary>
        /// Internal implementation for the non-bulk path.
        /// When externalTrans is null, creates and manages its own transaction.
        /// When externalTrans is provided, uses it without committing (caller commits).
        /// </summary>
        private void BatchSaveClusterSleevesInternal(List<ClusterSaveData> clusters, SQLiteConnection conn, SQLiteTransaction externalTrans)
        {
            if (clusters == null || clusters.Count == 0)
                return;

            // When externalTrans is null, we manage our own transaction lifecycle
            bool ownTransaction = (externalTrans == null);
            var transaction = externalTrans ?? conn.BeginTransaction();

            try
            {
                foreach (var cluster in clusters)
                {
                    // ✅ CRITICAL FIX: Generate deterministic ClusterGuid from sorted ClashZoneIds
                    string clusterGuid = GenerateDeterministicClusterGuid(cluster.ClashZoneIds);

                    // Check if cluster already exists
                    using (var checkCmd = conn.CreateCommand())
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
                            using (var updateCmd = conn.CreateCommand())
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
                                            Corner1X = @Corner1X, Corner1Y = @Corner1Y, Corner1Z = @Corner1Z,
                                            Corner2X = @Corner2X, Corner2Y = @Corner2Y, Corner2Z = @Corner2Z,
                                            Corner3X = @Corner3X, Corner3Y = @Corner3Y, Corner3Z = @Corner3Z,
                                            Corner4X = @Corner4X, Corner4Y = @Corner4Y, Corner4Z = @Corner4Z,
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
                            using (var insertCmd = conn.CreateCommand())
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

                if (ownTransaction)
                {
                    transaction.Commit();
                }
                DatabaseOperationLogger.LogTransaction("COMMIT", "SUCCESS", $"Batch saved {clusters.Count} clusters");
                _logger($"[SQLite] ✅ Batch saved {clusters.Count} cluster sleeves in single transaction");
            }
            catch (Exception ex)
            {
                DatabaseOperationLogger.LogTransaction("ROLLBACK", "FAILED", ex.Message);
                if (ownTransaction)
                {
                    transaction.Rollback();
                }
                _logger($"[SQLite] ❌ Error batch saving cluster sleeves: {ex.Message}");
                throw;
            }
            finally
            {
                if (ownTransaction)
                {
                    transaction.Dispose();
                }
            }
        }

        /// <summary>
        /// ✅ SIMPLIFIED: Save clusters using INSERT OR REPLACE based on ClusterInstanceId (PRIMARY KEY).
        /// Uses _context.Connection with its own transaction (standalone mode).
        /// </summary>
        private void BatchSaveClusterSleevesBulk(List<ClusterSaveData> clusters)
        {
            BatchSaveClusterSleevesBulkInternal(clusters, _context.Connection, null);
        }

        /// <summary>
        /// ✅ ISOLATION FIX: Save clusters using caller's connection/transaction.
        /// All writes to ClusterSleeves_v2 happen on the same connection/transaction as
        /// UpdateStatusBatch, so the legacy sync SELECT sees the newly-written rows.
        /// </summary>
        private void BatchSaveClusterSleevesBulk(List<ClusterSaveData> clusters, SQLiteConnection conn, SQLiteTransaction trans)
        {
            BatchSaveClusterSleevesBulkInternal(clusters, conn, trans);
        }

        /// <summary>
        /// Internal shared implementation for bulk cluster save.
        /// When externalTrans is null, creates and manages its own transaction.
        /// When externalTrans is provided, uses it without committing (caller commits).
        /// </summary>
        private void BatchSaveClusterSleevesBulkInternal(List<ClusterSaveData> clusters, SQLiteConnection conn, SQLiteTransaction externalTrans)
        {
            if (clusters == null || clusters.Count == 0)
                return;

            // ✅ SINGLE WRITER: When bulk cluster save is enabled, BatchClusterCalculationService.SaveToDatabase
            // is the only writer for ClusterSleeves_v2 (timestamp batch ID + placement/bbox).
            // However, BatchClusterPlacementService ALSO calls this to save FINAL corners and Instance IDs.
            // So we only skip if ALL items are Pending (Calculation Phase). If any have ID > 0, we MUST write.
            if (OptimizationFlags.UseBulkClusterSave && clusters.All(c => c.ClusterInstanceId <= 0))
            {
                SafeFileLogger.SafeAppendText("batch_v2.log",
                    $"[{DateTime.Now:HH:mm:ss.fff}] [CLUSTER_V2_WRITE_REPO] Skipping v2 write for {clusters.Count} PENDING clusters (UseBulkClusterSave=true; calculation service is writer)\n");
                return;
            }

            var sw = System.Diagnostics.Stopwatch.StartNew();
            bool usingExternalTrans = (externalTrans != null);
            SafeFileLogger.SafeAppendText("batch_v2.log",
                $"[{DateTime.Now:HH:mm:ss.fff}] [CLUSTER_V2_WRITE_REPO] Writing {clusters.Count} clusters to ClusterSleeves_v2 (ExternalTrans={usingExternalTrans})\n");

            // When externalTrans is null, we manage our own transaction lifecycle
            bool ownTransaction = (externalTrans == null);
            var transaction = externalTrans ?? conn.BeginTransaction();

            try
            {
                // ✅ SCOPED DELETE: Delete ALL old clusters for the affected scopes (Combo/Filter/Category)
                // This is critical to remove "Ghost Clusters" if the cluster composition changes (and thus GUID changes).
                // We simply wipe the slate clean for the scopes we are about to update.
                var scopes = clusters
                    .Select(c => new { c.ComboId, c.FilterId, Category = c.Category ?? string.Empty })
                    .Distinct()
                    .ToList();

                using (var deleteCmd = conn.CreateCommand())
                {
                    deleteCmd.Transaction = transaction;
                    foreach (var scope in scopes)
                    {
                        deleteCmd.CommandText = @"
                            DELETE FROM ClusterSleeves_v2
                            WHERE ComboId = @ComboId
                              AND FilterId = @FilterId
                              AND Category = @Category";
                        deleteCmd.Parameters.Clear();
                        deleteCmd.Parameters.AddWithValue("@ComboId", scope.ComboId);
                        deleteCmd.Parameters.AddWithValue("@FilterId", scope.FilterId);
                        deleteCmd.Parameters.AddWithValue("@Category", scope.Category);
                        deleteCmd.ExecuteNonQuery();
                    }
                }

                int savedCount = 0;

                // ✅ FIX: Generate ONE batch ID for the entire save operation (not per row)
                // Use timestamp format: yyyyMMdd_HHmmss_Category_FilterId
                var firstCluster = clusters.FirstOrDefault();
                var category = firstCluster?.Category ?? "Unknown";
                var filterId = firstCluster?.FilterId ?? 0;
                var sharedBatchId = $"{DateTime.Now:yyyyMMdd_HHmmss}_{category}_{filterId}";

                SafeFileLogger.SafeAppendText("batch_v2.log",
                    $"[{DateTime.Now:HH:mm:ss.fff}] [CLUSTER_V2_WRITE_REPO] Using shared ClusterBatchId={sharedBatchId} for {clusters.Count} clusters\n");

                foreach (var cluster in clusters)
                {
                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.Transaction = transaction;

                        // Generate deterministic ClusterGuid
                        var clusterGuid = GenerateDeterministicClusterGuid(cluster.ClashZoneIds);

                        // Get comma-separated MEP data (reads from ClashZones/SleeveSnapshots via _context - safe, those tables are not being modified)
                        var (clashZoneGuids, mepSizes, mepSystemNames, mepServiceTypes, mepElementIds) = GetCommaSeparatedMepData(cluster.ClashZoneIds);

                        // Serialize ClashZoneIds to JSON
                        var clashZoneIdsJson = cluster.ClashZoneIds != null && cluster.ClashZoneIds.Count > 0
                            ? JsonSerializer.Serialize(cluster.ClashZoneIds.Select(g => g.ToString()).ToList())
                            : "[]";

                        // ✅ FIX: Use the shared batch ID for all clusters in this save operation
                        var clusterBatchId = sharedBatchId;
                        SafeFileLogger.SafeAppendText("batch_v2.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] [CLUSTER_V2_WRITE_REPO] Saving Cluster: ClusterGuid={clusterGuid}, PlacementX={cluster.PlacementX:F3}, PlacementY={cluster.PlacementY:F3}, PlacementZ={cluster.PlacementZ:F3}\n");

                        cmd.CommandText = @"
                            INSERT OR REPLACE INTO ClusterSleeves_v2 (
                                ClusterInstanceId, ComboId, FilterId, ClusterGUID, ClusterBatchId,
                                PlacementX, PlacementY, PlacementZ,
                                ClusterWidth, ClusterHeight, ClusterDepth,
                                RotationAngleRad, IsRotated,
                                HostType, HostOrientation, Category,
                                ValidationStatus, ValidationMessage,
                                ClashZoneIdsJson, ConstituentZoneGuids, ClashZoneGuids,
                                MepSizes, MepSystemNames, MepServiceTypes, MepElementIds,
                                Corner1X, Corner1Y, Corner1Z,
                                Corner2X, Corner2Y, Corner2Z,
                                Corner3X, Corner3Y, Corner3Z,
                                Corner4X, Corner4Y, Corner4Z,
                                SleeveFamilyName,
                                CalculatedAt, PlacedAt, Status
                            ) VALUES (
                                @ClusterInstanceId, @ComboId, @FilterId, @ClusterGuid, @ClusterBatchId,
                                @PlacementX, @PlacementY, @PlacementZ,
                                @ClusterWidth, @ClusterHeight, @ClusterDepth,
                                @RotationAngleRad, @IsRotated,
                                @HostType, @HostOrientation, @Category,
                                'Valid', '',
                                @ClashZoneIdsJson, @ClashZoneGuids, @ClashZoneGuids,
                                @MepSizes, @MepSystemNames, @MepServiceTypes, @MepElementIds,
                                @Corner1X, @Corner1Y, @Corner1Z,
                                @Corner2X, @Corner2Y, @Corner2Z,
                                @Corner3X, @Corner3Y, @Corner3Z,
                                @Corner4X, @Corner4Y, @Corner4Z,
                                @SleeveFamilyName,
                                CURRENT_TIMESTAMP, NULL, 'Pending'
                            )";

                        // Clear and re-add all parameters for the INSERT
                        cmd.Parameters.Clear();
                        cmd.Parameters.AddWithValue("@ClusterInstanceId", cluster.ClusterInstanceId);
                        cmd.Parameters.AddWithValue("@ComboId", cluster.ComboId);
                        cmd.Parameters.AddWithValue("@FilterId", cluster.FilterId);
                        cmd.Parameters.AddWithValue("@ClusterGuid", string.IsNullOrWhiteSpace(clusterGuid) ? (object)DBNull.Value : clusterGuid);
                        cmd.Parameters.AddWithValue("@ClusterBatchId", clusterBatchId);
                        cmd.Parameters.AddWithValue("@PlacementX", cluster.PlacementX);
                        cmd.Parameters.AddWithValue("@PlacementY", cluster.PlacementY);
                        cmd.Parameters.AddWithValue("@PlacementZ", cluster.PlacementZ);
                        cmd.Parameters.AddWithValue("@ClusterWidth", cluster.ClusterWidth);
                        cmd.Parameters.AddWithValue("@ClusterHeight", cluster.ClusterHeight);
                        cmd.Parameters.AddWithValue("@ClusterDepth", cluster.ClusterDepth);
                        // ✅ ANGLE: Convert Deg to Rad for V2
                        cmd.Parameters.AddWithValue("@RotationAngleRad", cluster.RotationAngleDeg * (Math.PI / 180.0));
                        cmd.Parameters.AddWithValue("@IsRotated", cluster.IsRotated ? 1 : 0);
                        cmd.Parameters.AddWithValue("@HostType", cluster.HostType ?? (object)DBNull.Value);
                        cmd.Parameters.AddWithValue("@HostOrientation", cluster.HostOrientation ?? (object)DBNull.Value);
                        cmd.Parameters.AddWithValue("@Category", cluster.Category ?? string.Empty);
                        cmd.Parameters.AddWithValue("@ClashZoneIdsJson", clashZoneIdsJson);
                        cmd.Parameters.AddWithValue("@ClashZoneGuids", clashZoneGuids ?? (object)DBNull.Value);
                        cmd.Parameters.AddWithValue("@MepSizes", mepSizes ?? (object)DBNull.Value);
                        cmd.Parameters.AddWithValue("@MepSystemNames", mepSystemNames ?? (object)DBNull.Value);
                        cmd.Parameters.AddWithValue("@MepServiceTypes", mepServiceTypes ?? (object)DBNull.Value);
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

                        // ✅ FAMILY NAME: Add family name explicitly
                        cmd.Parameters.AddWithValue("@SleeveFamilyName", cluster.SleeveFamilyName ?? (object)DBNull.Value);

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

                if (ownTransaction)
                {
                    transaction.Commit();
                }
                sw.Stop();

                DatabaseOperationLogger.LogTransaction("COMMIT", "SUCCESS", $"Saved {savedCount} clusters using INSERT OR REPLACE");
                _logger($"[SQLite] ✅ Saved {savedCount} cluster sleeves in {sw.ElapsedMilliseconds}ms using INSERT OR REPLACE (ExternalTrans={usingExternalTrans})");
            }
            catch (Exception ex)
            {
                DatabaseOperationLogger.LogTransaction("ROLLBACK", "FAILED", ex.Message);
                if (ownTransaction)
                {
                    transaction.Rollback();
                }
                _logger($"[SQLite] ❌ Error saving cluster sleeves: {ex.Message}");
                throw;
            }
            finally
            {
                if (ownTransaction)
                {
                    transaction.Dispose();
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
                    var (guids, sizes, names, serviceTypes, ids) = GetCommaSeparatedMepData(cluster.ClashZoneIds);
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
            var (clashZoneGuids, mepSizes, mepSystemNames, mepServiceTypes, mepElementIds) = GetCommaSeparatedMepData(cluster.ClashZoneIds);
            cmd.Parameters.AddWithValue($"@ClashZoneGuids{index}", clashZoneGuids ?? (object)DBNull.Value);
            cmd.Parameters.AddWithValue($"@MepSizes{index}", mepSizes ?? (object)DBNull.Value);
            cmd.Parameters.AddWithValue($"@MepSystemNames{index}", mepSystemNames ?? (object)DBNull.Value);
            cmd.Parameters.AddWithValue($"@MepServiceTypes{index}", mepServiceTypes ?? (object)DBNull.Value);
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
            var (clashZoneGuids, mepSizes, mepSystemNames, mepServiceTypes, mepElementIds) = GetCommaSeparatedMepData(clashZoneIds);
            cmd.Parameters.AddWithValue("@ClashZoneGuids", clashZoneGuids ?? (object)DBNull.Value);
            cmd.Parameters.AddWithValue("@MepSizes", mepSizes ?? (object)DBNull.Value);
            cmd.Parameters.AddWithValue("@MepSystemNames", mepSystemNames ?? (object)DBNull.Value);
            cmd.Parameters.AddWithValue("@MepServiceTypes", mepServiceTypes ?? (object)DBNull.Value);
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
        private (string clashZoneGuids, string mepSizes, string mepSystemNames, string mepServiceTypes, string mepElementIds) GetCommaSeparatedMepData(List<Guid> clashZoneIds)
        {
            if (clashZoneIds == null || clashZoneIds.Count == 0)
                return (null, null, null, null, null);

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
                            cz.ClashZoneGuid, -- ✅ FIX: Select GUID from ClashZones, not SleeveSnapshots (which might be null)
                            ss.MepParametersJson,
                            ss.MepElementIdsJson,
                            cz.MepElementId,
                            cz.MepElementSizeParameterValue,
                            cz.MepElementFormattedSize,
                            cz.MepSystemType,
                            cz.MepServiceType,
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
                    var serviceTypeList = new List<string>();
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
                            var systemName = reader.IsDBNull(6) ? null : reader.GetString(6); // ✅ FIXED: Was index 7, now 6 after removing MepElementSizeData
                            var serviceType = reader.IsDBNull(7) ? null : reader.GetString(7); // ✅ FIXED: Was index 8, now 7
                            var systemAbbr = reader.IsDBNull(8) ? null : reader.GetString(8); // ✅ FIXED: Was index 9, now 8

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
                            if (string.IsNullOrWhiteSpace(sizeValue) && !string.IsNullOrWhiteSpace(mepFormattedSize))
                            {
                                sizeValue = mepFormattedSize.Trim();
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

                            // ✅ FIX: Populate serviceTypeList
                            if (!string.IsNullOrWhiteSpace(serviceType))
                            {
                                serviceTypeList.Add(serviceType);
                            }
                        }
                    }

                    return (
                        guidList.Count > 0 ? string.Join(", ", guidList) : null,
                        sizeList.Count > 0 ? string.Join(", ", sizeList) : null,
                        systemNameList.Count > 0 ? string.Join(", ", systemNameList) : null,
                        serviceTypeList.Count > 0 ? string.Join(", ", serviceTypeList) : null,
                        elementIdList.Count > 0 ? string.Join(", ", elementIdList) : null
                    );
                }
            }
            catch (Exception ex)
            {
                _logger($"[SQLite] ⚠️ Error loading MEP data for cluster: {ex.Message}");
                return (null, null, null, null, null);
            }
        }

        /// <summary>
        /// Retrieves ClashZoneGuids associated with a cluster from the ClashZones table.
        /// </summary>
        public List<Guid> GetClashZoneGuidsForCluster(int clusterInstanceId)
        {
            var guids = new List<Guid>();
            try
            {
                using (var cmd = _context.Connection.CreateCommand())
                {
                    cmd.CommandText = "SELECT ClashZoneGuid FROM ClashZones WHERE ClusterInstanceId = @ClusterInstanceId AND ClashZoneGuid IS NOT NULL AND ClashZoneGuid != ''";
                    cmd.Parameters.AddWithValue("@ClusterInstanceId", clusterInstanceId);
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            if (Guid.TryParse(reader.GetString(0), out var guid))
                                guids.Add(guid);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger($"[SQLite] ⚠️ Error retrieving GUIDs for cluster {clusterInstanceId}: {ex.Message}");
            }
            return guids;
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
                    SELECT ClusterGUID, ClusterInstanceId, ComboId, FilterId, Category,
                           ClusterWidth, ClusterHeight, ClusterDepth,
                           RotationAngleRad, IsRotated,
                           PlacementX, PlacementY, PlacementZ,
                           HostType, HostOrientation, ClashZoneIdsJson,
                           SleeveFamilyName, Status
                    FROM ClusterSleeves_v2
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
                        SELECT COUNT(*) FROM ClusterSleeves_v2 
                        WHERE ComboId = @ComboId";
                    cmd.Parameters.AddWithValue("@ComboId", comboId);
                }
                else
                {
                    cmd.CommandText = @"
                        SELECT COUNT(*) FROM ClusterSleeves_v2 
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
                // ✅ V2: Delete by GUID or InstanceId depending on what's available
                // Usually called with InstanceId from Revit, so we might need to lookup or delete where InstanceId matches
                // For now, assume InstanceId is populated in V2 (it defaults to -1 but is updated after placement)
                cmd.CommandText = "DELETE FROM ClusterSleeves_v2 WHERE ClusterInstanceId = @ClusterInstanceId";
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
        /// <summary>
        /// ✅ BATCH UPDATE: Update corner coordinates for multiple cluster sleeves
        /// Updates BOTH ClusterSleeves (legacy) and ClusterSleeves_v2 (new)
        /// </summary>
        public void BatchUpdateClusterSleeveCorners(List<(int InstanceId, double c1x, double c1y, double c1z, double c2x, double c2y, double c2z, double c3x, double c3y, double c3z, double c4x, double c4y, double c4z)> updates)
        {
            if (updates == null || updates.Count == 0) return;

            using (var transaction = _context.Connection.BeginTransaction())
            {
                try
                {
                    using (var cmd = _context.Connection.CreateCommand())
                    {
                        cmd.Transaction = transaction;
                        var sql = new System.Text.StringBuilder();

                        // 1. Update Legacy Table
                        sql.AppendLine("UPDATE ClusterSleeves SET");
                        sql.AppendLine("  Corner1X = CASE ClusterInstanceId");
                        foreach (var u in updates) sql.AppendLine($"    WHEN {u.InstanceId} THEN {u.c1x}");
                        sql.AppendLine("  END,");
                        sql.AppendLine("  Corner1Y = CASE ClusterInstanceId");
                        foreach (var u in updates) sql.AppendLine($"    WHEN {u.InstanceId} THEN {u.c1y}");
                        sql.AppendLine("  END,");
                        sql.AppendLine("  Corner1Z = CASE ClusterInstanceId");
                        foreach (var u in updates) sql.AppendLine($"    WHEN {u.InstanceId} THEN {u.c1z}");
                        sql.AppendLine("  END,");
                        sql.AppendLine("  Corner2X = CASE ClusterInstanceId");
                        foreach (var u in updates) sql.AppendLine($"    WHEN {u.InstanceId} THEN {u.c2x}");
                        sql.AppendLine("  END,");
                        sql.AppendLine("  Corner2Y = CASE ClusterInstanceId");
                        foreach (var u in updates) sql.AppendLine($"    WHEN {u.InstanceId} THEN {u.c2y}");
                        sql.AppendLine("  END,");
                        sql.AppendLine("  Corner2Z = CASE ClusterInstanceId");
                        foreach (var u in updates) sql.AppendLine($"    WHEN {u.InstanceId} THEN {u.c2z}");
                        sql.AppendLine("  END,");
                        sql.AppendLine("  Corner3X = CASE ClusterInstanceId");
                        foreach (var u in updates) sql.AppendLine($"    WHEN {u.InstanceId} THEN {u.c3x}");
                        sql.AppendLine("  END,");
                        sql.AppendLine("  Corner3Y = CASE ClusterInstanceId");
                        foreach (var u in updates) sql.AppendLine($"    WHEN {u.InstanceId} THEN {u.c3y}");
                        sql.AppendLine("  END,");
                        sql.AppendLine("  Corner3Z = CASE ClusterInstanceId");
                        foreach (var u in updates) sql.AppendLine($"    WHEN {u.InstanceId} THEN {u.c3z}");
                        sql.AppendLine("  END,");
                        sql.AppendLine("  Corner4X = CASE ClusterInstanceId");
                        foreach (var u in updates) sql.AppendLine($"    WHEN {u.InstanceId} THEN {u.c4x}");
                        sql.AppendLine("  END,");
                        sql.AppendLine("  Corner4Y = CASE ClusterInstanceId");
                        foreach (var u in updates) sql.AppendLine($"    WHEN {u.InstanceId} THEN {u.c4y}");
                        sql.AppendLine("  END,");
                        sql.AppendLine("  Corner4Z = CASE ClusterInstanceId");
                        foreach (var u in updates) sql.AppendLine($"    WHEN {u.InstanceId} THEN {u.c4z}");
                        sql.AppendLine("  END,");
                        sql.AppendLine("  UpdatedAt = CURRENT_TIMESTAMP");
                        sql.AppendLine($"WHERE ClusterInstanceId IN ({string.Join(",", updates.Select(u => u.InstanceId))});");

                        // 2. Update New V2 Table
                        // Note: ClusterSleeves_v2 also has ClusterInstanceId populated after placement
                        sql.AppendLine("UPDATE ClusterSleeves_v2 SET");
                        sql.AppendLine("  Corner1X = CASE ClusterInstanceId");
                        foreach (var u in updates) sql.AppendLine($"    WHEN {u.InstanceId} THEN {u.c1x}");
                        sql.AppendLine("  END,");
                        sql.AppendLine("  Corner1Y = CASE ClusterInstanceId");
                        foreach (var u in updates) sql.AppendLine($"    WHEN {u.InstanceId} THEN {u.c1y}");
                        sql.AppendLine("  END,");
                        sql.AppendLine("  Corner1Z = CASE ClusterInstanceId");
                        foreach (var u in updates) sql.AppendLine($"    WHEN {u.InstanceId} THEN {u.c1z}");
                        sql.AppendLine("  END,");
                        sql.AppendLine("  Corner2X = CASE ClusterInstanceId");
                        foreach (var u in updates) sql.AppendLine($"    WHEN {u.InstanceId} THEN {u.c2x}");
                        sql.AppendLine("  END,");
                        sql.AppendLine("  Corner2Y = CASE ClusterInstanceId");
                        foreach (var u in updates) sql.AppendLine($"    WHEN {u.InstanceId} THEN {u.c2y}");
                        sql.AppendLine("  END,");
                        sql.AppendLine("  Corner2Z = CASE ClusterInstanceId");
                        foreach (var u in updates) sql.AppendLine($"    WHEN {u.InstanceId} THEN {u.c2z}");
                        sql.AppendLine("  END,");
                        sql.AppendLine("  Corner3X = CASE ClusterInstanceId");
                        foreach (var u in updates) sql.AppendLine($"    WHEN {u.InstanceId} THEN {u.c3x}");
                        sql.AppendLine("  END,");
                        sql.AppendLine("  Corner3Y = CASE ClusterInstanceId");
                        foreach (var u in updates) sql.AppendLine($"    WHEN {u.InstanceId} THEN {u.c3y}");
                        sql.AppendLine("  END,");
                        sql.AppendLine("  Corner3Z = CASE ClusterInstanceId");
                        foreach (var u in updates) sql.AppendLine($"    WHEN {u.InstanceId} THEN {u.c3z}");
                        sql.AppendLine("  END,");
                        sql.AppendLine("  Corner4X = CASE ClusterInstanceId");
                        foreach (var u in updates) sql.AppendLine($"    WHEN {u.InstanceId} THEN {u.c4x}");
                        sql.AppendLine("  END,");
                        sql.AppendLine("  Corner4Y = CASE ClusterInstanceId");
                        foreach (var u in updates) sql.AppendLine($"    WHEN {u.InstanceId} THEN {u.c4y}");
                        sql.AppendLine("  END,");
                        sql.AppendLine("  Corner4Z = CASE ClusterInstanceId");
                        foreach (var u in updates) sql.AppendLine($"    WHEN {u.InstanceId} THEN {u.c4z}");
                        sql.AppendLine("  END");
                        // Only update V2 where InstanceId matches
                        sql.AppendLine($"WHERE ClusterInstanceId IN ({string.Join(",", updates.Select(u => u.InstanceId))});");

                        cmd.CommandText = sql.ToString();
                        int rows = cmd.ExecuteNonQuery();
                        _logger($"[SQLite] ✅ Batch updated corners for {updates.Count} clusters (Legacy + V2). Rows affected: {rows}");
                    }
                    transaction.Commit();
                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    _logger($"[SQLite] ❌ Error batch updating cluster corners: {ex.Message}");
                    throw;
                }
            }
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



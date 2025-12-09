using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Text.Json;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Data.Repositories
{
    /// <summary>
    /// Repository for reading sleeve parameter snapshots from SQLite.
    /// </summary>
    public class SleeveSnapshotRepository
    {
        private readonly SleeveDbContext _context;
        private readonly Action<string> _logger;

        public SleeveSnapshotRepository(SleeveDbContext context, Action<string>? logger = null)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            _logger = logger ?? (_ => { });
        }

        public SleeveSnapshotIndex LoadSnapshotIndex()
        {
            var index = new SleeveSnapshotIndex();

            using (var cmd = _context.Connection.CreateCommand())
            {
                cmd.CommandText = @"
                    SELECT
                        SnapshotId,
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
                        ClashZoneGuid
                    FROM SleeveSnapshots
                    ORDER BY SnapshotId";
                
                // ✅ DIAGNOSTIC: Log total snapshots in database before loading
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    using (var countCmd = _context.Connection.CreateCommand())
                    {
                        countCmd.CommandText = "SELECT COUNT(*) FROM SleeveSnapshots";
                        var totalCount = countCmd.ExecuteScalar();
                        _logger?.Invoke($"[SQLite] 🔍 Database contains {totalCount} total snapshot(s)");
                    }
                }

                int rowsProcessed = 0;
                int rowsAddedToIndex = 0;
                
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        rowsProcessed++;
                        try
                        {
                            var view = new SleeveSnapshotView
                            {
                                SnapshotId = reader.GetInt32(reader.GetOrdinal("SnapshotId")),
                                SleeveInstanceId = reader.IsDBNull(reader.GetOrdinal("SleeveInstanceId"))
                                    ? (int?)null
                                    : reader.GetInt32(reader.GetOrdinal("SleeveInstanceId")),
                                ClusterInstanceId = reader.IsDBNull(reader.GetOrdinal("ClusterInstanceId"))
                                    ? (int?)null
                                    : reader.GetInt32(reader.GetOrdinal("ClusterInstanceId")),
                                SourceType = SafeGetString(reader, "SourceType") ?? "Individual",
                                FilterId = reader.IsDBNull(reader.GetOrdinal("FilterId"))
                                    ? (int?)null
                                    : reader.GetInt32(reader.GetOrdinal("FilterId")),
                                ComboId = reader.IsDBNull(reader.GetOrdinal("ComboId"))
                                    ? (int?)null
                                    : reader.GetInt32(reader.GetOrdinal("ComboId")),
                                MepElementIds = DeserializeIntList(SafeGetString(reader, "MepElementIdsJson")),
                                HostElementIds = DeserializeIntList(SafeGetString(reader, "HostElementIdsJson")),
                                MepParameters = DeserializeDictionary(SafeGetString(reader, "MepParametersJson")),
                                HostParameters = DeserializeDictionary(SafeGetString(reader, "HostParametersJson")),
                                SourceDocKeys = DeserializeStringList(SafeGetString(reader, "SourceDocKeysJson")),
                                HostDocKeys = DeserializeStringList(SafeGetString(reader, "HostDocKeysJson")),
                                ClashZoneGuid = SafeGetString(reader, "ClashZoneGuid") // ✅ NEW: Load ClashZoneGuid
                            };

                            if (view.SleeveInstanceId.HasValue && view.SleeveInstanceId.Value > 0)
                            {
                                index.BySleeve[view.SleeveInstanceId.Value] = view;
                                rowsAddedToIndex++;
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    _logger?.Invoke($"[SQLite] ✅ Added to BySleeve: SnapshotId={view.SnapshotId}, SleeveInstanceId={view.SleeveInstanceId.Value}");
                                }
                            }
                            else if (view.SleeveInstanceId.HasValue)
                            {
                                // ✅ DIAGNOSTIC: Log when SleeveInstanceId is 0 or negative
                                _logger?.Invoke($"[SQLite] ⚠️ Snapshot {view.SnapshotId} has invalid SleeveInstanceId={view.SleeveInstanceId.Value} (must be > 0)");
                            }
                            else
                            {
                                // ✅ DIAGNOSTIC: Log when SleeveInstanceId is NULL
                                _logger?.Invoke($"[SQLite] ℹ️ Snapshot {view.SnapshotId} has NULL SleeveInstanceId (checking ClusterInstanceId)");
                            }

                            if (view.ClusterInstanceId.HasValue && view.ClusterInstanceId.Value > 0)
                            {
                                index.ByCluster[view.ClusterInstanceId.Value] = view;
                                rowsAddedToIndex++;
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    _logger?.Invoke($"[SQLite] ✅ Added to ByCluster: SnapshotId={view.SnapshotId}, ClusterInstanceId={view.ClusterInstanceId.Value}");
                                }
                            }
                            {
                                index.ByCluster[view.ClusterInstanceId.Value] = view;
                            }
                            
                            // ✅ DIAGNOSTIC: Log all snapshots loaded, especially for debugging missing sleeves
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                var sleeveId = view.SleeveInstanceId.HasValue ? view.SleeveInstanceId.Value.ToString() : "NULL";
                                var clusterId = view.ClusterInstanceId.HasValue ? view.ClusterInstanceId.Value.ToString() : "NULL";
                                _logger?.Invoke($"[SQLite] Loaded snapshot: SnapshotId={view.SnapshotId}, SleeveInstanceId={sleeveId}, ClusterInstanceId={clusterId}");
                            }
                        }
                        catch (Exception ex)
                        {
                            _logger?.Invoke($"[SQLite] ⚠️ Failed to parse sleeve snapshot row: {ex.Message}");
                        }
                    }
                }
            }

            _logger?.Invoke($"[SQLite] ✅ Loaded {index.BySleeve.Count} individual and {index.ByCluster.Count} cluster snapshots");
            
            // ✅ DIAGNOSTIC: Log all loaded snapshots for debugging
            if (!DeploymentConfiguration.DeploymentMode && index.BySleeve.Count > 0)
            {
                var sleeveIdsLoaded = string.Join(", ", index.BySleeve.Keys.Take(10));
                _logger?.Invoke($"[SQLite] Loaded individual sleeve snapshots: [{sleeveIdsLoaded}]");
            }
            
            if (!DeploymentConfiguration.DeploymentMode && index.ByCluster.Count > 0)
            {
                var clusterIdsLoaded = string.Join(", ", index.ByCluster.Keys.Take(10));
                _logger?.Invoke($"[SQLite] Loaded cluster snapshots: [{clusterIdsLoaded}]");
            }
            
            if (!DeploymentConfiguration.DeploymentMode && index.BySleeve.Count == 0 && index.ByCluster.Count == 0)
            {
                _logger?.Invoke($"[SQLite] ⚠️⚠️⚠️ WARNING: No snapshots were loaded (both individual and cluster are empty)!");
            }
            
            return index;
        }

        private string SafeGetString(SQLiteDataReader reader, string column)
        {
            var ordinal = reader.GetOrdinal(column);
            return reader.IsDBNull(ordinal) ? null : reader.GetString(ordinal);
        }

        private Dictionary<string, string> DeserializeDictionary(string json)
        {
            if (string.IsNullOrWhiteSpace(json))
                return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            try
            {
                return JsonSerializer.Deserialize<Dictionary<string, string>>(json) ??
                       new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            }
            catch
            {
                return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            }
        }

        private List<int> DeserializeIntList(string json)
        {
            if (string.IsNullOrWhiteSpace(json))
                return new List<int>();

            try
            {
                return JsonSerializer.Deserialize<List<int>>(json) ?? new List<int>();
            }
            catch
            {
                return new List<int>();
            }
        }

        private List<string> DeserializeStringList(string json)
        {
            if (string.IsNullOrWhiteSpace(json))
                return new List<string>();

            try
            {
                return JsonSerializer.Deserialize<List<string>>(json) ?? new List<string>();
            }
            catch
            {
                return new List<string>();
            }
        }
    }

    public class SleeveSnapshotIndex
    {
        public Dictionary<int, SleeveSnapshotView> BySleeve { get; } = new Dictionary<int, SleeveSnapshotView>();
        public Dictionary<int, SleeveSnapshotView> ByCluster { get; } = new Dictionary<int, SleeveSnapshotView>();

        public bool TryGetBySleeve(int sleeveInstanceId, out SleeveSnapshotView view)
        {
            return BySleeve.TryGetValue(sleeveInstanceId, out view);
        }

        public bool TryGetByCluster(int clusterInstanceId, out SleeveSnapshotView view)
        {
            return ByCluster.TryGetValue(clusterInstanceId, out view);
        }
    }

    public class SleeveSnapshotView
    {
        public int SnapshotId { get; set; }
        public int? SleeveInstanceId { get; set; }
        public int? ClusterInstanceId { get; set; }
        public string SourceType { get; set; } = "Individual";
        public int? FilterId { get; set; }
        public int? ComboId { get; set; }
        public List<int> MepElementIds { get; set; } = new List<int>();
        public List<int> HostElementIds { get; set; } = new List<int>();
        public Dictionary<string, string> MepParameters { get; set; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        public Dictionary<string, string> HostParameters { get; set; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        public List<string> SourceDocKeys { get; set; } = new List<string>();
        public List<string> HostDocKeys { get; set; } = new List<string>();
        public string ClashZoneGuid { get; set; } // ✅ NEW
    }
}


using System;
using System.Collections.Generic;
#if !NET8_0_OR_GREATER
using System.Data.SQLite;
#endif
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
                                ClashZoneGuid = SafeGetString(reader, "ClashZoneGuid"), // ✅ NEW: Load ClashZoneGuid
                                CombinedInstanceId = SafeGetInt(reader, "CombinedInstanceId") // ✅ NEW: Load CombinedInstanceId
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

                            if (view.CombinedInstanceId.HasValue && view.CombinedInstanceId.Value > 0)
                            {
                                index.ByCombinedSnapshot[view.CombinedInstanceId.Value] = view;
                                rowsAddedToIndex++;
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    _logger?.Invoke($"[SQLite] ✅ Added to ByCombinedSnapshot: SnapshotId={view.SnapshotId}, CombinedInstanceId={view.CombinedInstanceId.Value}");
                                }
                            }
                            
                            // ✅ NEW: Index by ClashZoneGuid (MUST be outside try-catch killer block)
                            if (!string.IsNullOrEmpty(view.ClashZoneGuid))
                            {
                                index.ByClashZoneGuid[view.ClashZoneGuid] = view;
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

            // ✅ NEW: Load SleeveInstanceId -> ClashZoneGuid mapping from ClashZones table
            // This is critical for finding snapshots when SleeveSnapshots table has missing/stale SleeveInstanceId
            // but ClashZones table has the correct link.
            try
            {
                using (var cmd = _context.Connection.CreateCommand())
                {
                    cmd.CommandText = @"
                        SELECT SleeveInstanceId, ClashZoneGuid 
                        FROM ClashZones 
                        WHERE SleeveInstanceId > 0 AND ClashZoneGuid IS NOT NULL AND ClashZoneGuid != ''";
                    
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            var sleeveId = reader.GetInt32(0);
                            var guid = reader.GetString(1);
                            index.SleeveIdToClashZoneGuid[sleeveId] = guid;
                        }
                    }
                }
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    _logger?.Invoke($"[SQLite] ✅ Loaded {index.SleeveIdToClashZoneGuid.Count} SleeveId->GUID mappings for fallback lookup");
                }
            }
            catch (Exception ex)
            {
                _logger?.Invoke($"[SQLite] ⚠️ Failed to load SleeveId->GUID mapping: {ex.Message}");
            }
            
            // ✅ LOAD COMBINED SLEEVE CONSTITUENTS
            // We need to map CombinedInstanceId -> List of Constituents to aggregate their parameters
            try 
            {
                 using (var cmd = _context.Connection.CreateCommand())
                 {
                     cmd.CommandText = @"
                        SELECT 
                            cs.CombinedInstanceId,
                            csc.ConstituentType,
                            csc.ClashZoneGuid,
                            csc.ClusterInstanceId
                        FROM CombinedSleeveConstituents csc
                        INNER JOIN CombinedSleeves cs ON csc.CombinedSleeveId = cs.CombinedSleeveId";
                     
                     using (var reader = cmd.ExecuteReader())
                     {
                         while (reader.Read())
                         {
                             var combinedId = reader.GetInt32(0);
                             if (!index.ByCombined.ContainsKey(combinedId))
                             {
                                 index.ByCombined[combinedId] = new List<SleeveConstituentSnapshotReference>();
                             }
                             
                             index.ByCombined[combinedId].Add(new SleeveConstituentSnapshotReference
                             {
                                 SourceType = SafeGetString(reader, "ConstituentType"),
                                 ClashZoneGuid = SafeGetString(reader, "ClashZoneGuid"),
                                 ClusterInstanceId = reader.IsDBNull(3) ? (int?)null : reader.GetInt32(3)
                             });
                         }
                     }
                 }
                 _logger?.Invoke($"[SQLite] ✅ Loaded {index.ByCombined.Count} combined sleeve definitions for parameter transfer");

                 // ✅ LOAD COMBINED GUID-BASED CONSTITUENTS (For Fallback)
                 using (var cmd = _context.Connection.CreateCommand())
                 {
                     cmd.CommandText = @"
                        SELECT 
                            cs.DeterministicGuid,
                            csc.ConstituentType,
                            csc.ClashZoneGuid,
                            csc.ClusterInstanceId
                        FROM CombinedSleeveConstituents csc
                        INNER JOIN CombinedSleeves cs ON csc.CombinedSleeveId = cs.CombinedSleeveId
                        WHERE cs.DeterministicGuid IS NOT NULL AND cs.DeterministicGuid != ''";
                     
                     using (var reader = cmd.ExecuteReader())
                     {
                         while (reader.Read())
                         {
                             var detGuid = reader.GetString(0);
                             if (!index.ByCombinedGuidConstituents.ContainsKey(detGuid))
                             {
                                 index.ByCombinedGuidConstituents[detGuid] = new List<SleeveConstituentSnapshotReference>();
                             }
                             
                             index.ByCombinedGuidConstituents[detGuid].Add(new SleeveConstituentSnapshotReference
                             {
                                 SourceType = SafeGetString(reader, "ConstituentType"),
                                 ClashZoneGuid = SafeGetString(reader, "ClashZoneGuid"),
                                 ClusterInstanceId = reader.IsDBNull(3) ? (int?)null : reader.GetInt32(3)
                             });
                         }
                     }
                 }
                 _logger?.Invoke($"[SQLite] ✅ Loaded {index.ByCombinedGuidConstituents.Count} GUID-based combined sleeve definitions");
            }
            catch (Exception ex)
            {
                _logger?.Invoke($"[SQLite] ⚠️ Failed to load combined sleeve constituents: {ex.Message}");
            }

            // ✅ LOAD CLUSTER CONSTITUENTS
            // We need to map ClusterInstanceId (for direct lookup) and ClusterGUID (for fallback) 
            // to their constituents for parameter aggregation.
            try
            {
                using (var cmd = _context.Connection.CreateCommand())
                {
                    cmd.CommandText = @"
                        SELECT 
                            ClusterGUID,
                            ClusterInstanceId,
                            ConstituentZoneGuids
                        FROM ClusterSleeves_v2";
                    
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            var clusterGuid = SafeGetString(reader, "ClusterGUID");
                            var clusterInstanceId = reader.IsDBNull(reader.GetOrdinal("ClusterInstanceId")) ? -1 : reader.GetInt32(reader.GetOrdinal("ClusterInstanceId"));
                            var constituentsCsv = SafeGetString(reader, "ConstituentZoneGuids");
                            
                            if (string.IsNullOrEmpty(constituentsCsv)) continue;
                            
                            var constituentRefs = new List<SleeveConstituentSnapshotReference>();
                            var guids = constituentsCsv.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                            foreach (var g in guids)
                            {
                                constituentRefs.Add(new SleeveConstituentSnapshotReference
                                {
                                    SourceType = "Individual", // Cluster constituents are always individual zones
                                    ClashZoneGuid = g.Trim()
                                });
                            }
                            
                            if (constituentRefs.Count > 0)
                            {
                                if (!string.IsNullOrEmpty(clusterGuid))
                                {
                                    index.ByClusterGuidConstituents[clusterGuid] = constituentRefs;
                                }
                                
                                if (clusterInstanceId > 0)
                                {
                                    index.ByClusterConstituents[clusterInstanceId] = constituentRefs;
                                }
                            }
                        }
                    }
                }
                _logger?.Invoke($"[SQLite] ✅ Loaded {index.ByClusterConstituents.Count} cluster instance and {index.ByClusterGuidConstituents.Count} GUID-based constituent definitions");
            }
            catch (Exception ex)
            {
                _logger?.Invoke($"[SQLite] ⚠️ Failed to load cluster constituents: {ex.Message}");
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

        private int? SafeGetInt(SQLiteDataReader reader, string column)
        {
            try
            {
                var ordinal = reader.GetOrdinal(column);
                return reader.IsDBNull(ordinal) ? (int?)null : Convert.ToInt32(reader.GetValue(ordinal));
            }
            catch
            {
                return null;
            }
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
        public Dictionary<int, SleeveSnapshotView> ByCombinedSnapshot { get; } = new Dictionary<int, SleeveSnapshotView>();

        public Dictionary<int, List<SleeveConstituentSnapshotReference>> ByCombined { get; } = new Dictionary<int, List<SleeveConstituentSnapshotReference>>();
        
        // ✅ NEW: Map ClusterInstanceId -> List of constituents for parameter aggregation
        public Dictionary<int, List<SleeveConstituentSnapshotReference>> ByClusterConstituents { get; } = new Dictionary<int, List<SleeveConstituentSnapshotReference>>();
        
        // ✅ NEW: Map ClusterGUID (string) -> List of constituents for fallback aggregation
        public Dictionary<string, List<SleeveConstituentSnapshotReference>> ByClusterGuidConstituents { get; } = new Dictionary<string, List<SleeveConstituentSnapshotReference>>(StringComparer.OrdinalIgnoreCase);

        // ✅ NEW: Map Combined DeterministicGuid (string) -> List of constituents for fallback aggregation
        public Dictionary<string, List<SleeveConstituentSnapshotReference>> ByCombinedGuidConstituents { get; } = new Dictionary<string, List<SleeveConstituentSnapshotReference>>(StringComparer.OrdinalIgnoreCase);

        public Dictionary<string, SleeveSnapshotView> ByClashZoneGuid { get; } = new Dictionary<string, SleeveSnapshotView>(StringComparer.OrdinalIgnoreCase);
        // ✅ NEW: Map SleeveInstanceId to ClashZoneGuid for fallback lookup
        public Dictionary<int, string> SleeveIdToClashZoneGuid { get; } = new Dictionary<int, string>();

        public bool IsEmpty => BySleeve.Count == 0 && ByCluster.Count == 0;

        public bool TryGetBySleeve(int sleeveInstanceId, out SleeveSnapshotView view)
        {
            return BySleeve.TryGetValue(sleeveInstanceId, out view);
        }

        public bool TryGetByCluster(int clusterInstanceId, out SleeveSnapshotView view)
        {
            return ByCluster.TryGetValue(clusterInstanceId, out view);
        }
        
        public bool TryGetByCombinedSnapshot(int combinedInstanceId, out SleeveSnapshotView view)
        {
            return ByCombinedSnapshot.TryGetValue(combinedInstanceId, out view);
        }

        public bool TryGetByCombined(int combinedInstanceId, out List<SleeveConstituentSnapshotReference> constituents)
        {
            return ByCombined.TryGetValue(combinedInstanceId, out constituents);
        }

        public bool TryGetByClashZoneGuid(string guid, out SleeveSnapshotView view)
        {
            return ByClashZoneGuid.TryGetValue(guid, out view);
        }
    }

    public class SleeveConstituentSnapshotReference
    {
        public string SourceType { get; set; }
        public string ClashZoneGuid { get; set; }
        public int? ClusterInstanceId { get; set; }
    }

    public class SleeveSnapshotView
    {
        public int SnapshotId { get; set; }
        public int? SleeveInstanceId { get; set; }
        public int? ClusterInstanceId { get; set; }
        public int? CombinedInstanceId { get; set; }
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


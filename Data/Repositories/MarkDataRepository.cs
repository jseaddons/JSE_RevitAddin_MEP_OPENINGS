using System;
using System.Collections.Generic;
using System.Data.SQLite;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Data.Repositories
{
    public class MarkDataRepository : IMarkDataRepository
    {
        private readonly SleeveDbContext _context;

        public MarkDataRepository(SleeveDbContext context)
        {
            _context = context;
        }

        public List<ClashZone> GetMarkableClashZones(string category)
        {
            var results = new List<ClashZone>();
            string sql = @"
                SELECT cz.ClashZoneId, cz.SleeveInstanceId, COALESCE(cz.MepCategory, cs.Category) AS MepCategory, cz.ClusterInstanceId, 
                       COALESCE(ss_cluster.MepParametersJson, ss_ind.MepParametersJson) AS ParamsJson,
                       cz.MepSystemType, cz.MepServiceType
                FROM ClashZones cz
                LEFT JOIN ClusterSleeves cs ON cz.ClusterInstanceId = cs.ClusterInstanceId
                LEFT JOIN SleeveSnapshots ss_cluster ON (cz.ClusterInstanceId > 0 AND cz.ClusterInstanceId = ss_cluster.ClusterInstanceId)
                LEFT JOIN SleeveSnapshots ss_ind ON (cz.SleeveInstanceId = ss_ind.SleeveInstanceId)
                WHERE (cz.MepCategory = @cat OR cs.Category = @cat)
                AND (cz.SleeveInstanceId > 0 OR cz.ClusterInstanceId > 0)";

            try
            {
                using (var cmd = _context.Connection.CreateCommand())
                {
                    cmd.CommandText = sql;
                    cmd.Parameters.AddWithValue("@cat", category);
                    RemarkDebugLogger.LogInfo($"Executing GetMarkableClashZones SQL for category: '{category}'");

                    using (var reader = cmd.ExecuteReader())
                    {
                        int rawCount = 0;

                        while (reader.Read())
                        {
                            int sleeveId = reader.GetInt32(1);
                            int clusterId = reader.IsDBNull(3) ? -1 : reader.GetInt32(3);

                            rawCount++;

                            var cz = new ClashZone
                            {
                                ClashZoneId = reader.GetInt32(0),
                                SleeveInstanceId = sleeveId,
                                MepElementCategory = reader.GetString(2),
                                ClusterInstanceId = clusterId,
                                MepSystemType = reader.IsDBNull(5) ? string.Empty : reader.GetString(5),
                                MepServiceType = reader.IsDBNull(6) ? string.Empty : reader.GetString(6)
                            };

                            // Load Snapshot if available
                            if (!reader.IsDBNull(4))
                            {
                                string json = reader.GetString(4);
                                if (!string.IsNullOrEmpty(json))
                                {
                                    try
                                    {
                                        var mepParams = System.Text.Json.JsonSerializer.Deserialize<List<SerializableKeyValue>>(json);
                                        if (mepParams != null)
                                        {
                                            cz.MepParameterValues = mepParams;
                                        }
                                    }
                                    catch { /* Ignore JSON errors */ }
                                }
                            }
                            results.Add(cz);
                        }
                        RemarkDebugLogger.LogInfo($"SQL returned {rawCount} raw rows, {results.Count} unique markable zones for '{category}'");
                    }
                }
            }
            catch (Exception ex)
            {
                // We let the service decide how to log, or log here to file?
                // For now, rethrow or log to SafeFileLogger if critical.
                string logPath = SafeFileLogger.GetLogFilePath("mepmark_debug.log");
                System.IO.File.AppendAllText(logPath, $"[MarkDataRepository] SQL Error: {ex.Message}\n");
                throw;
            }

            return results;
        }

        public List<ClashZone> GetSleevesForLevel(string levelName, string category)
        {
            var results = new List<ClashZone>();
            string sql = @"
                SELECT cz.ClashZoneId, cz.SleeveInstanceId, COALESCE(cz.MepCategory, cs.Category) AS MepCategory, 
                       cz.ClusterInstanceId, COALESCE(ss_cluster.MepParametersJson, ss_ind.MepParametersJson) AS ParamsJson,
                       cz.IntersectionPointX, cz.IntersectionPointY, cz.IntersectionPointZ,
                       cz.MepSystemType, cz.MepServiceType
                FROM ClashZones cz
                LEFT JOIN ClusterSleeves cs ON cz.ClusterInstanceId = cs.ClusterInstanceId
                LEFT JOIN SleeveSnapshots ss_cluster ON (cz.ClusterInstanceId > 0 AND cz.ClusterInstanceId = ss_cluster.ClusterInstanceId)
                LEFT JOIN SleeveSnapshots ss_ind ON (cz.SleeveInstanceId = ss_ind.SleeveInstanceId)
                WHERE cz.MepElementLevelName = @LevelName 
                AND (cz.SleeveInstanceId > 0 OR cz.ClusterInstanceId > 0)";

            bool isCombinedQuery = category != null && category.Equals("Combined", StringComparison.OrdinalIgnoreCase);

            if (isCombinedQuery)
            {
                sql += " AND cz.CombinedClusterSleeveInstanceId > 0";
            }
            else if (!string.IsNullOrWhiteSpace(category))
            {
                sql += " AND (cz.MepCategory = @cat OR cs.Category = @cat) AND (cz.CombinedClusterSleeveInstanceId IS NULL OR cz.CombinedClusterSleeveInstanceId <= 0)";
            }
            else
            {
                sql += " AND (cz.CombinedClusterSleeveInstanceId IS NULL OR cz.CombinedClusterSleeveInstanceId <= 0)";
            }

            try
            {
                using (var cmd = _context.Connection.CreateCommand())
                {
                    cmd.CommandText = sql;
                    cmd.Parameters.AddWithValue("@LevelName", levelName);
                    if (!isCombinedQuery && !string.IsNullOrWhiteSpace(category))
                    {
                        cmd.Parameters.AddWithValue("@cat", category);
                    }

                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            int sleeveId = reader.GetInt32(1);
                            int clusterId = reader.IsDBNull(3) ? -1 : reader.GetInt32(3);

                            var cz = new ClashZone
                            {
                                ClashZoneId = reader.GetInt32(0),
                                SleeveInstanceId = sleeveId,
                                MepElementCategory = reader.GetString(2),
                                ClusterInstanceId = clusterId,
                                IntersectionPointX = reader.GetDouble(5),
                                IntersectionPointY = reader.GetDouble(6),
                                IntersectionPointZ = reader.GetDouble(7),
                                MepSystemType = reader.IsDBNull(8) ? string.Empty : reader.GetString(8),
                                MepServiceType = reader.IsDBNull(9) ? string.Empty : reader.GetString(9)
                            };

                            if (!reader.IsDBNull(4))
                            {
                                string json = reader.GetString(4);
                                try
                                {
                                    var values = System.Text.Json.JsonSerializer.Deserialize<List<JSE_RevitAddin_MEP_OPENINGS.Models.SerializableKeyValue>>(json);
                                    if (values != null) cz.MepParameterValues = values;
                                }
                                catch { }
                            }
                            results.Add(cz);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                string logPath = SafeFileLogger.GetLogFilePath("mepmark_debug.log");
                System.IO.File.AppendAllText(logPath, $"[MarkDataRepository] GetSleevesForLevel SQL Error: {ex.Message}\n");
            }

            return results;
        }

        public Dictionary<int, string> GetCategoryLookup(IEnumerable<int> instanceIds)
        {
            var lookup = new Dictionary<int, string>();
            var idList = new List<int>(instanceIds);
            if (idList.Count == 0) return lookup;

            // Simple batch query
            string ids = string.Join(",", idList);
            string sql = $@"
                SELECT COALESCE(cz.SleeveInstanceId, -cz.ClusterInstanceId) AS TrackingId, 
                       COALESCE(cz.MepCategory, cs.Category) AS MepCategory
                FROM ClashZones cz
                LEFT JOIN ClusterSleeves cs ON cz.ClusterInstanceId = cs.ClusterInstanceId
                WHERE cz.SleeveInstanceId IN ({ids}) OR cz.ClusterInstanceId IN ({ids})";

            try
            {
                using (var cmd = _context.Connection.CreateCommand())
                {
                    cmd.CommandText = sql;
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            int id = reader.GetInt32(0);
                            string category = reader.GetString(1);
                            lookup[Math.Abs(id)] = category;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                string logPath = SafeFileLogger.GetLogFilePath("mepmark_debug.log");
                System.IO.File.AppendAllText(logPath, $"[MarkDataRepository] Lookup SQL Error: {ex.Message}\n");
            }

            return lookup;
        }
    }
}

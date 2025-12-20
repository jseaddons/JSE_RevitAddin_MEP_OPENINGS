using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Linq;
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Data.Repositories
{
    /// <summary>
    /// Repository implementation for cross-category combined sleeve operations.
    /// Handles database persistence for combined sleeves and their constituents.
    /// </summary>
    public class CombinedSleeveRepository : ICombinedSleeveRepository
    {
        private readonly SleeveDbContext _context;
        private readonly Action<string> _logger;
        
        public CombinedSleeveRepository(SleeveDbContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            _logger = msg => DebugLogger.Info(msg);
        }
        
        // ============================================================================
        // CREATE OPERATIONS
        // ============================================================================
        
        public int SaveCombinedSleeve(CombinedSleeve combinedSleeve)
        {
            if (combinedSleeve == null)
                throw new ArgumentNullException(nameof(combinedSleeve));
            
            using (var transaction = _context.Connection.BeginTransaction())
            {
                try
                {
                    int id = SaveCombinedSleeveInternal(combinedSleeve, transaction);
                    transaction.Commit();
                    _logger($"[SQLite] ✅ Saved combined sleeve {id}");
                    return id;
                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    _logger($"[SQLite] ❌ Error saving combined sleeve: {ex.Message}");
                    DatabaseOperationLogger.LogOperation("ROLLBACK", "CombinedSleeves", null, 0, $"Error: {ex.Message}");
                    throw;
                }
            }
        }

        public void SaveCombinedSleevesBatch(List<CombinedSleeve> combinedSleeves)
        {
            if (combinedSleeves == null || combinedSleeves.Count == 0) return;

            using (var transaction = _context.Connection.BeginTransaction())
            {
                try
                {
                    int count = 0;
                    foreach (var sleeve in combinedSleeves)
                    {
                        // 1. Save Sleeve
                        int id = SaveCombinedSleeveInternal(sleeve, transaction);
                        sleeve.CombinedSleeveId = id;

                        // 2. Mark Constituents (Logic duplicated/inlined for transactional safety)
                        if (sleeve.Constituents != null)
                        {
                            foreach (var constituent in sleeve.Constituents)
                            {
                                if (constituent.Type == ConstituentType.Individual && constituent.ClashZoneGuid.HasValue)
                                {
                                    using (var cmd = _context.Connection.CreateCommand())
                                    {
                                        cmd.Transaction = transaction;
                                        cmd.CommandText = @"UPDATE ClashZones 
                                                            SET IsCombinedResolved = 1, 
                                                                IsResolvedFlag = 0, 
                                                                IsClusterResolvedFlag = 0, 
                                                                SleeveInstanceId = -1, 
                                                                ClusterInstanceId = -1, 
                                                                CombinedClusterSleeveInstanceId = @CId 
                                                            WHERE ClashZoneGuid = @Guid";
                                        cmd.Parameters.AddWithValue("@Guid", constituent.ClashZoneGuid.Value.ToString());
                                        cmd.Parameters.AddWithValue("@CId", sleeve.CombinedInstanceId);
                                        cmd.ExecuteNonQuery();
                                    }
                                }
                                else if (constituent.Type == ConstituentType.Cluster && constituent.ClusterInstanceId.HasValue)
                                {
                                    // ✅ FIX: Update CombinedClusterSleeveInstanceId for parameter transfer lookup
                                    // NOTE: IsCombinedResolved is NOT in ClusterSleeves table - only in ClashZones
                                    using (var cmd = _context.Connection.CreateCommand())
                                    {
                                        cmd.Transaction = transaction;
                                        cmd.CommandText = @"UPDATE ClusterSleeves SET CombinedClusterSleeveInstanceId = @CId WHERE ClusterInstanceId = @Id";
                                        cmd.Parameters.AddWithValue("@Id", constituent.ClusterInstanceId.Value);
                                        cmd.Parameters.AddWithValue("@CId", sleeve.CombinedInstanceId);
                                        cmd.ExecuteNonQuery();
                                    }
                                }
                            }
                        }
                        count++;
                    }
                    transaction.Commit();
                    _logger($"[SQLite] ✅ Batch saved {count} combined sleeves and updated flags.");
                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    _logger($"[SQLite] ❌ Batch save failed: {ex.Message}");
                    throw;
                }
            }
        }

        private int SaveCombinedSleeveInternal(CombinedSleeve combinedSleeve, SQLiteTransaction transaction)
        {
             int combinedSleeveId;
             using (var cmd = _context.Connection.CreateCommand())
             {
                 cmd.Transaction = transaction;
                 cmd.CommandText = @"
                            INSERT INTO CombinedSleeves (
                                CombinedInstanceId, ComboId, FilterId, Categories,
                                BoundingBoxMinX, BoundingBoxMinY, BoundingBoxMinZ,
                                BoundingBoxMaxX, BoundingBoxMaxY, BoundingBoxMaxZ,
                                CombinedWidth, CombinedHeight, CombinedDepth,
                                PlacementX, PlacementY, PlacementZ,
                                RotationAngleDeg, HostType, HostOrientation,
                                Corner1X, Corner1Y, Corner1Z,
                                Corner2X, Corner2Y, Corner2Z,
                                Corner3X, Corner3Y, Corner3Z,
                                Corner4X, Corner4Y, Corner4Z
                            ) VALUES (
                                @CombinedInstanceId, @ComboId, @FilterId, @Categories,
                                @BoundingBoxMinX, @BoundingBoxMinY, @BoundingBoxMinZ,
                                @BoundingBoxMaxX, @BoundingBoxMaxY, @BoundingBoxMaxZ,
                                @CombinedWidth, @CombinedHeight, @CombinedDepth,
                                @PlacementX, @PlacementY, @PlacementZ,
                                @RotationAngleDeg, @HostType, @HostOrientation,
                                @Corner1X, @Corner1Y, @Corner1Z,
                                @Corner2X, @Corner2Y, @Corner2Z,
                                @Corner3X, @Corner3Y, @Corner3Z,
                                @Corner4X, @Corner4Y, @Corner4Z
                            );
                            SELECT last_insert_rowid();";

                        // Add parameters
                        cmd.Parameters.AddWithValue("@CombinedInstanceId", combinedSleeve.CombinedInstanceId);
                        cmd.Parameters.AddWithValue("@ComboId", combinedSleeve.ComboId > 0 ? (object)combinedSleeve.ComboId : DBNull.Value);
                        cmd.Parameters.AddWithValue("@FilterId", combinedSleeve.FilterId > 0 ? (object)combinedSleeve.FilterId : DBNull.Value);
                        cmd.Parameters.AddWithValue("@Categories", string.Join(",", combinedSleeve.Categories));
                        
                        cmd.Parameters.AddWithValue("@BoundingBoxMinX", combinedSleeve.BoundingBoxMinX);
                        cmd.Parameters.AddWithValue("@BoundingBoxMinY", combinedSleeve.BoundingBoxMinY);
                        cmd.Parameters.AddWithValue("@BoundingBoxMinZ", combinedSleeve.BoundingBoxMinZ);
                        cmd.Parameters.AddWithValue("@BoundingBoxMaxX", combinedSleeve.BoundingBoxMaxX);
                        cmd.Parameters.AddWithValue("@BoundingBoxMaxY", combinedSleeve.BoundingBoxMaxY);
                        cmd.Parameters.AddWithValue("@BoundingBoxMaxZ", combinedSleeve.BoundingBoxMaxZ);
                        
                        cmd.Parameters.AddWithValue("@CombinedWidth", combinedSleeve.CombinedWidth);
                        cmd.Parameters.AddWithValue("@CombinedHeight", combinedSleeve.CombinedHeight);
                        cmd.Parameters.AddWithValue("@CombinedDepth", combinedSleeve.CombinedDepth);
                        
                        cmd.Parameters.AddWithValue("@PlacementX", combinedSleeve.PlacementX);
                        cmd.Parameters.AddWithValue("@PlacementY", combinedSleeve.PlacementY);
                        cmd.Parameters.AddWithValue("@PlacementZ", combinedSleeve.PlacementZ);
                        cmd.Parameters.AddWithValue("@RotationAngleDeg", combinedSleeve.RotationAngleDeg);
                        
                        cmd.Parameters.AddWithValue("@HostType", combinedSleeve.HostType ?? (object)DBNull.Value);
                        cmd.Parameters.AddWithValue("@HostOrientation", combinedSleeve.HostOrientation ?? (object)DBNull.Value);
                        
                        cmd.Parameters.AddWithValue("@Corner1X", combinedSleeve.Corner1X);
                        cmd.Parameters.AddWithValue("@Corner1Y", combinedSleeve.Corner1Y);
                        cmd.Parameters.AddWithValue("@Corner1Z", combinedSleeve.Corner1Z);
                        cmd.Parameters.AddWithValue("@Corner2X", combinedSleeve.Corner2X);
                        cmd.Parameters.AddWithValue("@Corner2Y", combinedSleeve.Corner2Y);
                        cmd.Parameters.AddWithValue("@Corner2Z", combinedSleeve.Corner2Z);
                        cmd.Parameters.AddWithValue("@Corner3X", combinedSleeve.Corner3X);
                        cmd.Parameters.AddWithValue("@Corner3Y", combinedSleeve.Corner3Y);
                        cmd.Parameters.AddWithValue("@Corner3Z", combinedSleeve.Corner3Z);
                        cmd.Parameters.AddWithValue("@Corner4X", combinedSleeve.Corner4X);
                        cmd.Parameters.AddWithValue("@Corner4Y", combinedSleeve.Corner4Y);
                        cmd.Parameters.AddWithValue("@Corner4Z", combinedSleeve.Corner4Z);

                        combinedSleeveId = Convert.ToInt32(cmd.ExecuteScalar());
             }
             
             if (combinedSleeve.Constituents != null && combinedSleeve.Constituents.Count > 0)
             {
                 SaveConstituentsInternal(combinedSleeveId, combinedSleeve.Constituents, transaction);
             }
             return combinedSleeveId;
        }
        
        public void SaveConstituents(int combinedSleeveId, List<SleeveConstituent> constituents)
        {
            using (var transaction = _context.Connection.BeginTransaction())
            {
                try
                {
                    SaveConstituentsInternal(combinedSleeveId, constituents, transaction);
                    transaction.Commit();
                    _logger($"[SQLite] ✅ Saved {constituents.Count} constituents for combined sleeve {combinedSleeveId}");
                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    _logger($"[SQLite] ❌ Error saving constituents: {ex.Message}");
                    throw;
                }
            }
        }
        
        private void SaveConstituentsInternal(int combinedSleeveId, List<SleeveConstituent> constituents, SQLiteTransaction transaction)
        {
            foreach (var constituent in constituents)
            {
                using (var cmd = _context.Connection.CreateCommand())
                {
                    cmd.Transaction = transaction;
                    cmd.CommandText = @"
                        INSERT INTO CombinedSleeveConstituents (
                            CombinedSleeveId, ConstituentType, Category,
                            ClashZoneId, ClashZoneGuid,
                            ClusterSleeveId, ClusterInstanceId
                        ) VALUES (
                            @CombinedSleeveId, @ConstituentType, @Category,
                            @ClashZoneId, @ClashZoneGuid,
                            @ClusterSleeveId, @ClusterInstanceId
                        )";
                    
                    cmd.Parameters.AddWithValue("@CombinedSleeveId", combinedSleeveId);
                    cmd.Parameters.AddWithValue("@ConstituentType", constituent.Type.ToString());
                    cmd.Parameters.AddWithValue("@Category", constituent.Category);
                    
                    cmd.Parameters.AddWithValue("@ClashZoneId", constituent.ClashZoneId ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@ClashZoneGuid", constituent.ClashZoneGuid?.ToString() ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@ClusterSleeveId", constituent.ClusterSleeveId ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@ClusterInstanceId", constituent.ClusterInstanceId ?? (object)DBNull.Value);
                    
                    cmd.ExecuteNonQuery();
                }
            }
            
            DatabaseOperationLogger.LogOperation("INSERT", "CombinedSleeveConstituents",
                new Dictionary<string, object>
                {
                    { "CombinedSleeveId", combinedSleeveId },
                    { "Count", constituents.Count }
                });
        }
        
        // ============================================================================
        // READ OPERATIONS
        // ============================================================================
        
        public CombinedSleeve GetCombinedSleeveById(int combinedSleeveId)
        {
            using (var cmd = _context.Connection.CreateCommand())
            {
                cmd.CommandText = @"
                    SELECT * FROM CombinedSleeves 
                    WHERE CombinedSleeveId = @CombinedSleeveId";
                
                cmd.Parameters.AddWithValue("@CombinedSleeveId", combinedSleeveId);
                
                using (var reader = cmd.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        var combinedSleeve = MapReaderToCombinedSleeve(reader);
                        combinedSleeve.Constituents = GetConstituents(combinedSleeveId);
                        return combinedSleeve;
                    }
                }
            }
            
            return null;
        }
        
        public CombinedSleeve GetCombinedSleeveByInstanceId(int instanceId)
        {
            using (var cmd = _context.Connection.CreateCommand())
            {
                cmd.CommandText = @"
                    SELECT * FROM CombinedSleeves 
                    WHERE CombinedInstanceId = @CombinedInstanceId";
                
                cmd.Parameters.AddWithValue("@CombinedInstanceId", instanceId);
                
                using (var reader = cmd.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        var combinedSleeve = MapReaderToCombinedSleeve(reader);
                        combinedSleeve.Constituents = GetConstituents(combinedSleeve.CombinedSleeveId);
                        return combinedSleeve;
                    }
                }
            }
            
            return null;
        }
        
        public List<CombinedSleeve> GetCombinedSleevesForCombo(int comboId, int filterId)
        {
            var combinedSleeves = new List<CombinedSleeve>();
            
            using (var cmd = _context.Connection.CreateCommand())
            {
                cmd.CommandText = @"
                    SELECT * FROM CombinedSleeves 
                    WHERE ComboId = @ComboId AND FilterId = @FilterId";
                
                cmd.Parameters.AddWithValue("@ComboId", comboId);
                cmd.Parameters.AddWithValue("@FilterId", filterId);
                
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        var combinedSleeve = MapReaderToCombinedSleeve(reader);
                        combinedSleeve.Constituents = GetConstituents(combinedSleeve.CombinedSleeveId);
                        combinedSleeves.Add(combinedSleeve);
                    }
                }
            }
            
            return combinedSleeves;
        }
        
        public List<SleeveConstituent> GetConstituents(int combinedSleeveId)
        {
            var constituents = new List<SleeveConstituent>();
            
            using (var cmd = _context.Connection.CreateCommand())
            {
                cmd.CommandText = @"
                    SELECT * FROM CombinedSleeveConstituents 
                    WHERE CombinedSleeveId = @CombinedSleeveId";
                
                cmd.Parameters.AddWithValue("@CombinedSleeveId", combinedSleeveId);
                
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        constituents.Add(MapReaderToConstituent(reader));
                    }
                }
            }
            
            return constituents;
        }
        
        // ============================================================================
        // UPDATE OPERATIONS
        // ============================================================================
        
        public void UpdateCombinedSleeveCorners(int combinedSleeveId,
            double c1x, double c1y, double c1z,
            double c2x, double c2y, double c2z,
            double c3x, double c3y, double c3z,
            double c4x, double c4y, double c4z)
        {
            using (var transaction = _context.Connection.BeginTransaction())
            {
                try
                {
                    using (var cmd = _context.Connection.CreateCommand())
                    {
                        cmd.Transaction = transaction;
                        cmd.CommandText = @"
                            UPDATE CombinedSleeves SET
                                Corner1X = @Corner1X, Corner1Y = @Corner1Y, Corner1Z = @Corner1Z,
                                Corner2X = @Corner2X, Corner2Y = @Corner2Y, Corner2Z = @Corner2Z,
                                Corner3X = @Corner3X, Corner3Y = @Corner3Y, Corner3Z = @Corner3Z,
                                Corner4X = @Corner4X, Corner4Y = @Corner4Y, Corner4Z = @Corner4Z,
                                UpdatedAt = CURRENT_TIMESTAMP
                            WHERE CombinedSleeveId = @CombinedSleeveId";
                        
                        cmd.Parameters.AddWithValue("@CombinedSleeveId", combinedSleeveId);
                        cmd.Parameters.AddWithValue("@Corner1X", c1x);
                        cmd.Parameters.AddWithValue("@Corner1Y", c1y);
                        cmd.Parameters.AddWithValue("@Corner1Z", c1z);
                        cmd.Parameters.AddWithValue("@Corner2X", c2x);
                        cmd.Parameters.AddWithValue("@Corner2Y", c2y);
                        cmd.Parameters.AddWithValue("@Corner2Z", c2z);
                        cmd.Parameters.AddWithValue("@Corner3X", c3x);
                        cmd.Parameters.AddWithValue("@Corner3Y", c3y);
                        cmd.Parameters.AddWithValue("@Corner3Z", c3z);
                        cmd.Parameters.AddWithValue("@Corner4X", c4x);
                        cmd.Parameters.AddWithValue("@Corner4Y", c4y);
                        cmd.Parameters.AddWithValue("@Corner4Z", c4z);
                        
                        var rowsAffected = cmd.ExecuteNonQuery();
                        
                        if (rowsAffected == 0)
                        {
                            _logger($"[SQLite] ⚠️ UpdateCombinedSleeveCorners: No rows updated for CombinedSleeveId {combinedSleeveId}");
                        }
                        else
                        {
                            DatabaseOperationLogger.LogOperation("UPDATE", "CombinedSleeves",
                                new Dictionary<string, object>
                                {
                                    { "CombinedSleeveId", combinedSleeveId },
                                    { "Corner1", $"({c1x:F2},{c1y:F2},{c1z:F2})" }
                                });
                            _logger($"[SQLite] ✅ Updated corners for CombinedSleeveId {combinedSleeveId}");
                        }
                    }
                    
                    transaction.Commit();
                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    _logger($"[SQLite] ❌ Error updating corners for combined sleeve {combinedSleeveId}: {ex.Message}");
                    throw;
                }
            }
        }
        
        // ============================================================================
        // DELETE OPERATIONS
        // ============================================================================
        
        public void DeleteCombinedSleeve(int combinedSleeveId)
        {
            using (var transaction = _context.Connection.BeginTransaction())
            {
                try
                {
                    using (var cmd = _context.Connection.CreateCommand())
                    {
                        cmd.Transaction = transaction;
                        cmd.CommandText = @"
                            DELETE FROM CombinedSleeves 
                            WHERE CombinedSleeveId = @CombinedSleeveId";
                        
                        cmd.Parameters.AddWithValue("@CombinedSleeveId", combinedSleeveId);
                        
                        var rowsAffected = cmd.ExecuteNonQuery();
                        
                        DatabaseOperationLogger.LogOperation("DELETE", "CombinedSleeves",
                            new Dictionary<string, object>
                            {
                                { "CombinedSleeveId", combinedSleeveId },
                                { "RowsAffected", rowsAffected }
                            });
                    }
                    
                    transaction.Commit();
                    _logger($"[SQLite] ✅ Deleted combined sleeve {combinedSleeveId}");
                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    _logger($"[SQLite] ❌ Error deleting combined sleeve {combinedSleeveId}: {ex.Message}");
                    throw;
                }
            }
        }
        
        // ============================================================================
        // FLAG OPERATIONS
        // ============================================================================
        
        public void MarkConstituentsAsResolved(List<SleeveConstituent> constituents, int combinedInstanceId)
        {
            using (var transaction = _context.Connection.BeginTransaction())
            {
                try
                {
                    int updatedCount = 0;
                    foreach (var constituent in constituents)
                    {
                        if (constituent.Type == ConstituentType.Individual && constituent.ClashZoneGuid.HasValue)
                        {
                            // Update ClashZones.IsCombinedResolved AND CombinedClusterSleeveInstanceId
                            using (var cmd = _context.Connection.CreateCommand())
                            {
                                cmd.Transaction = transaction;
                                cmd.CommandText = @"
                                    UPDATE ClashZones 
                                    SET IsCombinedResolved = 1,
                                        IsResolvedFlag = 0,
                                        IsClusterResolvedFlag = 0,
                                        SleeveInstanceId = -1,
                                        ClusterInstanceId = -1,
                                        CombinedClusterSleeveInstanceId = @CombinedInstanceId
                                    WHERE ClashZoneGuid = @ClashZoneGuid";
                                
                                cmd.Parameters.AddWithValue("@ClashZoneGuid", constituent.ClashZoneGuid.Value.ToString());
                                cmd.Parameters.AddWithValue("@CombinedInstanceId", combinedInstanceId);
                                int rows = cmd.ExecuteNonQuery();
                                if (rows > 0)
                                {
                                    updatedCount++;
                                    _logger($"[SQLite]   -> Flag set for ClashZone {constituent.ClashZoneGuid}");
                                }
                            }
                        }
                        else if (constituent.Type == ConstituentType.Cluster && constituent.ClusterInstanceId.HasValue)
                        {
                            // ✅ FIX: Update CombinedClusterSleeveInstanceId for parameter transfer lookup
                            // NOTE: IsCombinedResolved is NOT in ClusterSleeves table - only in ClashZones
                            using (var cmd = _context.Connection.CreateCommand())
                            {
                                cmd.Transaction = transaction;
                                cmd.CommandText = @"
                                    UPDATE ClusterSleeves 
                                    SET CombinedClusterSleeveInstanceId = @CombinedInstanceId
                                    WHERE ClusterInstanceId = @ClusterInstanceId";
                                
                                cmd.Parameters.AddWithValue("@ClusterInstanceId", constituent.ClusterInstanceId.Value);
                                cmd.Parameters.AddWithValue("@CombinedInstanceId", combinedInstanceId);
                                int rows = cmd.ExecuteNonQuery();
                                if (rows > 0) 
                                {
                                    updatedCount++;
                                    _logger($"[SQLite]   -> CombinedInstanceId set for ClusterSleeve {constituent.ClusterInstanceId}");
                                }
                            }
                        }
                        else
                        {
                            _logger($"[SQLite]   ⚠️ Skipping flag update for constituent: Type={constituent.Type}, No ID available");
                        }
                    }
                    
                    transaction.Commit();
                    _logger($"[SQLite] ✅ Successfully marked {updatedCount}/{constituents.Count} constituents as resolved (CombinedID={combinedInstanceId})");
                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    _logger($"[SQLite] ❌ Error marking constituents as resolved: {ex.Message}");
                    throw;
                }
            }
        }
        
        // ============================================================================
        // HELPER METHODS
        // ============================================================================
        
        private CombinedSleeve MapReaderToCombinedSleeve(SQLiteDataReader reader)
        {
            return new CombinedSleeve
            {
                CombinedSleeveId = reader.GetInt32(reader.GetOrdinal("CombinedSleeveId")),
                CombinedInstanceId = reader.GetInt32(reader.GetOrdinal("CombinedInstanceId")),
                ComboId = reader.GetInt32(reader.GetOrdinal("ComboId")),
                FilterId = reader.GetInt32(reader.GetOrdinal("FilterId")),
                Categories = reader.GetString(reader.GetOrdinal("Categories")).Split(',').ToList(),
                
                BoundingBoxMinX = reader.GetDouble(reader.GetOrdinal("BoundingBoxMinX")),
                BoundingBoxMinY = reader.GetDouble(reader.GetOrdinal("BoundingBoxMinY")),
                BoundingBoxMinZ = reader.GetDouble(reader.GetOrdinal("BoundingBoxMinZ")),
                BoundingBoxMaxX = reader.GetDouble(reader.GetOrdinal("BoundingBoxMaxX")),
                BoundingBoxMaxY = reader.GetDouble(reader.GetOrdinal("BoundingBoxMaxY")),
                BoundingBoxMaxZ = reader.GetDouble(reader.GetOrdinal("BoundingBoxMaxZ")),
                
                CombinedWidth = reader.GetDouble(reader.GetOrdinal("CombinedWidth")),
                CombinedHeight = reader.GetDouble(reader.GetOrdinal("CombinedHeight")),
                CombinedDepth = reader.GetDouble(reader.GetOrdinal("CombinedDepth")),
                
                PlacementX = reader.GetDouble(reader.GetOrdinal("PlacementX")),
                PlacementY = reader.GetDouble(reader.GetOrdinal("PlacementY")),
                PlacementZ = reader.GetDouble(reader.GetOrdinal("PlacementZ")),
                RotationAngleDeg = reader.GetDouble(reader.GetOrdinal("RotationAngleDeg")),
                
                HostType = reader.IsDBNull(reader.GetOrdinal("HostType")) ? null : reader.GetString(reader.GetOrdinal("HostType")),
                HostOrientation = reader.IsDBNull(reader.GetOrdinal("HostOrientation")) ? null : reader.GetString(reader.GetOrdinal("HostOrientation")),
                
                Corner1X = reader.GetDouble(reader.GetOrdinal("Corner1X")),
                Corner1Y = reader.GetDouble(reader.GetOrdinal("Corner1Y")),
                Corner1Z = reader.GetDouble(reader.GetOrdinal("Corner1Z")),
                Corner2X = reader.GetDouble(reader.GetOrdinal("Corner2X")),
                Corner2Y = reader.GetDouble(reader.GetOrdinal("Corner2Y")),
                Corner2Z = reader.GetDouble(reader.GetOrdinal("Corner2Z")),
                Corner3X = reader.GetDouble(reader.GetOrdinal("Corner3X")),
                Corner3Y = reader.GetDouble(reader.GetOrdinal("Corner3Y")),
                Corner3Z = reader.GetDouble(reader.GetOrdinal("Corner3Z")),
                Corner4X = reader.GetDouble(reader.GetOrdinal("Corner4X")),
                Corner4Y = reader.GetDouble(reader.GetOrdinal("Corner4Y")),
                Corner4Z = reader.GetDouble(reader.GetOrdinal("Corner4Z")),
                
                CreatedAt = reader.GetDateTime(reader.GetOrdinal("CreatedAt")),
                UpdatedAt = reader.GetDateTime(reader.GetOrdinal("UpdatedAt"))
            };
        }
        
        private SleeveConstituent MapReaderToConstituent(SQLiteDataReader reader)
        {
            return new SleeveConstituent
            {
                ConstituentId = reader.GetInt32(reader.GetOrdinal("ConstituentId")),
                CombinedSleeveId = reader.GetInt32(reader.GetOrdinal("CombinedSleeveId")),
                Type = (ConstituentType)Enum.Parse(typeof(ConstituentType), reader.GetString(reader.GetOrdinal("ConstituentType"))),
                Category = reader.GetString(reader.GetOrdinal("Category")),
                
                ClashZoneId = reader.IsDBNull(reader.GetOrdinal("ClashZoneId")) ? (int?)null : reader.GetInt32(reader.GetOrdinal("ClashZoneId")),
                ClashZoneGuid = reader.IsDBNull(reader.GetOrdinal("ClashZoneGuid")) ? (Guid?)null : Guid.Parse(reader.GetString(reader.GetOrdinal("ClashZoneGuid"))),
                ClusterSleeveId = reader.IsDBNull(reader.GetOrdinal("ClusterSleeveId")) ? (int?)null : reader.GetInt32(reader.GetOrdinal("ClusterSleeveId")),
                ClusterInstanceId = reader.IsDBNull(reader.GetOrdinal("ClusterInstanceId")) ? (int?)null : reader.GetInt32(reader.GetOrdinal("ClusterInstanceId")),
                
                CreatedAt = reader.GetDateTime(reader.GetOrdinal("CreatedAt"))
            };
        }
    }
}

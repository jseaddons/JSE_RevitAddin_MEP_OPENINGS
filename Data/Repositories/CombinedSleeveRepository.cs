using System;
using System.Collections.Generic;
#if !NET8_0_OR_GREATER
using System.Data.SQLite;
#endif
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
        
        public long SaveCombinedSleeve(CombinedSleeve combinedSleeve)
        {
            if (combinedSleeve == null)
                throw new ArgumentNullException(nameof(combinedSleeve));
            
            using (var transaction = _context.Connection.BeginTransaction())
            {
                try
                {
                    long id = SaveCombinedSleeveInternal(combinedSleeve, transaction);
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
                        long id = SaveCombinedSleeveInternal(sleeve, transaction);
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
                                        // ✅ CRITICAL FIX: Do NOT reset IsResolvedFlag or IsClusterResolvedFlag to 0!
                                        // Keep them at 1 (maintaining hierarchy). Only refresh should reset when combined is deleted.
                                        cmd.CommandText = @"UPDATE ClashZones 
                                                            SET IsCombinedResolved = 1, 
                                                                CombinedClusterSleeveInstanceId = @CId 
                                                            WHERE ClashZoneGuid = @Guid";
                                        cmd.Parameters.AddWithValue("@Guid", constituent.ClashZoneGuid.Value.ToString());
                                        cmd.Parameters.AddWithValue("@CId", sleeve.CombinedInstanceId);
                                        cmd.ExecuteNonQuery();
                                    }
                                }
                                else if (constituent.Type == ConstituentType.Cluster && constituent.ClusterInstanceId.HasValue)
                                {
                                    // ✅ FIX: Update CombinedClusterSleeveInstanceId in ClusterSleeves table for parameter transfer lookup
                                    using (var cmd = _context.Connection.CreateCommand())
                                    {
                                        cmd.Transaction = transaction;
                                        cmd.CommandText = @"UPDATE ClusterSleeves SET CombinedClusterSleeveInstanceId = @CId WHERE ClusterInstanceId = @Id";
                                        cmd.Parameters.AddWithValue("@Id", constituent.ClusterInstanceId.Value);
                                        cmd.Parameters.AddWithValue("@CId", sleeve.CombinedInstanceId);
                                        cmd.ExecuteNonQuery();
                                    }
                                    
                                    // ✅ CRITICAL FIX: Also update all ClashZones rows that belong to this cluster
                                    // NOTE: Do NOT reset ClusterInstanceId - maintain the link to cluster data!
                                    // NOTE: Do NOT reset IsResolvedFlag or IsClusterResolvedFlag - keep hierarchy!
                                    using (var cmd = _context.Connection.CreateCommand())
                                    {
                                        cmd.Transaction = transaction;
                                        cmd.CommandText = @"UPDATE ClashZones 
                                                            SET IsCombinedResolved = 1, 
                                                                CombinedClusterSleeveInstanceId = @CId 
                                                            WHERE ClusterInstanceId = @Id";
                                        cmd.Parameters.AddWithValue("@Id", constituent.ClusterInstanceId.Value);
                                        cmd.Parameters.AddWithValue("@CId", sleeve.CombinedInstanceId);
                                        int rows = cmd.ExecuteNonQuery();
                                        _logger($"[SQLite] 🔍 Updated {rows} ClashZones for ClusterInstanceId={constituent.ClusterInstanceId.Value}");
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

        private long SaveCombinedSleeveInternal(CombinedSleeve combinedSleeve, SQLiteTransaction transaction)
        {
             // ✅ DEDUPLICATE: Ensure constituents are unique before processing/hashing
             if (combinedSleeve.Constituents != null && combinedSleeve.Constituents.Count > 0)
             {
                 combinedSleeve.Constituents = DeduplicateConstituents(combinedSleeve.Constituents);
             }

             // 1. Generate Deterministic GUID if not present
             if (string.IsNullOrEmpty(combinedSleeve.DeterministicGuid) && combinedSleeve.Constituents != null)
             {
                 combinedSleeve.DeterministicGuid = GenerateDeterministicGuid(combinedSleeve.Constituents);
             }

             long combinedSleeveId = -1;
             
             // 2a. Check for existing record with this GUID
             if (!string.IsNullOrEmpty(combinedSleeve.DeterministicGuid))
             {
                 using (var cmd = _context.Connection.CreateCommand())
                 {
                     cmd.Transaction = transaction;
                     cmd.CommandText = "SELECT CombinedSleeveId FROM CombinedSleeves WHERE DeterministicGuid = @Guid";
                     cmd.Parameters.AddWithValue("@Guid", combinedSleeve.DeterministicGuid);
                     var result = cmd.ExecuteScalar();
                     if (result != null && result != DBNull.Value)
                     {
                         combinedSleeveId = Convert.ToInt64(result);
                         // Fall through to update
                     }
                 }
             }

             // 2b. ✅ CRITICAL DUPLICATE FIX: Check for existing record by Revit Instance ID
             // If hash changed but it's the same Revit element, we must UPDATE, not insert.
             if (combinedSleeveId == -1 && combinedSleeve.CombinedInstanceId > 0)
             {
                 using (var cmd = _context.Connection.CreateCommand())
                 {
                     cmd.Transaction = transaction;
                     cmd.CommandText = "SELECT CombinedSleeveId FROM CombinedSleeves WHERE CombinedInstanceId = @Id";
                     cmd.Parameters.AddWithValue("@Id", combinedSleeve.CombinedInstanceId);
                     var result = cmd.ExecuteScalar();
                     if (result != null && result != DBNull.Value)
                     {
                         combinedSleeveId = Convert.ToInt64(result);
                         _logger($"[SQLite] ⚠️ Hash mismatch but found existing CombinedSleeve {combinedSleeveId} by RevitID {combinedSleeve.CombinedInstanceId}. Updating...");
                     }
                 }
             }

             // UPDATE EXISTING
             if (combinedSleeveId != -1)
             {
                 _logger($"[SQLite] ♻️ Updating existing CombinedSleeve {combinedSleeveId} (RevitID: {combinedSleeve.CombinedInstanceId})");
                 
                 // Update the existing record
                 UpdateCombinedSleeveInternal(combinedSleeveId, combinedSleeve, transaction);
                 
                 // Delete old constituents to replace them
                 DeleteConstituentsInternal(combinedSleeveId, transaction);
                 
                 if (combinedSleeve.Constituents != null && combinedSleeve.Constituents.Count > 0)
                 {
                     SaveConstituentsInternal(combinedSleeveId, combinedSleeve.Constituents, transaction);
                 }
                 
                 return combinedSleeveId;
             }

             // 3. Insert New Record
             using (var cmd = _context.Connection.CreateCommand())
             {
                 cmd.Transaction = transaction;
                 cmd.CommandText = @"
                            INSERT INTO CombinedSleeves (
                                CombinedInstanceId, DeterministicGuid, ComboId, FilterId, Categories,
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
                                @CombinedInstanceId, @DeterministicGuid, @ComboId, @FilterId, @Categories,
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
                        cmd.Parameters.AddWithValue("@DeterministicGuid", combinedSleeve.DeterministicGuid ?? (object)DBNull.Value);
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

                        combinedSleeveId = Convert.ToInt64(cmd.ExecuteScalar());
             }
             
             if (combinedSleeve.Constituents != null && combinedSleeve.Constituents.Count > 0)
             {
                 SaveConstituentsInternal(combinedSleeveId, combinedSleeve.Constituents, transaction);
             }
             return combinedSleeveId;
        }
        
        private void UpdateCombinedSleeveInternal(long combinedSleeveId, CombinedSleeve s, SQLiteTransaction transaction)
        {
            using (var cmd = _context.Connection.CreateCommand())
            {
                cmd.Transaction = transaction;
                cmd.CommandText = @"
                    UPDATE CombinedSleeves SET
                        CombinedInstanceId = @CombinedInstanceId,
                        BoundingBoxMinX = @MinX, BoundingBoxMinY = @MinY, BoundingBoxMinZ = @MinZ,
                        BoundingBoxMaxX = @MaxX, BoundingBoxMaxY = @MaxY, BoundingBoxMaxZ = @MaxZ,
                        CombinedWidth = @W, CombinedHeight = @H, CombinedDepth = @D,
                        PlacementX = @PX, PlacementY = @PY, PlacementZ = @PZ,
                        RotationAngleDeg = @Rot, HostType = @HT, HostOrientation = @HO,
                        Corner1X=@C1X, Corner1Y=@C1Y, Corner1Z=@C1Z,
                        Corner2X=@C2X, Corner2Y=@C2Y, Corner2Z=@C2Z,
                        Corner3X=@C3X, Corner3Y=@C3Y, Corner3Z=@C3Z,
                        Corner4X=@C4X, Corner4Y=@C4Y, Corner4Z=@C4Z,
                        UpdatedAt = CURRENT_TIMESTAMP
                    WHERE CombinedSleeveId = @Id";

                cmd.Parameters.AddWithValue("@Id", combinedSleeveId);
                cmd.Parameters.AddWithValue("@CombinedInstanceId", s.CombinedInstanceId);
                cmd.Parameters.AddWithValue("@MinX", s.BoundingBoxMinX);
                cmd.Parameters.AddWithValue("@MinY", s.BoundingBoxMinY);
                cmd.Parameters.AddWithValue("@MinZ", s.BoundingBoxMinZ);
                cmd.Parameters.AddWithValue("@MaxX", s.BoundingBoxMaxX);
                cmd.Parameters.AddWithValue("@MaxY", s.BoundingBoxMaxY);
                cmd.Parameters.AddWithValue("@MaxZ", s.BoundingBoxMaxZ);
                cmd.Parameters.AddWithValue("@W", s.CombinedWidth);
                cmd.Parameters.AddWithValue("@H", s.CombinedHeight);
                cmd.Parameters.AddWithValue("@D", s.CombinedDepth);
                cmd.Parameters.AddWithValue("@PX", s.PlacementX);
                cmd.Parameters.AddWithValue("@PY", s.PlacementY);
                cmd.Parameters.AddWithValue("@PZ", s.PlacementZ);
                cmd.Parameters.AddWithValue("@Rot", s.RotationAngleDeg);
                cmd.Parameters.AddWithValue("@HT", s.HostType ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@HO", s.HostOrientation ?? (object)DBNull.Value);
                
                cmd.Parameters.AddWithValue("@C1X", s.Corner1X); cmd.Parameters.AddWithValue("@C1Y", s.Corner1Y); cmd.Parameters.AddWithValue("@C1Z", s.Corner1Z);
                cmd.Parameters.AddWithValue("@C2X", s.Corner2X); cmd.Parameters.AddWithValue("@C2Y", s.Corner2Y); cmd.Parameters.AddWithValue("@C2Z", s.Corner2Z);
                cmd.Parameters.AddWithValue("@C3X", s.Corner3X); cmd.Parameters.AddWithValue("@C3Y", s.Corner3Y); cmd.Parameters.AddWithValue("@C3Z", s.Corner3Z);
                cmd.Parameters.AddWithValue("@C4X", s.Corner4X); cmd.Parameters.AddWithValue("@C4Y", s.Corner4Y); cmd.Parameters.AddWithValue("@C4Z", s.Corner4Z);

                cmd.ExecuteNonQuery();
            }
        }

        private void DeleteConstituentsInternal(long combinedSleeveId, SQLiteTransaction transaction)
        {
            using (var cmd = _context.Connection.CreateCommand())
            {
                cmd.Transaction = transaction;
                cmd.CommandText = "DELETE FROM CombinedSleeveConstituents WHERE CombinedSleeveId = @Id";
                cmd.Parameters.AddWithValue("@Id", combinedSleeveId);
                cmd.ExecuteNonQuery();
            }
        }

        /// <summary>
        /// Generates a deterministic GUID based on the sorted constituent IDs.
        /// Ensure duplicates of the same combined sleeve (same constituents) map to the same GUID.
        /// </summary>
        public string GenerateDeterministicGuid(List<SleeveConstituent> constituents)
        {
            if (constituents == null || constituents.Count == 0)
                return Guid.NewGuid().ToString();

            // Collect identifiers - Use HashSet for implicit deduplication of duplicate inputs
            var uniqueIds = new HashSet<string>();
            foreach (var c in constituents)
            {
                if (c.Type == ConstituentType.Individual && c.ClashZoneGuid.HasValue)
                {
                    uniqueIds.Add($"I:{c.ClashZoneGuid.Value}");
                }
                else if (c.Type == ConstituentType.Cluster && c.ClusterInstanceId.HasValue)
                {
                    // Use ClusterInstanceId as the identifier for the cluster instance
                    uniqueIds.Add($"C:{c.ClusterInstanceId.Value}"); 
                }
                else if (c.Type == ConstituentType.Cluster && c.ClusterSleeveId.HasValue)
                {
                     // Fallback to DB ID if instance ID missing (rare)
                     uniqueIds.Add($"C_DB:{c.ClusterSleeveId.Value}");
                }
            }

            // Sort to ensure order independence
            var sortedIds = uniqueIds.OrderBy(x => x).ToList();
            
            // Join and Hash
            // Join and Hash
            var combinedString = string.Join("|", sortedIds);
            using (var sha = System.Security.Cryptography.SHA256.Create())
            {
                var bytes = System.Text.Encoding.UTF8.GetBytes(combinedString);
                var hash = sha.ComputeHash(bytes);
                // Return as hex string or base64. Hex is safer for DB viewing.
                return BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant();
            }
        }

        public void SaveConstituents(long combinedSleeveId, List<SleeveConstituent> constituents)
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
        
        private void SaveConstituentsInternal(long combinedSleeveId, List<SleeveConstituent> constituents, SQLiteTransaction transaction)
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
        
        public CombinedSleeve GetCombinedSleeveById(long combinedSleeveId)
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
        
        public CombinedSleeve GetCombinedSleeveByInstanceId(long instanceId)
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
        
        public List<CombinedSleeve> GetCombinedSleevesForCombo(long comboId, int filterId)
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

        public List<CombinedSleeve> GetAllCombinedSleeves()
        {
            var combinedSleeves = new List<CombinedSleeve>();
            
            try
            {
                using (var cmd = _context.Connection.CreateCommand())
                {
                    cmd.CommandText = "SELECT * FROM CombinedSleeves";
                    
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            var combinedSleeve = MapReaderToCombinedSleeve(reader);
                            // Note: Pre-loading constituents might be expensive if there are many sessions.
                            // For MarkParameterService cache, we mostly need the CombinedInstanceId.
                            combinedSleeves.Add(combinedSleeve);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                 _logger($"[SQLite] ⚠️ GetAllCombinedSleeves failed: {ex.Message}");
            }
            
            return combinedSleeves;
        }
        
        public List<SleeveConstituent> GetConstituents(long combinedSleeveId)
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
        
        public void UpdateCombinedSleeveCorners(long combinedSleeveId,
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
        
        public void DeleteCombinedSleeve(long combinedSleeveId)
        {
            using (var transaction = _context.Connection.BeginTransaction())
            {
                try
                {
                    // ✅ CRITICAL FIX: Get the CombinedInstanceId BEFORE deleting
                    // This is needed to reset ClashZones flags
                    long? combinedInstanceId = null;
                    using (var cmd = _context.Connection.CreateCommand())
                    {
                        cmd.Transaction = transaction;
                        cmd.CommandText = "SELECT CombinedInstanceId FROM CombinedSleeves WHERE CombinedSleeveId = @CombinedSleeveId";
                        cmd.Parameters.AddWithValue("@CombinedSleeveId", combinedSleeveId);
                        var result = cmd.ExecuteScalar();
                        if (result != null && result != DBNull.Value)
                        {
                            combinedInstanceId = Convert.ToInt64(result);
                        }
                    }
                    
                    // ✅ CRITICAL FIX: Reset ClashZones flags for zones linked to this combined sleeve
                    if (combinedInstanceId.HasValue && combinedInstanceId.Value > 0)
                    {
                        using (var cmd = _context.Connection.CreateCommand())
                        {
                            cmd.Transaction = transaction;
                            cmd.CommandText = @"
                                UPDATE ClashZones 
                                SET IsCombinedResolved = 0,
                                    CombinedClusterSleeveInstanceId = NULL,
                                    ReadyForPlacementFlag = 1
                                WHERE CombinedClusterSleeveInstanceId = @CombinedInstanceId";
                            cmd.Parameters.AddWithValue("@CombinedInstanceId", combinedInstanceId.Value);
                            var resetCount = cmd.ExecuteNonQuery();
                            
                            if (resetCount > 0)
                            {
                                _logger($"[SQLite] ✅ Reset {resetCount} ClashZones for deleted combined sleeve {combinedSleeveId} (InstanceId: {combinedInstanceId.Value})");
                            }
                        }
                        
                        // ✅ CRITICAL FIX: Also reset ClusterSleeves table
                        using (var cmd = _context.Connection.CreateCommand())
                        {
                            cmd.Transaction = transaction;
                            cmd.CommandText = @"
                                UPDATE ClusterSleeves 
                                SET CombinedClusterSleeveInstanceId = -1
                                WHERE CombinedClusterSleeveInstanceId = @CombinedInstanceId";
                            cmd.Parameters.AddWithValue("@CombinedInstanceId", combinedInstanceId.Value);
                            cmd.ExecuteNonQuery();
                        }
                    }
                    
                    // Delete the combined sleeve (constituents will be cascade deleted)
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
        
        public void MarkConstituentsAsResolved(List<SleeveConstituent> constituents, long combinedInstanceId)
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
                            // ✅ CRITICAL FIX: Do NOT reset IsResolvedFlag, IsClusterResolvedFlag, SleeveInstanceId, or ClusterInstanceId!
                            // Keep them at their current values (preserving hierarchy). Only refresh should reset when combined is deleted.
                            using (var cmd = _context.Connection.CreateCommand())
                            {
                                cmd.Transaction = transaction;
                                cmd.CommandText = @"
                                    UPDATE ClashZones 
                                    SET IsCombinedResolved = 1,
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
        
        /// <summary>
        /// ✅ ROBUSTNESS FIX: Resets all combined sleeve flags in ClashZones and ClusterSleeves.
        /// Call this at the start of a new run when NOT clearing the database.
        /// This ensures clean state for combined sleeve processing.
        /// </summary>
        public void ResetAllCombinedFlags()
        {
            using (var transaction = _context.Connection.BeginTransaction())
            {
                try
                {
                    int clashZonesReset = 0;
                    int clusterSleevesReset = 0;
                    
                    // Reset ClashZones
                    using (var cmd = _context.Connection.CreateCommand())
                    {
                        cmd.Transaction = transaction;
                        cmd.CommandText = @"
                            UPDATE ClashZones 
                            SET IsCombinedResolved = 0,
                                CombinedClusterSleeveInstanceId = NULL,
                                ReadyForPlacementFlag = CASE 
                                    WHEN IsResolvedFlag = 0 AND IsClusterResolvedFlag = 0 THEN 1 
                                    ELSE ReadyForPlacementFlag 
                                END
                            WHERE IsCombinedResolved = 1 OR CombinedClusterSleeveInstanceId IS NOT NULL";
                        clashZonesReset = cmd.ExecuteNonQuery();
                    }
                    
                    // Reset ClusterSleeves
                    using (var cmd = _context.Connection.CreateCommand())
                    {
                        cmd.Transaction = transaction;
                        cmd.CommandText = @"
                            UPDATE ClusterSleeves 
                            SET CombinedClusterSleeveInstanceId = -1
                            WHERE CombinedClusterSleeveInstanceId > 0";
                        clusterSleevesReset = cmd.ExecuteNonQuery();
                    }
                    
                    transaction.Commit();
                    _logger($"[SQLite] ✅ ResetAllCombinedFlags: {clashZonesReset} ClashZones, {clusterSleevesReset} ClusterSleeves");
                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    _logger($"[SQLite] ❌ Error in ResetAllCombinedFlags: {ex.Message}");
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
                CombinedInstanceId = GetLong(reader, "CombinedInstanceId", -1),
                ComboId = GetLong(reader, "ComboId", -1),
                FilterId = reader.IsDBNull(reader.GetOrdinal("FilterId")) ? 0 : reader.GetInt32(reader.GetOrdinal("FilterId")),
                DeterministicGuid = reader.IsDBNull(reader.GetOrdinal("DeterministicGuid")) ? null : reader.GetString(reader.GetOrdinal("DeterministicGuid")),
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
                ClusterInstanceId = reader.IsDBNull(reader.GetOrdinal("ClusterInstanceId")) ? (long?)null : GetLong(reader, "ClusterInstanceId", -1),
                
                CreatedAt = reader.GetDateTime(reader.GetOrdinal("CreatedAt"))
            };
        }
        private List<SleeveConstituent> DeduplicateConstituents(List<SleeveConstituent> original)
        {
            if (original == null) return null;
            var unique = new List<SleeveConstituent>();
            var seen = new HashSet<string>();
            
            foreach (var c in original)
            {
                string key = "";
                if (c.Type == ConstituentType.Individual && c.ClashZoneGuid.HasValue)
                    key = $"I:{c.ClashZoneGuid.Value}";
                else if (c.Type == ConstituentType.Cluster && c.ClusterInstanceId.HasValue)
                    key = $"C:{c.ClusterInstanceId.Value}";
                else if (c.Type == ConstituentType.Cluster && c.ClusterSleeveId.HasValue)
                    key = $"C_DB:{c.ClusterSleeveId.Value}";
                
                if (string.IsNullOrEmpty(key)) continue; 
                
                if (seen.Add(key))
                {
                    unique.Add(c);
                }
            }
            return unique;
        }


        private static long GetLong(SQLiteDataReader reader, string column, long defaultValue = 0)
        {
            var ordinal = reader.GetOrdinal(column);
            return reader.IsDBNull(ordinal) ? defaultValue : Convert.ToInt64(reader.GetValue(ordinal));
        }
    }
}

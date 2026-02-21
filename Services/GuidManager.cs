using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Centralized GUID management for clash zones.
    /// Eliminates redundant GUID operations across multiple services.
    /// Uses SQLite database as the sole source of truth via <see cref="ClashZoneRepository"/>.
    /// </summary>
    public class GuidManager
    {
        private readonly Document _document;
        
        public GuidManager(Document document)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
        }

        /// <summary>
        /// ✅ BATCH OPTIMIZATION: Fetches existing GUIDs for a collection of intersections using database-first approach.
        /// Returns a dictionary mapping (MepId, HostId, PointKey) -> Guid.
        /// </summary>
        public Dictionary<(long MepId, long HostId, string PointKey), Guid> BatchFetchGuidsDatabaseFirst(
            IEnumerable<(long MepId, long HostId, double X, double Y, double Z)> targets, 
            double tolerance = 0.1)
        {
            var results = new Dictionary<(long MepId, long HostId, string PointKey), Guid>();
            if (targets == null || !targets.Any()) return results;

            try
            {
                using (var context = new SleeveDbContext(_document, msg =>
                {
                    if (!DeploymentConfiguration.DeploymentMode && OptimizationFlags.UseDiagnosticMode)
                        DebugLogger.Info($"[GUID-MANAGER][SQLite] {msg}");
                }))
                {
                    var repository = new ClashZoneRepository(context, msg =>
                    {
                        if (!DeploymentConfiguration.DeploymentMode && OptimizationFlags.UseDiagnosticMode)
                            DebugLogger.Info($"[GUID-MANAGER][SQLite] {msg}");
                    });

                    return repository.FindGuidsByMepHostAndPointsBulk(targets, tolerance);
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning($"[GUID-MANAGER] ⚠️ Batch database lookup failed: {ex.Message}");
                
                // Fallback: Generate deterministic GUIDs individually if database fails
                foreach (var target in targets)
                {
                    string pointKey = $"{Math.Round(target.X, 4)}_{Math.Round(target.Y, 4)}_{Math.Round(target.Z, 4)}";
                    results[(target.MepId, target.HostId, pointKey)] = GenerateDeterministicGuid(target.MepId, target.HostId, target.X, target.Y, target.Z, tolerance);
                }
            }
            return results;
        }
        
        /// <summary>
        /// ✅ DATABASE-FIRST GUID MANAGEMENT: Gets or creates deterministic GUID using database-first approach.
        /// Checks database first (fast, indexed), then generates deterministic GUID if not found.
        /// </summary>
        /// <param name="mepId">MEP element ID (integer value)</param>
        /// <param name="hostId">Host/Structural element ID (integer value)</param>
        /// <param name="intersectionPointX">Intersection point X coordinate</param>
        /// <param name="intersectionPointY">Intersection point Y coordinate</param>
        /// <param name="intersectionPointZ">Intersection point Z coordinate</param>
        /// <param name="tolerance">Tolerance for rounding coordinates (default 0.1ft = ~30mm)</param>
        /// <returns>Deterministic GUID that is stable for the same 3-point combo</returns>
        public Guid GetOrCreateDeterministicGuidDatabaseFirst(long mepId, long hostId, double intersectionPointX, double intersectionPointY, double intersectionPointZ, double tolerance = 0.1)
        {
            if (mepId <= 0 || hostId <= 0)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning($"[GUID-MANAGER] Invalid IDs for deterministic GUID (MEP={mepId}, Host={hostId}) - using random GUID");
                return Guid.NewGuid();
            }

            try
            {
                // Step 1: Check database first (fast, indexed lookup)
                using (var context = new SleeveDbContext(_document, msg =>
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[GUID-MANAGER][SQLite] {msg}");
                }))
                {
                    var repository = new ClashZoneRepository(context, msg =>
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[GUID-MANAGER][SQLite] {msg}");
                    });

                    var existingGuid = repository.FindGuidByMepHostAndPoint(mepId, hostId, intersectionPointX, intersectionPointY, intersectionPointZ, tolerance);
                    if (existingGuid.HasValue)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[GUID-MANAGER] ✅ Found existing GUID {existingGuid.Value} in database for MEP={mepId}, Host={hostId}, Point=({intersectionPointX:F3},{intersectionPointY:F3},{intersectionPointZ:F3})");
                        return existingGuid.Value;
                    }

                    // Step 2: GUID not in database - generate deterministic GUID and store it
                    var deterministicGuid = GenerateDeterministicGuid(mepId, hostId, intersectionPointX, intersectionPointY, intersectionPointZ, tolerance);
                    
                    // Store in database (will be stored when clash zone is inserted, but we can pre-store it)
                    try
                    {
                        repository.GetOrCreateDeterministicGuid(mepId, hostId, intersectionPointX, intersectionPointY, intersectionPointZ, GenerateDeterministicGuid, tolerance);
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[GUID-MANAGER] ✅ Stored deterministic GUID {deterministicGuid} in database for MEP={mepId}, Host={hostId}");
                    }
                    catch (Exception dbEx)
                    {
                        // Non-blocking: Log but continue (GUID will be stored when clash zone is inserted)
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Warning($"[GUID-MANAGER] ⚠️ Could not pre-store GUID in database: {dbEx.Message}");
                    }

                    return deterministicGuid;
                }
            }
            catch (Exception ex)
            {
                // Fallback: Generate deterministic GUID even if database fails
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning($"[GUID-MANAGER] ⚠️ Database lookup failed, generating deterministic GUID: {ex.Message}");
                return GenerateDeterministicGuid(mepId, hostId, intersectionPointX, intersectionPointY, intersectionPointZ, tolerance);
            }
        }

        // ✅ PERFORMANCE OPTIMIZATION: FNV-1a hash constants (much faster than MD5)
        private const uint FNV1a_OffsetBasis = 2166136261;
        private const uint FNV1a_Prime = 16777619;
        
        /// <summary>
        /// ✅ PERFORMANCE OPTIMIZED: Fast FNV-1a hash (10-20x faster than MD5)
        /// Deterministic GUID from stable identifiers (MEP+Host+Point)
        /// </summary>
        private static byte[] Fnv1aHash128(byte[] data)
        {
            // Use two independent 64-bit hashes for 128-bit output
            uint h1 = FNV1a_OffsetBasis;
            uint h2 = FNV1a_OffsetBasis;
            uint h3 = FNV1a_OffsetBasis;
            uint h4 = FNV1a_OffsetBasis;
            
            for (int i = 0; i < data.Length; i++)
            {
                byte b = data[i];
                h1 ^= b;
                h1 *= FNV1a_Prime;
                h2 ^= (byte)(b + 1);
                h2 *= FNV1a_Prime;
                h3 ^= (byte)(b * 7);
                h3 *= FNV1a_Prime;
                h4 ^= (byte)(b ^ 0x5A);
                h4 *= FNV1a_Prime;
            }
            
            // Pack into 16 bytes
            var result = new byte[16];
            BitConverter.GetBytes(h1).CopyTo(result, 0);
            BitConverter.GetBytes(h2).CopyTo(result, 4);
            BitConverter.GetBytes(h3).CopyTo(result, 8);
            BitConverter.GetBytes(h4).CopyTo(result, 12);
            return result;
        }

        /// <summary>
        /// ✅ CRITICAL: Generates a deterministic GUID from stable identifiers (MEP+Host+Point)
        /// This ensures the same intersection always gets the same GUID across detection runs
        /// Uses FAST FNV-1a hash instead of slow MD5 (10-20x performance improvement)
        /// </summary>
        /// <param name="mepId">MEP element ID (integer value)</param>
        /// <param name="hostId">Host/Structural element ID (integer value)</param>
        /// <param name="intersectionPointX">Intersection point X coordinate</param>
        /// <param name="intersectionPointY">Intersection point Y coordinate</param>
        /// <param name="intersectionPointZ">Intersection point Z coordinate</param>
        /// <param name="tolerance">Tolerance for rounding coordinates (default 0.1ft = ~30mm)</param>
        /// <returns>Deterministic GUID that is stable for the same 3-point combo</returns>
        public Guid GenerateDeterministicGuid(long mepId, long hostId, double intersectionPointX, double intersectionPointY, double intersectionPointZ, double tolerance = 0.1)
        {
            if (mepId <= 0 || hostId <= 0)
            {
                // Invalid IDs - fallback to random GUID
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning($"[GUID-MANAGER] Invalid IDs for deterministic GUID (MEP={mepId}, Host={hostId}) - using random GUID");
                return Guid.NewGuid();
            }
            
            // Round coordinates to tolerance to ensure stable matching
            // Convert to integers to avoid floating-point string formatting overhead
            int roundedX = (int)Math.Round(intersectionPointX / tolerance);
            int roundedY = (int)Math.Round(intersectionPointY / tolerance);
            int roundedZ = (int)Math.Round(intersectionPointZ / tolerance);
            
            // ✅ PERFORMANCE OPT: Direct binary serialization instead of string formatting
            // 8 + 8 + 4 + 4 + 4 = 28 bytes vs ~60+ bytes for string
            var hashInput = new byte[28];
            BitConverter.GetBytes(mepId).CopyTo(hashInput, 0);
            BitConverter.GetBytes(hostId).CopyTo(hashInput, 8);
            BitConverter.GetBytes(roundedX).CopyTo(hashInput, 16);
            BitConverter.GetBytes(roundedY).CopyTo(hashInput, 20);
            BitConverter.GetBytes(roundedZ).CopyTo(hashInput, 24);
            
            // ✅ PERFORMANCE OPT: Fast FNV-1a hash instead of slow MD5
            byte[] hashBytes = Fnv1aHash128(hashInput);
            
            // Convert hash bytes to GUID format
            Guid deterministicGuid = new Guid(hashBytes);
            
            if (!DeploymentConfiguration.DeploymentMode && OptimizationFlags.UseDiagnosticMode)
            {
                DebugLogger.Info($"[GUID-MANAGER] Generated deterministic GUID {deterministicGuid} for MEP={mepId}, Host={hostId}, Point=({roundedX},{roundedY},{roundedZ})");
            }
            
            return deterministicGuid;
        }
        
        /// <summary>
        /// Generates a new random GUID (for backward compatibility)
        /// </summary>
        /// <returns>A new unique GUID</returns>
        public Guid GenerateNewGuid()
        {
            return Guid.NewGuid();
        }
        
        /// <summary>
        /// Finds an existing clash zone by MEP element ID and structural element ID.
        /// </summary>
        /// <param name="clashZones">List of clash zones to search</param>
        /// <param name="mepId">MEP element ID</param>
        /// <param name="hostId">Structural element (host) ID</param>
        /// <returns>The matching ClashZone if found, null otherwise</returns>
        public ClashZone? FindByMepAndHost(List<ClashZone> clashZones, ElementId mepId, ElementId hostId)
        {
            if (clashZones == null || clashZones.Count == 0)
                return null;
                
            long mepIdValue = mepId?.GetIntegerValue() ?? -1;
            long hostIdValue = hostId?.GetIntegerValue() ?? -1;
            
            if (mepIdValue <= 0 || hostIdValue <= 0)
                return null;
            
            var match = clashZones.FirstOrDefault(cz => 
            {
                if (cz == null) return false;
                long czMepId = cz.MepElementId?.GetIntegerValue() ?? cz.MepElementIdValue;
                long czStructuralId = cz.StructuralElementId?.GetIntegerValue() ?? cz.StructuralElementIdValue;
                return czMepId == mepIdValue && czStructuralId == hostIdValue;
            });
            
            if (match != null)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[GUID-MANAGER] Found existing ClashZone {match.Id} for MEP={mepIdValue}, Structural={hostIdValue}");
            }
            
            return match;
        }
        
        /// <summary>
        /// Finds existing clash zone by MEP+Host+IntersectionPoint with tolerance matching.
        /// This ensures same intersection reuses existing GUID instead of creating new one.
        /// </summary>
        /// <param name="clashZones">List of clash zones to search</param>
        /// <param name="mepIdValue">MEP element ID (integer value)</param>
        /// <param name="structuralIdValue">Structural element ID (integer value)</param>
        /// <param name="intersectionPoint">Intersection point (with tolerance matching)</param>
        /// <returns>The matching ClashZone if found, null otherwise</returns>
        public ClashZone? FindByMepHostAndPoint(List<ClashZone> clashZones, long mepIdValue, long structuralIdValue, XYZ intersectionPoint)
        {
            if (clashZones == null || clashZones.Count == 0)
                return null;
                
            if (mepIdValue <= 0 || structuralIdValue <= 0)
                return null;
            
            const double pointTolerance = 0.1; // 0.1 feet = ~30mm tolerance for intersection point matching
            
            ClashZone? bestMatch = null;
            double closestDistance = double.MaxValue;
            
            foreach (var cz in clashZones)
            {
                if (cz == null) continue;
                
                long czMepId = cz.MepElementId?.GetIntegerValue() ?? cz.MepElementIdValue;
                long czStructuralId = cz.StructuralElementId?.GetIntegerValue() ?? cz.StructuralElementIdValue;
                
                // First check: MEP and Host must match
                if (czMepId != mepIdValue || czStructuralId != structuralIdValue)
                    continue;
                
                // Second check: Intersection point must be close (within tolerance)
                if (intersectionPoint != null && 
                    Math.Abs(cz.IntersectionPointX) > 1e-9 && 
                    Math.Abs(cz.IntersectionPointY) > 1e-9 && 
                    Math.Abs(cz.IntersectionPointZ) > 1e-9)
                {
                    var existingPoint = new XYZ(cz.IntersectionPointX, cz.IntersectionPointY, cz.IntersectionPointZ);
                    var distance = intersectionPoint.DistanceTo(existingPoint);
                    
                    if (distance <= pointTolerance)
                    {
                        // Found match - keep the closest one if multiple matches
                        if (distance < closestDistance)
                        {
                            closestDistance = distance;
                            bestMatch = cz;
                        }
                    }
                }
                else
                {
                    // If existing clash zone has zero/invalid intersection point, still match by MEP+Host only
                    if (bestMatch == null || closestDistance == double.MaxValue)
                    {
                        bestMatch = cz;
                        closestDistance = 0; // Exact match by MEP+Host
                    }
                }
            }
            
            if (bestMatch != null)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[GUID-MANAGER] ✅ FOUND BY POINT MATCH: ClashZone {bestMatch.Id} (MEP={mepIdValue}, Structural={structuralIdValue}, Distance={closestDistance:F3}ft) - REUSING EXISTING GUID");
            }
            
            return bestMatch;
        }
        
        /// <summary>
        /// ✅ LEGACY STUB: Finds existing clash zone by checking placed Revit sleeves for GUID matches.
        /// This method is a stub for backward compatibility after GlobalIndexService removal.
        /// </summary>
        /// <param name="clashZones">List of clash zones to search</param>
        /// <param name="mepIdValue">MEP element ID (integer value)</param>
        /// <param name="structuralIdValue">Structural element ID (integer value)</param>
        /// <param name="intersectionPoint">Intersection point (with tolerance matching)</param>
        /// <returns>The matching ClashZone if found, null otherwise</returns>
        public ClashZone? FindByRevitSleeveGuid(List<ClashZone> clashZones, int mepIdValue, int structuralIdValue, XYZ intersectionPoint)
        {
            // ✅ LEGACY STUB: This was used for searching placed sleeves by GUID.
            // Now we rely on database-first approach via FindByMepHostAndPoint.
            // Return null to fall through to the next priority check.
            return null;
        }
        
        /// <summary>
        /// ✅ LEGACY STUB: Removes entries from Global XML (now no-op since XML is deprecated).
        /// This method is a stub for backward compatibility after GlobalIndexService removal.
        /// </summary>
        /// <param name="clashZones">List of clash zones to remove</param>
        public void RemoveFromGlobalXml(IEnumerable<ClashZone> clashZones)
        {
            // ✅ LEGACY STUB: Global XML is deprecated. This is a no-op.
            // The database is the sole source of truth.
        }
        
        /// <summary>
        /// ✅ LEGACY STUB: Removes entry from Global XML by GUID and category (now no-op since XML is deprecated).
        /// This method is a stub for backward compatibility after GlobalIndexService removal.
        /// </summary>
        /// <param name="id">ClashZone GUID to remove</param>
        /// <param name="category">MEP category</param>
        public void RemoveFromGlobalXml(Guid id, string category)
        {
            // ✅ LEGACY STUB: Global XML is deprecated. This is a no-op.
            // The database is the sole source of truth.
        }
    }
}

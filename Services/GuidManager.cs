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
        public Dictionary<(int MepId, int HostId, string PointKey), Guid> BatchFetchGuidsDatabaseFirst(
            IEnumerable<(int MepId, int HostId, double X, double Y, double Z)> targets, 
            double tolerance = 0.1)
        {
            var results = new Dictionary<(int MepId, int HostId, string PointKey), Guid>();
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
        public Guid GetOrCreateDeterministicGuidDatabaseFirst(int mepId, int hostId, double intersectionPointX, double intersectionPointY, double intersectionPointZ, double tolerance = 0.1)
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

        /// <summary>
        /// ✅ CRITICAL: Generates a deterministic GUID from stable identifiers (MEP+Host+Point)
        /// This ensures the same intersection always gets the same GUID across detection runs
        /// Uses MD5 hash of MEP ID + Host ID + rounded intersection point coordinates
        /// Follows industry best practices for stable clash identification
        /// </summary>
        /// <param name="mepId">MEP element ID (integer value)</param>
        /// <param name="hostId">Host/Structural element ID (integer value)</param>
        /// <param name="intersectionPointX">Intersection point X coordinate</param>
        /// <param name="intersectionPointY">Intersection point Y coordinate</param>
        /// <param name="intersectionPointZ">Intersection point Z coordinate</param>
        /// <param name="tolerance">Tolerance for rounding coordinates (default 0.1ft = ~30mm)</param>
        /// <returns>Deterministic GUID that is stable for the same 3-point combo</returns>
        public Guid GenerateDeterministicGuid(int mepId, int hostId, double intersectionPointX, double intersectionPointY, double intersectionPointZ, double tolerance = 0.1)
        {
            if (mepId <= 0 || hostId <= 0)
            {
                // Invalid IDs - fallback to random GUID
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning($"[GUID-MANAGER] Invalid IDs for deterministic GUID (MEP={mepId}, Host={hostId}) - using random GUID");
                return Guid.NewGuid();
            }
            
            // Round coordinates to tolerance to ensure stable matching
            // This ensures slight coordinate variations don't generate different GUIDs
            double roundedX = Math.Round(intersectionPointX / tolerance) * tolerance;
            double roundedY = Math.Round(intersectionPointY / tolerance) * tolerance;
            double roundedZ = Math.Round(intersectionPointZ / tolerance) * tolerance;
            
            // Create deterministic hash input from stable identifiers
            // Format: "MEP_ID|HOST_ID|X|Y|Z" with high precision
            string hashInput = $"{mepId}|{hostId}|{roundedX:F6}|{roundedY:F6}|{roundedZ:F6}";
            
            // Generate MD5 hash (deterministic - same input always produces same output)
            byte[] hashBytes;
            using (var md5 = System.Security.Cryptography.MD5.Create())
            {
                hashBytes = md5.ComputeHash(System.Text.Encoding.UTF8.GetBytes(hashInput));
            }
            
            // Convert hash bytes to GUID format (version 3 UUID-like)
            // MD5 produces 16 bytes, which is exactly what we need for a GUID
            Guid deterministicGuid = new Guid(hashBytes);
            
            if (!DeploymentConfiguration.DeploymentMode && OptimizationFlags.UseDiagnosticMode)
            {
                DebugLogger.Info($"[GUID-MANAGER] Generated deterministic GUID {deterministicGuid} for MEP={mepId}, Host={hostId}, Point=({roundedX:F3},{roundedY:F3},{roundedZ:F3})");
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
                
            int mepIdValue = mepId?.GetIntegerValue() ?? -1;
            int hostIdValue = hostId?.GetIntegerValue() ?? -1;
            
            if (mepIdValue <= 0 || hostIdValue <= 0)
                return null;
            
            var match = clashZones.FirstOrDefault(cz => 
            {
                if (cz == null) return false;
                int czMepId = cz.MepElementId?.GetIntegerValue() ?? cz.MepElementIdValue;
                int czStructuralId = cz.StructuralElementId?.GetIntegerValue() ?? cz.StructuralElementIdValue;
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
        public ClashZone? FindByMepHostAndPoint(List<ClashZone> clashZones, int mepIdValue, int structuralIdValue, XYZ intersectionPoint)
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
                
                int czMepId = cz.MepElementId?.GetIntegerValue() ?? cz.MepElementIdValue;
                int czStructuralId = cz.StructuralElementId?.GetIntegerValue() ?? cz.StructuralElementIdValue;
                
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

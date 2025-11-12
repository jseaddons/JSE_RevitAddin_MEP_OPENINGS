using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Centralized GUID management for clash zones.
    /// Eliminates redundant GUID operations across multiple services.
    /// Uses GlobalIndexService for all Global XML operations.
    /// </summary>
    public class GuidManager
    {
        private readonly Document _document;
        
        public GuidManager(Document document)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
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
        /// Matches by IntegerValue to handle XML deserialization cases where ElementId objects may not be initialized.
        /// </summary>
        /// <param name="clashZones">List of clash zones to search</param>
        /// <param name="mepId">MEP element ID</param>
        /// <param name="hostId">Structural element (host) ID</param>
        /// <returns>The matching ClashZone if found, null otherwise</returns>
        public ClashZone? FindByMepAndHost(List<ClashZone> clashZones, ElementId mepId, ElementId hostId)
        {
            if (clashZones == null || clashZones.Count == 0)
                return null;
                
            int mepIdValue = mepId?.IntegerValue ?? -1;
            int hostIdValue = hostId?.IntegerValue ?? -1;
            
            if (mepIdValue <= 0 || hostIdValue <= 0)
                return null;
            
            var match = clashZones.FirstOrDefault(cz => 
            {
                if (cz == null) return false;
                int czMepId = cz.MepElementId?.IntegerValue ?? cz.MepElementIdValue;
                int czStructuralId = cz.StructuralElementId?.IntegerValue ?? cz.StructuralElementIdValue;
                return czMepId == mepIdValue && czStructuralId == hostIdValue;
            });
            
            if (match != null)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[GUID-MANAGER] Found existing ClashZone {match.Id} for MEP={mepIdValue}, Structural={hostIdValue}");
            }
            else
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[GUID-MANAGER] No existing ClashZone found for MEP={mepIdValue}, Structural={hostIdValue}");
            }
            
            return match;
        }
        
        /// <summary>
        /// ✅ CRITICAL: Finds existing clash zone by matching MEP+Host+Point in Global XML (not GUID parameter on sleeve).
        /// This prevents duplicate GUID creation when the same intersection is detected again.
        /// Priority: Check Global XML FIRST (most reliable source of truth) - no GUID parameter needed on sleeves.
        /// </summary>
        /// <param name="clashZones">List of clash zones loaded from XML to search by GUID</param>
        /// <param name="mepIdValue">MEP element ID (integer value)</param>
        /// <param name="structuralIdValue">Structural element ID (integer value)</param>
        /// <param name="intersectionPoint">Intersection point (with tolerance matching)</param>
        /// <returns>The matching ClashZone if found in Global XML, null otherwise</returns>
        public ClashZone? FindByRevitSleeveGuid(List<ClashZone> clashZones, int mepIdValue, int structuralIdValue, XYZ intersectionPoint)
        {
            if (_document == null)
                return null;
                
            if (mepIdValue <= 0 || structuralIdValue <= 0)
                return null;
            
            try
            {
                // ✅ CRITICAL: Match by MEP+Host+Point in Global XML (not GUID parameter on sleeve)
                // This enables cross-filter matching without needing GUID parameter on sleeves
                // Find which category this clash zone belongs to (check first clash zone's category)
                string category = clashZones?.FirstOrDefault()?.MepElementCategory;
                if (string.IsNullOrWhiteSpace(category))
                    return null;
                
                // Match by MEP+Host+Point in Global XML
                var globalEntry = GlobalIndexService.FindByMepHostAndPoint(
                    _document, 
                    category, 
                    mepIdValue, 
                    structuralIdValue, 
                    intersectionPoint?.X ?? 0, 
                    intersectionPoint?.Y ?? 0, 
                    intersectionPoint?.Z ?? 0);
                
                if (globalEntry != null && Guid.TryParse(globalEntry.Id, out Guid matchedGuid))
                {
                    // Find clash zone in XML by GUID from Global XML
                    var clashZone = clashZones?.FirstOrDefault(cz => cz.Id == matchedGuid);
                    if (clashZone != null)
                    {
                                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[GUID-MANAGER] ✅ FOUND VIA GLOBAL XML: ClashZone {clashZone.Id} (MEP={mepIdValue}, Structural={structuralIdValue}) - REUSING EXISTING GUID");
                        return clashZone;
                    }
                }
                
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[GUID-MANAGER] No matching clash zone found in Global XML for MEP={mepIdValue}, Structural={structuralIdValue}");
                return null;
            }
            catch (Exception ex)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[GUID-MANAGER] Error finding clash zone by Global XML matching: {ex.Message}");
                return null;
            }
        }
        
        /// <summary>
        /// Finds existing clash zone by MEP+Host+IntersectionPoint with tolerance matching.
        /// This ensures same intersection reuses existing GUID from XML instead of creating new one.
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
                
                int czMepId = cz.MepElementId?.IntegerValue ?? cz.MepElementIdValue;
                int czStructuralId = cz.StructuralElementId?.IntegerValue ?? cz.StructuralElementIdValue;
                
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
                    // (fallback for old XML data without intersection points)
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
            else
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[GUID-MANAGER] No clash zone found by point match for MEP={mepIdValue}, Structural={structuralIdValue}");
            }
            
            return bestMatch;
        }
        
        /// <summary>
        /// ✅ CRITICAL: Ensures a Global XML entry exists for a clash zone with MEP+Host+Point data.
        /// This enables O(1) matching by intersection point instead of GUID, allowing cross-filter matching.
        /// ✅ CRITICAL: Stores FilterName to identify which Filter XML file contains placement data.
        /// ✅ CRITICAL: Extracts LinkedFile/HostFile from ClashZone to group entries by file combo in hierarchical structure.
        /// </summary>
        /// <param name="clashZone">The clash zone to ensure an entry for</param>
        /// <param name="category">MEP element category name</param>
        /// <param name="filterName">Filter name that contains this clash zone's placement data (optional, for backward compatibility)</param>
        public void EnsureGlobalXmlEntry(ClashZone clashZone, string category, string filterName = null)
        {
            if (clashZone == null)
                throw new ArgumentNullException(nameof(clashZone));
                
            if (string.IsNullOrWhiteSpace(category))
                throw new ArgumentException("Category cannot be null or empty", nameof(category));
            
            try
            {
                int mepId = clashZone.MepElementId?.IntegerValue ?? clashZone.MepElementIdValue;
                int hostId = clashZone.StructuralElementId?.IntegerValue ?? clashZone.StructuralElementIdValue;
                
                // ✅ CRITICAL: Extract LinkedFile and HostFile from ClashZone to group entries by file combo
                // ✅ FIX: Normalize file names to match UI display format used in ProcessedFileCombo
                string linkedFile = null;
                string hostFile = null;
                
                // Helper function to normalize file names (matches GetNormalizedKey logic)
                Func<string, string> normalizeFile = s =>
                {
                    if (string.IsNullOrWhiteSpace(s)) return string.Empty;
                    var trimmed = s.Trim();
                    // Extract filename from full path
                    trimmed = System.IO.Path.GetFileNameWithoutExtension(trimmed);
                    // Remove parentheses content (e.g., "PH-00001 (50 elements)" -> "PH-00001")
                    var idxParen = trimmed.IndexOf('(');
                    if (idxParen >= 0) trimmed = trimmed.Substring(0, idxParen).Trim();
                    return trimmed;
                };
                
                // Get LinkedFile from SourceDocKey (MEP element's document)
                // SourceDocKey is full path (e.g., "C:\Users\...\PH-00001.rvt"), normalize to match UI format
                if (!string.IsNullOrWhiteSpace(clashZone.SourceDocKey))
                {
                    linkedFile = normalizeFile(clashZone.SourceDocKey);
                }
                
                // Get HostFile from HostDocKey or StructuralElementDocumentTitle (structural element's document)
                if (!string.IsNullOrWhiteSpace(clashZone.HostDocKey))
                {
                    hostFile = normalizeFile(clashZone.HostDocKey);
                }
                else if (!string.IsNullOrWhiteSpace(clashZone.StructuralElementDocumentTitle))
                {
                    hostFile = normalizeFile(clashZone.StructuralElementDocumentTitle);
                }
                
                // ✅ DEBUG: Log when LinkedFile/HostFile are missing
                if ((string.IsNullOrWhiteSpace(linkedFile) || string.IsNullOrWhiteSpace(hostFile)) && !DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Warning($"[GUID-MANAGER] ⚠️ ClashZone {clashZone.Id} missing file combo data - SourceDocKey='{clashZone.SourceDocKey ?? "NULL"}', HostDocKey='{clashZone.HostDocKey ?? "NULL"}', StructuralElementDocumentTitle='{clashZone.StructuralElementDocumentTitle ?? "NULL"}'. Entry will go to flat structure only.");
                }
                else if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[GUID-MANAGER] ClashZone {clashZone.Id} file combo - SourceDocKey='{clashZone.SourceDocKey ?? "NULL"}' -> LinkedFile='{linkedFile}', HostDocKey='{clashZone.HostDocKey ?? "NULL"}' -> HostFile='{hostFile}'");
                }
                
                // ✅ CRITICAL: Only pass LinkedFile/HostFile if BOTH are available
                // This ensures entries are grouped by file combo, not placed in placeholder combos
                // If either is missing, entries will go to flat structure only (no placeholder FileComboGroup)
                
                // ✅ CRITICAL: Store MEP+Host+Point in Global XML for O(1) matching
                // ✅ CRITICAL: Store FilterName to identify which Filter XML file contains placement data
                // ✅ CRITICAL: Pass LinkedFile/HostFile to group entries by file combo in hierarchical structure
                GlobalIndexService.EnsureEntriesWithClashZoneData(_document, category, new[] 
                { 
                    (clashZone.Id, mepId, hostId, clashZone.IntersectionPointX, clashZone.IntersectionPointY, clashZone.IntersectionPointZ)
                }, filterName, linkedFile, hostFile);
                
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[GUID-MANAGER] Ensured Global XML entry for ClashZone {clashZone.Id} (MEP={mepId}, Host={hostId}, Point=({clashZone.IntersectionPointX:F3},{clashZone.IntersectionPointY:F3},{clashZone.IntersectionPointZ:F3}), Filter='{filterName ?? "N/A"}', LinkedFile='{linkedFile ?? "N/A"}', HostFile='{hostFile ?? "N/A"}') in category '{category}'");
            }
            catch (Exception ex)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[GUID-MANAGER] Error ensuring Global XML entry for ClashZone {clashZone.Id}: {ex.Message}");
                throw;
            }
        }
        
        /// <summary>
        /// Removes a clash zone entry from Global XML.
        /// Called when a clash zone is invalidated (e.g., during 3-point validation failure).
        /// </summary>
        /// <param name="guid">The GUID of the clash zone to remove</param>
        /// <param name="category">MEP element category name</param>
        public void RemoveFromGlobalXml(Guid guid, string category)
        {
            if (guid == Guid.Empty)
                throw new ArgumentException("GUID cannot be empty", nameof(guid));
                
            if (string.IsNullOrWhiteSpace(category))
                throw new ArgumentException("Category cannot be null or empty", nameof(category));
            
            try
            {
                var globalIndex = GlobalIndexService.LoadOrCreate(_document, category);
                
                // ✅ CRITICAL FIX: Remove from BOTH hierarchical and flat structures
                // Entries can be in Filters → FileCombos → Entries OR in flat Entries list
                var guidString = guid.ToString();
                int removedCount = 0;
                
                // Remove from hierarchical structure
                if (globalIndex.Filters != null && globalIndex.Filters.Count > 0)
                {
                    foreach (var filterGroup in globalIndex.Filters)
                    {
                        if (filterGroup.FileCombos != null)
                        {
                            foreach (var fileCombo in filterGroup.FileCombos)
                            {
                                if (fileCombo.Entries != null)
                                {
                                    var beforeCount = fileCombo.Entries.Count;
                                    fileCombo.Entries.RemoveAll(e => string.Equals(e.Id, guidString, StringComparison.OrdinalIgnoreCase));
                                    removedCount += (beforeCount - fileCombo.Entries.Count);
                                }
                            }
                        }
                    }
                }
                
                // Remove from flat structure
                if (globalIndex.Entries != null && globalIndex.Entries.Count > 0)
                {
                    var beforeCount = globalIndex.Entries.Count;
                    globalIndex.Entries.RemoveAll(e => string.Equals(e.Id, guidString, StringComparison.OrdinalIgnoreCase));
                    removedCount += (beforeCount - globalIndex.Entries.Count);
                }
                
                if (removedCount > 0)
                {
                    GlobalIndexService.Save(_document, globalIndex);
                                        if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[GUID-MANAGER] Removed ClashZone {guid} from Global XML for category '{category}' ({removedCount} entry removed)");
                }
                else
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[GUID-MANAGER] ClashZone {guid} not found in Global XML for category '{category}' (no removal needed)");
                }
            }
            catch (Exception ex)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[GUID-MANAGER] Error removing ClashZone {guid} from Global XML: {ex.Message}");
                throw;
            }
        }
    }
}


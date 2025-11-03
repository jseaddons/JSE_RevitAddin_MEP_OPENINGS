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
        /// Generates a new GUID for a new clash zone.
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
        /// ✅ CRITICAL: Finds existing clash zone by reading GUID from Revit sleeve elements.
        /// This prevents duplicate GUID creation when the same intersection is detected again.
        /// Priority: Check Revit sleeves FIRST (most reliable source of truth).
        /// </summary>
        /// <param name="clashZones">List of clash zones loaded from XML to search by GUID</param>
        /// <param name="mepIdValue">MEP element ID (integer value)</param>
        /// <param name="structuralIdValue">Structural element ID (integer value)</param>
        /// <param name="intersectionPoint">Intersection point (with tolerance matching)</param>
        /// <returns>The matching ClashZone if found in Revit sleeve, null otherwise</returns>
        public ClashZone? FindByRevitSleeveGuid(List<ClashZone> clashZones, int mepIdValue, int structuralIdValue, XYZ intersectionPoint)
        {
            if (_document == null)
                return null;
                
            if (mepIdValue <= 0 || structuralIdValue <= 0)
                return null;
            
            try
            {
                // Find all sleeves with matching MEP_ElementId parameter
                // ✅ CRITICAL: Check both individual sleeves (MEP_ElementId) and cluster sleeves (MEP_ElementIds comma-separated)
                var sleeves = new FilteredElementCollector(_document)
                    .OfClass(typeof(FamilyInstance))
                    .WhereElementIsNotElementType()
                    .Cast<FamilyInstance>()
                    .Where(s => 
                    {
                        // Check individual sleeve MEP_ElementId
                        var mepParam = s.LookupParameter("MEP_ElementId");
                        if (mepParam != null && mepParam.HasValue && mepParam.AsInteger() == mepIdValue)
                            return true;
                        
                        // Check cluster sleeve MEP_ElementIds (comma-separated list)
                        var mepIdsParam = s.LookupParameter("MEP_ElementIds");
                        if (mepIdsParam != null && mepIdsParam.HasValue)
                        {
                            string mepIdsString = mepIdsParam.AsString();
                            if (!string.IsNullOrWhiteSpace(mepIdsString))
                            {
                                // Check if mepIdValue is in the comma-separated list
                                var mepIds = mepIdsString.Split(',')
                                    .Select(id => id.Trim())
                                    .Where(id => !string.IsNullOrWhiteSpace(id))
                                    .Select(id => long.TryParse(id, out long parsed) ? parsed : (long?)null)
                                    .Where(id => id.HasValue)
                                    .Select(id => id.Value)
                                    .ToList();
                                
                                if (mepIds.Contains(mepIdValue))
                                    return true;
                            }
                        }
                        
                        return false;
                    })
                    .ToList();
                
                if (sleeves.Count == 0)
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[GUID-MANAGER] No sleeves found with MEP_ElementId={mepIdValue}");
                    return null;
                }
                
                const double pointTolerance = 0.1; // 0.1 feet = ~30mm tolerance
                
                // Check each sleeve for matching intersection point and read its GUID
                foreach (var sleeve in sleeves)
                {
                    try
                    {
                        // Read GUID from sleeve parameter
                        var guidParam = sleeve.LookupParameter("ClashZone_GUID");
                        if (guidParam == null || !guidParam.HasValue)
                            continue;
                            
                        string guidString = guidParam.AsString();
                        if (string.IsNullOrWhiteSpace(guidString) || !Guid.TryParse(guidString, out Guid sleeveGuid))
                            continue;
                        
                        // Find clash zone in XML by this GUID
                        var clashZone = clashZones?.FirstOrDefault(cz => cz.Id == sleeveGuid);
                        if (clashZone == null)
                            continue;
                        
                        // Verify structural element ID matches
                        int czStructuralId = clashZone.StructuralElementId?.IntegerValue ?? clashZone.StructuralElementIdValue;
                        if (czStructuralId != structuralIdValue)
                            continue;
                        
                        // Verify intersection point is close (within tolerance)
                        if (intersectionPoint != null &&
                            Math.Abs(clashZone.IntersectionPointX) > 1e-9 &&
                            Math.Abs(clashZone.IntersectionPointY) > 1e-9 &&
                            Math.Abs(clashZone.IntersectionPointZ) > 1e-9)
                        {
                            var existingPoint = new XYZ(
                                clashZone.IntersectionPointX,
                                clashZone.IntersectionPointY,
                                clashZone.IntersectionPointZ
                            );
                            var distance = intersectionPoint.DistanceTo(existingPoint);
                            
                            if (distance > pointTolerance)
                                continue; // Point doesn't match, try next sleeve
                        }
                        
                        // ✅ MATCH FOUND: Reuse existing GUID from Revit sleeve
                                                if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[GUID-MANAGER] ✅ FOUND VIA REVIT SLEEVE: ClashZone {clashZone.Id} (sleeve {sleeve.Id}, MEP={mepIdValue}, Structural={structuralIdValue}) - REUSING EXISTING GUID");
                        return clashZone;
                    }
                    catch (Exception ex)
                    {
                                                if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Warning($"[GUID-MANAGER] Error reading GUID from sleeve {sleeve.Id}: {ex.Message}");
                        continue;
                    }
                }
                
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[GUID-MANAGER] No matching clash zone found in Revit sleeves for MEP={mepIdValue}, Structural={structuralIdValue}");
                return null;
            }
            catch (Exception ex)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[GUID-MANAGER] Error finding clash zone by Revit sleeve GUID: {ex.Message}");
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
        /// Ensures a Global XML entry exists for a clash zone (creates if missing).
        /// Called when a new clash zone is detected during refresh.
        /// </summary>
        /// <param name="clashZone">The clash zone to ensure an entry for</param>
        /// <param name="category">MEP element category name</param>
        public void EnsureGlobalXmlEntry(ClashZone clashZone, string category)
        {
            if (clashZone == null)
                throw new ArgumentNullException(nameof(clashZone));
                
            if (string.IsNullOrWhiteSpace(category))
                throw new ArgumentException("Category cannot be null or empty", nameof(category));
            
            try
            {
                GlobalIndexService.EnsureEntries(_document, category, new[] { clashZone.Id });
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[GUID-MANAGER] Ensured Global XML entry for ClashZone {clashZone.Id} in category '{category}'");
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
                var beforeCount = globalIndex.Entries.Count;
                
                globalIndex.Entries.RemoveAll(e => 
                    string.Equals(e.Id, guid.ToString(), StringComparison.OrdinalIgnoreCase));
                
                var afterCount = globalIndex.Entries.Count;
                var removedCount = beforeCount - afterCount;
                
                if (removedCount > 0)
                {
                    GlobalIndexService.Save(_document, globalIndex);
                                        if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[GUID-MANAGER] Removed ClashZone {guid} from Global XML for category '{category}'");
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


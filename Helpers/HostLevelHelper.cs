using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;

namespace JSE_RevitAddin_MEP_OPENINGS.Helpers
{
    public static class HostLevelHelper
    {
        /// <summary>
        /// Gets the reference Level for a host element (pipe, duct, cable tray, damper).
        /// Returns null if not found.
        /// </summary>
            public static Level? GetHostReferenceLevel(Document doc, Element? host)
    {
        if (host == null) return null;
        
        int hostId = host.Id.GetIntegerValue();
        // JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Log($"[HostLevelHelper] DEBUG: Starting GetHostReferenceLevel for host {hostId} (Document: '{host.Document.Title}', IsLinked: {host.Document.IsLinked})");
        
        // FIXED: Always try to get the level from the linked document first for consistency
        if (host.Document.IsLinked)
        {
            try
            {
                // JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Log($"[HostLevelHelper] DEBUG: Host {hostId} - Processing linked document level lookup");
                
                // Get the level from the linked document
                Parameter linkedRefLevelParam = host.LookupParameter("Reference Level") ?? host.LookupParameter("Level");
                // JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Log($"[HostLevelHelper] DEBUG: Host {hostId} - Found linked param: {linkedRefLevelParam?.Definition.Name}, StorageType: {linkedRefLevelParam?.StorageType}");
                
                if (linkedRefLevelParam != null && linkedRefLevelParam.StorageType == StorageType.ElementId)
                {
                    ElementId linkedLevelId = linkedRefLevelParam.AsElementId();
                    // JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Log($"[HostLevelHelper] DEBUG: Host {hostId} - Linked level ElementId: {linkedLevelId.GetIntegerValue()} (Valid: {linkedLevelId != ElementId.InvalidElementId})");
                    
                    if (linkedLevelId != ElementId.InvalidElementId)
                    {
                        // Get the level from the linked document
                        Level? linkedLevel = host.Document.GetElement(linkedLevelId) as Level;
                        if (linkedLevel != null)
                        {
                            // JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Log($"[HostLevelHelper] DEBUG: Host {hostId} - Found level in linked doc: '{linkedLevel.Name}' (ID: {linkedLevel.Id.GetIntegerValue()})");
                            
                            // Find a matching level in the active document by name
                            var matchingLevel = new FilteredElementCollector(doc)
                                .OfClass(typeof(Level))
                                .Cast<Level>()
                                .FirstOrDefault(l => string.Equals(l.Name, linkedLevel.Name, StringComparison.OrdinalIgnoreCase));
                            
                            if (matchingLevel != null)
                            {
                                JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Log($"[HostLevelHelper] DEBUG: Host {hostId} - Found matching level in active doc: '{matchingLevel.Name}' (ID: {matchingLevel.Id.GetIntegerValue()}) - RETURNING THIS");
                                return matchingLevel;
                            }
                            else
                            {
                                // JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Log($"[HostLevelHelper] DEBUG: Host {hostId} - NO matching level found in active doc for '{linkedLevel.Name}'");
                                
                                // Log all available levels in active document for debugging
                                var allLevels = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().ToList();
                                // JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Log($"[HostLevelHelper] DEBUG: Host {hostId} - Available levels in active doc: {string.Join(", ", allLevels.Select(l => $"'{l.Name}' (ID: {l.Id.GetIntegerValue()})"))}");
                            }
                        }
                        else
                        {
                            JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Log($"[HostLevelHelper] DEBUG: Host {hostId} - Could not get level from linked document for ElementId {linkedLevelId.GetIntegerValue()}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Log error but don't fail
                JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Log($"[HostLevelHelper] ERROR: Host {hostId} - Error getting level from linked document: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"Error getting level from linked document: {ex.Message}");
            }
        }
        
        // Fallback: try to get level from active document (this was causing the inconsistency)
        JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Log($"[HostLevelHelper] DEBUG: Host {hostId} - Trying fallback to active document level lookup");
        Parameter refLevelParam = host.LookupParameter("Reference Level") ?? host.LookupParameter("Level");
        if (refLevelParam != null && refLevelParam.StorageType == StorageType.ElementId)
        {
            ElementId levelId = refLevelParam.AsElementId();
            JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Log($"[HostLevelHelper] DEBUG: Host {hostId} - Fallback ElementId: {levelId.GetIntegerValue()} (Valid: {levelId != ElementId.InvalidElementId})");
            if (levelId != ElementId.InvalidElementId)
            {
                Level? level = doc.GetElement(levelId) as Level;
                if (level != null)
                {
                    JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Log($"[HostLevelHelper] DEBUG: Host {hostId} - Found fallback level in active doc: '{level.Name}' (ID: {level.Id.GetIntegerValue()}) - RETURNING THIS");
                    return level;
                }
            }
        }
        
        JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Log($"[HostLevelHelper] DEBUG: Host {hostId} - Returning null (no level found)");
        return null;
    }
        /// <summary>
        /// ✅ NEW: Gets the elevation of the reference Level for a host element.
        /// CRITICAL: Returns the elevation from the LINKED document directly if the element is linked.
        /// This avoids the need for a matching level to exist in the active document just to get the elevation value.
        /// Returns null if not found.
        /// </summary>
        public static double? GetHostReferenceLevelElevation(Document doc, Element? host)
        {
            if (host == null) return null;

            int hostId = host.Id.GetIntegerValue();

            // ✅ PRIORITY 1: Try to get from linked document first
            if (host.Document.IsLinked)
            {
                try
                {
                    // Get the level from the linked document
                    Parameter linkedRefLevelParam = host.LookupParameter("Reference Level") ?? host.LookupParameter("Level");
                    
                    if (linkedRefLevelParam != null && linkedRefLevelParam.StorageType == StorageType.ElementId)
                    {
                        ElementId linkedLevelId = linkedRefLevelParam.AsElementId();
                        
                        if (linkedLevelId != ElementId.InvalidElementId)
                        {
                            // Get the level from the linked document
                            Level? linkedLevel = host.Document.GetElement(linkedLevelId) as Level;
                            if (linkedLevel != null)
                            {
                                // JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Log($"[HostLevelHelper] DEBUG: Host {hostId} - Found level in linked doc: '{linkedLevel.Name}', Elevation: {linkedLevel.Elevation}");
                                return linkedLevel.Elevation;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    // Log error but don't fail
                    JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Log($"[HostLevelHelper] ERROR: Host {hostId} - Error getting level elevation from linked document: {ex.Message}");
                }
            }

            // ✅ PRIORITY 2: Fallback to active document level lookup (existing method)
            // This is useful for non-linked elements or if linked lookup failed
            Level? level = GetHostReferenceLevel(doc, host);
            if (level != null)
            {
                return level.Elevation;
            }

            return null;
        }

        /// <summary>
        /// ✅ NEW: Gets Reference Level elevation from ClashZone.MepParameterValues.
        /// Prioritizes "Reference Level" as primary source, falls back to other level parameters if not found.
        /// Returns null if not found.
        /// </summary>
        /// <param name="doc">Active document (for level lookup by name)</param>
        /// <param name="mepParameterValues">MEP parameter values from ClashZone (captured during refresh)</param>
        /// <returns>Reference Level elevation if found, null otherwise</returns>
        public static double? GetReferenceLevelElevationFromParameters(Document doc, System.Collections.Generic.List<Models.SerializableKeyValue>? mepParameterValues)
        {
            if (mepParameterValues == null || mepParameterValues.Count == 0)
                return null;

            // ✅ PRIMARY: Try "Reference Level" first
            var refLevelParam = mepParameterValues
                .FirstOrDefault(p => string.Equals(p.Key, "Reference Level", StringComparison.OrdinalIgnoreCase));
            
            if (refLevelParam != null && !string.IsNullOrWhiteSpace(refLevelParam.Value))
            {
                var levelByName = new FilteredElementCollector(doc)
                    .OfClass(typeof(Level))
                    .Cast<Level>()
                    .FirstOrDefault(l => string.Equals(l.Name, refLevelParam.Value, StringComparison.OrdinalIgnoreCase));
                
                if (levelByName != null)
                {
                    return levelByName.Elevation;
                }
            }

            // ✅ FALLBACK: If "Reference Level" not found, try other level parameters
            var fallbackLevelParam = mepParameterValues
                .FirstOrDefault(p => string.Equals(p.Key, "Level", StringComparison.OrdinalIgnoreCase) ||
                                     string.Equals(p.Key, "Schedule Level", StringComparison.OrdinalIgnoreCase) ||
                                     string.Equals(p.Key, "Schedule of Level", StringComparison.OrdinalIgnoreCase));
            
            if (fallbackLevelParam != null && !string.IsNullOrWhiteSpace(fallbackLevelParam.Value))
            {
                var levelByName = new FilteredElementCollector(doc)
                    .OfClass(typeof(Level))
                    .Cast<Level>()
                    .FirstOrDefault(l => string.Equals(l.Name, fallbackLevelParam.Value, StringComparison.OrdinalIgnoreCase));
                
                if (levelByName != null)
                {
                    return levelByName.Elevation;
                }
            }

            return null;
        }
    }
}
using Autodesk.Revit.DB;

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
        
        int hostId = host.Id.IntegerValue;
        JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Log($"[HostLevelHelper] DEBUG: Starting GetHostReferenceLevel for host {hostId} (Document: '{host.Document.Title}', IsLinked: {host.Document.IsLinked})");
        
        // FIXED: Always try to get the level from the linked document first for consistency
        if (host.Document.IsLinked)
        {
            try
            {
                JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Log($"[HostLevelHelper] DEBUG: Host {hostId} - Processing linked document level lookup");
                
                // Get the level from the linked document
                Parameter linkedRefLevelParam = host.LookupParameter("Reference Level") ?? host.LookupParameter("Level");
                JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Log($"[HostLevelHelper] DEBUG: Host {hostId} - Found linked param: {linkedRefLevelParam?.Definition.Name}, StorageType: {linkedRefLevelParam?.StorageType}");
                
                if (linkedRefLevelParam != null && linkedRefLevelParam.StorageType == StorageType.ElementId)
                {
                    ElementId linkedLevelId = linkedRefLevelParam.AsElementId();
                    JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Log($"[HostLevelHelper] DEBUG: Host {hostId} - Linked level ElementId: {linkedLevelId.IntegerValue} (Valid: {linkedLevelId != ElementId.InvalidElementId})");
                    
                    if (linkedLevelId != ElementId.InvalidElementId)
                    {
                        // Get the level from the linked document
                        Level? linkedLevel = host.Document.GetElement(linkedLevelId) as Level;
                        if (linkedLevel != null)
                        {
                            JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Log($"[HostLevelHelper] DEBUG: Host {hostId} - Found level in linked doc: '{linkedLevel.Name}' (ID: {linkedLevel.Id.IntegerValue})");
                            
                            // Find a matching level in the active document by name
                            var matchingLevel = new FilteredElementCollector(doc)
                                .OfClass(typeof(Level))
                                .Cast<Level>()
                                .FirstOrDefault(l => string.Equals(l.Name, linkedLevel.Name, StringComparison.OrdinalIgnoreCase));
                            
                            if (matchingLevel != null)
                            {
                                JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Log($"[HostLevelHelper] DEBUG: Host {hostId} - Found matching level in active doc: '{matchingLevel.Name}' (ID: {matchingLevel.Id.IntegerValue}) - RETURNING THIS");
                                return matchingLevel;
                            }
                            else
                            {
                                JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Log($"[HostLevelHelper] DEBUG: Host {hostId} - NO matching level found in active doc for '{linkedLevel.Name}'");
                                
                                // Log all available levels in active document for debugging
                                var allLevels = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().ToList();
                                JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Log($"[HostLevelHelper] DEBUG: Host {hostId} - Available levels in active doc: {string.Join(", ", allLevels.Select(l => $"'{l.Name}' (ID: {l.Id.IntegerValue})"))}");
                            }
                        }
                        else
                        {
                            JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Log($"[HostLevelHelper] DEBUG: Host {hostId} - Could not get level from linked document for ElementId {linkedLevelId.IntegerValue}");
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
            JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Log($"[HostLevelHelper] DEBUG: Host {hostId} - Fallback ElementId: {levelId.IntegerValue} (Valid: {levelId != ElementId.InvalidElementId})");
            if (levelId != ElementId.InvalidElementId)
            {
                Level? level = doc.GetElement(levelId) as Level;
                if (level != null)
                {
                    JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Log($"[HostLevelHelper] DEBUG: Host {hostId} - Found fallback level in active doc: '{level.Name}' (ID: {level.Id.IntegerValue}) - RETURNING THIS");
                    return level;
                }
            }
        }
        
        JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Log($"[HostLevelHelper] DEBUG: Host {hostId} - Returning null (no level found)");
        return null;
    }
    }
}

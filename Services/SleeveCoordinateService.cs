using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Models;

using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Service to handle sleeve coordinate operations
    /// </summary>
    public class SleeveCoordinateService
    {
        private readonly Document _doc;
        private Dictionary<long, ClashZone> _clashZoneCache;
        
        public SleeveCoordinateService(Document doc)
        {
            _doc = doc;
            _clashZoneCache = new Dictionary<long, ClashZone>();
        }
        
        /// <summary>
        /// Get all sleeve coordinates from the model and log them
        /// </summary>
        public void LogAllSleeveCoordinates()
        {
            try
            {
                // Get all sleeves in the model
                var sleeves = new FilteredElementCollector(_doc)
                    .OfClass(typeof(FamilyInstance))
                    .Cast<FamilyInstance>()
                    .Where(s => s.Symbol.FamilyName.Contains("Opening"))
                    .ToList();
                
                if (sleeves.Count == 0)
                {
                    string sleeveCoordinatesLogPath = SafeFileLogger.GetLogFilePath("sleeve_coordinates.log");
                    System.IO.File.AppendAllText(sleeveCoordinatesLogPath, "No sleeves found in model!\n");
                    return;
                }
                
                // Log all sleeve coordinates
                string logPath = SafeFileLogger.GetLogFilePath("all_sleeve_coordinates.log");
                System.IO.File.WriteAllText(logPath, $"ALL SLEEVE COORDINATES - {DateTime.Now}\n");
                System.IO.File.AppendAllText(logPath, $"Found {sleeves.Count} sleeves\n\n");
                
                foreach (var sleeve in sleeves)
                {
                    var bbox = sleeve.get_BoundingBox(null);
                    if (bbox != null)
                    {
                        var min = bbox.Min;
                        var max = bbox.Max;
                        
                        System.IO.File.AppendAllText(logPath, 
                            $"SLEEVE {sleeve.Id.IntegerValue}:\n");
                        System.IO.File.AppendAllText(logPath, 
                            $"  Min: ({min.X:F6}, {min.Y:F6}, {min.Z:F6})\n");
                        System.IO.File.AppendAllText(logPath, 
                            $"  Max: ({max.X:F6}, {max.Y:F6}, {max.Z:F6})\n");
                        System.IO.File.AppendAllText(logPath, 
                            $"  Host: {sleeve.Host?.Id?.IntegerValue}\n");
                        System.IO.File.AppendAllText(logPath, 
                            $"  Family: {sleeve.Symbol.FamilyName}\n\n");
                    }
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[DEBUG] Exception: {ex.Message}\n");
                }
            }
        }
        
        /// <summary>
        /// Update sleeve coordinates in XML files with correct coordinates from model
        /// </summary>
        public void UpdateSleeveCoordinatesInXml(string xmlFilePath = null)
        {
            try
            {
                // ✅ Direct file logging to placement_debug.log
                var placementDebugPath = SafeFileLogger.GetLogFilePath("placement_debug.log");
                try { System.IO.File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [UpdateSleeveCoordinatesInXml] CALLED with xmlFilePath: {xmlFilePath ?? "NULL"}\n"); } catch { }
                
                // Get all sleeves
                var sleeves = new FilteredElementCollector(_doc)
                    .OfClass(typeof(FamilyInstance))
                    .Cast<FamilyInstance>()
                    .Where(s => s.Symbol.FamilyName.Contains("Opening"))
                    .ToList();
                
                try { System.IO.File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [UpdateSleeveCoordinatesInXml] Found {sleeves.Count} sleeves in Revit model\n"); } catch { }
                
                var coordinateUpdater = new SleeveCoordinateUpdater(_doc);
                
                // ✅ DYNAMIC: Load clash zones from specific XML file
                var clashZones = LoadClashZonesFromXml(xmlFilePath);
                
                // ✅ CRITICAL DEBUG: Log how many have SleeveInstanceId (calculate outside condition for use later)
                var withSleeveId = clashZones.Count(cz => cz.SleeveInstanceId > 0);
                
                try 
                { 
                    System.IO.File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [UpdateSleeveCoordinatesInXml] Loaded {clashZones.Count} clash zones from XML\n");
                    System.IO.File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [UpdateSleeveCoordinatesInXml] {withSleeveId} out of {clashZones.Count} clash zones have SleeveInstanceId > 0\n");
                    
                    // ✅ CRITICAL: Log which sleeve IDs we're looking for
                    if (withSleeveId > 0)
                    {
                        var sleeveIds = clashZones.Where(cz => cz.SleeveInstanceId > 0).Select(cz => cz.SleeveInstanceId).Take(10).ToList();
                        System.IO.File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [UpdateSleeveCoordinatesInXml] Looking for sleeves with IDs: [{string.Join(", ", sleeveIds)}...]\n");
                    }
                } 
                catch { }
                
                // Update coordinates
                coordinateUpdater.UpdateSleeveCoordinates(clashZones);
                
                // ✅ CRITICAL DEBUG: Log how many were actually updated after UpdateSleeveCoordinates - Direct file write
                try
                {
                    var withBbox = clashZones.Count(cz => cz.SleeveInstanceId > 0 && 
                        !(cz.SleeveBoundingBoxMinX == 0.0 && cz.SleeveBoundingBoxMinY == 0.0 && cz.SleeveBoundingBoxMinZ == 0.0 &&
                          cz.SleeveBoundingBoxMaxX == 0.0 && cz.SleeveBoundingBoxMaxY == 0.0 && cz.SleeveBoundingBoxMaxZ == 0.0));
                    System.IO.File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [UpdateSleeveCoordinatesInXml] After UpdateSleeveCoordinates: {withBbox} out of {withSleeveId} clash zones with SleeveInstanceId now have bounding boxes\n");
                }
                catch { }
                
                // Save updated XML
                SaveClashZonesToXml(clashZones, xmlFilePath);
                
                // ✅ Direct file write
                try { System.IO.File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [UpdateSleeveCoordinatesInXml] COMPLETED - Updated coordinates for {sleeves.Count} sleeves\n"); } catch { }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Error($"[UpdateSleeveCoordinatesInXml] EXCEPTION: {ex.Message}\n");
                    DebugLogger.Error($"[UpdateSleeveCoordinatesInXml] Stack trace: {ex.StackTrace}\n");
                }
            }
        }
        
        private List<ClashZone> LoadClashZonesFromXml(string xmlFilePath = null)
        {
            var clashZones = new List<ClashZone>();
            
            try
            {
                // ✅ STRICT: Use ONLY the specified file path - no file searching allowed
                if (string.IsNullOrEmpty(xmlFilePath) || !File.Exists(xmlFilePath))
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[LOAD-XML] ERROR: xmlFilePath not provided or file doesn't exist: {xmlFilePath}\n");
                    }
                    return clashZones;
                }
                
                var xmlFiles = new[] { xmlFilePath };
                
                // ✅ DEPLOYMENT: Wrapped in deployment mode check
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[LOAD-XML] Processing ONLY file: {xmlFilePath}\n");
                }
                
                foreach (var xmlFile in xmlFiles)
                {
                    try
                    {
                        var xmlDoc = new System.Xml.XmlDocument();
                        xmlDoc.Load(xmlFile);
                        
                        var clashNodes = xmlDoc.SelectNodes("//ClashZone");
                        if (clashNodes != null)
                        {
                            foreach (System.Xml.XmlNode node in clashNodes)
                            {
                                var clashZone = new ClashZone();
                                
                                // ✅ CRITICAL: Load ClashZone ID for matching
                                if (Guid.TryParse(node.SelectSingleNode("Id")?.InnerText, out Guid clashZoneId))
                                    clashZone.Id = clashZoneId;
                                
                                // Load basic properties
                                if (int.TryParse(node.SelectSingleNode("SleeveInstanceId")?.InnerText, out int sleeveId))
                                    clashZone.SleeveInstanceId = sleeveId;
                                
                                // Load placement point coordinates (for position matching)
                                if (double.TryParse(node.SelectSingleNode("SleevePlacementPointX")?.InnerText, out double placeX))
                                    clashZone.SleevePlacementPointX = placeX;
                                if (double.TryParse(node.SelectSingleNode("SleevePlacementPointY")?.InnerText, out double placeY))
                                    clashZone.SleevePlacementPointY = placeY;
                                if (double.TryParse(node.SelectSingleNode("SleevePlacementPointZ")?.InnerText, out double placeZ))
                                    clashZone.SleevePlacementPointZ = placeZ;
                                
                                // Load bounding box coordinates
                                if (double.TryParse(node.SelectSingleNode("SleeveBoundingBoxMinX")?.InnerText, out double minX))
                                    clashZone.SleeveBoundingBoxMinX = minX;
                                if (double.TryParse(node.SelectSingleNode("SleeveBoundingBoxMinY")?.InnerText, out double minY))
                                    clashZone.SleeveBoundingBoxMinY = minY;
                                if (double.TryParse(node.SelectSingleNode("SleeveBoundingBoxMinZ")?.InnerText, out double minZ))
                                    clashZone.SleeveBoundingBoxMinZ = minZ;
                                if (double.TryParse(node.SelectSingleNode("SleeveBoundingBoxMaxX")?.InnerText, out double maxX))
                                    clashZone.SleeveBoundingBoxMaxX = maxX;
                                if (double.TryParse(node.SelectSingleNode("SleeveBoundingBoxMaxY")?.InnerText, out double maxY))
                                    clashZone.SleeveBoundingBoxMaxY = maxY;
                                if (double.TryParse(node.SelectSingleNode("SleeveBoundingBoxMaxZ")?.InnerText, out double maxZ))
                                    clashZone.SleeveBoundingBoxMaxZ = maxZ;
                                
                                // ✅ NEW: Load cluster sleeve instance ID
                                if (int.TryParse(node.SelectSingleNode("ClusterSleeveInstanceId")?.InnerText, out int clusterSleeveInstanceId))
                                    clashZone.ClusterSleeveInstanceId = clusterSleeveInstanceId;
                                
                                // ✅ DEBUG: Log cluster data loading
                                // ✅ DEPLOYMENT: Wrapped in deployment mode check
                                if (!DeploymentConfiguration.DeploymentMode && clashZone.ClusterSleeveInstanceId > 0)
                                {
                                    DebugLogger.Info($"[LOAD-XML] Loaded ClusterSleeveInstanceId={clashZone.ClusterSleeveInstanceId} from XML for ClashZone {clashZone.Id}\n");
                                }
                                
                                // ✅ NEW: Load cluster sleeve bounding box coordinates
                                if (double.TryParse(node.SelectSingleNode("ClusterSleeveBoundingBoxMinX")?.InnerText, out double clusterMinX))
                                    clashZone.ClusterSleeveBoundingBoxMinX = clusterMinX;
                                if (double.TryParse(node.SelectSingleNode("ClusterSleeveBoundingBoxMinY")?.InnerText, out double clusterMinY))
                                    clashZone.ClusterSleeveBoundingBoxMinY = clusterMinY;
                                if (double.TryParse(node.SelectSingleNode("ClusterSleeveBoundingBoxMinZ")?.InnerText, out double clusterMinZ))
                                    clashZone.ClusterSleeveBoundingBoxMinZ = clusterMinZ;
                                if (double.TryParse(node.SelectSingleNode("ClusterSleeveBoundingBoxMaxX")?.InnerText, out double clusterMaxX))
                                    clashZone.ClusterSleeveBoundingBoxMaxX = clusterMaxX;
                                if (double.TryParse(node.SelectSingleNode("ClusterSleeveBoundingBoxMaxY")?.InnerText, out double clusterMaxY))
                                    clashZone.ClusterSleeveBoundingBoxMaxY = clusterMaxY;
                                if (double.TryParse(node.SelectSingleNode("ClusterSleeveBoundingBoxMaxZ")?.InnerText, out double clusterMaxZ))
                                    clashZone.ClusterSleeveBoundingBoxMaxZ = clusterMaxZ;
                                
                                clashZones.Add(clashZone);
                            }
                        }
                        
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[DEBUG] Loaded {clashNodes?.Count ?? 0} clash zones from {Path.GetFileName(xmlFile)}\n");
                        }
                    }
                    catch (Exception ex)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[DEBUG] Error loading {xmlFile}: {ex.Message}\n");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[DEBUG] Error in LoadClashZonesFromXml: {ex.Message}\n");
                }
            }
            
            return clashZones;
        }
        
        public void SaveClashZonesToXml(List<ClashZone> clashZones, string xmlFilePath = null)
        {
            try
            {
                // ✅ STRICT: Use ONLY the specified file path - no file searching allowed
                if (string.IsNullOrEmpty(xmlFilePath) || !File.Exists(xmlFilePath))
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[SAVE-XML] ERROR: xmlFilePath not provided or file doesn't exist: {xmlFilePath}\n");
                    }
                    return;
                }
                
                var xmlFiles = new[] { xmlFilePath };
                
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[SAVE-XML] Saving to ONLY file: {xmlFilePath}\n");
                }
                
                foreach (var xmlFile in xmlFiles)
                {
                    try
                    {
                        var xmlDoc = new System.Xml.XmlDocument();
                        xmlDoc.Load(xmlFile);
                        
                        var clashNodes = xmlDoc.SelectNodes("//ClashZone");
                        if (clashNodes != null)
                        {
                            foreach (System.Xml.XmlNode node in clashNodes)
                            {
                                // ✅ FIX: Match by ClashZone ID instead of wrong SleeveInstanceId
                                var clashZoneIdText = node.SelectSingleNode("Id")?.InnerText;
                                if (!string.IsNullOrEmpty(clashZoneIdText) && Guid.TryParse(clashZoneIdText, out Guid clashZoneId))
                                {
                                    // Find matching clash zone by ID
                                    var matchingClashZone = clashZones.FirstOrDefault(cz => cz.Id == clashZoneId);
                                    if (matchingClashZone != null)
                                    {
                                        // ✅ CRITICAL FIX: Update cluster sleeve bounding boxes FIRST (independent of individual sleeve)
                                        if (matchingClashZone.ClusterSleeveInstanceId > 0)
                                        {
                                            UpdateXmlNode(node, "ClusterSleeveInstanceId", matchingClashZone.ClusterSleeveInstanceId.ToString());
                                            UpdateXmlNode(node, "ClusterSleeveBoundingBoxMinX", matchingClashZone.ClusterSleeveBoundingBoxMinX.ToString("F6"));
                                            UpdateXmlNode(node, "ClusterSleeveBoundingBoxMinY", matchingClashZone.ClusterSleeveBoundingBoxMinY.ToString("F6"));
                                            UpdateXmlNode(node, "ClusterSleeveBoundingBoxMinZ", matchingClashZone.ClusterSleeveBoundingBoxMinZ.ToString("F6"));
                                            UpdateXmlNode(node, "ClusterSleeveBoundingBoxMaxX", matchingClashZone.ClusterSleeveBoundingBoxMaxX.ToString("F6"));
                                            UpdateXmlNode(node, "ClusterSleeveBoundingBoxMaxY", matchingClashZone.ClusterSleeveBoundingBoxMaxY.ToString("F6"));
                                            UpdateXmlNode(node, "ClusterSleeveBoundingBoxMaxZ", matchingClashZone.ClusterSleeveBoundingBoxMaxZ.ToString("F6"));
                                            
                                            // ✅ DEPLOYMENT: Wrapped in deployment mode check
                                            if (!DeploymentConfiguration.DeploymentMode)
                                            {
                                                DebugLogger.Info($"[XML-UPDATE-CLUSTER] Updated cluster sleeve {matchingClashZone.ClusterSleeveInstanceId} bbox\n");
                                            }
                                        }
                                        
                                        // ✅ CRITICAL FIX: Update individual sleeve if it exists - always save bounding box if SleeveInstanceId > 0
                                        // The issue was: checking "bbox != 0" prevented saving valid bounding boxes (coordinates can be negative or small)
                                        // Solution: Check if bbox was actually set (any non-zero value in any coordinate) OR if coordinates differ from default zeros
                                        if (matchingClashZone.SleeveInstanceId > 0)
                                        {
                                            // ✅ CRITICAL LOGGING: Log bounding box values BEFORE saving to XML - Direct file write
                                            var placementDebugPath = SafeFileLogger.GetLogFilePath("placement_debug.log");
                                            try 
                                            { 
                                                System.IO.File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [BOUNDING_BOX_BEFORE_XML_SAVE] ClashZone {clashZoneId}, SleeveInstanceId={matchingClashZone.SleeveInstanceId}: " +
                                                    $"MinX={matchingClashZone.SleeveBoundingBoxMinX:F6}, MinY={matchingClashZone.SleeveBoundingBoxMinY:F6}, MinZ={matchingClashZone.SleeveBoundingBoxMinZ:F6}, " +
                                                    $"MaxX={matchingClashZone.SleeveBoundingBoxMaxX:F6}, MaxY={matchingClashZone.SleeveBoundingBoxMaxY:F6}, MaxZ={matchingClashZone.SleeveBoundingBoxMaxZ:F6}\n");
                                            } 
                                            catch { }
                                            
                                            // ✅ CRITICAL: Check if bounding box has been set (any coordinate is non-zero, or all are set but different from 0)
                                            // Use a more robust check: if ALL coordinates are exactly 0.0, then bbox wasn't set
                                            bool hasValidBbox = !(matchingClashZone.SleeveBoundingBoxMinX == 0.0 && 
                                                                matchingClashZone.SleeveBoundingBoxMinY == 0.0 && 
                                                                matchingClashZone.SleeveBoundingBoxMinZ == 0.0 &&
                                                                matchingClashZone.SleeveBoundingBoxMaxX == 0.0 && 
                                                                matchingClashZone.SleeveBoundingBoxMaxY == 0.0 && 
                                                                matchingClashZone.SleeveBoundingBoxMaxZ == 0.0);
                                            
                                            if (hasValidBbox)
                                            {
                                                // ✅ CRITICAL: Update SleeveInstanceId with correct Revit element ID
                                                UpdateXmlNode(node, "SleeveInstanceId", matchingClashZone.SleeveInstanceId.ToString());
                                                
                                                // Update bounding box coordinates
                                                UpdateXmlNode(node, "SleeveBoundingBoxMinX", matchingClashZone.SleeveBoundingBoxMinX.ToString("F6"));
                                                UpdateXmlNode(node, "SleeveBoundingBoxMinY", matchingClashZone.SleeveBoundingBoxMinY.ToString("F6"));
                                                UpdateXmlNode(node, "SleeveBoundingBoxMinZ", matchingClashZone.SleeveBoundingBoxMinZ.ToString("F6"));
                                                UpdateXmlNode(node, "SleeveBoundingBoxMaxX", matchingClashZone.SleeveBoundingBoxMaxX.ToString("F6"));
                                                UpdateXmlNode(node, "SleeveBoundingBoxMaxY", matchingClashZone.SleeveBoundingBoxMaxY.ToString("F6"));
                                                UpdateXmlNode(node, "SleeveBoundingBoxMaxZ", matchingClashZone.SleeveBoundingBoxMaxZ.ToString("F6"));
                                                
                                                // ✅ CRITICAL LOGGING: Log bounding box values AFTER saving to XML
                                                // ✅ DEPLOYMENT: Wrapped in deployment mode check
                                                if (!DeploymentConfiguration.DeploymentMode)
                                                {
                                                    var savedMinX = node.SelectSingleNode("SleeveBoundingBoxMinX")?.InnerText ?? "NULL";
                                                    var savedMinY = node.SelectSingleNode("SleeveBoundingBoxMinY")?.InnerText ?? "NULL";
                                                    var savedMinZ = node.SelectSingleNode("SleeveBoundingBoxMinZ")?.InnerText ?? "NULL";
                                                    var savedMaxX = node.SelectSingleNode("SleeveBoundingBoxMaxX")?.InnerText ?? "NULL";
                                                    var savedMaxY = node.SelectSingleNode("SleeveBoundingBoxMaxY")?.InnerText ?? "NULL";
                                                    var savedMaxZ = node.SelectSingleNode("SleeveBoundingBoxMaxZ")?.InnerText ?? "NULL";
                                                    
                                                    try 
                                                    { 
                                                        System.IO.File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [BOUNDING_BOX_AFTER_XML_SAVE] ClashZone {clashZoneId}, SleeveInstanceId={matchingClashZone.SleeveInstanceId}: " +
                                                            $"XML now has - MinX={savedMinX}, MinY={savedMinY}, MinZ={savedMinZ}, " +
                                                            $"MaxX={savedMaxX}, MaxY={savedMaxY}, MaxZ={savedMaxZ}\n");
                                                    } 
                                                    catch { }
                                                }
                                                
                                                // ✅ DEPLOYMENT: Wrapped in deployment mode check
                                                if (!DeploymentConfiguration.DeploymentMode)
                                                {
                                                    DebugLogger.Info($"[XML-UPDATE] Updated ClashZone {clashZoneId} with SleeveInstanceId {matchingClashZone.SleeveInstanceId} and bounding box coordinates\n");
                                                }
                                            }
                                            else
                                            {
                                                // Bounding box not set yet - log warning but don't skip SleeveInstanceId update
                                                UpdateXmlNode(node, "SleeveInstanceId", matchingClashZone.SleeveInstanceId.ToString());
                                                if (!DeploymentConfiguration.DeploymentMode)
                                                {
                                                    DebugLogger.Warning($"[XML-WARNING] ClashZone {clashZoneId} has SleeveInstanceId={matchingClashZone.SleeveInstanceId} but bounding box is still zero. Run 'Update XML' to refresh from Revit.\n");
                                                }
                                            }
                                        }
                                        else
                                        {
                                            // ✅ DEPLOYMENT: Wrapped in deployment mode check
                                            if (!DeploymentConfiguration.DeploymentMode)
                                            {
                                                DebugLogger.Info($"[XML-SKIP] Skipped ClashZone {clashZoneId} - SleeveInstanceId={matchingClashZone.SleeveInstanceId} (no sleeve placed yet)\n");
                                            }
                                        }
                                    }
                                    else
                                    {
                                        // ✅ DEPLOYMENT: Wrapped in deployment mode check
                                        if (!DeploymentConfiguration.DeploymentMode)
                                        {
                                            DebugLogger.Info($"[XML-NO-MATCH] No matching clash zone found for Id {clashZoneId}\n");
                                        }
                                    }
                                }
                                else
                                {
                                    // ✅ DEPLOYMENT: Wrapped in deployment mode check
                                    if (!DeploymentConfiguration.DeploymentMode)
                                    {
                                        DebugLogger.Info($"[XML-NO-ID] ClashZone node has no valid Id\n");
                                    }
                                }
                            }
                            
                            // ✅ CRITICAL: Save the updated XML with error handling
                            try
                            {
                                xmlDoc.Save(xmlFile);
                                
                                // ✅ Direct file write to confirm save
                                var placementDebugPath = SafeFileLogger.GetLogFilePath("placement_debug.log");
                                try 
                                { 
                                    System.IO.File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [XML_FILE_SAVED] Successfully saved {Path.GetFileName(xmlFile)} to disk\n"); 
                                } 
                                catch { }
                                
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    DebugLogger.Info($"[DEBUG] Updated coordinates in {Path.GetFileName(xmlFile)}\n");
                                }
                            }
                            catch (Exception saveEx)
                            {
                                // ✅ CRITICAL: Log save failure
                                var placementDebugPath = SafeFileLogger.GetLogFilePath("placement_debug.log");
                                try 
                                { 
                                    System.IO.File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [XML_SAVE_ERROR] FAILED to save {Path.GetFileName(xmlFile)}: {saveEx.Message}\n"); 
                                    System.IO.File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [XML_SAVE_ERROR] Stack trace: {saveEx.StackTrace}\n"); 
                                } 
                                catch { }
                                
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    DebugLogger.Error($"[SAVE-XML] ERROR saving {Path.GetFileName(xmlFile)}: {saveEx.Message}\n");
                                    DebugLogger.Error($"[SAVE-XML] Stack trace: {saveEx.StackTrace}\n");
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[DEBUG] Error updating {xmlFile}: {ex.Message}\n");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[DEBUG] Error in SaveClashZonesToXml: {ex.Message}\n");
                }
            }
        }
        
        private void UpdateXmlNode(System.Xml.XmlNode parentNode, string nodeName, string value)
        {
            var node = parentNode.SelectSingleNode(nodeName);
            if (node != null)
            {
                node.InnerText = value;
            }
            else
            {
                // Create node if it doesn't exist
                var newNode = parentNode.OwnerDocument.CreateElement(nodeName);
                newNode.InnerText = value;
                parentNode.AppendChild(newNode);
            }
        }
        
        /// <summary>
        /// MASTER DEBUGGER FIX: Regenerate _CLUSTER.xml with actual Revit coordinates after sleeve placement
        /// Waits for Revit to update, then collects real coordinates via API
        /// </summary>
        public void RegenerateClusterXmlAfterPlacement(string filterName = null)
        {
            try
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[REGENERATE-CLUSTER-XML] Starting regeneration after sleeve placement\n");
                }
                
                // ✅ CRITICAL: Load clash zone cache first
                LoadClashZoneCache();
                
                // ✅ CRITICAL: Wait for Revit to update the model
                System.Threading.Thread.Sleep(500); // Reduced to 0.5 seconds since timing is working
                
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[COORDINATE-SERVICE] {DateTime.Now:HH:mm:ss.fff} - Starting coordinate collection after 0.5-second wait\n");
                }
                
                // ✅ MASTER FIX: Collect ALL sleeves with actual Revit coordinates
                var allSleeves = new FilteredElementCollector(_doc)
                    .OfClass(typeof(FamilyInstance))
                    .Cast<FamilyInstance>()
                    .Where(s => s.Symbol.FamilyName.Contains("Opening"))
                    .ToList();
                
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[REGENERATE-CLUSTER-XML] Found {allSleeves.Count} sleeves in Revit model\n");
                    DebugLogger.Info($"[COORDINATE-SERVICE] {DateTime.Now:HH:mm:ss.fff} - Found {allSleeves.Count} total sleeves in Revit model\n");
                }
                
                // Group sleeves by category for separate XML files
                var groupedSleeves = allSleeves.GroupBy(s => GetSleeveCategory(s)).ToList();
                
                foreach (var group in groupedSleeves)
                {
                    var category = group.Key;
                    var sleeves = group.ToList();
                    
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[COORDINATE-SERVICE] {DateTime.Now:HH:mm:ss.fff} - Processing {sleeves.Count} sleeves for category: {category}\n");
                        DebugLogger.Info($"[REGENERATE-CLUSTER-XML] Processing {sleeves.Count} sleeves for category: {category}\n");
                    }
                    
                    // Create SleeveDataList with actual Revit coordinates
                    var sleeveDataList = new List<SleeveData>();
                    
                    foreach (var sleeve in sleeves)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            // ✅ CRITICAL LOGGING: Log each sleeve found after wait time
                            DebugLogger.Info($"[SLEEVE-FOUND] {DateTime.Now:HH:mm:ss.fff} - RevitElementId = {sleeve.Id.IntegerValue}, Category = {category}, Family = {sleeve.Symbol.FamilyName}\n");
                        }
                        
                        var bbox = sleeve.get_BoundingBox(null);
                        if (bbox != null)
                        {
                            // ✅ MASTER FIX: Get actual Revit coordinates
                            var corner1 = new XYZ(bbox.Min.X, bbox.Min.Y, bbox.Min.Z);
                            var corner2 = new XYZ(bbox.Max.X, bbox.Min.Y, bbox.Min.Z);
                            var corner3 = new XYZ(bbox.Max.X, bbox.Max.Y, bbox.Max.Z);
                            var corner4 = new XYZ(bbox.Min.X, bbox.Max.Y, bbox.Max.Z);
                            
                            var sleeveData = new SleeveData
                            {
                                SleeveInstanceId = sleeve.Id.IntegerValue,
                                Corner1 = corner1,
                                Corner2 = corner2,
                                Corner3 = corner3,
                                Corner4 = corner4,
                                Width = UnitUtils.ConvertFromInternalUnits(bbox.Max.X - bbox.Min.X, UnitTypeId.Meters),
                                Height = UnitUtils.ConvertFromInternalUnits(bbox.Max.Y - bbox.Min.Y, UnitTypeId.Meters),
                                Depth = UnitUtils.ConvertFromInternalUnits(bbox.Max.Z - bbox.Min.Z, UnitTypeId.Meters),
                                HostType = GetHostType(sleeve),
                                Orientation = GetSleeveOrientation(sleeve),
                                Category = category,
                                CreatedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
                            };
                            
                            sleeveDataList.Add(sleeveData);
                            
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Info($"[REGENERATE-CLUSTER-XML] Sleeve {sleeve.Id.IntegerValue}: Min=({bbox.Min.X:F6}, {bbox.Min.Y:F6}, {bbox.Min.Z:F6}), Max=({bbox.Max.X:F6}, {bbox.Max.Y:F6}, {bbox.Max.Z:F6})\n");
                            }
                        }
                    }
                    
                    // ✅ MASTER FIX: Save to _CLUSTER.xml with actual coordinates
                    SaveSleeveDataToClusterXml(sleeveDataList, category, filterName);
                    
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        // ✅ CRITICAL LOGGING: Log summary for this category
                        DebugLogger.Info($"[CATEGORY-SUMMARY] {DateTime.Now:HH:mm:ss.fff} - Category '{category}': Saved {sleeveDataList.Count} sleeves to _CLUSTER.xml\n");
                    }
                }
                
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[COORDINATE-SERVICE] {DateTime.Now:HH:mm:ss.fff} - COMPLETED: Processed {groupedSleeves.Count} categories\n");
                    DebugLogger.Info($"[REGENERATE-CLUSTER-XML] Completed regeneration for {groupedSleeves.Count} categories\n");
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[REGENERATE-CLUSTER-XML] ERROR: {ex.Message}\n");
                }
            }
        }
        
        private string GetSleeveCategory(FamilyInstance sleeve)
        {
            // ✅ MASTER DEBUGGER FIX: Determine category based on MEP element, not just family name
            try
            {
                // Get the MEP element that this sleeve is penetrating
                var mepElementId = sleeve.LookupParameter("MEP_ElementId");
                if (mepElementId != null && mepElementId.HasValue)
                {
                    var mepElement = _doc.GetElement(mepElementId.AsElementId());
                    if (mepElement != null)
                    {
                        var mepCategory = mepElement.Category?.Name;
                        
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            // ✅ CRITICAL DEBUGGING: Log MEP element details
                            DebugLogger.Info($"[CATEGORY-DEBUG] Sleeve {sleeve.Id.IntegerValue}: MEP_ElementId={mepElementId.AsInteger()}, MEP_Category='{mepCategory}', MEP_Type='{mepElement.GetType().Name}'\n");
                        }
                        
                        if (!string.IsNullOrEmpty(mepCategory))
                        {
                            // Map Revit categories to our system categories
                            if (mepCategory.Contains("Pipe")) return "Pipes";
                            if (mepCategory.Contains("Duct")) return "Ducts";
                            if (mepCategory.Contains("Cable")) return "Cable Trays";
                            
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Info($"[CATEGORY-DEBUG] Sleeve {sleeve.Id.IntegerValue}: MEP Category='{mepCategory}', Family='{sleeve.Symbol.FamilyName}'\n");
                            }
                        }
                    }
                }
                
                // Fallback: Check family name
                var familyName = sleeve.Symbol.FamilyName.ToLower();
                
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    // ✅ CRITICAL DEBUGGING: Log fallback logic
                    DebugLogger.Info($"[CATEGORY-FALLBACK] Sleeve {sleeve.Id.IntegerValue}: FamilyName='{familyName}', Using fallback logic\n");
                }
                
                if (familyName.Contains("duct")) return "Ducts";
                if (familyName.Contains("pipe")) return "Pipes";
                if (familyName.Contains("cable")) return "Cable Trays";
                
                // ✅ CRITICAL FIX: Try to determine category from sleeve placement context
                // Check if this sleeve was placed for a specific MEP category by looking at nearby elements
                var category = DetermineCategoryFromContext(sleeve);
                
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[CATEGORY-CONTEXT] Sleeve {sleeve.Id.IntegerValue}: Determined category from context: '{category}'\n");
                }
                
                return category;
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[CATEGORY-ERROR] Sleeve {sleeve.Id.IntegerValue}: {ex.Message}\n");
                }
                
                // ✅ CRITICAL FIX: Try context-based detection even in error case
                try
                {
                    var category = DetermineCategoryFromContext(sleeve);
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[CATEGORY-ERROR-RECOVERY] Sleeve {sleeve.Id.IntegerValue}: Recovered with context category: '{category}'\n");
                    }
                    return category;
                }
                catch
                {
                    return "Unknown"; // Only use Unknown as absolute last resort
                }
            }
        }
        
        /// <summary>
        /// Determine category from sleeve placement context when MEP element detection fails
        /// </summary>
        private string DetermineCategoryFromContext(FamilyInstance sleeve)
        {
            try
            {
                // Method 1: Check if sleeve has MEP_Category parameter (set during placement)
                var mepCategoryParam = sleeve.LookupParameter("MEP_Category");
                if (mepCategoryParam != null)
                {
                    var categoryValue = mepCategoryParam.AsString();
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[CATEGORY-PARAM-CHECK] Sleeve {sleeve.Id.IntegerValue}: MEP_Category parameter exists, value='{categoryValue}', IsReadOnly={mepCategoryParam.IsReadOnly}\n");
                    }
                    
                    if (!string.IsNullOrEmpty(categoryValue))
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[CATEGORY-PARAM-SUCCESS] Sleeve {sleeve.Id.IntegerValue}: Using MEP_Category parameter: '{categoryValue}'\n");
                        }
                        return categoryValue;
                    }
                }
                else
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[CATEGORY-PARAM-MISSING] Sleeve {sleeve.Id.IntegerValue}: MEP_Category parameter not found\n");
                    }
                }
                
                // Method 2: Check sleeve family name more intelligently
                var familyName = sleeve.Symbol.FamilyName.ToLower();
                if (familyName.Contains("circular") || familyName.Contains("round"))
                {
                    // Circular openings are typically for pipes
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[CATEGORY-CIRCULAR] Sleeve {sleeve.Id.IntegerValue}: Circular family detected, assuming Pipes\n");
                    }
                    return "Pipes";
                }
                
                // Method 3: Check host element type
                var host = sleeve.Host;
                if (host != null)
                {
                    var hostCategory = host.Category?.Name?.ToLower();
                    if (hostCategory != null)
                    {
                        if (hostCategory.Contains("wall"))
                        {
                            // Wall-hosted sleeves are often for pipes
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Info($"[CATEGORY-WALL] Sleeve {sleeve.Id.IntegerValue}: Wall-hosted, assuming Pipes\n");
                            }
                            return "Pipes";
                        }
                        else if (hostCategory.Contains("floor") || hostCategory.Contains("slab"))
                        {
                            // Floor-hosted sleeves are often for ducts
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Info($"[CATEGORY-FLOOR] Sleeve {sleeve.Id.IntegerValue}: Floor-hosted, assuming Ducts\n");
                            }
                            return "Ducts";
                        }
                    }
                }
                
                // Method 4: Check sleeve dimensions (pipes are typically smaller)
                var bbox = sleeve.get_BoundingBox(null);
                if (bbox != null)
                {
                    var width = bbox.Max.X - bbox.Min.X;
                    var height = bbox.Max.Y - bbox.Min.Y;
                    var avgSize = (width + height) / 2;
                    
                    // Small openings (< 0.5m) are typically pipes
                    if (avgSize < 0.5)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[CATEGORY-SIZE] Sleeve {sleeve.Id.IntegerValue}: Small size ({avgSize:F2}m), assuming Pipes\n");
                        }
                        return "Pipes";
                    }
                }
                
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[CATEGORY-CONTEXT-FAILED] Sleeve {sleeve.Id.IntegerValue}: All context methods failed, using Unknown\n");
                }
                return "Unknown";
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[CATEGORY-CONTEXT-ERROR] Sleeve {sleeve.Id.IntegerValue}: Context detection error: {ex.Message}\n");
                }
                return "Unknown";
            }
        }
        
        private string GetHostType(FamilyInstance sleeve)
        {
            var host = sleeve.Host;
            if (host == null) return "Unknown";
            
            var hostType = host.GetType().Name;
            if (hostType.Contains("Wall")) return "Wall";
            if (hostType.Contains("Floor")) return "Floor";
            if (hostType.Contains("Structural")) return "Structural Framing";
            return "Unknown";
        }
        
        private string GetSleeveOrientation(FamilyInstance sleeve)
        {
            // ✅ CORRECT: Get orientation from clash zone data, not recalculate
            try
            {
                // Get the MEP element ID from the sleeve
                var mepElementId = sleeve.LookupParameter("MEP_ElementId");
                if (mepElementId != null && mepElementId.HasValue)
                {
                    long mepElementIdValue = mepElementId.AsInteger();
                    
                    // Find the clash zone for this MEP element
                    var clashZone = FindClashZoneByMepElementId(mepElementIdValue);
                    if (clashZone != null)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[ORIENTATION-DEBUG] Sleeve {sleeve.Id.IntegerValue}: Found clash zone, Orientation='{clashZone.MepElementOrientationDirection}'\n");
                        }
                        return clashZone.MepElementOrientationDirection ?? "";
                    }
                    else
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[ORIENTATION-DEBUG] Sleeve {sleeve.Id.IntegerValue}: No clash zone found for MEP_ElementId {mepElementIdValue}\n");
                        }
                    }
                }
                
                return ""; // Default fallback
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[ORIENTATION-ERROR] Sleeve {sleeve.Id.IntegerValue}: {ex.Message}\n");
                }
                return ""; // Safe default
            }
        }
        
        private void SaveSleeveDataToClusterXml(List<SleeveData> sleeveDataList, string category, string filterName = null)
        {
            try
            {
                string filtersDirectory = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "JSE_MEP_Openings", "Projects", "Default", "Filters");
                
                if (!Directory.Exists(filtersDirectory))
                    Directory.CreateDirectory(filtersDirectory);
                
                // ✅ MASTER DEBUGGER FIX: Create filename that matches clustering service expectations
                // ✅ CRITICAL: Filter name MUST be provided - no hardcoded fallback
                if (string.IsNullOrEmpty(filterName))
                {
                    var errorMsg = "Filter name is required for cluster XML file creation.";
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[REGENERATE-CLUSTER-XML] ERROR: {errorMsg}\n");
                    }
                    throw new InvalidOperationException(errorMsg);
                }
                var fileName = $"{filterName}_{category.Replace(" ", "_").ToLower()}_CLUSTER.xml";
                
                // ✅ CRITICAL: Also create the plural version that clustering service expects
                var pluralCategory = category switch
                {
                    "Pipes" => "Pipes",
                    "Ducts" => "Ducts", 
                    "Cable Trays" => "Cable Trays",
                    _ => category
                };
                var pluralFileName = $"{filterName}_{pluralCategory.Replace(" ", "_").ToLower()}_CLUSTER.xml";
                var filePath = Path.Combine(filtersDirectory, fileName);
                
                // Create XML document
                var xmlDoc = new System.Xml.XmlDocument();
                var root = xmlDoc.CreateElement("SleeveDataList");
                xmlDoc.AppendChild(root);
                
                foreach (var sleeveData in sleeveDataList)
                {
                    var sleeveElement = xmlDoc.CreateElement("SleeveData");
                    
                    AddXmlElement(xmlDoc, sleeveElement, "SleeveInstanceId", sleeveData.SleeveInstanceId.ToString());
                    AddXmlElement(xmlDoc, sleeveElement, "Corner1X", sleeveData.Corner1.X.ToString("F6"));
                    AddXmlElement(xmlDoc, sleeveElement, "Corner1Y", sleeveData.Corner1.Y.ToString("F6"));
                    AddXmlElement(xmlDoc, sleeveElement, "Corner1Z", sleeveData.Corner1.Z.ToString("F6"));
                    AddXmlElement(xmlDoc, sleeveElement, "Corner2X", sleeveData.Corner2.X.ToString("F6"));
                    AddXmlElement(xmlDoc, sleeveElement, "Corner2Y", sleeveData.Corner2.Y.ToString("F6"));
                    AddXmlElement(xmlDoc, sleeveElement, "Corner2Z", sleeveData.Corner2.Z.ToString("F6"));
                    AddXmlElement(xmlDoc, sleeveElement, "Corner3X", sleeveData.Corner3.X.ToString("F6"));
                    AddXmlElement(xmlDoc, sleeveElement, "Corner3Y", sleeveData.Corner3.Y.ToString("F6"));
                    AddXmlElement(xmlDoc, sleeveElement, "Corner3Z", sleeveData.Corner3.Z.ToString("F6"));
                    AddXmlElement(xmlDoc, sleeveElement, "Corner4X", sleeveData.Corner4.X.ToString("F6"));
                    AddXmlElement(xmlDoc, sleeveElement, "Corner4Y", sleeveData.Corner4.Y.ToString("F6"));
                    AddXmlElement(xmlDoc, sleeveElement, "Corner4Z", sleeveData.Corner4.Z.ToString("F6"));
                    AddXmlElement(xmlDoc, sleeveElement, "Width", sleeveData.Width.ToString("F6"));
                    AddXmlElement(xmlDoc, sleeveElement, "Height", sleeveData.Height.ToString("F6"));
                    AddXmlElement(xmlDoc, sleeveElement, "Depth", sleeveData.Depth.ToString("F6"));
                    AddXmlElement(xmlDoc, sleeveElement, "HostType", sleeveData.HostType);
                    AddXmlElement(xmlDoc, sleeveElement, "Orientation", sleeveData.Orientation);
                    AddXmlElement(xmlDoc, sleeveElement, "Category", sleeveData.Category);
                    AddXmlElement(xmlDoc, sleeveElement, "CreatedAt", sleeveData.CreatedAt);
                    
                    root.AppendChild(sleeveElement);
                }
                
                // Save XML file
                xmlDoc.Save(filePath);
                
                // ✅ CRITICAL: Also save the plural version for clustering service compatibility
                var pluralFilePath = Path.Combine(filtersDirectory, pluralFileName);
                xmlDoc.Save(pluralFilePath);
                
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[REGENERATE-CLUSTER-XML] Saved {sleeveDataList.Count} sleeves to {fileName} and {pluralFileName}\n");
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[REGENERATE-CLUSTER-XML] Save error: {ex.Message}\n");
                }
            }
        }
        
        private void AddXmlElement(System.Xml.XmlDocument xmlDoc, System.Xml.XmlElement parent, string name, string value)
        {
            var element = xmlDoc.CreateElement(name);
            element.InnerText = value;
            parent.AppendChild(element);
        }
        
        /// <summary>
        /// Load clash zone cache from XML files for orientation lookup
        /// </summary>
        private void LoadClashZoneCache()
        {
            try
            {
                _clashZoneCache.Clear();
                
                var filtersDirectory = ProjectPathService.GetFiltersDirectory(_doc);
                var allXmlFiles = Directory.GetFiles(filtersDirectory, "*.xml");
                var xmlFiles = allXmlFiles
                    .Where(f => f.Contains("_ducts.xml") || f.Contains("_pipes.xml") || f.Contains("_cable_trays.xml") || 
                                f.Contains("_duct_accessories.xml") || f.Contains("_pipe_accessories.xml") || f.Contains("_cable_tray_accessories.xml"))
                    .ToList();
                
                // ✅ CRITICAL DEBUGGING: Log file discovery
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[CLASH-CACHE-DEBUG] Filters directory: {filtersDirectory}\n");
                    DebugLogger.Info($"[CLASH-CACHE-DEBUG] Found {allXmlFiles.Length} total XML files\n");
                    DebugLogger.Info($"[CLASH-CACHE-DEBUG] Filtered to {xmlFiles.Count} relevant files\n");
                    
                    foreach (var file in xmlFiles)
                    {
                        DebugLogger.Info($"[CLASH-CACHE-DEBUG] Processing file: {Path.GetFileName(file)}\n");
                    }
                }
                
                foreach (var xmlFile in xmlFiles)
                {
                    try
                    {
                        var xmlDoc = new System.Xml.XmlDocument();
                        xmlDoc.Load(xmlFile);
                        
                        var clashNodes = xmlDoc.SelectNodes("//ClashZone");
                        if (clashNodes != null)
                        {
                            foreach (System.Xml.XmlNode node in clashNodes)
                            {
                                // Get MEP element ID
                                if (long.TryParse(node.SelectSingleNode("MepElementId")?.InnerText, out long mepElementIdLong))
                                {
                                    var clashZone = new ClashZone
                                    {
                                        MepElementId = new ElementId((int)mepElementIdLong),
                                        MepElementOrientationDirection = node.SelectSingleNode("MepElementOrientationDirection")?.InnerText ?? ""
                                    };
                                    
                                    _clashZoneCache[mepElementIdLong] = clashZone;
                                }
                            }
                        }
                        
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[DEBUG] Loaded {clashNodes?.Count ?? 0} clash zones from {Path.GetFileName(xmlFile)}\n");
                        }
                    }
                    catch (Exception ex)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[DEBUG] Error loading {xmlFile}: {ex.Message}\n");
                        }
                    }
                }
                
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[CLASH-CACHE] Loaded {_clashZoneCache.Count} clash zones for orientation lookup\n");
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[CLASH-CACHE] Error loading clash zone cache: {ex.Message}\n");
                }
            }
        }
        
        /// <summary>
        /// Find clash zone by MEP element ID
        /// </summary>
        private ClashZone FindClashZoneByMepElementId(long mepElementId)
        {
            return _clashZoneCache.TryGetValue(mepElementId, out var clashZone) ? clashZone : null;
        }
    }
    
    /// <summary>
    /// Helper service to update sleeve coordinates AFTER placement
    /// Fixes timing issue where immediate bounding box is wrong
    /// </summary>
    public class SleeveCoordinateUpdater
    {
        private readonly Document _doc;
        
        public SleeveCoordinateUpdater(Document doc)
        {
            _doc = doc;
        }
        
        public void UpdateSleeveCoordinates(List<ClashZone> clashZones)
        {
            // ✅ Declare placementDebugPath once at method start
            var placementDebugPath = SafeFileLogger.GetLogFilePath("placement_debug.log");
            
            // ✅ DEPLOYMENT: Wrapped in deployment mode check
            if (!DeploymentConfiguration.DeploymentMode)
            {
                DebugLogger.Info($"[UpdateSleeveCoordinates] ⚠️ CALLED - Processing {clashZones.Count} clash zones for coordinate update\n");
            }
            
            // Get all sleeves in the model for position matching
            var allSleeves = new FilteredElementCollector(_doc)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(s => s.Symbol.FamilyName.Contains("Opening"))
                .ToList();
            
            // ✅ DEPLOYMENT: Wrapped in deployment mode check
            if (!DeploymentConfiguration.DeploymentMode)
            {
                DebugLogger.Info($"[UpdateSleeveCoordinates] Found {allSleeves.Count} sleeves in model for position matching\n");
                
                if (allSleeves.Count > 0)
                {
                    var firstFewIds = allSleeves.Take(10).Select(s => s.Id.IntegerValue).ToList();
                    DebugLogger.Info($"[UpdateSleeveCoordinates] First 10 sleeve IDs in Revit: [{string.Join(", ", firstFewIds)}]\n");
                }
            }
            
            // ✅ CRITICAL DEBUG: Log sleeve IDs we're looking for vs what we found - Direct file write
            try
            {
                var clashZoneSleeveIds = clashZones.Where(cz => cz.SleeveInstanceId > 0).Select(cz => cz.SleeveInstanceId).ToList();
                var revitSleeveIds = allSleeves.Select(s => s.Id.IntegerValue).ToList();
                var matchingIds = clashZoneSleeveIds.Intersect(revitSleeveIds).ToList();
                System.IO.File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [UPDATE_COORD_DEBUG] Looking for {clashZoneSleeveIds.Count} sleeve IDs: [{string.Join(", ", clashZoneSleeveIds.Take(10))}...]\n");
                System.IO.File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [UPDATE_COORD_DEBUG] Found {revitSleeveIds.Count} sleeves in Revit: [{string.Join(", ", revitSleeveIds.Take(10))}...]\n");
                System.IO.File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [UPDATE_COORD_DEBUG] Matching IDs: {matchingIds.Count} out of {clashZoneSleeveIds.Count}\n");
            }
            catch { }
            
            foreach (var clashZone in clashZones)
            {
                try
                {
                    FamilyInstance matchedSleeve = null;
                    
                    // ✅ DEBUG: Log if this clashZone has cluster information
                    if (clashZone.ClusterSleeveInstanceId > 0 && !DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[CLUSTER-CHECK] ClashZone has ClusterSleeveInstanceId={clashZone.ClusterSleeveInstanceId}, SleeveInstanceId={clashZone.SleeveInstanceId}\n");
                    }
                    
                    // ✅ CRITICAL FIX: Try direct ID match first (for individual sleeves)
                    // This MUST succeed if SleeveInstanceId > 0, as the sleeve was just placed
                    if (clashZone.SleeveInstanceId > 0)
                    {
                        try
                        {
                            var sleeveElement = _doc.GetElement(new ElementId(clashZone.SleeveInstanceId));
                            if (sleeveElement is FamilyInstance sleeve && sleeve.Symbol.FamilyName.Contains("Opening"))
                            {
                                matchedSleeve = sleeve;
                                try { System.IO.File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [DIRECT-ID-MATCH] Found individual sleeve {clashZone.SleeveInstanceId} by ID\n"); } catch { }
                            }
                            else if (sleeveElement == null)
                            {
                                try { System.IO.File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [DIRECT-ID-MATCH] Sleeve {clashZone.SleeveInstanceId} not found in document - may have been deleted\n"); } catch { }
                            }
                            else
                            {
                                try { System.IO.File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [DIRECT-ID-MATCH] Element {clashZone.SleeveInstanceId} is not a FamilyInstance with Opening name (Type={sleeveElement?.GetType()?.Name})\n"); } catch { }
                            }
                        }
                        catch (Exception ex)
                        {
                            try { System.IO.File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [DIRECT-ID-MATCH] Error getting sleeve {clashZone.SleeveInstanceId}: {ex.Message}\n"); } catch { }
                        }
                    }
                    
                    // ✅ NEW: Handle cluster sleeves
                    if (matchedSleeve == null && clashZone.ClusterSleeveInstanceId > 0)
                    {
                        // ✅ DEPLOYMENT: Wrapped in deployment mode check
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[CLUSTER-LOOKUP] Looking for cluster sleeve ID {clashZone.ClusterSleeveInstanceId} in Revit\n");
                        }
                        
                        var clusterSleeveElement = _doc.GetElement(new ElementId(clashZone.ClusterSleeveInstanceId));
                        if (clusterSleeveElement == null)
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Info($"[CLUSTER-LOOKUP] ERROR: Cluster sleeve ID {clashZone.ClusterSleeveInstanceId} NOT found in Revit!\n");
                            }
                        }
                        else if (clusterSleeveElement is FamilyInstance clusterSleeve && clusterSleeve.Symbol.FamilyName.Contains("Opening"))
                        {
                            // ✅ DEPLOYMENT: Wrapped in deployment mode check
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Info($"[CLUSTER-LOOKUP] ✓ Found cluster sleeve {clashZone.ClusterSleeveInstanceId} in Revit (Family={clusterSleeve.Symbol.FamilyName})\n");
                            }
                            
                            // Get bounding box for cluster sleeve
                            var bbox = clusterSleeve.get_BoundingBox(null);
                            if (bbox != null)
                            {
                                // ✅ NEW: Update cluster sleeve bounding box coordinates
                                clashZone.ClusterSleeveBoundingBoxMinX = bbox.Min.X;
                                clashZone.ClusterSleeveBoundingBoxMinY = bbox.Min.Y;
                                clashZone.ClusterSleeveBoundingBoxMinZ = bbox.Min.Z;
                                clashZone.ClusterSleeveBoundingBoxMaxX = bbox.Max.X;
                                clashZone.ClusterSleeveBoundingBoxMaxY = bbox.Max.Y;
                                clashZone.ClusterSleeveBoundingBoxMaxZ = bbox.Max.Z;
                                
                                // ✅ DEPLOYMENT: Wrapped in deployment mode check
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    DebugLogger.Info($"[CLUSTER-SLEEVE] Updated cluster sleeve {clashZone.ClusterSleeveInstanceId} bbox: Min=({bbox.Min.X:F6}, {bbox.Min.Y:F6}), Max=({bbox.Max.X:F6}, {bbox.Max.Y:F6})\n");
                                }
                            }
                        }
                        else if (clusterSleeveElement != null)
                        {
                            // ✅ DEPLOYMENT: Wrapped in deployment mode check
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Info($"[CLUSTER-LOOKUP] Element {clashZone.ClusterSleeveInstanceId} is not a FamilyInstance with Opening name\n");
                            }
                        }
                    }
                    
                    // ✅ CRITICAL FIX: If no direct match, try position matching
                    if (matchedSleeve == null && clashZone.SleevePlacementPointX != 0 && clashZone.SleevePlacementPointY != 0)
                    {
                        foreach (var sleeve in allSleeves)
                        {
                            var bbox = sleeve.get_BoundingBox(null);
                            if (bbox != null)
                            {
                                // Check if sleeve position matches clash zone placement point (within 1mm tolerance)
                                var distance = Math.Sqrt(
                                    Math.Pow(bbox.Min.X - clashZone.SleevePlacementPointX, 2) +
                                    Math.Pow(bbox.Min.Y - clashZone.SleevePlacementPointY, 2) +
                                    Math.Pow(bbox.Min.Z - clashZone.SleevePlacementPointZ, 2));
                                
                                if (distance < 0.00328) // 1mm tolerance in feet
                                {
                                    matchedSleeve = sleeve;
                                    clashZone.SleeveInstanceId = sleeve.Id.IntegerValue; // ✅ FIX: Update the instance ID
                                    // ✅ DEPLOYMENT: Wrapped in deployment mode check
                                    if (!DeploymentConfiguration.DeploymentMode)
                                    {
                                        DebugLogger.Info($"[POSITION-MATCH] Found sleeve {sleeve.Id.IntegerValue} by position (distance: {distance:F6}ft)\n");
                                    }
                                    break;
                                }
                            }
                        }
                    }
                    
                    // Update coordinates if sleeve found
                    if (matchedSleeve != null)
                    {
                        var bbox = matchedSleeve.get_BoundingBox(null);
                        if (bbox != null)
                        {
                            // ✅ CRITICAL LOGGING: Log BEFORE SetSleeveBoundingBox - Direct file write to placement_debug.log
                            try 
                            { 
                                System.IO.File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [BOUNDING_BOX_BEFORE_SET] ClashZone {clashZone.Id}, Sleeve {clashZone.SleeveInstanceId}: " +
                                    $"BEFORE - ClashZone has: MinX={clashZone.SleeveBoundingBoxMinX:F6}, MinY={clashZone.SleeveBoundingBoxMinY:F6}, MinZ={clashZone.SleeveBoundingBoxMinZ:F6}, " +
                                    $"MaxX={clashZone.SleeveBoundingBoxMaxX:F6}, MaxY={clashZone.SleeveBoundingBoxMaxY:F6}, MaxZ={clashZone.SleeveBoundingBoxMaxZ:F6}\n");
                                
                                System.IO.File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [BOUNDING_BOX_BEFORE_SET] Revit bbox from sleeve {matchedSleeve.Id.IntegerValue}: " +
                                    $"Min=({bbox.Min.X:F6}, {bbox.Min.Y:F6}, {bbox.Min.Z:F6}), Max=({bbox.Max.X:F6}, {bbox.Max.Y:F6}, {bbox.Max.Z:F6})\n");
                            } 
                            catch { }
                            
                            // ✅ CORRECT: Update bounding box and instance ID
                            clashZone.SetSleeveBoundingBox(bbox);
                            
                            // ✅ CRITICAL LOGGING: Log AFTER SetSleeveBoundingBox - Direct file write to placement_debug.log
                            try 
                            { 
                                System.IO.File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [BOUNDING_BOX_AFTER_SET] ClashZone {clashZone.Id}, Sleeve {clashZone.SleeveInstanceId}: " +
                                    $"AFTER - ClashZone now has: MinX={clashZone.SleeveBoundingBoxMinX:F6}, MinY={clashZone.SleeveBoundingBoxMinY:F6}, MinZ={clashZone.SleeveBoundingBoxMinZ:F6}, " +
                                    $"MaxX={clashZone.SleeveBoundingBoxMaxX:F6}, MaxY={clashZone.SleeveBoundingBoxMaxY:F6}, MaxZ={clashZone.SleeveBoundingBoxMaxZ:F6}\n");
                                
                                System.IO.File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [UPDATED] Sleeve {clashZone.SleeveInstanceId}: Min=({bbox.Min.X:F6}, {bbox.Min.Y:F6}, {bbox.Min.Z:F6}), Max=({bbox.Max.X:F6}, {bbox.Max.Y:F6}, {bbox.Max.Z:F6})\n");
                            } 
                            catch { }
                        }
                        else
                        {
                            // ✅ DEPLOYMENT: Wrapped in deployment mode check
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Info($"[NO-BBOX] Sleeve {clashZone.SleeveInstanceId} has no bounding box\n");
                            }
                        }
                    }
                    else
                    {
                        // ✅ DEPLOYMENT: Wrapped in deployment mode check
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[NO-MATCH] ClashZone {clashZone.Id} - no matching sleeve found (SleeveInstanceId={clashZone.SleeveInstanceId}, PlacePoint=({clashZone.SleevePlacementPointX:F3}, {clashZone.SleevePlacementPointY:F3}, {clashZone.SleevePlacementPointZ:F3}))\n");
                        }
                    }
                }
                catch (Exception ex)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[ERROR] Clash zone error: {ex.Message}\n");
                    }
                }
            }
        }
    }
    
    /// <summary>
    /// Data structure for sleeve information
    /// </summary>
    public class SleeveData
    {
        public int SleeveInstanceId { get; set; }
        public XYZ Corner1 { get; set; }
        public XYZ Corner2 { get; set; }
        public XYZ Corner3 { get; set; }
        public XYZ Corner4 { get; set; }
        public double Width { get; set; }
        public double Height { get; set; }
        public double Depth { get; set; }
        public string HostType { get; set; }
        public string Orientation { get; set; }
        public string Category { get; set; }
        public string CreatedAt { get; set; }
    }
}

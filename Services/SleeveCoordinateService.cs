using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories;
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
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        File.AppendAllText(sleeveCoordinatesLogPath, "No sleeves found in model!\n");
                    }
                    return;
                }
                
                // Log all sleeve coordinates
                string logPath = SafeFileLogger.GetLogFilePath("all_sleeve_coordinates.log");
                System.IO.File.WriteAllText(logPath, $"ALL SLEEVE COORDINATES - {DateTime.Now}\n");
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    File.AppendAllText(logPath, $"Found {sleeves.Count} sleeves\n\n");
                }
                
                foreach (var sleeve in sleeves)
                {
                    var bbox = sleeve.get_BoundingBox(null);
                    if (bbox != null)
                    {
                        var min = bbox.Min;
                        var max = bbox.Max;
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            File.AppendAllText(logPath, 
                            $"SLEEVE {sleeve.Id.IntegerValue}:\n");
                        }
                        System.IO.File.AppendAllText(logPath, 
                            $"  Min: ({min.X:F6}, {min.Y:F6}, {min.Z:F6})\n");
                        System.IO.File.AppendAllText(logPath, 
                            $"  Max: ({max.X:F6}, {max.Y:F6}, {max.Z:F6})\n");
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            File.AppendAllText(logPath, 
                            $"  Host: {sleeve.Host?.Id?.IntegerValue}\n");
                        }
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            File.AppendAllText(logPath, 
                            $"  Family: {sleeve.Symbol.FamilyName}\n\n");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
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
                try {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        System.IO.File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [UpdateSleeveCoordinatesInXml] CALLED with xmlFilePath: {xmlFilePath ?? "NULL"}\n");
                    }
                } catch { }
                
                // Get all sleeves
                var sleeves = new FilteredElementCollector(_doc)
                    .OfClass(typeof(FamilyInstance))
                    .Cast<FamilyInstance>()
                    .Where(s => s.Symbol.FamilyName.Contains("Opening"))
                    .ToList();
                
                try {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        System.IO.File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [UpdateSleeveCoordinatesInXml] Found {sleeves.Count} sleeves in Revit model\n");
                    }
                } catch { }
                
                var coordinateUpdater = new SleeveCoordinateUpdater(_doc);
                
                // ✅ PHASE 2: DATABASE-FIRST - Load clash zones from database first (primary source of truth)
                // Fallback to XML only if database has no data (backward compatibility)
                var clashZones = new List<ClashZone>();
                
                // Try to determine category from XML file path or load from database
                string category = null;
                if (!string.IsNullOrEmpty(xmlFilePath))
                {
                    // Extract category from filename (e.g., "Electrical_cable_trays.xml" -> "Cable Trays")
                    var fileName = Path.GetFileNameWithoutExtension(xmlFilePath);
                    if (fileName.Contains("_"))
                    {
                        var parts = fileName.Split('_');
                        if (parts.Length >= 2)
                        {
                            // Try to map to category name - take last part and capitalize properly
                            var categoryPart = parts[parts.Length - 1];
                            // Convert "cable_trays" -> "Cable Trays"
                            var words = categoryPart.Split('_');
                            category = string.Join(" ", words.Select(w => 
                                w.Length > 0 ? char.ToUpper(w[0]) + (w.Length > 1 ? w.Substring(1).ToLower() : "") : w));
                        }
                    }
                }
                
                // ✅ DATABASE-FIRST: Load from database if category is known
                if (!string.IsNullOrEmpty(category))
                {
                    try
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [UpdateSleeveCoordinatesInXml] 🔍 Attempting to load clash zones from database for category '{category}' (extracted from filename: {Path.GetFileName(xmlFilePath ?? "NULL")})\n");
                        }
                        
                        using (var dbContext = new Data.SleeveDbContext(_doc))
                        {
                            var repository = new Data.Repositories.ClashZoneRepository(dbContext);
                            var dbZones = repository.GetClashZonesByCategory(category);
                            
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [UpdateSleeveCoordinatesInXml] 🔍 Database query returned {dbZones?.Count ?? 0} zones for category '{category}'\n");
                            }
                            
                            if (dbZones != null && dbZones.Count > 0)
                            {
                                // ✅ CRITICAL FIX: Load ALL zones (not just those with SleeveInstanceId > 0)
                                // This includes zones with ClusterSleeveInstanceId > 0 (after clustering)
                                // SleeveInstanceId may not be saved to DB yet after placement, so we need to match by position
                                clashZones = dbZones.Where(z => z != null).ToList();
                                
                                foreach (var cz in clashZones)
                                {
                                    // ✅ CRITICAL: Reconstruct SleevePlacementPoint from database properties
                                    cz.EnsureSleevePlacementPointReconstructed();
                                }
                                
                                // ✅ DEBUG: Log sample zone data to diagnose matching issues
                                var zonesWithPlacementPoint = clashZones.Count(z => z != null && (z.SleevePlacementPointX != 0 || z.SleevePlacementPointY != 0 || z.SleevePlacementPointZ != 0));
                                var zonesWithSleeveId = clashZones.Count(z => z != null && z.SleeveInstanceId > 0);
                                var zonesWithClusterId = clashZones.Count(z => z != null && z.ClusterSleeveInstanceId > 0);
                                
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    DebugLogger.Info($"[SleeveCoordinateService] ✅ DATABASE-FIRST: Loaded {clashZones.Count} clash zones from database for category '{category}' (all zones, will match by position)");
                                    File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [UpdateSleeveCoordinatesInXml] ✅ DATABASE-FIRST: Loaded {clashZones.Count} clash zones from database for category '{category}' (all zones)\n");
                                    File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [UpdateSleeveCoordinatesInXml] DEBUG: Zones with placement point: {zonesWithPlacementPoint}/{clashZones.Count}, Zones with SleeveInstanceId>0: {zonesWithSleeveId}/{clashZones.Count}, Zones with ClusterSleeveInstanceId>0: {zonesWithClusterId}/{clashZones.Count}\n");
                                    
                                    // Log sample placement points
                                    var sampleZones = clashZones.Where(z => z != null && (z.SleevePlacementPointX != 0 || z.SleevePlacementPointY != 0)).Take(3).ToList();
                                    foreach (var sample in sampleZones)
                                    {
                                        File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [UpdateSleeveCoordinatesInXml] SAMPLE: Zone {sample.Id}, SPP=({sample.SleevePlacementPointX:F3}, {sample.SleevePlacementPointY:F3}, {sample.SleevePlacementPointZ:F3}), SleeveId={sample.SleeveInstanceId}, ClusterId={sample.ClusterSleeveInstanceId}\n");
                                    }
                                }
                            }
                            else
                            {
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [UpdateSleeveCoordinatesInXml] ⚠️ Database returned 0 zones for category '{category}' - will fall back to XML\n");
                                }
                            }
                        }
                    }
                    catch (Exception dbEx)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Warning($"[SleeveCoordinateService] ⚠️ Database load failed, falling back to XML: {dbEx.Message}");
                            File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [UpdateSleeveCoordinatesInXml] ⚠️ Database load failed: {dbEx.Message} (falling back to XML)\n");
                            File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [UpdateSleeveCoordinatesInXml] ⚠️ Stack trace: {dbEx.StackTrace}\n");
                        }
                    }
                }
                else
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [UpdateSleeveCoordinatesInXml] ⚠️ Could not extract category from filename '{xmlFilePath ?? "NULL"}' - will fall back to XML\n");
                    }
                }
                
                // ✅ FALLBACK: Load from XML if database has no data
                if (clashZones.Count == 0)
                {
                    clashZones = LoadClashZonesFromXml(xmlFilePath);
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [UpdateSleeveCoordinatesInXml] ⚠️ Loaded {clashZones.Count} clash zones from XML (fallback)\n");
                    }
                }
                
                // ✅ CRITICAL DEBUG: Log how many have SleeveInstanceId (calculate outside condition for use later)
                var withSleeveId = clashZones.Count(cz => cz.SleeveInstanceId > 0);
                
                try 
                { 
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [UpdateSleeveCoordinatesInXml] Loaded {clashZones.Count} clash zones from XML\n");
                    }
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [UpdateSleeveCoordinatesInXml] {withSleeveId} out of {clashZones.Count} clash zones have SleeveInstanceId > 0\n");
                    }
                    
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
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [UpdateSleeveCoordinatesInXml] After UpdateSleeveCoordinates: {withBbox} out of {withSleeveId} clash zones with SleeveInstanceId now have bounding boxes\n");
                    }
                }
                catch { }
                
                // ✅ PHASE 2: DATABASE-FIRST - Save SleeveInstanceId and bounding boxes to database (primary source of truth)
                // ⚠️ PROTECTED CODE: DO NOT MODIFY THIS SECTION WITHOUT UNDERSTANDING THE IMPACT
                // This code is critical for clustering to work correctly. It saves:
                // 1. SleeveInstanceId - Required for clustering to find sleeves in the database
                // 2. Bounding box coordinates - Required for clustering to calculate cluster bounding boxes
                // 3. ClusterSleeveInstanceId and cluster bounding boxes - Required after clustering
                // If this code is broken, clustering will fail because it won't find sleeves with valid bounding boxes.
                try
                {
                    using (var dbContext = new Data.SleeveDbContext(_doc))
                    {
                        var repository = new Data.Repositories.ClashZoneRepository(dbContext);
                        
                        int dbUpdatedCount = 0;
                        int dbSleeveIdUpdatedCount = 0;
                        int dbClusterUpdatedCount = 0;
                        foreach (var cz in clashZones)
                        {
                            if (cz == null) continue;
                            
                            // ✅ INDIVIDUAL SLEEVE: Save individual sleeve data
                            if (cz.SleeveInstanceId > 0)
                            {
                                // ✅ CRITICAL: Save SleeveInstanceId first (needed for clustering to find sleeves)
                                // This must be saved immediately after placement so clustering can find the sleeves
                                repository.UpdateSleeveInstanceId(cz.Id, cz.SleeveInstanceId);
                                dbSleeveIdUpdatedCount++;
                                
                                // ✅ Save bounding boxes if they're not zero
                                // Bounding boxes are required for clustering to calculate cluster bounding boxes
                                // Without these, clustering will find 0 sleeves with valid bounding boxes
                                if (!(cz.SleeveBoundingBoxMinX == 0.0 && cz.SleeveBoundingBoxMinY == 0.0 && cz.SleeveBoundingBoxMinZ == 0.0 &&
                                      cz.SleeveBoundingBoxMaxX == 0.0 && cz.SleeveBoundingBoxMaxY == 0.0 && cz.SleeveBoundingBoxMaxZ == 0.0))
                                {
                                    repository.UpdateSleeveBoundingBoxes(
                                        cz.Id,
                                        cz.SleeveBoundingBoxMinX, cz.SleeveBoundingBoxMinY, cz.SleeveBoundingBoxMinZ,
                                        cz.SleeveBoundingBoxMaxX, cz.SleeveBoundingBoxMaxY, cz.SleeveBoundingBoxMaxZ);
                                    dbUpdatedCount++;
                                }
                            }
                            
                            // ✅ CLUSTER SLEEVE: Save cluster sleeve data (after clustering)
                            if (cz.ClusterSleeveInstanceId > 0 && 
                                !(cz.ClusterSleeveBoundingBoxMinX == 0.0 && cz.ClusterSleeveBoundingBoxMinY == 0.0 && cz.ClusterSleeveBoundingBoxMinZ == 0.0 &&
                                  cz.ClusterSleeveBoundingBoxMaxX == 0.0 && cz.ClusterSleeveBoundingBoxMaxY == 0.0 && cz.ClusterSleeveBoundingBoxMaxZ == 0.0))
                            {
                                // Get ClashZoneId (int) from GUID
                                using (var cmd = dbContext.Connection.CreateCommand())
                                {
                                    cmd.CommandText = @"
                                        SELECT ClashZoneId FROM ClashZones
                                        WHERE UPPER(ClashZoneGuid) = UPPER(@ClashZoneGuid)
                                          AND ClashZoneGuid != '' AND ClashZoneGuid IS NOT NULL
                                        LIMIT 1";
                                    cmd.Parameters.AddWithValue("@ClashZoneGuid", cz.Id.ToString().ToUpperInvariant());
                                    var clashZoneIdResult = cmd.ExecuteScalar();
                                    
                                    if (clashZoneIdResult != null)
                                    {
                                        int clashZoneId = Convert.ToInt32(clashZoneIdResult);
                                        
                                        // ✅ Calculate cluster sleeve placement point (midpoint of cluster bbox)
                                        double placementX = (cz.ClusterSleeveBoundingBoxMinX + cz.ClusterSleeveBoundingBoxMaxX) / 2.0;
                                        double placementY = (cz.ClusterSleeveBoundingBoxMinY + cz.ClusterSleeveBoundingBoxMaxY) / 2.0;
                                        double placementZ = (cz.ClusterSleeveBoundingBoxMinZ + cz.ClusterSleeveBoundingBoxMaxZ) / 2.0;
                                        
                                        // ✅ Get rotated bounding box if available (from GetClusterBoundingBoxWithRotatedCoordinates)
                                        double? rotatedMinX = cz.RotatedBoundingBoxMinX;
                                        double? rotatedMinY = cz.RotatedBoundingBoxMinY;
                                        double? rotatedMinZ = cz.RotatedBoundingBoxMinZ;
                                        double? rotatedMaxX = cz.RotatedBoundingBoxMaxX;
                                        double? rotatedMaxY = cz.RotatedBoundingBoxMaxY;
                                        double? rotatedMaxZ = cz.RotatedBoundingBoxMaxZ;
                                        
                                        // ✅ Get flags
                                        bool? isClustered = cz.MarkedForClusteringSleeveProcess;
                                        bool? markedForCluster = cz.MarkedForClusteringSleeveProcess;
                                        
                                        repository.UpdateClusterPlacement(
                                            clashZoneId,
                                            cz.ClusterSleeveInstanceId,
                                            cz.ClusterSleeveBoundingBoxMinX, cz.ClusterSleeveBoundingBoxMinY, cz.ClusterSleeveBoundingBoxMinZ,
                                            cz.ClusterSleeveBoundingBoxMaxX, cz.ClusterSleeveBoundingBoxMaxY, cz.ClusterSleeveBoundingBoxMaxZ,
                                            placementX, placementY, placementZ,
                                            rotatedMinX, rotatedMinY, rotatedMinZ,
                                            rotatedMaxX, rotatedMaxY, rotatedMaxZ,
                                            isClustered, markedForCluster);
                                        dbClusterUpdatedCount++;
                                        
                                        if (!DeploymentConfiguration.DeploymentMode)
                                        {
                                            File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [UpdateSleeveCoordinatesInXml] ✅ Saved cluster data for ClashZone {cz.Id}: ClusterId={cz.ClusterSleeveInstanceId}, Placement=({placementX:F6}, {placementY:F6}, {placementZ:F6}), RotatedBbox={rotatedMinX?.ToString("F6") ?? "NULL"}\n");
                                        }
                                    }
                                    else if (!DeploymentConfiguration.DeploymentMode)
                                    {
                                        File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [UpdateSleeveCoordinatesInXml] ⚠️ Could not find ClashZoneId for GUID {cz.Id} to update cluster bounding box\n");
                                    }
                                }
                            }
                        }
                        
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[SleeveCoordinateService] ✅ DATABASE: Updated SleeveInstanceId for {dbSleeveIdUpdatedCount} clash zones, bounding boxes for {dbUpdatedCount} clash zones, cluster bounding boxes for {dbClusterUpdatedCount} clash zones in database");
                            File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [UpdateSleeveCoordinatesInXml] ✅ DATABASE: Updated SleeveInstanceId for {dbSleeveIdUpdatedCount} zones, bounding boxes for {dbUpdatedCount} zones, cluster bounding boxes for {dbClusterUpdatedCount} zones\n");
                        }
                    }
                }
                catch (Exception dbEx)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Warning($"[SleeveCoordinateService] ⚠️ Failed to save SleeveInstanceId/bounding boxes to database: {dbEx.Message}");
                        File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [UpdateSleeveCoordinatesInXml] ⚠️ Database save failed: {dbEx.Message}\n");
                    }
                }
                
                // ✅ FALLBACK: Save updated XML (only if XML creation is enabled)
                if (!DeploymentConfiguration.DisableXmlCreation)
                {
                    SaveClashZonesToXml(clashZones, xmlFilePath);
                }
                else
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[SleeveCoordinateService] ⚠️ XML creation disabled - skipping SaveClashZonesToXml (database only mode)");
                        File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [UpdateSleeveCoordinatesInXml] ⚠️ XML creation disabled - skipping XML save (database only mode)\n");
                    }
                }
                
                // ✅ Direct file write
                try {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        System.IO.File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [UpdateSleeveCoordinatesInXml] COMPLETED - Updated coordinates for {sleeves.Count} sleeves\n");
                    }
                } catch { }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Error($"[UpdateSleeveCoordinatesInXml] EXCEPTION: {ex.Message}\n");
                                        if (!DeploymentConfiguration.DeploymentMode)
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
                
                // ✅ HIERARCHICAL STRUCTURE ONLY: Use XmlSerializer to load from Filters → FileCombo → ClashZones
                var serializer = new System.Xml.Serialization.XmlSerializer(typeof(OpeningFilter));
                using (var reader = new StreamReader(xmlFilePath))
                {
                    var filter = (OpeningFilter)serializer.Deserialize(reader);
                    
                    // ✅ LOAD FROM HIERARCHICAL STRUCTURE: Filters → FileCombo → ClashZones
                    if (filter?.ClashZoneStorage?.Filters != null)
                    {
                        foreach (var filterGroup in filter.ClashZoneStorage.Filters)
                        {
                            if (filterGroup?.FileCombos != null)
                            {
                                foreach (var fileCombo in filterGroup.FileCombos)
                                {
                                    if (fileCombo?.ClashZones != null)
                                    {
                                        foreach (var cz in fileCombo.ClashZones)
                                        {
                                            // ✅ CRITICAL: Reconstruct SleevePlacementPoint from XML-serializable properties
                                            cz.EnsureSleevePlacementPointReconstructed();
                                            clashZones.Add(cz);
                                        }
                                    }
                                }
                            }
                        }
                    }
                    
                    // ✅ FALLBACK: Also check flat structure for backward compatibility (but shouldn't be used)
                    if (clashZones.Count == 0 && filter?.ClashZoneStorage?.AllZones != null)
                    {
                        foreach (var cz in filter.ClashZoneStorage.AllZones)
                        {
                            cz.EnsureSleevePlacementPointReconstructed();
                            clashZones.Add(cz);
                        }
                    }
                }
                
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[LOAD-XML] Loaded {clashZones.Count} clash zones from hierarchical structure in {Path.GetFileName(xmlFilePath)}\n");
                    try
                    {
                        var logPath = SafeFileLogger.GetLogFilePath("placement_debug.log");
                        foreach (var zone in clashZones.Take(5))
                        {
                            System.IO.File.AppendAllText(logPath,
                                $"[{DateTime.Now:HH:mm:ss}] [COORD-LOAD-DETAIL] Zone {zone.Id} → SleeveId={zone.SleeveInstanceId}\n");
                        }
                    }
                    catch { }
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[LOAD-XML] Error loading {xmlFilePath}: {ex.Message}\n");
                }
            }
            
            return clashZones;
        }
        
        /// <summary>
        /// ✅ CORRECT APPROACH: Use ClashZonePersistenceService to save updated clash zones with bounding boxes
        /// This ensures the tree structure is maintained correctly using the dedicated service
        /// </summary>
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
                
                // ✅ STEP 1: Load Filter XML to get OpeningFilter object
                var serializer = new System.Xml.Serialization.XmlSerializer(typeof(OpeningFilter));
                OpeningFilter filter;
                
                using (var reader = new StreamReader(xmlFilePath))
                {
                    filter = (OpeningFilter)serializer.Deserialize(reader);
                }
                
                if (filter?.ClashZoneStorage == null)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Warning($"[SAVE-XML] Filter or ClashZoneStorage is null in {Path.GetFileName(xmlFilePath)}\n");
                    }
                    return;
                }
                
                // ✅ STEP 2: Extract base filter name from filter (e.g., "Plumbing" from "Plumbing_pipes.xml")
                var baseFilterName = FilterNameHelper.NormalizeBaseName(null, filter?.Name);
                if (string.IsNullOrEmpty(baseFilterName))
                {
                    // Try to extract from filename
                    var fileName = Path.GetFileNameWithoutExtension(xmlFilePath);
                    var parts = fileName.Split('_');
                    baseFilterName = parts.Length > 0 ? parts[0] : "Unknown";
                }
                
                // ✅ STEP 3: Extract ALL clash zones from tree structure to get complete list
                var allClashZones = new List<ClashZone>();
                if (filter.ClashZoneStorage.Filters != null)
                {
                    foreach (var filterGroup in filter.ClashZoneStorage.Filters)
                    {
                        if (filterGroup?.FileCombos != null)
                        {
                            foreach (var fileCombo in filterGroup.FileCombos)
                            {
                                if (fileCombo?.ClashZones != null)
                                {
                                    foreach (var cz in fileCombo.ClashZones)
                                    {
                                        // ✅ CRITICAL: Update bounding boxes from updated clash zones
                                        var updatedZone = clashZones.FirstOrDefault(c => c.Id == cz.Id);
                                        if (updatedZone != null)
                                        {
                                            // Update individual sleeve bounding boxes only
                                            // ✅ CONSOLIDATION: Cluster bounding boxes are now handled by UniversalClusterService
                                            // SleeveCoordinateService only handles individual sleeve bounding boxes
                                            cz.SleeveBoundingBoxMinX = updatedZone.SleeveBoundingBoxMinX;
                                            cz.SleeveBoundingBoxMinY = updatedZone.SleeveBoundingBoxMinY;
                                            cz.SleeveBoundingBoxMinZ = updatedZone.SleeveBoundingBoxMinZ;
                                            cz.SleeveBoundingBoxMaxX = updatedZone.SleeveBoundingBoxMaxX;
                                            cz.SleeveBoundingBoxMaxY = updatedZone.SleeveBoundingBoxMaxY;
                                            cz.SleeveBoundingBoxMaxZ = updatedZone.SleeveBoundingBoxMaxZ;

                                            if (!DeploymentConfiguration.DeploymentMode)
                                            {
                                                try
                                                {
                                                    var logPath = SafeFileLogger.GetLogFilePath("placement_debug.log");
                                                    System.IO.File.AppendAllText(logPath,
                                                        $"[{DateTime.Now:HH:mm:ss}] [COORD-MERGE-DETAIL] Zone {cz.Id} → BBoxMin=({updatedZone.SleeveBoundingBoxMinX:F3},{updatedZone.SleeveBoundingBoxMinY:F3},{updatedZone.SleeveBoundingBoxMinZ:F3}), BBoxMax=({updatedZone.SleeveBoundingBoxMaxX:F3},{updatedZone.SleeveBoundingBoxMaxY:F3},{updatedZone.SleeveBoundingBoxMaxZ:F3})\n");
                                                }
                                                catch { }
                                            }
                                             
                                            // ✅ REMOVED: Cluster bounding box update - now handled by UniversalClusterService immediately after placement
                                        }
                                        
                                        allClashZones.Add(cz);
                                    }
                                }
                            }
                        }
                    }
                }
                
                // ✅ STEP 4: Use ClashZonePersistenceService to save updated clash zones
                // This ensures tree structure is maintained correctly
                var guidManager = new GuidManager(_doc);
                var persistenceService = new ClashZonePersistenceService(_doc, guidManager);
                persistenceService.SaveClashZones(allClashZones, baseFilterName, filter, allowStructuralUpdates: false);
                
                // ✅ STEP 5: Save Filter XML file using FilterManagementService
                var filterManagementService = new FilterManagementService(_doc, null, null);
                filterManagementService.SaveFilterToXmlFile(filter, xmlFilePath);
                
                // ✅ LOGGING: Confirm save
                var placementDebugPath = SafeFileLogger.GetLogFilePath("placement_debug.log");
                try 
                { 
                    var updatedCount = clashZones.Count(cz => cz.SleeveInstanceId > 0 && 
                        !(cz.SleeveBoundingBoxMinX == 0.0 && cz.SleeveBoundingBoxMinY == 0.0 && cz.SleeveBoundingBoxMinZ == 0.0 &&
                          cz.SleeveBoundingBoxMaxX == 0.0 && cz.SleeveBoundingBoxMaxY == 0.0 && cz.SleeveBoundingBoxMaxZ == 0.0));
                    System.IO.File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [XML_FILE_SAVED] Successfully saved {updatedCount} bounding box updates using ClashZonePersistenceService in {Path.GetFileName(xmlFilePath)}\n"); 
                } 
                catch { }
                
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[SAVE-XML] ✅ Updated clash zones with bounding boxes using ClashZonePersistenceService: {Path.GetFileName(xmlFilePath)}\n");
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Error($"[SAVE-XML] Error saving using ClashZonePersistenceService: {ex.Message}\n");
                    DebugLogger.Error($"[SAVE-XML] Stack trace: {ex.StackTrace}\n");
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
                                        if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[REGENERATE-CLUSTER-XML] Starting regeneration after sleeve placement\n");
                }
                
                // ✅ CRITICAL: Load clash zone cache first
                LoadClashZoneCache();
                
                // ✅ CRITICAL: Wait for Revit to update the model
                System.Threading.Thread.Sleep(500); // Reduced to 0.5 seconds since timing is working
                
                if (!DeploymentConfiguration.DeploymentMode)
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
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
                                        if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[REGENERATE-CLUSTER-XML] Found {allSleeves.Count} sleeves in Revit model\n");
                                        if (!DeploymentConfiguration.DeploymentMode)
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
                                                if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[COORDINATE-SERVICE] {DateTime.Now:HH:mm:ss.fff} - Processing {sleeves.Count} sleeves for category: {category}\n");
                                                if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[REGENERATE-CLUSTER-XML] Processing {sleeves.Count} sleeves for category: {category}\n");
                    }
                    
                    // Create SleeveDataList with actual Revit coordinates
                    var sleeveDataList = new List<SleeveData>();
                    
                    foreach (var sleeve in sleeves)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            // ✅ CRITICAL LOGGING: Log each sleeve found after wait time
                                                        if (!DeploymentConfiguration.DeploymentMode)
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
                                                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Info($"[REGENERATE-CLUSTER-XML] Sleeve {sleeve.Id.IntegerValue}: Min=({bbox.Min.X:F6}, {bbox.Min.Y:F6}, {bbox.Min.Z:F6}), Max=({bbox.Max.X:F6}, {bbox.Max.Y:F6}, {bbox.Max.Z:F6})\n");
                            }
                        }
                    }
                    
                    // ✅ MASTER FIX: Save to _CLUSTER.xml with actual coordinates
                    SaveSleeveDataToClusterXml(sleeveDataList, category, filterName);
                    
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        // ✅ CRITICAL LOGGING: Log summary for this category
                                                if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[CATEGORY-SUMMARY] {DateTime.Now:HH:mm:ss.fff} - Category '{category}': Saved {sleeveDataList.Count} sleeves to _CLUSTER.xml\n");
                    }
                }
                
                if (!DeploymentConfiguration.DeploymentMode)
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[COORDINATE-SERVICE] {DateTime.Now:HH:mm:ss.fff} - COMPLETED: Processed {groupedSleeves.Count} categories\n");
                                        if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[REGENERATE-CLUSTER-XML] Completed regeneration for {groupedSleeves.Count} categories\n");
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
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
                                                        if (!DeploymentConfiguration.DeploymentMode)
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
                                                                if (!DeploymentConfiguration.DeploymentMode)
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
                                        if (!DeploymentConfiguration.DeploymentMode)
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
                                        if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[CATEGORY-CONTEXT] Sleeve {sleeve.Id.IntegerValue}: Determined category from context: '{category}'\n");
                }
                
                return category;
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[CATEGORY-ERROR] Sleeve {sleeve.Id.IntegerValue}: {ex.Message}\n");
                }
                
                // ✅ CRITICAL FIX: Try context-based detection even in error case
                try
                {
                    var category = DetermineCategoryFromContext(sleeve);
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                                                if (!DeploymentConfiguration.DeploymentMode)
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
                                                if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[CATEGORY-PARAM-CHECK] Sleeve {sleeve.Id.IntegerValue}: MEP_Category parameter exists, value='{categoryValue}', IsReadOnly={mepCategoryParam.IsReadOnly}\n");
                    }
                    
                    if (!string.IsNullOrEmpty(categoryValue))
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                                                        if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[CATEGORY-PARAM-SUCCESS] Sleeve {sleeve.Id.IntegerValue}: Using MEP_Category parameter: '{categoryValue}'\n");
                        }
                        return categoryValue;
                    }
                }
                else
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                                                if (!DeploymentConfiguration.DeploymentMode)
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
                                                if (!DeploymentConfiguration.DeploymentMode)
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
                                                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Info($"[CATEGORY-WALL] Sleeve {sleeve.Id.IntegerValue}: Wall-hosted, assuming Pipes\n");
                            }
                            return "Pipes";
                        }
                        else if (hostCategory.Contains("floor") || hostCategory.Contains("slab"))
                        {
                            // Floor-hosted sleeves are often for ducts
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                                                if (!DeploymentConfiguration.DeploymentMode)
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
                                                        if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[CATEGORY-SIZE] Sleeve {sleeve.Id.IntegerValue}: Small size ({avgSize:F2}m), assuming Pipes\n");
                        }
                        return "Pipes";
                    }
                }
                
                if (!DeploymentConfiguration.DeploymentMode)
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[CATEGORY-CONTEXT-FAILED] Sleeve {sleeve.Id.IntegerValue}: All context methods failed, using Unknown\n");
                }
                return "Unknown";
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
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
                                                        if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[ORIENTATION-DEBUG] Sleeve {sleeve.Id.IntegerValue}: Found clash zone, Orientation='{clashZone.MepElementOrientationDirection}'\n");
                        }
                        return clashZone.MepElementOrientationDirection ?? "";
                    }
                    else
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                                                        if (!DeploymentConfiguration.DeploymentMode)
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
                                        if (!DeploymentConfiguration.DeploymentMode)
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
                                                if (!DeploymentConfiguration.DeploymentMode)
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
                                        if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[REGENERATE-CLUSTER-XML] Saved {sleeveDataList.Count} sleeves to {fileName} and {pluralFileName}\n");
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
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
                                        if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[CLASH-CACHE-DEBUG] Filters directory: {filtersDirectory}\n");
                                        if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[CLASH-CACHE-DEBUG] Found {allXmlFiles.Length} total XML files\n");
                                        if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[CLASH-CACHE-DEBUG] Filtered to {xmlFiles.Count} relevant files\n");
                    
                    foreach (var file in xmlFiles)
                    {
                                                if (!DeploymentConfiguration.DeploymentMode)
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
                                                        if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[DEBUG] Loaded {clashNodes?.Count ?? 0} clash zones from {Path.GetFileName(xmlFile)}\n");
                        }
                    }
                    catch (Exception ex)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                                                        if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[DEBUG] Error loading {xmlFile}: {ex.Message}\n");
                        }
                    }
                }
                
                if (!DeploymentConfiguration.DeploymentMode)
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[CLASH-CACHE] Loaded {_clashZoneCache.Count} clash zones for orientation lookup\n");
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
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
                                if (!DeploymentConfiguration.DeploymentMode)
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
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[UpdateSleeveCoordinates] Found {allSleeves.Count} sleeves in model for position matching\n");
                
                if (allSleeves.Count > 0)
                {
                    var firstFewIds = allSleeves.Take(10).Select(s => s.Id.IntegerValue).ToList();
                                        if (!DeploymentConfiguration.DeploymentMode)
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
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [UPDATE_COORD_DEBUG] Matching IDs: {matchingIds.Count} out of {clashZoneSleeveIds.Count}\n");
                }
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
                                                if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[CLUSTER-CHECK] ClashZone has ClusterSleeveInstanceId={clashZone.ClusterSleeveInstanceId}, SleeveInstanceId={clashZone.SleeveInstanceId}\n");
                    }
                    
                    // ✅ CRITICAL FIX: Try direct ID match first (for individual sleeves OR cluster sleeves)
                    // This MUST succeed if SleeveInstanceId > 0 (individual) or ClusterSleeveInstanceId > 0 (cluster)
                    bool isClusterSleeve = clashZone.ClusterSleeveInstanceId > 0 && clashZone.SleeveInstanceId <= 0;
                    int targetSleeveId = isClusterSleeve ? clashZone.ClusterSleeveInstanceId : clashZone.SleeveInstanceId;
                    
                    if (targetSleeveId > 0)
                    {
                        try
                        {
                            var sleeveElement = _doc.GetElement(new ElementId(targetSleeveId));
                            if (sleeveElement is FamilyInstance sleeve && sleeve.Symbol.FamilyName.Contains("Opening"))
                            {
                                matchedSleeve = sleeve;
                                try {
                                    if (!DeploymentConfiguration.DeploymentMode)
                                    {
                                        string sleeveType = isClusterSleeve ? "cluster" : "individual";
                                        System.IO.File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [DIRECT-ID-MATCH] Found {sleeveType} sleeve {targetSleeveId} by ID\n");
                                    }
                } catch { }
                            }
                            else if (sleeveElement == null)
                            {
                                try {
                                    if (!DeploymentConfiguration.DeploymentMode)
                                    {
                                        string sleeveType = isClusterSleeve ? "cluster" : "individual";
                                        System.IO.File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [DIRECT-ID-MATCH] {sleeveType} sleeve {targetSleeveId} not found in document - may have been deleted\n");
                                    }
                } catch { }
                            }
                            else
                            {
                                try { System.IO.File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [DIRECT-ID-MATCH] Element {targetSleeveId} is not a FamilyInstance with Opening name (Type={sleeveElement?.GetType()?.Name})\n"); } catch { }
                            }
                        }
                        catch (Exception ex)
                        {
                            try {
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    System.IO.File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [DIRECT-ID-MATCH] Error getting sleeve {targetSleeveId}: {ex.Message}\n");
                                }
                } catch { }
                        }
                    }
                    
                    // ✅ CRITICAL FIX: If no direct match for cluster sleeve, try position matching
                    // For cluster sleeves, we need to update cluster bounding boxes, not individual sleeve bounding boxes
                    // Use SleevePlacementPoint if available, otherwise fall back to IntersectionPoint
                    double matchX = clashZone.SleevePlacementPointX;
                    double matchY = clashZone.SleevePlacementPointY;
                    double matchZ = clashZone.SleevePlacementPointZ;
                    string matchPointType = "SleevePlacementPoint";
                    
                    // ✅ FALLBACK: Use IntersectionPoint if SleevePlacementPoint is zero (for zones not yet placed)
                    if (matchX == 0 && matchY == 0 && clashZone.IntersectionPoint != null)
                    {
                        matchX = clashZone.IntersectionPoint.X;
                        matchY = clashZone.IntersectionPoint.Y;
                        matchZ = clashZone.IntersectionPoint.Z;
                        matchPointType = "IntersectionPoint";
                    }
                    
                    if (matchedSleeve == null && (matchX != 0 || matchY != 0))
                    {
                        double minDistance = double.MaxValue;
                        FamilyInstance closestSleeve = null;
                        
                        foreach (var sleeve in allSleeves)
                        {
                            var bbox = sleeve.get_BoundingBox(null);
                            if (bbox != null)
                            {
                                // Check if sleeve position matches clash zone point (within 1mm tolerance)
                                var distance = Math.Sqrt(
                                    Math.Pow(bbox.Min.X - matchX, 2) +
                                    Math.Pow(bbox.Min.Y - matchY, 2) +
                                    Math.Pow(bbox.Min.Z - matchZ, 2));
                                
                                if (distance < minDistance)
                                {
                                    minDistance = distance;
                                    closestSleeve = sleeve;
                                }
                                
                                if (distance < 0.00328) // 1mm tolerance in feet
                                {
                                    matchedSleeve = sleeve;
                                    clashZone.SleeveInstanceId = sleeve.Id.IntegerValue; // ✅ FIX: Update the instance ID
                                    // ✅ DEPLOYMENT: Wrapped in deployment mode check
                                    if (!DeploymentConfiguration.DeploymentMode)
                                    {
                                        System.IO.File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [POSITION-MATCH] Found sleeve {sleeve.Id.IntegerValue} by position for zone {clashZone.Id} using {matchPointType}, distance: {distance:F6}ft, MatchPoint=({matchX:F3}, {matchY:F3}, {matchZ:F3}), bbox.Min=({bbox.Min.X:F3}, {bbox.Min.Y:F3}, {bbox.Min.Z:F3})\n");
                                    }
                                    break;
                                }
                            }
                        }
                        
                        // ✅ DEBUG: Log if no match found but there was a closest sleeve
                        if (matchedSleeve == null && closestSleeve != null && !DeploymentConfiguration.DeploymentMode)
                        {
                            System.IO.File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [POSITION-MATCH-FAILED] Zone {clashZone.Id}: No match within 1mm tolerance using {matchPointType}. Closest sleeve {closestSleeve.Id.IntegerValue} at distance {minDistance:F6}ft ({minDistance * 304.8:F2}mm). MatchPoint=({matchX:F3}, {matchY:F3}, {matchZ:F3})\n");
                        }
                    }
                    else if (matchedSleeve == null && matchX == 0 && matchY == 0 && !DeploymentConfiguration.DeploymentMode)
                    {
                        System.IO.File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [POSITION-MATCH-SKIP] Zone {clashZone.Id}: Both SleevePlacementPoint and IntersectionPoint are zero, cannot match by position. SPP=({clashZone.SleevePlacementPointX:F3}, {clashZone.SleevePlacementPointY:F3}, {clashZone.SleevePlacementPointZ:F3}), IP=({clashZone.IntersectionPoint?.X:F3}, {clashZone.IntersectionPoint?.Y:F3}, {clashZone.IntersectionPoint?.Z:F3})\n");
                    }
                    
                    // Update coordinates if sleeve found
                    if (matchedSleeve != null)
                    {
                        var bbox = matchedSleeve.get_BoundingBox(null);
                        if (bbox != null)
                        {
                            if (isClusterSleeve)
                            {
                                // ✅ CLUSTER SLEEVE: Update cluster bounding box coordinates
                                clashZone.ClusterSleeveBoundingBoxMinX = bbox.Min.X;
                                clashZone.ClusterSleeveBoundingBoxMinY = bbox.Min.Y;
                                clashZone.ClusterSleeveBoundingBoxMinZ = bbox.Min.Z;
                                clashZone.ClusterSleeveBoundingBoxMaxX = bbox.Max.X;
                                clashZone.ClusterSleeveBoundingBoxMaxY = bbox.Max.Y;
                                clashZone.ClusterSleeveBoundingBoxMaxZ = bbox.Max.Z;
                                
                                try 
                                { 
                                    if (!DeploymentConfiguration.DeploymentMode)
                                    {
                                        File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [CLUSTER-BBOX-UPDATED] ClashZone {clashZone.Id}, Cluster Sleeve {targetSleeveId}: " +
                                        $"Min=({bbox.Min.X:F6}, {bbox.Min.Y:F6}, {bbox.Min.Z:F6}), Max=({bbox.Max.X:F6}, {bbox.Max.Y:F6}, {bbox.Max.Z:F6})\n");
                                    }
                                } 
                                catch { }
                            }
                            else
                            {
                                // ✅ INDIVIDUAL SLEEVE: Update individual sleeve bounding box
                                // ✅ CRITICAL LOGGING: Log BEFORE SetSleeveBoundingBox - Direct file write to placement_debug.log
                                try 
                                { 
                                    if (!DeploymentConfiguration.DeploymentMode)
                                    {
                                        File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [BOUNDING_BOX_BEFORE_SET] ClashZone {clashZone.Id}, Sleeve {clashZone.SleeveInstanceId}: " +
                                        $"BEFORE - ClashZone has: MinX={clashZone.SleeveBoundingBoxMinX:F6}, MinY={clashZone.SleeveBoundingBoxMinY:F6}, MinZ={clashZone.SleeveBoundingBoxMinZ:F6}, " +
                                        $"MaxX={clashZone.SleeveBoundingBoxMaxX:F6}, MaxY={clashZone.SleeveBoundingBoxMaxY:F6}, MaxZ={clashZone.SleeveBoundingBoxMaxZ:F6}\n");
                                    }
                                    
                                    System.IO.File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [BOUNDING_BOX_BEFORE_SET] Revit bbox from sleeve {matchedSleeve.Id.IntegerValue}: " +
                                        $"Min=({bbox.Min.X:F6}, {bbox.Min.Y:F6}, {bbox.Min.Z:F6}), Max=({bbox.Max.X:F6}, {bbox.Max.Y:F6}, {bbox.Max.Z:F6})\n");
                                } 
                                catch { }
                                
                                // ✅ CORRECT: Update bounding box and instance ID
                                clashZone.SetSleeveBoundingBox(bbox);
                                
                                // ✅ CRITICAL LOGGING: Log AFTER SetSleeveBoundingBox - Direct file write to placement_debug.log
                                try 
                                { 
                                    if (!DeploymentConfiguration.DeploymentMode)
                                    {
                                        File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [BOUNDING_BOX_AFTER_SET] ClashZone {clashZone.Id}, Sleeve {clashZone.SleeveInstanceId}: " +
                                        $"AFTER - ClashZone now has: MinX={clashZone.SleeveBoundingBoxMinX:F6}, MinY={clashZone.SleeveBoundingBoxMinY:F6}, MinZ={clashZone.SleeveBoundingBoxMinZ:F6}, " +
                                        $"MaxX={clashZone.SleeveBoundingBoxMaxX:F6}, MaxY={clashZone.SleeveBoundingBoxMaxY:F6}, MaxZ={clashZone.SleeveBoundingBoxMaxZ:F6}\n");
                                    }
                                    
                                    System.IO.File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [UPDATED] Sleeve {clashZone.SleeveInstanceId}: Min=({bbox.Min.X:F6}, {bbox.Min.Y:F6}, {bbox.Min.Z:F6}), Max=({bbox.Max.X:F6}, {bbox.Max.Y:F6}, {bbox.Max.Z:F6})\n");
                                } 
                                catch { }
                            }
                        }
                        else
                        {
                            // ✅ DEPLOYMENT: Wrapped in deployment mode check
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Info($"[NO-BBOX] Sleeve {targetSleeveId} ({(isClusterSleeve ? "cluster" : "individual")}) has no bounding box\n");
                            }
                        }
                    }
                    else
                    {
                        // ✅ DEPLOYMENT: Wrapped in deployment mode check
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                                                        if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[NO-MATCH] ClashZone {clashZone.Id} - no matching sleeve found (SleeveInstanceId={clashZone.SleeveInstanceId}, PlacePoint=({clashZone.SleevePlacementPointX:F3}, {clashZone.SleevePlacementPointY:F3}, {clashZone.SleevePlacementPointZ:F3}))\n");
                        }
                    }
                }
                catch (Exception ex)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                                                if (!DeploymentConfiguration.DeploymentMode)
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

using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Service to monitor and copy levels from linked files to ensure proper level association for sleeve placement
    /// </summary>
    public class LevelMonitoringService
    {
        private readonly Document _hostDocument;
        private readonly Dictionary<string, Level> _copiedLevels;

        public LevelMonitoringService(Document hostDocument)
        {
            _hostDocument = hostDocument ?? throw new ArgumentNullException(nameof(hostDocument));
            _copiedLevels = new Dictionary<string, Level>();
            DebugLogger.SetServiceContext("LevelMonitoring");
        }

        /// <summary>
        /// Get or create a level in the host document based on a linked level
        /// </summary>
        /// <param name="linkedLevel">The level from the linked document</param>
        /// <param name="linkedDocument">The linked document containing the level</param>
        /// <returns>The corresponding level in the host document</returns>
        public Level GetOrCreateHostLevel(Level linkedLevel, Document linkedDocument)
        {
            if (linkedLevel == null || linkedDocument == null)
                return null;

            try
            {
                string levelKey = $"{linkedDocument.Title}_{linkedLevel.Name}_{linkedLevel.Elevation:F3}";
                
                // Check if we already copied this level
                if (_copiedLevels.ContainsKey(levelKey))
                {
                    DebugLogger.Info($"[LevelMonitoring] Using cached level: {linkedLevel.Name} (elevation: {linkedLevel.Elevation:F2})");
                    return _copiedLevels[levelKey];
                }

                // First, try to find existing level by name and elevation
                var existingLevel = FindExistingLevel(linkedLevel.Name, linkedLevel.Elevation);
                if (existingLevel != null)
                {
                    _copiedLevels[levelKey] = existingLevel;
                    DebugLogger.Info($"[LevelMonitoring] Found existing level: {existingLevel.Name} (elevation: {existingLevel.Elevation:F2})");
                    return existingLevel;
                }

                // Create new level if not found
                var newLevel = CreateLevelFromLinked(linkedLevel);
                if (newLevel != null)
                {
                    _copiedLevels[levelKey] = newLevel;
                    DebugLogger.Info($"[LevelMonitoring] Created new level: {newLevel.Name} (elevation: {newLevel.Elevation:F2})");
                    return newLevel;
                }

                DebugLogger.Warning($"[LevelMonitoring] Failed to get or create level for: {linkedLevel.Name}");
                return null;
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[LevelMonitoring] Error getting/creating level for {linkedLevel.Name}: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Find existing level by name and elevation (with tolerance)
        /// </summary>
        private Level FindExistingLevel(string levelName, double elevation)
        {
            const double tolerance = 0.01; // 1cm tolerance

            var existingLevels = new FilteredElementCollector(_hostDocument)
                .OfClass(typeof(Level))
                .Cast<Level>()
                .ToList();

            // First try exact name match
            var exactMatch = existingLevels.FirstOrDefault(l => 
                string.Equals(l.Name, levelName, StringComparison.OrdinalIgnoreCase) &&
                Math.Abs(l.Elevation - elevation) <= tolerance);

            if (exactMatch != null)
                return exactMatch;

            // Then try elevation-only match (in case names are different)
            var elevationMatch = existingLevels.FirstOrDefault(l => 
                Math.Abs(l.Elevation - elevation) <= tolerance);

            if (elevationMatch != null)
            {
                DebugLogger.Info($"[LevelMonitoring] Found level by elevation: {elevationMatch.Name} matches elevation {elevation:F2}");
                return elevationMatch;
            }

            return null;
        }

        /// <summary>
        /// Create a new level in the host document based on the linked level
        /// </summary>
        private Level CreateLevelFromLinked(Level linkedLevel)
        {
            try
            {
                using (var transaction = new Transaction(_hostDocument, "Copy Level from Linked File"))
                {
                    transaction.Start();

                    // Create new level with the same elevation
                    var newLevel = Level.Create(_hostDocument, linkedLevel.Elevation);
                    
                    if (newLevel != null)
                    {
                        // Set the name to match the linked level
                        var nameParam = newLevel.get_Parameter(BuiltInParameter.DATUM_TEXT);
                        if (nameParam != null && !nameParam.IsReadOnly)
                        {
                            nameParam.Set(linkedLevel.Name);
                        }

                        transaction.Commit();
                        DebugLogger.Info($"[LevelMonitoring] Successfully created level: {linkedLevel.Name} at elevation {linkedLevel.Elevation:F2}");
                        return newLevel;
                    }
                    else
                    {
                        transaction.RollBack();
                        DebugLogger.Error($"[LevelMonitoring] Failed to create level: {linkedLevel.Name}");
                        return null;
                    }
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[LevelMonitoring] Exception creating level {linkedLevel.Name}: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Get level for a specific elevation (find closest or create if needed)
        /// </summary>
        /// <param name="elevation">The elevation to find a level for</param>
        /// <param name="preferredName">Preferred name for the level if creating new</param>
        /// <returns>The appropriate level for the given elevation</returns>
        public Level GetLevelForElevation(double elevation, string preferredName = null)
        {
            try
            {
                // First try to find existing level close to this elevation
                var existingLevels = new FilteredElementCollector(_hostDocument)
                    .OfClass(typeof(Level))
                    .Cast<Level>()
                    .OrderBy(l => Math.Abs(l.Elevation - elevation))
                    .ToList();

                if (existingLevels.Count > 0)
                {
                    var closestLevel = existingLevels.First();
                    const double maxDistance = 1.0; // 1 foot tolerance

                    if (Math.Abs(closestLevel.Elevation - elevation) <= maxDistance)
                    {
                        DebugLogger.Info($"[LevelMonitoring] Using closest existing level: {closestLevel.Name} (elevation: {closestLevel.Elevation:F2}) for target elevation {elevation:F2}");
                        return closestLevel;
                    }
                }

                // Create new level if no suitable existing level found
                if (!string.IsNullOrEmpty(preferredName))
                {
                    return CreateLevelAtElevation(elevation, preferredName);
                }
                else
                {
                    // Generate a name based on elevation
                    var levelName = $"Level {elevation:F1}";
                    return CreateLevelAtElevation(elevation, levelName);
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[LevelMonitoring] Error getting level for elevation {elevation:F2}: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Create a new level at the specified elevation
        /// </summary>
        private Level CreateLevelAtElevation(double elevation, string levelName)
        {
            try
            {
                using (var transaction = new Transaction(_hostDocument, "Create Level for Sleeve Placement"))
                {
                    transaction.Start();

                    var newLevel = Level.Create(_hostDocument, elevation);
                    
                    if (newLevel != null)
                    {
                        var nameParam = newLevel.get_Parameter(BuiltInParameter.DATUM_TEXT);
                        if (nameParam != null && !nameParam.IsReadOnly)
                        {
                            nameParam.Set(levelName);
                        }

                        transaction.Commit();
                        DebugLogger.Info($"[LevelMonitoring] Created new level: {levelName} at elevation {elevation:F2}");
                        return newLevel;
                    }
                    else
                    {
                        transaction.RollBack();
                        DebugLogger.Error($"[LevelMonitoring] Failed to create level: {levelName}");
                        return null;
                    }
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[LevelMonitoring] Exception creating level {levelName}: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Get statistics about monitored levels
        /// </summary>
        public string GetStatistics()
        {
            var totalHostLevels = new FilteredElementCollector(_hostDocument)
                .OfClass(typeof(Level))
                .GetElementCount();

            return $"Level Monitoring Statistics: {_copiedLevels.Count} levels cached, {totalHostLevels} total levels in host document";
        }
    }
}





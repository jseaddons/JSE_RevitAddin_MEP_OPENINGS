using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces;
using JSE_RevitAddin_MEP_OPENINGS.Utils;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Service for filtering eligible clash zones.
    /// Extracted from UniversalSleevePlacerService to adhere to Single Responsibility Principle.
    /// </summary>
    public class ZoneFilterService : IZoneFilterService
    {
        /// <summary>
        /// Pre-filter eligible clash zones based on UI selection and flags.
        /// Currently performs basic validation and optional parallel processing.
        /// </summary>
        public List<ClashZone> PreFilterEligibleClashZones(Document doc, List<ClashZone> allClashZones)
        {
            if (allClashZones == null || allClashZones.Count == 0)
                return new List<ClashZone>();

            // Access settings
            var settings = ApplicationProfileService.Instance.GetCurrentSettings();
            double minWallThickness = settings.MinWallThickness; // In internal units (feet) or mm?
            // SettingsModel says: "Minimum wall thickness in mm"
            // We need to convert to internal units (feet) for comparison if ClashZone stores internal units.
            // ClashZone typically stores internal units.
            
            // However, SettingsModel usually stores raw values from UI. If UI is mm, this is mm.
            // Let's assume it is mm based on comments in SettingsModel.
            double minThicknessInternal = 0;
            if (minWallThickness > 0)
            {
                minThicknessInternal = UnitUtils.ConvertToInternalUnits(minWallThickness, UnitTypeId.Millimeters);
            }

            // ✅ PERFORMANCE: Pre-filter and validate clash zones (reserved for future optimization)
            var sw = System.Diagnostics.Stopwatch.StartNew();
            
            var validZones = new List<ClashZone>();
            int excludedCount = 0;

            foreach (var z in allClashZones)
            {
                if (z == null) continue;

                // Basic validation
                if (z.IntersectionPointX == 0 && z.IntersectionPointY == 0 && z.IntersectionPointZ == 0) continue;

                // ✅ WALL THICKNESS CHECK
                if (minThicknessInternal > 0)
                {
                    // Check if host is a wall
                    if (z.StructuralElementType != null && (z.StructuralElementType.Contains("Wall") || z.StructuralElementType.Contains("Walls")))
                    {
                        // Use WallThickness if available, otherwise StructuralElementThickness
                        double thickness = z.WallThickness > 0 ? z.WallThickness : z.StructuralElementThickness;
                        
                        if (thickness > 0 && thickness < minThicknessInternal)
                        {
                            excludedCount++;
                            continue; // Skip this zone
                        }
                    }
                }

                validZones.Add(z);
            }
            
            sw.Stop();
            if (!DeploymentConfiguration.DeploymentMode)
            {
                if (sw.ElapsedMilliseconds > 5 || excludedCount > 0)
                {
                    SafeFileLogger.SafeAppendText("placement_performance.log",
                        $"[{DateTime.Now:HH:mm:ss}] ⚡ Pre-colored {allClashZones.Count} clash zones → {validZones.Count} valid ({excludedCount} filtered by thickness < {minWallThickness}mm) in {sw.ElapsedMilliseconds}ms\n");
                }
            }
            
            return validZones;
        }

        /// <summary>
        /// Filter zones by resolved status (optional helper method for future use).
        /// </summary>
        public List<ClashZone> FilterUnresolvedZones(List<ClashZone> zones)
        {
            if (zones == null || zones.Count == 0)
                return new List<ClashZone>();

            return zones
                .Where(z => z != null)
                .Where(z => !z.IsResolved && !z.IsClusterResolved)
                .ToList();
        }

        /// <summary>
        /// Filter zones by category (optional helper method for future use).
        /// </summary>
        public List<ClashZone> FilterByCategory(List<ClashZone> zones, string category)
        {
            if (zones == null || zones.Count == 0 || string.IsNullOrWhiteSpace(category))
                return new List<ClashZone>();

            return zones
                .Where(z => z != null)
                .Where(z => string.Equals(z.MepElementCategory, category, StringComparison.OrdinalIgnoreCase))
                .ToList();
        }

        /// <summary>
        /// Filter zones by host type (optional helper method for future use).
        /// </summary>
        public List<ClashZone> FilterByHostType(List<ClashZone> zones, string hostType)
        {
            if (zones == null || zones.Count == 0 || string.IsNullOrWhiteSpace(hostType))
                return new List<ClashZone>();

            return zones
                .Where(z => z != null)
                .Where(z => z.StructuralElementType != null && 
                           z.StructuralElementType.Contains(hostType, StringComparison.OrdinalIgnoreCase))
                .ToList();
        }
    }
}

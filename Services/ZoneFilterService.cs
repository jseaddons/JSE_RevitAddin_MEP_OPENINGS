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

            // ✅ PERFORMANCE: Pre-filter and validate clash zones (reserved for future optimization)
            // Note: Actual clearance calculation happens during placement, not here
            // ClashZone does not have SleeveDepth or cached clearance properties
            // This method can be used for parallel validation in the future
            
            if (!OptimizationFlags.UseParallelClearanceCalculation || allClashZones.Count <= 10)
            {
                // Skip parallel processing for small lists
                return allClashZones;
            }

            var sw = System.Diagnostics.Stopwatch.StartNew();
            
            // Currently just returns the list as-is
            // Future optimization: Add parallel validation logic here
            // (e.g., check required properties are set, validate coordinates, etc.)
            
            // Basic validation: filter out zones with invalid data
            var validZones = allClashZones
                .Where(z => z != null)
                .Where(z => z.IntersectionPointX != 0 || z.IntersectionPointY != 0 || z.IntersectionPointZ != 0)
                .ToList();
            
            sw.Stop();
            if (!DeploymentConfiguration.DeploymentMode && sw.ElapsedMilliseconds > 5)
            {
                SafeFileLogger.SafeAppendText("placement_performance.log",
                    $"[{DateTime.Now:HH:mm:ss}] ⚡ Pre-filtered {allClashZones.Count} clash zones → {validZones.Count} valid in {sw.ElapsedMilliseconds}ms\n");
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

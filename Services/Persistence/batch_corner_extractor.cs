using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Persistence
{
    /// <summary>
    /// âœ… BATCH CORNER EXTRACTOR: Extracts actual corner coordinates from placed Revit sleeves.
    /// NO CALCULATION - Pure extraction from Revit geometry after placement + regeneration.
    /// 
    /// Usage:
    /// 1. Place sleeves in Revit
    /// 2. Call doc.Regenerate()
    /// 3. Call ExtractCorners() to get corner data
    /// 4. Save to database using Repository
    /// </summary>
    public class BatchSleeveCornerExtractor
    {
        private readonly Document _doc;

        public BatchSleeveCornerExtractor(Document doc)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
        }
        public BatchSleeveCornerExtractor()
        {
        }

        /// <summary>
        /// Extracts corners for a list of clash zones from the Revit document.
        /// Returns a list of tuples compatible with ClashZoneRepository.BatchUpdateSleeveCorners.
        /// </summary>
        public List<(Guid ClashZoneGuid,
            double Corner1X, double Corner1Y, double Corner1Z,
            double Corner2X, double Corner2Y, double Corner2Z,
            double Corner3X, double Corner3Y, double Corner3Z,
            double Corner4X, double Corner4Y, double Corner4Z)> ExtractCorners(Document doc, List<ClashZone> zones)
        {
            var results = new List<(Guid, double, double, double, double, double, double, double, double, double, double, double, double)>();
            
            if (zones == null || zones.Count == 0) return results;

            SafeFileLogger.SafeAppendText("corner_extraction.log",
                $"[{DateTime.Now:HH:mm:ss}] [START] BATCH EXTRACTION: Processing {zones.Count} zones...\n");

            int successCount = 0;
            int failedCount = 0;

            foreach (var zone in zones)
            {
                // Skip if no sleeve ID
                if (zone.SleeveInstanceId <= 0) continue;

                try
                {
                    var sleeveId = new ElementId(zone.SleeveInstanceId);
                    var sleeve = doc.GetElement(sleeveId) as FamilyInstance;

                    if (sleeve == null)
                    {
                        failedCount++;
                        SafeFileLogger.SafeAppendText("corner_extraction.log",
                            $"[{DateTime.Now:HH:mm:ss}] [WARN] Sleeve {zone.SleeveInstanceId} not found in Revit\n");
                        continue;
                    }

                    var corners = ExtractCornersFromRevitElement(sleeve);

                    if (corners == null)
                    {
                        failedCount++;
                        SafeFileLogger.SafeAppendText("corner_extraction.log",
                            $"[{DateTime.Now:HH:mm:ss}] [ERROR] Failed to extract corners for sleeve {zone.SleeveInstanceId}\n");
                        continue;
                    }

                    // Add to results
                    results.Add((
                        zone.Id,
                        corners.Value.c1.X, corners.Value.c1.Y, corners.Value.c1.Z,
                        corners.Value.c2.X, corners.Value.c2.Y, corners.Value.c2.Z,
                        corners.Value.c3.X, corners.Value.c3.Y, corners.Value.c3.Z,
                        corners.Value.c4.X, corners.Value.c4.Y, corners.Value.c4.Z
                    ));
                    
                    successCount++;
                    
                    // Verbose log for first few
                    if (successCount <= 5)
                    {
                        SafeFileLogger.SafeAppendText("corner_extraction.log",
                            $"[{DateTime.Now:HH:mm:ss}] [SUCCESS] Extracted corners for {zone.SleeveInstanceId}\n");
                    }
                }
                catch (Exception ex)
                {
                    failedCount++;
                    SafeFileLogger.SafeAppendText("corner_extraction.log",
                        $"[{DateTime.Now:HH:mm:ss}] [ERROR] Exception for sleeve {zone.SleeveInstanceId}: {ex.Message}\n");
                }
            }

            SafeFileLogger.SafeAppendText("corner_extraction.log",
                $"[{DateTime.Now:HH:mm:ss}] [SUCCESS] BATCH EXTRACTION FINISHED: {successCount} extracted, {failedCount} failed\n");

            return results;
        }

        private (XYZ c1, XYZ c2, XYZ c3, XYZ c4)? ExtractCornersFromRevitElement(FamilyInstance sleeve)
        {
            try
            {
                // ✅ FIX: Void families have no positive volume solids
                // Use element BoundingBox directly - it's already in World Coordinates
                BoundingBoxXYZ bbox = sleeve.get_BoundingBox(null);
                
                if (bbox == null)
                {
                    SafeFileLogger.SafeAppendText("corner_extraction.log",
                        $"[{DateTime.Now:HH:mm:ss}] [WARN] No BoundingBox for sleeve {sleeve.Id}\n");
                    return null;
                }
                
                // ✅ FIX: Use Min.Z for corners 1-2 (bottom), Max.Z for corners 3-4 (top)
                // This captures the full Z range for proper 3D overlap detection
                XYZ c1 = new XYZ(bbox.Min.X, bbox.Min.Y, bbox.Min.Z);  // Bottom-left
                XYZ c2 = new XYZ(bbox.Max.X, bbox.Min.Y, bbox.Min.Z);  // Bottom-right
                XYZ c3 = new XYZ(bbox.Max.X, bbox.Max.Y, bbox.Max.Z);  // Top-right
                XYZ c4 = new XYZ(bbox.Min.X, bbox.Max.Y, bbox.Max.Z);  // Top-left
                
                SafeFileLogger.SafeAppendText("corner_extraction.log",
                    $"[{DateTime.Now:HH:mm:ss}] [SUCCESS] Extracted corners for {sleeve.Id}: " +
                    $"C1=({c1.X:F4},{c1.Y:F4},{c1.Z:F4}) C3=({c3.X:F4},{c3.Y:F4},{c3.Z:F4}) ZRange={bbox.Min.Z:F4}-{bbox.Max.Z:F4}\n");

                return (c1, c2, c3, c4);
            }
            catch (Exception ex)
            {
                SafeFileLogger.SafeAppendText("corner_extraction.log",
                    $"[{DateTime.Now:HH:mm:ss}] [ERROR] Exception extracting corners for sleeve {sleeve.Id}: {ex.Message}\n");
                return null;
            }
        }
    }
}

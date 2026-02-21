using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Persistence
{
    /// <summary>
    /// BATCH CORNER EXTRACTOR: Only allowed source for persisting SleeveCorner1-4 X/Y/Z (no calculation).
    /// Extracts corners from placed Revit sleeve geometry via CalculateCornersFromInstance.
    /// See REVIT_GEOMETRY_RULES.md. Usage: Place sleeves, Regenerate(), ExtractCorners(), then Repository.BatchUpdateSleeveCorners.
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
                    var sleeveId = ElementIdCompat.FromValue(zone.SleeveInstanceId);
                    var sleeve = doc.GetElement(sleeveId) as FamilyInstance;

                    if (sleeve == null)
                    {
                        failedCount++;
                        SafeFileLogger.SafeAppendText("corner_extraction.log",
                            $"[{DateTime.Now:HH:mm:ss}] [WARN] Sleeve {zone.SleeveInstanceId} not found in Revit\n");
                        continue;
                    }

                    var corners = ExtractCornersFromRevitElement(sleeve, zone);

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

        private (XYZ c1, XYZ c2, XYZ c3, XYZ c4)? ExtractCornersFromRevitElement(FamilyInstance sleeve, ClashZone zone)
        {
            try
            {
                // ✅ REUSE: Use existing SleeveCornerCalculationService (Methodology compliant)
                var calculator = new JSE_RevitAddin_MEP_OPENINGS.Services.Geometry.SleeveCornerCalculationService();
                
                // ✅ CRITICAL: Pass HostOrientation AND StructuralElementType from database
                // This ensures correct strategy (Wall vs Floor) is used even if sleeve.Host is null
                var corners = calculator.CalculateCornersFromInstance(sleeve, zone?.HostOrientation, zone?.StructuralElementType);
                
                if (corners.HasValue)
                {
                     var c = corners.Value;
                     SafeFileLogger.SafeAppendText("corner_extraction.log",
                        $"[{DateTime.Now:HH:mm:ss}] [SUCCESS] Extracted corners using SleeveCornerCalculationService for {sleeve.Id}\n" + 
                        $"   -> C1=({c.corner1.X:F3},{c.corner1.Y:F3},{c.corner1.Z:F3}), C2=({c.corner2.X:F3},{c.corner2.Y:F3},{c.corner2.Z:F3})\n" +
                        $"   -> C3=({c.corner3.X:F3},{c.corner3.Y:F3},{c.corner3.Z:F3}), C4=({c.corner4.X:F3},{c.corner4.Y:F3},{c.corner4.Z:F3})\n");
                     return corners;
                }
                
                // Fallback to AABB if service fails (e.g. no solids)
                SafeFileLogger.SafeAppendText("corner_extraction.log",
                    $"[{DateTime.Now:HH:mm:ss}] [WARN] SleeveCornerCalculationService returned null for {sleeve.Id}. Extraction Failed (AABB Fallback DISABLED).\n");
                    
                // DEBUG: Why did it fail?
                try 
                {
                    var pW = sleeve.LookupParameter("Element Width");
                    var pH = sleeve.LookupParameter("Element Height");
                    var pD = sleeve.LookupParameter("Element Diameter");
                    var loc = sleeve.Location as LocationPoint;
                    SafeFileLogger.SafeAppendText("corner_extraction.log",
                        $"   -> DEBUG: W={pW?.AsDouble() ?? -1}, H={pH?.AsDouble() ?? -1}, Dia={pD?.AsDouble() ?? -1}, Loc={loc?.Point?.ToString() ?? "NULL"}\n");
                }
                catch {}

                // USER REQUEST: NO FALLBACK AABB. Strict refusal if strategy fails.
                // The Strategy Pattern in SleeveCornerCalculationService handles both Void and Wall logic now.
                // If it returns null, we respect it and do NOT fabricate generic bounds.
                return null;
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

using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Refactored
{
    /// <summary>
    /// ✅ SOLID REFACTORED: Service for checking if clash zones are visible in the current section box.
    /// Extracted from UniversalSleevePlacementCommand to adhere to SRP.
    /// </summary>
    public class SectionBoxCheckerService : ISectionBoxChecker
    {
        public bool IsClashZoneVisibleInCurrentSectionBox(Document doc, ClashZone clashZone, string logPrefix)
        {
            try
            {
                // Check if we have an active 3D view with section box
                if (!(doc.ActiveView is View3D view3D) || !view3D.IsSectionBoxActive)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"{logPrefix} No active 3D section box - ClashZone {clashZone.Id} considered visible");
                    }
                    return true; // No section box = all visible
                }

                // Get section box bounds
                var sectionBox = SectionBoxHelper.GetSectionBoxBounds(view3D);
                if (sectionBox == null)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"{logPrefix} Could not get section box bounds - ClashZone {clashZone.Id} considered visible");
                    }
                    return true; // Can't get bounds = all visible
                }

                // Tolerance to avoid precision misses (in feet ~ 30mm)
                const double tol = 0.1;

                // Prefer checking the clash zone's intersection bounding box if present
                var czBox = clashZone.ClashBoundingBox;
                if (czBox != null)
                {
                    // Expand section box slightly by tol
                    bool overlaps =
                        (czBox.Min.X <= sectionBox.Max.X + tol) && (czBox.Max.X >= sectionBox.Min.X - tol) &&
                        (czBox.Min.Y <= sectionBox.Max.Y + tol) && (czBox.Max.Y >= sectionBox.Min.Y - tol) &&
                        (czBox.Min.Z <= sectionBox.Max.Z + tol) && (czBox.Max.Z >= sectionBox.Min.Z - tol);

                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"{logPrefix} SectionBox test (BBox) CZ={clashZone.Id} overlaps={overlaps}");
                    }
                    return overlaps;
                }

                // Fallback to single point test with tolerance
                var p = clashZone.IntersectionPoint;
                bool inside =
                    p.X >= sectionBox.Min.X - tol && p.X <= sectionBox.Max.X + tol &&
                    p.Y >= sectionBox.Min.Y - tol && p.Y <= sectionBox.Max.Y + tol &&
                    p.Z >= sectionBox.Min.Z - tol && p.Z <= sectionBox.Max.Z + tol;

                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"{logPrefix} SectionBox test (Point) CZ={clashZone.Id} inside={inside}");
                }
                return inside;
            }
            catch (System.Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Error($"{logPrefix} Error checking ClashZone {clashZone.Id} section box visibility: {ex.Message} - considering visible");
                }
                return true; // Consider visible on error
            }
        }
    }
}


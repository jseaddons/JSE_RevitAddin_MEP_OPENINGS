using System;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Proximity
{
    /// <summary>
    /// Proximity checker that uses SleeveCorner data for accurate results.
    /// Recommended for all rectangular sleeves where Revit bounding boxes might be erratic.
    /// </summary>
    public class CornerProximityChecker : IProximityChecker
    {
        private readonly SleeveCornerProximityHelper _helper;

        public CornerProximityChecker()
        {
            _helper = new SleeveCornerProximityHelper();
        }

        public bool CheckProximity(dynamic sleeve1, dynamic sleeve2, double tolerance)
        {
            try
            {
                // ✅ CRASH-SAFE: Handled both wrapper objects (ClashZoneWorkItem) and direct ClashZone objects
                ClashZone cz1 = null;
                ClashZone cz2 = null;

                // Try to get ClashZone from sleeve1
                if (sleeve1 is ClashZone c1) cz1 = c1;
                else if (sleeve1 != null) { try { cz1 = sleeve1.ClashZone as ClashZone; } catch { } }

                // Try to get ClashZone from sleeve2
                if (sleeve2 is ClashZone c2) cz2 = c2;
                else if (sleeve2 != null) { try { cz2 = sleeve2.ClashZone as ClashZone; } catch { } }

                if (cz1 == null || cz2 == null)
                    return false;

                bool isNearby = _helper.AreWithinProximity(cz1, cz2, tolerance);

                return isNearby;
            }
            catch (Exception ex)
            {
                SafeFileLogger.SafeAppendText("geometry_errors.log",
                    $"[CornerProximityChecker] Exception: {ex.Message}");
                return false;
            }
        }

        public double? CalculateDistance(dynamic sleeve1, dynamic sleeve2)
        {
            try
            {
                // ✅ CRASH-SAFE: Handled both wrapper objects (ClashZoneWorkItem) and direct ClashZone objects
                ClashZone cz1 = null;
                ClashZone cz2 = null;

                // Try to get ClashZone from sleeve1
                if (sleeve1 is ClashZone c1) cz1 = c1;
                else if (sleeve1 != null) { try { cz1 = sleeve1.ClashZone as ClashZone; } catch { } }

                // Try to get ClashZone from sleeve2
                if (sleeve2 is ClashZone c2) cz2 = c2;
                else if (sleeve2 != null) { try { cz2 = sleeve2.ClashZone as ClashZone; } catch { } }

                if (cz1 == null || cz2 == null)
                    return null;

                return _helper.GetDistanceBetweenCorners(cz1, cz2);
            }
            catch (Exception ex)
            {
                SafeFileLogger.SafeAppendText("geometry_errors.log",
                    $"[CornerProximityChecker] CalculateDistance Exception: {ex.Message}");
                return null;
            }
        }
    }
}

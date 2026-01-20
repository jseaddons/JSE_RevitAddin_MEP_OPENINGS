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
                if (sleeve1?.ClashZone == null || sleeve2?.ClashZone == null)
                    return false;

                ClashZone cz1 = sleeve1.ClashZone as ClashZone;
                ClashZone cz2 = sleeve2.ClashZone as ClashZone;

                if (cz1 == null || cz2 == null)
                    return false;

                bool isNearby = _helper.AreWithinProximity(cz1, cz2, tolerance);

                // ✅ DIAGNOSTIC LOGGING: For dampers or rectangular ducts
                string category = cz1.MepElementCategory ?? "";
                if (category.Contains("Damper") || category.Contains("Duct"))
                {
                    double toleranceMM = tolerance * 304.8;
                    
                    SafeFileLogger.SafeAppendText("cluster_debug.log",
                        $"[CornerChecker] 📐 {category} Corner Proximity: Result={isNearby}, Tol={toleranceMM:F1}mm, " +
                        $"ID1={cz1.SleeveInstanceId}, ID2={cz2.SleeveInstanceId}\n");
                }

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
                if (sleeve1?.ClashZone == null || sleeve2?.ClashZone == null)
                    return null;

                ClashZone cz1 = sleeve1.ClashZone as ClashZone;
                ClashZone cz2 = sleeve2.ClashZone as ClashZone;

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

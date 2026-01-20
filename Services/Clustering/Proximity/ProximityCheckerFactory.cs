using System;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Proximity
{
    /// <summary>
    /// Factory for creating appropriate proximity checker based on sleeve characteristics.
    /// Phase 2: Decision logic extracted from UniversalClusterService.
    /// Decision tree: round pipes/ducts → rotated → bounding box
    /// </summary>
    public static class ProximityCheckerFactory
    {
        /// <summary>
        /// Create appropriate proximity checker based on sleeve characteristics.
        /// </summary>
        /// <param name="sleeve1">First sleeve (for type detection)</param>
        /// <param name="sleeve2">Second sleeve (for type detection)</param>
        /// <param name="rotationAngle">Rotation angle in radians (for rotated checker, if needed)</param>
        /// <param name="isRotated">Whether sleeves are rotated (non-axis-aligned)</param>
        /// <returns>Appropriate proximity checker instance</returns>
        public static IProximityChecker CreateChecker(
            dynamic sleeve1,
            dynamic sleeve2,
            double rotationAngle = 0.0,
            bool isRotated = false)
        {
            try
            {
                // ✅ CRASH-SAFE: Validate inputs
                if (sleeve1 == null || sleeve2 == null)
                {
                    return new BoundingBoxProximityChecker(); // Default fallback
                }

                // ✅ DECISION 1: Check for rectangular sleeves (all categories) → use corner-based proximity
                // This is the most accurate for any rectangular geometry where Revit BBoxes are erratic
                var helper = new SleeveCornerProximityHelper();
                var cz1 = sleeve1.ClashZone as Models.ClashZone;
                
                if (cz1 != null && helper.IsRectangularSleeve(cz1))
                {
                    return new CornerProximityChecker();
                }

                // ✅ DECISION 2: Check for round pipes/ducts → use edge-to-edge distance
                string systemType = sleeve1.SystemType ?? "";
                bool isRoundPipeOrDuct = (systemType.IndexOf("Pipe", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                         systemType.IndexOf("Duct", StringComparison.OrdinalIgnoreCase) >= 0) &&
                                        (sleeve1.IsCircular == true || sleeve2.IsCircular == true);

                if (isRoundPipeOrDuct)
                {
                    // Round pipes/ducts: Use edge-to-edge distance (accounts for sleeve diameter)
                    return new EdgeToEdgeProximityChecker();
                }

                // ✅ DECISION 3: Check for rotated sleeves → use rotated proximity checker
                if (isRotated && Math.Abs(rotationAngle) > 1e-6)
                {
                    // Rotated sleeves: Use rotated proximity checker with rotation angle
                    return new RotatedProximityChecker(rotationAngle);
                }

                // ✅ DECISION 4: Default fallback → use bounding box proximity checker
                // This handles axis-aligned rectangular sleeves on floors, walls, etc.
                return new BoundingBoxProximityChecker();
            }
            catch (Exception ex)
            {
                // ✅ CRASH-SAFE: Return safe default on exception
                SafeFileLogger.SafeAppendText("geometry_errors.log",
                    $"[ProximityCheckerFactory] Exception in CreateChecker: {ex.Message}, StackTrace: {ex.StackTrace}");
                return new CornerProximityChecker(); // 🏆 Preferred fallback now is Corners if possible
            }
        }
    }
}


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

                // ✅ Based on SLEEVE SHAPE only (opening in wall/floor), not MEP element shape.
                var helper = new SleeveCornerProximityHelper();
                var cz1 = sleeve1.ClashZone as Models.ClashZone;
                var cz2 = sleeve2.ClashZone as Models.ClashZone;

                // DECISION 1: Both sleeves have rectangular opening (corner geometry) → corner-based proximity
                if (cz1 != null && cz2 != null && helper.HasRectangularSleeveShape(cz1) && helper.HasRectangularSleeveShape(cz2))
                {
                    return new CornerProximityChecker();
                }

                // DECISION 2: Both sleeves have circular opening (diameter, no corners) → edge-to-edge
                if (cz1 != null && cz2 != null && helper.HasCircularSleeveShape(cz1) && helper.HasCircularSleeveShape(cz2))
                {
                    return new EdgeToEdgeProximityChecker();
                }

                // ✅ DECISION 3: Mixed types (one rectangular, one circular) → use robust MixedTypeProximityChecker
                // This is specifically for the user's issue with bounding boxes on angled walls.
                if (cz1 != null && cz2 != null && 
                    ((helper.HasRectangularSleeveShape(cz1) && helper.HasCircularSleeveShape(cz2)) ||
                     (helper.HasCircularSleeveShape(cz1) && helper.HasRectangularSleeveShape(cz2))))
                {
                    return new MixedTypeProximityChecker();
                }

                // ✅ DECISION 4: Check for rotated sleeves → use rotated proximity checker
                if (isRotated && Math.Abs(rotationAngle) > 1e-6)
                {
                    // Rotated sleeves: Use rotated proximity checker with rotation angle
                    return new RotatedProximityChecker(rotationAngle);
                }

                // ✅ DECISION 5: Default fallback → use robust MixedTypeProximityChecker
                // It handles Floors (WCS) and Walls (RCS) more safely than the basic BoundingBox checker.
                return new MixedTypeProximityChecker();
            }
            catch (Exception ex)
            {
                // ✅ CRASH-SAFE: Return safe default on exception
                SafeFileLogger.SafeAppendText("geometry_errors.log",
                    $"[ProximityCheckerFactory] Exception in CreateChecker: {ex.Message}, StackTrace: {ex.StackTrace}");
                return new MixedTypeProximityChecker(); // 🏆 Preferred fallback is now MixedType (RCS-aware)
            }
        }
    }
}


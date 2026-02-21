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
                // ✅ CRASH-SAFE: Handle both wrapper objects (ClashZoneWorkItem) and direct ClashZone objects
                var helper = new SleeveCornerProximityHelper();
                Models.ClashZone cz1 = null;
                Models.ClashZone cz2 = null;

                // Try to get ClashZone from sleeve1
                if (sleeve1 is Models.ClashZone c1)
                {
                    cz1 = c1;
                }
                else if (sleeve1 != null)
                {
                    try { cz1 = sleeve1.ClashZone as Models.ClashZone; } catch { }
                }

                // Try to get ClashZone from sleeve2
                if (sleeve2 is Models.ClashZone c2)
                {
                    cz2 = c2;
                }
                else if (sleeve2 != null)
                {
                    try { cz2 = sleeve2.ClashZone as Models.ClashZone; } catch { }
                }

                // ==========================================================================================
                // 🛑 CRITICAL / DO NOT MOVE: ROTATION CHECK MUST BE FIRST! 🛑
                // ==========================================================================================
                // Reason: Rotated sleeves (especially on FLOORS like Cable Trays) are often "Rectangular".
                // If we check for Rectangular shape first (Decision 2), they will be routed to 
                // 'CornerProximityChecker', which assumes AXIS-ALIGNED geometry.
                // This causes rotated elements to fail clustering or crash.
                //
                // ALWAYS check for rotation (isRotated) before checking shape (Rectangular/Circular).
                // This ensures the specialized 'RotatedProximityChecker' is used, which handles
                // the rotation matrix and true dimensions (not inflated bounding boxes).
                // ==========================================================================================
                if (isRotated && Math.Abs(rotationAngle) > 1e-6)
                {
                    return new RotatedProximityChecker(rotationAngle);
                }

                // DECISION 2: Both sleeves have rectangular opening (corner geometry) → corner-based proximity
                if (cz1 != null && cz2 != null && helper.HasRectangularSleeveShape(cz1) && helper.HasRectangularSleeveShape(cz2))
                {
                    return new CornerProximityChecker();
                }

                // DECISION 3: Both sleeves have circular opening (diameter, no corners) → edge-to-edge
                if (cz1 != null && cz2 != null && helper.HasCircularSleeveShape(cz1) && helper.HasCircularSleeveShape(cz2))
                {
                    return new EdgeToEdgeProximityChecker();
                }

                // ✅ DECISION 4: Mixed types (one rectangular, one circular) → use MixedTypeProximityChecker
                if (cz1 != null && cz2 != null && 
                    ((helper.HasRectangularSleeveShape(cz1) && helper.HasCircularSleeveShape(cz2)) ||
                     (helper.HasCircularSleeveShape(cz1) && helper.HasRectangularSleeveShape(cz2))))
                {
                    SafeFileLogger.SafeAppendText("proximity_checker_selection.log",
                        $"[{DateTime.Now:HH:mm:ss}] MIXED: {cz1.ClashZoneGuid?.Substring(0,8)} ({cz1.MepElementCategory}) vs {cz2.ClashZoneGuid?.Substring(0,8)} ({cz2.MepElementCategory}) - Host={cz1.StructuralElementType}, Orient={cz1.HostOrientation}\n");
                    return new MixedTypeProximityChecker();
                }

                // ✅ DECISION 5: Default fallback → Bounding Box for others
                return new BoundingBoxProximityChecker();
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


using System;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Placement
{
    /// <summary>
    /// ✅ SOLID SRP: Service for determining sleeve rotation angle.
    /// Responsible ONLY for calculating correct rotation based on host type and MEP orientation.
    /// 
    /// Rotation logic:
    /// - FLOOR + CIRCULAR MEP (Pipes, Round Ducts): Always 0° (no rotation) - MEP orientation is meaningless for circular elements
    /// - FLOOR + RECTANGULAR MEP: Uses MepElementRotationAngle (pre-calculated during refresh)
    /// - X-WALL: Applies +90° rotation (π/2 radians) - regardless of MEP shape
    /// - Y-WALL: Applies 0° rotation (no rotation needed) - regardless of MEP shape
    /// - X-FRAMING: Applies +90° rotation (π/2 radians) - matches X-wall behavior
    /// - Y-FRAMING: Applies 0° rotation (no rotation needed) - matches Y-wall behavior
    /// 
    /// ✅ CRITICAL: For floor-hosted circular elements (pipes, round ducts), rotation is always 0°.
    ///    This avoids unnecessary MEP orientation detection and ensures straight placement.
    ///    Walls/framing still use MEP orientation for rotation (circular elements on walls still need rotation).
    /// 
    /// ✅ TESTABILITY: No Revit API calls, pure calculation logic
    /// ✅ REUSABILITY: Used by NewSleevePlacerService and clustering services
    /// </summary>
    public class SleeveRotationService
    {
        /// <summary>
        /// Determine sleeve rotation angle in radians based on host type and orientation.
        /// </summary>
        /// <param name="clashZone">Clash zone containing host type and orientation information</param>
        /// <returns>Rotation angle in radians (0, π/2, or pre-calculated angle)</returns>
        public double DetermineRotation(ClashZone clashZone)
        {
            // User requirement (2026‑03‑04): 
            // "for all cats and clusters no rotation needed" – families encode orientation.
            // So we always return 0.0 radians and let family choice control orientation.
            return 0.0;
        }
        
        /// <summary>
        /// Get rotation angle in degrees (for logging/debugging).
        /// </summary>
        public double GetRotationInDegrees(ClashZone clashZone)
        {
            return DetermineRotation(clashZone) * 180.0 / Math.PI;
        }
    }
}

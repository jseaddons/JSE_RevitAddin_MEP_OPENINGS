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
            if (clashZone == null)
                return 0.0; // Default: no rotation

            // ✅ WALL ROTATION: Based on wall orientation direction (X or Y)
            // ✅ CRITICAL FIX: Use HostOrientation (wall orientation), not MepElementOrientationDirection
            // X-WALL: Apply +90° rotation (π/2 radians)
            // Y-WALL: Apply 0° rotation (no rotation)
            if (clashZone.StructuralElementType?.Contains("Wall") == true)
            {
                // ✅ PRIORITY 1: Check HostOrientation (most reliable - set during refresh)
                string hostOrientation = (clashZone.HostOrientation ?? "").Trim();
                
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    SafeFileLogger.SafeAppendText("cluster_debug.log",
                           $"[{DateTime.Now:HH:mm:ss}] 🔍 WALL ROTATION CHECK: HostOri='{clashZone.HostOrientation}' (Trimmed='{hostOrientation}'), MepOri='{clashZone.MepElementOrientationDirection}'\n");
                }

                if (string.Equals(hostOrientation, "X", StringComparison.OrdinalIgnoreCase) || 
                    string.Equals(hostOrientation, "X-WALL", StringComparison.OrdinalIgnoreCase))
                {
                    return Math.PI / 2.0; // 90 degrees in radians for X-WALL
                }
                else if (string.Equals(hostOrientation, "Y", StringComparison.OrdinalIgnoreCase) || 
                         string.Equals(hostOrientation, "Y-WALL", StringComparison.OrdinalIgnoreCase))
                {
                    return 0.0; // No rotation for Y-WALL
                }
                
                // ✅ FALLBACK: Check MepElementOrientationDirection (if HostOrientation not set)
                // Note: This is a fallback - HostOrientation should be set during refresh
                if ((clashZone.MepElementOrientationDirection ?? "").Trim() == "X")
                {
                    return Math.PI / 2.0; // 90 degrees in radians for X-WALL
                }
                else if (clashZone.MepElementOrientationDirection == "Y")
                {
                    return 0.0; // No rotation for Y-WALL
                }
                
                // Default fallback for walls (if orientation cannot be determined)
                return 0.0;
            }
            
            // ⚠️⚠️⚠️ CRITICAL FIX: Structural framing needs rotation based on X vs Y orientation (same as walls)
            // X-FRAMING: Apply +90° rotation (π/2 radians) - matches X-wall behavior
            // Y-FRAMING: Apply 0° rotation (no rotation) - matches Y-wall behavior
            if (clashZone.StructuralElementType?.Contains("Structural Framing") == true)
            {
                string hostOrientation = (clashZone.HostOrientation ?? "").Trim();
                
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    SafeFileLogger.SafeAppendText("cluster_debug.log",
                           $"[{DateTime.Now:HH:mm:ss}] 🔍 FRAMING ROTATION CHECK: HostOri='{clashZone.HostOrientation}' (Trimmed='{hostOrientation}'), MepOri='{clashZone.MepElementOrientationDirection}'\n");
                }

                if (string.Equals(hostOrientation, "X", StringComparison.OrdinalIgnoreCase))
                {
                    return Math.PI / 2.0; // 90 degrees in radians for X-framing
                }
                else if (string.Equals(hostOrientation, "Y", StringComparison.OrdinalIgnoreCase))
                {
                    return 0.0; // No rotation for Y-framing
                }
                
                // Fallback: Check MepElementOrientationDirection (may be set for framing)
                if ((clashZone.MepElementOrientationDirection ?? "").Trim() == "X")
                {
                    return Math.PI / 2.0; // 90 degrees in radians for X-framing
                }
                else if ((clashZone.MepElementOrientationDirection ?? "").Trim() == "Y")
                {
                    return 0.0; // No rotation for Y-framing
                }

                if (!DeploymentConfiguration.DeploymentMode)
                {
                    SafeFileLogger.SafeAppendText("cluster_debug.log",
                           $"[{DateTime.Now:HH:mm:ss}] ⚠️ FRAMING ROTATION UNKNOWN: HostOri='{clashZone.HostOrientation}', MepOri='{clashZone.MepElementOrientationDirection}', using 0.0\n");
                }
                
                // Default fallback for framing (if orientation cannot be determined)
                return 0.0;
            }
            
            // ✅ FLOOR ROTATION: For circular MEP elements (pipes, round ducts), always place straight (no rotation)
            // MEP orientation is meaningless for circular elements - they have the same dimensions in all directions
            // This applies ONLY to floors - walls/framing still need rotation based on MEP orientation
            if (clashZone.StructuralElementType?.Contains("Floor") == true)
            {
                // ✅ CRITICAL: Check if MEP element is circular (pipe or round duct)
                bool isPipe = clashZone.MepElementCategory?.IndexOf("Pipe", StringComparison.OrdinalIgnoreCase) >= 0;
                bool isRoundDuct = clashZone.MepElementCategory?.IndexOf("Duct", StringComparison.OrdinalIgnoreCase) >= 0 &&
                                  (string.Equals(clashZone.DuctShape, "Round", StringComparison.OrdinalIgnoreCase) ||
                                   (clashZone.MepElementSizeData != null &&
                                    (string.Equals(clashZone.MepElementSizeData.Shape, "Round", StringComparison.OrdinalIgnoreCase) ||
                                     string.Equals(clashZone.MepElementSizeData.Shape, "Circular", StringComparison.OrdinalIgnoreCase))));
                
                // ✅ FOR FLOOR CIRCULAR ELEMENTS: Skip rotation calculation, place straight (0°)
                if (isPipe || isRoundDuct)
                {
                    return 0.0; // No rotation for circular elements on floors
                }
                
                // ✅ FOR FLOOR RECTANGULAR ELEMENTS: Use pre-calculated MepElementRotationAngle
                return clashZone.MepElementRotationAngle; // Already calculated during refresh
            }
            
            // Default: no rotation
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

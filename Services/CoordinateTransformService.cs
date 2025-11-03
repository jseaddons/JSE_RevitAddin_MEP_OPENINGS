using System;
using Autodesk.Revit.DB;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Centralized service for coordinate transformation operations.
    /// This eliminates code duplication across RefreshService and ThreePointValidator.
    /// </summary>
    public static class CoordinateTransformService
    {
        /// <summary>
        /// Transforms coordinates from a linked document to the active document coordinate system.
        /// Currently returns the point as-is (identity transformation).
        /// TODO: Implement proper coordinate transformation if needed for linked document workflows.
        /// </summary>
        /// <param name="point">The point to transform</param>
        /// <param name="linkedDocument">The linked document from which the point originates</param>
        /// <returns>The transformed point (currently returns the original point)</returns>
        public static XYZ TransformToActiveDocumentCoordinates(XYZ point, Document linkedDocument)
        {
            if (point == null)
                return XYZ.Zero;

            if (linkedDocument == null)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning("[CoordinateTransformService] Linked document is null, returning point as-is");
                return point;
            }

            try
            {
                // TODO: Implement proper coordinate transformation if needed
                // For now, return the point as-is since coordinate transformation is complex
                // This matches the current implementation in RefreshService and ThreePointValidator
                return point;
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[CoordinateTransformService] Error transforming coordinates: {ex.Message}");
                return point; // Return original point as fallback
            }
        }
    }
}


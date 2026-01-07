using System;
using Autodesk.Revit.DB;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Placement
{
    /// <summary>
    /// Result of pre-calculating sleeve placement data for a single zone.
    /// Contains all data needed for placement decision-making without Revit API calls.
    /// </summary>
    public class SleevePreCalculationResult
    {
        /// <summary>
        /// ClashZone identifier.
        /// </summary>
        public Guid ClashZoneId { get; set; }
        
        /// <summary>
        /// Whether the pre-calculation succeeded and the zone is valid for placement.
        /// </summary>
        public bool IsValid { get; set; }
        
        /// <summary>
        /// Error message if calculation failed.
        /// </summary>
        public string ErrorMessage { get; set; }
        
        /// <summary>
        /// Reason for skipping this zone (e.g., "AlreadyResolved", "ClusterResolved", "InvalidHostId").
        /// </summary>
        public string SkipReason { get; set; }
        
        /// <summary>
        /// MEP category (Ducts, Pipes, Cable Trays, Duct Accessories).
        /// </summary>
        public string Category { get; set; }
        
        // =====================================================
        // DIMENSION DATA
        // =====================================================
        
        /// <summary>
        /// Calculated target width in feet (includes clearance and insulation).
        /// </summary>
        public double TargetWidth { get; set; }
        
        /// <summary>
        /// Calculated target height in feet (includes clearance and insulation).
        /// </summary>
        public double TargetHeight { get; set; }
        
        /// <summary>
        /// Applied clearance in feet.
        /// </summary>
        public double ClearanceFeet { get; set; }
        
        /// <summary>
        /// Rotation angle in degrees.
        /// </summary>
        public double RotationAngleDeg { get; set; }
        
        // =====================================================
        // RISK CLASSIFICATION
        // =====================================================
        
        /// <summary>
        /// Risk classification based on clearance ratio.
        /// </summary>
        public ClearanceRiskClassification Risk { get; set; }
        
        // =====================================================
        // HOST VALIDATION
        // =====================================================
        
        /// <summary>
        /// Host structural element type (Wall, Floor, Structural Framing).
        /// </summary>
        public string HostType { get; set; }
        
        /// <summary>
        /// Whether the host element ID is valid.
        /// </summary>
        public bool HostExists { get; set; }
        
        /// <summary>
        /// Whether the host element type is valid.
        /// </summary>
        public bool HostIsValid { get; set; }
        
        // =====================================================
        // PLACEMENT POINT
        // =====================================================
        
        /// <summary>
        /// Placement point X coordinate.
        /// </summary>
        public double PlacementPointX { get; set; }
        
        /// <summary>
        /// Placement point Y coordinate.
        /// </summary>
        public double PlacementPointY { get; set; }
        
        /// <summary>
        /// Placement point Z coordinate.
        /// </summary>
        public double PlacementPointZ { get; set; }
        
        /// <summary>
        /// Whether the placement point is valid (not NaN, not zero).
        /// </summary>
        public bool PlacementPointIsValid { get; set; }
        
        /// <summary>
        /// Get placement point as XYZ.
        /// </summary>
        public XYZ GetPlacementPoint()
        {
            return new XYZ(PlacementPointX, PlacementPointY, PlacementPointZ);
        }
        
        /// <summary>
        /// Generate a log summary line for diagnostics.
        /// </summary>
        public string GetLogSummary()
        {
            if (!string.IsNullOrEmpty(SkipReason))
                return $"SKIP Zone={ClashZoneId} Reason={SkipReason}";
            
            if (!IsValid)
                return $"INVALID Zone={ClashZoneId} Error={ErrorMessage}";
            
            return $"PLAN Zone={ClashZoneId} Category={Category} W={TargetWidth * 304.8:F1}mm H={TargetHeight * 304.8:F1}mm " +
                   $"Clearance={ClearanceFeet * 304.8:F1}mm Risk={Risk} Host={HostType}";
        }
    }
}

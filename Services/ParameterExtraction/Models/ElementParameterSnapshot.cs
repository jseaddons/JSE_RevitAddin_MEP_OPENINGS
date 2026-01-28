using System.Collections.Generic;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.ParameterExtraction
{
    /// <summary>
    /// Unified snapshot of all parameters needed for sleeve placement.
    /// Contains pre-calculated values following "dump once, use many times" principle.
    /// </summary>
    public class ElementParameterSnapshot
    {
        #region MEP Dimensions (Critical for Sizing)

        /// <summary>MEP element width in feet</summary>
        public double? Width { get; set; }

        /// <summary>MEP element height in feet</summary>
        public double? Height { get; set; }

        /// <summary>Pipe outer diameter in feet</summary>
        public double? OuterDiameter { get; set; }

        /// <summary>Pipe nominal diameter in feet</summary>
        public double? NominalDiameter { get; set; }

        /// <summary>Size parameter as string (e.g., "20 mmø")</summary>
        public string? SizeParameterValue { get; set; }

        /// <summary>Formatted size string for display</summary>
        public string? FormattedSize { get; set; }

        #endregion

        #region Orientation (Critical for Floors)

        /// <summary>MEP orientation direction (X/Y)</summary>
        public string? OrientationDirection { get; set; }

        /// <summary>MEP orientation vector X component</summary>
        public double? OrientationX { get; set; }

        /// <summary>MEP orientation vector Y component</summary>
        public double? OrientationY { get; set; }

        /// <summary>MEP orientation vector Z component</summary>
        public double? OrientationZ { get; set; }

        /// <summary>Rotation angle in radians</summary>
        public double? RotationAngleRad { get; set; }

        /// <summary>Rotation angle in degrees</summary>
        public double? RotationAngleDeg { get; set; }

        /// <summary>Pre-calculated cos(angle) - avoids Math.Cos during placement</summary>
        public double? RotationCos { get; set; }

        /// <summary>Pre-calculated sin(angle) - avoids Math.Sin during placement</summary>
        public double? RotationSin { get; set; }

        /// <summary>Angle to X-axis in radians</summary>
        public double? AngleToXRad { get; set; }

        /// <summary>Angle to X-axis in degrees</summary>
        public double? AngleToXDeg { get; set; }

        /// <summary>Angle to Y-axis in radians</summary>
        public double? AngleToYRad { get; set; }

        /// <summary>Angle to Y-axis in degrees</summary>
        public double? AngleToYDeg { get; set; }

        #endregion

        #region Host/Structural Info

        /// <summary>Host orientation (X/Y/Z)</summary>
        public string? HostOrientation { get; set; }

        /// <summary>Structural type (Wall/Floor/StructuralFraming)</summary>
        public string? StructuralType { get; set; }

        /// <summary>Host thickness in feet</summary>
        public double? Thickness { get; set; }

        /// <summary>Wall-specific thickness in feet</summary>
        public double? WallThickness { get; set; }

        /// <summary>Framing-specific thickness in feet</summary>
        public double? FramingThickness { get; set; }

        #endregion

        #region Level Info

        /// <summary>Reference level name</summary>
        public string? LevelName { get; set; }

        /// <summary>Reference level elevation in feet</summary>
        public double? LevelElevation { get; set; }

        #endregion

        #region Damper-Specific

        /// <summary>Damper type name (MSD, MSFD, Standard, etc.)</summary>
        public string? TypeName { get; set; }

        /// <summary>Damper family name</summary>
        public string? FamilyName { get; set; }

        /// <summary>True if damper has connected MEP connector</summary>
        public bool HasMepConnector { get; set; }

        /// <summary>Connector side (+Y, -Y, Left, Right, etc.)</summary>
        public string? ConnectorSide { get; set; }

        /// <summary>True if standard Fire Damper (not MSFD/Motorized)</summary>
        public bool IsStandardDamper { get; set; }

        #endregion

        #region Insulation

        /// <summary>True if element is insulated</summary>
        public bool IsInsulated { get; set; }

        /// <summary>Insulation thickness in feet</summary>
        public double? InsulationThickness { get; set; }

        #endregion

        #region All Parameters (for transfer)

        /// <summary>All MEP parameters as key-value pairs</summary>
        public Dictionary<string, object> AllParameters { get; set; } = new();

        #endregion
    }
}

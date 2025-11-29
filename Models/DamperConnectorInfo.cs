using System;

namespace JSE_RevitAddin_MEP_OPENINGS.Models
{
    /// <summary>
    /// Immutable value object containing damper connector detection results.
    /// Represents complete information about a damper's connector presence and direction.
    /// Thread-safe and unit-testable.
    /// </summary>
    public class DamperConnectorInfo
    {
        /// <summary>
        /// Whether this damper has an MEP connector
        /// </summary>
        public bool HasMepConnector { get; }

        /// <summary>
        /// Connector side direction ("Left", "Right", "Top", "Bottom", or empty if not detected)
        /// </summary>
        public string ConnectorSide { get; }

        /// <summary>
        /// Whether this is a standard damper (uses symmetric clearance on all sides)
        /// </summary>
        public bool IsStandardDamper { get; }

        /// <summary>
        /// Damper type classification ("MSFD", "MSD", "MD", "Motorized", or "Standard")
        /// </summary>
        public string DamperType { get; }

        /// <summary>
        /// Whether this damper requires MEP-side clearance (non-standard types)
        /// </summary>
        public bool RequiresMepSideClearance { get; }

        /// <summary>
        /// Creates a new instance of DamperConnectorInfo.
        /// </summary>
        public DamperConnectorInfo(
            bool hasMepConnector,
            string connectorSide,
            bool isStandardDamper,
            string damperType,
            bool requiresMepSideClearance)
        {
            HasMepConnector = hasMepConnector;
            ConnectorSide = connectorSide ?? string.Empty;
            IsStandardDamper = isStandardDamper;
            DamperType = damperType ?? "Standard";
            RequiresMepSideClearance = requiresMepSideClearance;
        }

        /// <summary>
        /// Creates a default instance for non-damper elements or elements without connectors.
        /// </summary>
        public static DamperConnectorInfo None => new DamperConnectorInfo(
            hasMepConnector: false,
            connectorSide: string.Empty,
            isStandardDamper: false,
            damperType: "Standard",
            requiresMepSideClearance: false);

        public override string ToString()
        {
            if (!HasMepConnector)
                return $"No Connector (Type: {DamperType})";
            
            return $"Connector: {ConnectorSide} (Type: {DamperType}, Standard: {IsStandardDamper}, MEP Clearance: {RequiresMepSideClearance})";
        }
    }
}


namespace JSE_RevitAddin_MEP_OPENINGS.Services.DamperDetection
{
    /// <summary>
    /// Interface for detecting damper type from family/type name.
    /// Follows Interface Segregation Principle - focused interface for type detection only.
    /// </summary>
    public interface IDamperTypeDetector
    {
        /// <summary>
        /// Detects the damper type from the family type name.
        /// Returns: "MSFD", "MSD", "MD", "Motorized", or "Standard"
        /// </summary>
        string DetectDamperType(string familyTypeName);

        /// <summary>
        /// Determines if the damper type requires MEP-side clearance.
        /// Non-standard types (MSFD, MSD, MD, Motorized) require MEP-side clearance.
        /// </summary>
        bool RequiresMepSideClearance(string damperType);

        /// <summary>
        /// Determines if the damper type is standard (uses symmetric clearance on all sides).
        /// </summary>
        bool IsStandardDamper(string damperType);
    }
}


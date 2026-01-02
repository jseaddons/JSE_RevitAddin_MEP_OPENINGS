using System.Collections.Generic;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.ParameterExtraction
{
    /// <summary>
    /// Options for controlling parameter extraction behavior.
    /// </summary>
    public class ParameterExtractionOptions
    {
        /// <summary>
        /// If true, capture ALL parameters into AllParameters dictionary.
        /// If false, only capture essential parameters (Width, Height, etc.)
        /// Default: true
        /// </summary>
        public bool CaptureAllParameters { get; set; } = true;

        /// <summary>
        /// If true, pre-calculate rotation cos/sin values.
        /// Only needed for floor sleeves.
        /// Default: true
        /// </summary>
        public bool PreCalculateRotation { get; set; } = true;

        /// <summary>
        /// Specific parameter names to capture (if CaptureAllParameters is false).
        /// If null or empty, uses default essential parameters.
        /// </summary>
        public List<string>? ParameterWhitelist { get; set; }

        /// <summary>
        /// If true, capture host parameters (thickness, orientation).
        /// Applicable when extracting from host elements.
        /// Default: true
        /// </summary>
        public bool CaptureHostParameters { get; set; } = true;

        /// <summary>
        /// If true, capture level parameters (name, elevation).
        /// Default: true
        /// </summary>
        public bool CaptureLevelParameters { get; set; } = true;

        /// <summary>
        /// If true, capture damper-specific parameters (connector, type).
        /// Only applicable for damper elements.
        /// Default: true
        /// </summary>
        public bool CaptureDamperParameters { get; set; } = true;

        /// <summary>
        /// Default options for most use cases.
        /// </summary>
        public static ParameterExtractionOptions Default => new();

        /// <summary>
        /// Minimal options for performance-critical scenarios.
        /// Only captures essential dimensions.
        /// </summary>
        public static ParameterExtractionOptions Minimal => new()
        {
            CaptureAllParameters = false,
            PreCalculateRotation = false,
            CaptureHostParameters = false,
            CaptureLevelParameters = false,
            CaptureDamperParameters = false
        };

        /// <summary>
        /// Options optimized for damper extraction.
        /// </summary>
        public static ParameterExtractionOptions ForDamper => new()
        {
            CaptureAllParameters = true,
            PreCalculateRotation = false, // Dampers are in walls, not floors
            CaptureHostParameters = true,
            CaptureLevelParameters = true,
            CaptureDamperParameters = true
        };

        /// <summary>
        /// Options optimized for floor sleeve extraction.
        /// </summary>
        public static ParameterExtractionOptions ForFloor => new()
        {
            CaptureAllParameters = true,
            PreCalculateRotation = true, // Floors need rotation
            CaptureHostParameters = true,
            CaptureLevelParameters = true,
            CaptureDamperParameters = false
        };
    }
}

using Autodesk.Revit.DB;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Validation
{
    /// <summary>
    /// Result of a clash zone validation operation.
    /// Contains validation status and any updates needed to the clash zone.
    /// </summary>
    public class ValidationResult
    {
        /// <summary>
        /// True if the clash zone passed validation, false otherwise
        /// </summary>
        public bool IsValid { get; set; }

        /// <summary>
        /// Message describing why validation failed (if IsValid is false)
        /// </summary>
        public string FailureReason { get; set; }

        /// <summary>
        /// Updated intersection point if it changed during validation (null if no update needed)
        /// </summary>
        public XYZ UpdatedIntersectionPoint { get; set; }

        /// <summary>
        /// Updated active document coordinates if intersection point changed (null if no update needed)
        /// </summary>
        public XYZ UpdatedActiveDocumentPoint { get; set; }

        /// <summary>
        /// Distance the intersection point moved (in feet), if it changed
        /// </summary>
        public double? IntersectionPointMovement { get; set; }

        public ValidationResult()
        {
            IsValid = true;
            FailureReason = null;
            UpdatedIntersectionPoint = null;
            UpdatedActiveDocumentPoint = null;
            IntersectionPointMovement = null;
        }

        /// <summary>
        /// Creates a validation result indicating failure
        /// </summary>
        public static ValidationResult Invalid(string reason)
        {
            return new ValidationResult
            {
                IsValid = false,
                FailureReason = reason
            };
        }

        /// <summary>
        /// Creates a validation result indicating success with updated intersection point
        /// </summary>
        public static ValidationResult ValidWithUpdate(XYZ newIntersectionPoint, XYZ newActiveDocumentPoint, double movementDistance)
        {
            return new ValidationResult
            {
                IsValid = true,
                UpdatedIntersectionPoint = newIntersectionPoint,
                UpdatedActiveDocumentPoint = newActiveDocumentPoint,
                IntersectionPointMovement = movementDistance
            };
        }

        /// <summary>
        /// Creates a validation result indicating success with no updates needed
        /// </summary>
        public static ValidationResult Valid()
        {
            return new ValidationResult
            {
                IsValid = true
            };
        }
    }
}


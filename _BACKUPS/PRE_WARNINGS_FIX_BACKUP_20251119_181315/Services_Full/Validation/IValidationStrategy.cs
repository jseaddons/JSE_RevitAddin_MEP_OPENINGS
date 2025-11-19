using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Validation
{
    /// <summary>
    /// Interface for validation strategies following the Strategy pattern.
    /// Allows different validation approaches (3-Point, Flag Reset, etc.)
    /// </summary>
    public interface IValidationStrategy
    {
        /// <summary>
        /// Validates a clash zone and returns the validation result.
        /// </summary>
        /// <param name="clashZone">The clash zone to validate</param>
        /// <param name="document">The Revit document</param>
        /// <returns>ValidationResult indicating if the clash zone is valid and any updates needed</returns>
        ValidationResult Validate(ClashZone clashZone, Document document);
    }
}


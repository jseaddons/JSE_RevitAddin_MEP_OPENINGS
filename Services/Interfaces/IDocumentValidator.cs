using Autodesk.Revit.DB;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces
{
    /// <summary>
    /// Interface for validating document state before operations.
    /// Extracted from UniversalSleevePlacementCommand to adhere to SRP.
    /// </summary>
    public interface IDocumentValidator
    {
        /// <summary>
        /// Validate that the document can be modified.
        /// </summary>
        bool ValidateDocument(Document doc, string logPrefix);
    }
}


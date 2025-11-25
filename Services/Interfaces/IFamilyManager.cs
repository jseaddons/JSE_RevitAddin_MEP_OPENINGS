using Autodesk.Revit.DB;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces
{
    /// <summary>
    /// Interface for family loading and management.
    /// Extracted from UniversalSleevePlacerService to adhere to SRP.
    /// </summary>
    public interface IFamilyManager
    {
        /// <summary>
        /// Load a universal family into the document if not already present.
        /// </summary>
        bool LoadFamily(Document doc, string familyName);

        /// <summary>
        /// Get family name for a clash zone based on structural element type.
        /// </summary>
        string GetFamilyNameForClashZone(string hostType, string mepShape);

        /// <summary>
        /// Pre-cache family symbols for all clash zones.
        /// </summary>
        void PreCacheFamilySymbols(Document doc, System.Collections.Generic.List<JSE_RevitAddin_MEP_OPENINGS.Models.ClashZone> clashZones);
    }
}

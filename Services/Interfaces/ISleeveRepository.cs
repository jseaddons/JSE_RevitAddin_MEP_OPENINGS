using System.Collections.Generic;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces
{
    /// <summary>
    /// Interface for sleeve data persistence operations (XML/Database).
    /// Extracted from UniversalSleevePlacerService to adhere to SRP.
    /// </summary>
    public interface ISleeveRepository
    {
        /// <summary>
        /// Save sleeve data to _CLUSTER.xml file.
        /// </summary>
        void SaveSleeveDataToClusterXml(string xmlFilePath, List<ClashZone> clashZones, Document doc);

        /// <summary>
        /// Regenerate _CLUSTER.xml files with actual Revit coordinates after all sleeve placement completed.
        /// </summary>
        void RegenerateClusterXmlFiles(Document doc, List<ClashZone> placedZones);

        /// <summary>
        /// Load Duct Accessories clash zones from XML file for proximity checking.
        /// </summary>
        List<ClashZone> LoadDuctAccessoriesClashZones(string xmlFilePath);

        /// <summary>
        /// ✅ PRE-PLACEMENT PERSISTENCE: Save calculated sleeve data (dimensions, coordinates) to database BEFORE placement.
        /// This ensures data is saved even if placement fails or crashes. Uses ClashZoneGuid for lookup.
        /// </summary>
        void UpdateSleeveCalculatedData(ClashZone zone);
    }
}

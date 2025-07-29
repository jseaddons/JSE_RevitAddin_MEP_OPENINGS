using Autodesk.Revit.DB;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    public class FireDamperSleevePlacerDebug_v2
    {
        readonly Document _doc;

        public FireDamperSleevePlacerDebug_v2(Document doc)
        {
            _doc = doc;
        }

        public void PlaceFireDamperSleeveDebug(
            FamilyInstance accessory,
            XYZ intersection,
            double width,
            double height,
            XYZ accessoryDirection,
            FamilySymbol sleeveSymbol,
            Wall linkedWall)
        {
            // Copied current code as backing up as bk2
        }
    }
}

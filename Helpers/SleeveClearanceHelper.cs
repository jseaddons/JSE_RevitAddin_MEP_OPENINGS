using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Electrical;
using System.Linq;

namespace JSE_RevitAddin_MEP_OPENINGS.Helpers
{
    public static class SleeveClearanceHelper
    {
        // Returns clearance in internal units (feet)
        public static double GetClearance(Element mepElement)
        {
            bool hasInsulation = false;

            if (mepElement is Duct duct)
            {
                hasInsulation = new FilteredElementCollector(duct.Document)
                    .OfClass(typeof(DuctInsulation))
                    .Cast<DuctInsulation>()
                    .Any(ins => ins.HostElementId == duct.Id);
            }
            else if (mepElement is Pipe pipe)
            {
                hasInsulation = new FilteredElementCollector(pipe.Document)
                    .OfClass(typeof(PipeInsulation))
                    .Cast<PipeInsulation>()
                    .Any(ins => ins.HostElementId == pipe.Id);
            }
            else if (mepElement is CableTray tray)
            {
                // No insulation for cable tray in standard Revit API
                hasInsulation = false;
            }

            double clearanceMM = hasInsulation ? 25.0 : 50.0;
            return UnitUtils.ConvertToInternalUnits(clearanceMM, UnitTypeId.Millimeters);
        }
    }
}

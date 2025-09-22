using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Services.ClearanceProviders;
using System;
using System.IO;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Electrical;
using System.Linq;

namespace JSE_RevitAddin_MEP_OPENINGS.Helpers
{
    public static class SleeveClearanceHelper
    {
        // Returns clearance in internal units (feet) - gets UI clearance values
        public static double GetClearance(Element mepElement)
        {
            if (mepElement is Duct duct)
            {
                // Get UI clearance values from ClearanceManager
                var uiClearances = ClearanceManager.Instance.GetUIClearances();
                
                if (uiClearances != null && uiClearances.Count > 0)
                {
                    // Use UI clearance values
                    if (uiClearances.TryGetValue("ducts_normal_clearance", out double uiClearanceMM))
                    {
                        double clearanceInInternalUnits = UnitUtils.ConvertToInternalUnits(uiClearanceMM, UnitTypeId.Millimeters);
                        DebugLogger.Info($"[SleeveClearanceHelper] Using UI clearance for duct {duct.Id}: {uiClearanceMM}mm");
                        return clearanceInInternalUnits;
                    }
                }
                
                // Default if no UI clearances available
                double defaultClearance = UnitUtils.ConvertToInternalUnits(50.0, UnitTypeId.Millimeters);
                DebugLogger.Info($"[SleeveClearanceHelper] Using default clearance for duct {duct.Id}: 50mm");
                return defaultClearance;
            }
            else if (mepElement is Pipe pipe)
            {
                var insulation = new FilteredElementCollector(pipe.Document)
                    .OfClass(typeof(PipeInsulation))
                    .Cast<PipeInsulation>()
                    .FirstOrDefault(ins => ins.HostElementId == pipe.Id);
                if (insulation != null)
                {
                    double insulationThickness = insulation.get_Parameter(BuiltInParameter.RBS_PIPE_INSULATION_THICKNESS)?.AsDouble() ?? 0.0;
                    double clearanceMM = 25.0 + UnitUtils.ConvertFromInternalUnits(insulationThickness, UnitTypeId.Millimeters);
                    return UnitUtils.ConvertToInternalUnits(clearanceMM, UnitTypeId.Millimeters);
                }
            }
            // No insulation for cable tray in standard Revit API
            // Default: 50mm per side for non-insulated
            double defaultClearanceMM2 = 50.0;
            return UnitUtils.ConvertToInternalUnits(defaultClearanceMM2, UnitTypeId.Millimeters);
        }
    }
}







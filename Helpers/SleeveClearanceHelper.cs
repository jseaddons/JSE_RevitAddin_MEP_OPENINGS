using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using System;
using System.IO;
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
            if (mepElement is Duct duct)
            {
                double insulationThickness = 0.0;
                int ductId = (int)duct.Id.Value;
                string logDir = @"C:\JSE_CSharp_Projects\JSE_RevitAddin_MEP_OPENINGS\JSE_RevitAddin_MEP_OPENINGS\Log";
                if (!Directory.Exists(logDir)) Directory.CreateDirectory(logDir);
                string logFile = Path.Combine(logDir, $"SleeveClearanceHelperLog_{DateTime.Now:yyyyMMdd}.log");
                void Log(string msg) { if (JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.IsEnabled) File.AppendAllText(logFile, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {msg}{Environment.NewLine}"); }
                Log($"Processing Duct {ductId}");
                // Search for DuctInsulation in host document
                var insulation = new FilteredElementCollector(duct.Document)
                    .OfClass(typeof(DuctInsulation))
                    .Cast<DuctInsulation>()
                    .FirstOrDefault(ins => ins.HostElementId == duct.Id);
                if (insulation != null)
                {
                    Log($"Duct {ductId}: Found insulation in host doc, insulationId={insulation.Id.Value}");
                }
                // If not found, search in all visible linked documents
                if (insulation == null)
                {
                    var app = duct.Document.Application;
                    foreach (Document linkedDoc in app.Documents)
                    {
                        if (linkedDoc.IsLinked && !linkedDoc.IsFamilyDocument)
                        {
                            var linkIns = new FilteredElementCollector(linkedDoc)
                                .OfClass(typeof(DuctInsulation))
                                .Cast<DuctInsulation>()
                                .FirstOrDefault(ins => ins.HostElementId == duct.Id);
                            if (linkIns != null)
                            {
                                Log($"Duct {ductId}: Found insulation in linked doc, insulationId={linkIns.Id.Value}");
                            }
                            if (linkIns != null)
                            {
                                insulation = linkIns;
                                break;
                            }
                        }
                    }
                }
                if (insulation != null)
                {
                    // Fallback to brute force search for display value
                    foreach (Parameter param in insulation.Parameters)
                    {
                        string display = param.AsValueString();
                        if (!string.IsNullOrEmpty(display) && display.Contains("75"))
                        {
                            Log($"Duct {ductId}: FOUND 75: {param.Definition.Name} = {display}");
                        }
                        if (!string.IsNullOrEmpty(display) && display.Contains("mm"))
                        {
                            // Try to parse a number from the display string
                            string digits = new string(display.Where(c => char.IsDigit(c) || c == '.' || c == ',').ToArray());
                            if (double.TryParse(digits, out double parsedMM) && parsedMM > 0.0)
                            {
                                insulationThickness = UnitUtils.ConvertToInternalUnits(parsedMM, UnitTypeId.Millimeters);
                                Log($"Duct {ductId}: Parsed insulation thickness from display string: {display} -> {parsedMM} mm");
                                break;
                            }
                        }
                    }
                    // Geometric fallback: try to get 'Overall Size' and calculate thickness
                    if (insulationThickness == 0.0)
                    {
                        string overallSize = null;
                        foreach (Parameter param in insulation.Parameters)
                        {
                            if (param.Definition != null && param.Definition.Name.ToLower().Contains("overall size"))
                            {
                                overallSize = param.AsValueString();
                                break;
                            }
                        }
                        if (!string.IsNullOrEmpty(overallSize) && overallSize.Contains("mm"))
                        {
                            // Expecting format like "250 mm x 300 mm" or "250mmx300mm"
                            string[] dims = overallSize.Replace("mm", "").Replace("x", " ").Replace("X", " ").Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                            if (dims.Length == 2 && double.TryParse(dims[0].Replace(",", "."), out double overallW) && double.TryParse(dims[1].Replace(",", "."), out double overallH))
                            {
                                double ductW = 0.0, ductH = 0.0;
                                Parameter widthParam = insulation.Document.GetElement(insulation.HostElementId)?.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM);
                                Parameter heightParam = insulation.Document.GetElement(insulation.HostElementId)?.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM);
                                if (widthParam != null) ductW = UnitUtils.ConvertFromInternalUnits(widthParam.AsDouble(), UnitTypeId.Millimeters);
                                if (heightParam != null) ductH = UnitUtils.ConvertFromInternalUnits(heightParam.AsDouble(), UnitTypeId.Millimeters);
                                if (ductW > 0.0 && overallW > ductW)
                                {
                                    double thickW = (overallW - ductW) / 2.0;
                                    insulationThickness = UnitUtils.ConvertToInternalUnits(thickW, UnitTypeId.Millimeters);
                                    Log($"Duct {ductId}: Calculated insulation thickness from Overall Size (width): {overallW} - {ductW} / 2 = {thickW} mm");
                                }
                                if (insulationThickness == 0.0 && ductH > 0.0 && overallH > ductH)
                                {
                                    double thickH = (overallH - ductH) / 2.0;
                                    insulationThickness = UnitUtils.ConvertToInternalUnits(thickH, UnitTypeId.Millimeters);
                                    Log($"Duct {ductId}: Calculated insulation thickness from Overall Size (height): {overallH} - {ductH} / 2 = {thickH} mm");
                                }
                            }
                        }
                    }
                }
                else
                {
                    Log($"Duct {ductId}: No insulation found");
                }
                // Always return per-side clearance: insulation thickness (if any) + 25mm for insulated, or 50mm for non-insulated
                double baseClearance = UnitUtils.ConvertToInternalUnits(25.0, UnitTypeId.Millimeters);
                double clearance;
                if (insulationThickness > 0.0)
                {
                    // Insulated: insulation thickness + 25mm per side
                    clearance = insulationThickness + baseClearance;
                    Log($"Duct {ductId}: Returning clearance (insulated) = {UnitUtils.ConvertFromInternalUnits(clearance, UnitTypeId.Millimeters):F2} mm");
                    return clearance;
                }
                else
                {
                    // Non-insulated: 50mm per side
                    clearance = UnitUtils.ConvertToInternalUnits(50.0, UnitTypeId.Millimeters);
                    Log($"Duct {ductId}: Returning clearance (non-insulated) = 50.00 mm");
                    return clearance;
                }
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







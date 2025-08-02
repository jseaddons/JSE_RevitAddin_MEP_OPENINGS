using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Electrical;
using System;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Helpers
{
    public static class SleeveRiserOrientationHelper
    {
        /// <summary>
        /// Determines if a vertical duct or cable tray sleeve should be rotated 90 degrees based on bounding box and width/height parameters.
        /// Returns true if rotation is needed, false otherwise.
        /// </summary>
        public static bool ShouldRotateRiserSleeve(Element element, XYZ placePoint, double width, double height)
        {
            if (!JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.IsEnabled) return false;
            {
                BoundingBoxXYZ bbox = element.get_BoundingBox(null);
                string logFile = @"C:\JSE_CSharp_Projects\JSE_RevitAddin_MEP_OPENINGS\JSE_RevitAddin_MEP_OPENINGS\Log\riser_debug.log";
                using (var writer = new System.IO.StreamWriter(logFile, true)) // append mode
                {
                    string elementId = element?.Id?.Value.ToString() ?? "<null>";
                    if (bbox == null)
                    {
                        writer.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] ElementId={elementId} Bounding box is null, cannot determine orientation. No rotation applied.");
                        return false;
                    }
                    double dx = bbox.Max.X - bbox.Min.X;
                    double dy = bbox.Max.Y - bbox.Min.Y;
                    const double tol = 1e-3; // Loosened tolerance for real-world modeling
                    bool dxIsWidth = Math.Abs(dx - width) < tol;
                    bool dxIsHeight = Math.Abs(dx - height) < tol;
                    bool dyIsWidth = Math.Abs(dy - width) < tol;
                    bool dyIsHeight = Math.Abs(dy - height) < tol;
                    writer.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] ElementId={elementId} dx={dx}, dy={dy}, width={width}, height={height}");
                    string result = "";
                    if (Math.Abs(dx - dy) < tol)
                    {
                        result = "Bounding box is square, no rotation applied.";
                        writer.WriteLine(result);
                        return false;
                    }
                    else if (dxIsWidth && dyIsHeight)
                    {
                        result = "dx matches width, dy matches height, no rotation applied.";
                        writer.WriteLine(result);
                        return false;
                    }
                    else if (dxIsHeight && dyIsWidth)
                    {
                        result = "dx matches height, dy matches width, rotation needed.";
                        writer.WriteLine(result);
                        return true;
                    }
                    else
                    {
                        result = "Ambiguous bounding box/parameter match, no rotation applied.";
                        writer.WriteLine(result);
                        return false;
                    }
                }
            }
            }

        public static void LogRiserDebugInfo(string elementType, int elementId, BoundingBoxXYZ bbox, double width, double height, XYZ placePoint, Element hostElement, XYZ direction, bool shouldRotate)
        {
            if (!JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.IsEnabled) return;
            if (!JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.IsEnabled) return;
            {
                string logFile = @"C:\JSE_CSharp_Projects\JSE_RevitAddin_MEP_OPENINGS\JSE_RevitAddin_MEP_OPENINGS\Log\riser_debug.log";
                System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(logFile));
                // Convert all values to mm for clarity
                double minXmm = bbox?.Min.X * 304.8 ?? 0;
                double minYmm = bbox?.Min.Y * 304.8 ?? 0;
                double minZmm = bbox?.Min.Z * 304.8 ?? 0;
                double maxXmm = bbox?.Max.X * 304.8 ?? 0;
                double maxYmm = bbox?.Max.Y * 304.8 ?? 0;
                double maxZmm = bbox?.Max.Z * 304.8 ?? 0;
                double dxmm = (bbox?.Max.X - bbox?.Min.X ?? 0) * 304.8;
                double dymm = (bbox?.Max.Y - bbox?.Min.Y ?? 0) * 304.8;
                double widthmm = width * 304.8;
                double heightmm = height * 304.8;
                string msg = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {elementType} Id={elementId} | Host={hostElement?.GetType().Name} Id={hostElement?.Id} | Place=({placePoint.X:F3},{placePoint.Y:F3},{placePoint.Z:F3}) | Dir=({direction.X:F3},{direction.Y:F3},{direction.Z:F3}) | BBoxMM=Min({minXmm:F1},{minYmm:F1},{minZmm:F1}) Max({maxXmm:F1},{maxYmm:F1},{maxZmm:F1}) | dxMM={dxmm:F1} dyMM={dymm:F1} | {(elementType == "Duct" ? "DuctWidthMM" : "TrayWidthMM")}={widthmm:F1} {(elementType == "Duct" ? "DuctHeightMM" : "TrayHeightMM")}={heightmm:F1} | Rotation={(shouldRotate ? "True" : "False")}";
                if (DebugLogger.IsEnabled)
                {
                    System.IO.File.AppendAllText(logFile, msg + System.Environment.NewLine);
                    // Also log to MEP_Sleeve_Placement_Compare.log for direct comparison
                    string compareLogFile = @"C:\JSE_CSharp_Projects\JSE_RevitAddin_MEP_OPENINGS\JSE_RevitAddin_MEP_OPENINGS\Log\MEP_Sleeve_Placement_Compare.log";
                    System.IO.File.AppendAllText(compareLogFile, msg + System.Environment.NewLine);
                }
            }
        }
    }
}
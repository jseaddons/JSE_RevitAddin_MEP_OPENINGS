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
            BoundingBoxXYZ bbox = element.get_BoundingBox(null);
            if (bbox == null)
            {
                DebugLogger.Log("[SleeveRiserOrientationHelper] Bounding box is null, cannot determine orientation. No rotation applied.");
                return false;
            }
            double dx = bbox.Max.X - bbox.Min.X;
            double dy = bbox.Max.Y - bbox.Min.Y;
            const double tol = 1e-6;
            bool dxIsWidth = Math.Abs(dx - width) < tol;
            bool dxIsHeight = Math.Abs(dx - height) < tol;
            bool dyIsWidth = Math.Abs(dy - width) < tol;
            bool dyIsHeight = Math.Abs(dy - height) < tol;
            DebugLogger.Log($"[SleeveRiserOrientationHelper] dx={dx}, dy={dy}, width={width}, height={height}");
            if (Math.Abs(dx - dy) < tol)
            {
                DebugLogger.Log("[SleeveRiserOrientationHelper] Bounding box is square, no rotation applied.");
                return false;
            }
            else if (dxIsWidth && dyIsHeight)
            {
                DebugLogger.Log("[SleeveRiserOrientationHelper] dx matches width, dy matches height, no rotation applied.");
                return false;
            }
            else if (dxIsHeight && dyIsWidth)
            {
                DebugLogger.Log("[SleeveRiserOrientationHelper] dx matches height, dy matches width, rotation needed.");
                return true;
            }
            else
            {
                DebugLogger.Log("[SleeveRiserOrientationHelper] Ambiguous bounding box/parameter match, no rotation applied.");
                return false;
            }
        }

        public static void LogRiserDebugInfo(string elementType, int elementId, BoundingBoxXYZ bbox, double width, double height, XYZ placePoint, Element hostElement, XYZ direction, bool shouldRotate)
        {
            string logFile = @"C:\JSE_CSharp_Projects\JSE_CSharp_Projects\JSE_RevitAddin_MEP_OPENINGS\JSE_RevitAddin_MEP_OPENINGS\Log\riser_debug.log";
            System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(logFile));
            // TEMP: Unconditional log to confirm method is called
            System.IO.File.AppendAllText(logFile, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] CALLED LogRiserDebugInfo for {elementType} Id={elementId}\n");
            string msg = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {elementType} Id={elementId} | Host={hostElement?.GetType().Name} Id={hostElement?.Id} | Place=({placePoint.X:F3},{placePoint.Y:F3},{placePoint.Z:F3}) | Dir=({direction.X:F3},{direction.Y:F3},{direction.Z:F3}) | BBox=Min({bbox?.Min.X:F3},{bbox?.Min.Y:F3},{bbox?.Min.Z:F3}) Max({bbox?.Max.X:F3},{bbox?.Max.Y:F3},{bbox?.Max.Z:F3}) | dx={(bbox?.Max.X-bbox?.Min.X):F3} dy={(bbox?.Max.Y-bbox?.Min.Y):F3} | {(elementType=="Duct"?"DuctWidth":"TrayWidth")}={width:F3} {(elementType=="Duct"?"DuctHeight":"TrayHeight")}={height:F3} | Rotation={(shouldRotate ? "True" : "False")}";
            System.IO.File.AppendAllText(logFile, msg + System.Environment.NewLine);
        }
    }
} 
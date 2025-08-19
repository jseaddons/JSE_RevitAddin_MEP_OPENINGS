using Autodesk.Revit.DB;
using System;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    public static class SleeveRotationHelper
    {
        /// <summary>
        /// Aligns a sleeve instance on a floor to the MEP element's orientation.
        /// </summary>
        /// <param name="doc">Revit document</param>
        /// <param name="sleeveInstance">The sleeve FamilyInstance</param>
        /// <param name="placementPoint">Insertion point</param>
        /// <param name="bbox">Bounding box of the sleeve instance</param>
        /// <param name="elemWidth">Width of the duct/cable tray (internal units)</param>
        /// <param name="elemHeight">Height of the duct/cable tray (internal units)</param>
        /// <param name="elementId">Id of the duct/cable tray (for logging)</param>
        /// <returns>True if rotation was applied, false otherwise</returns>
        public static bool AlignSleeveToMepElement(
            Document doc,
            FamilyInstance sleeveInstance,
            XYZ placementPoint,
            BoundingBoxXYZ bbox,
            double elemWidth,
            double elemHeight,
            ElementId elementId)
        {
            double dx = bbox.Max.X - bbox.Min.X;
            double dy = bbox.Max.Y - bbox.Min.Y;
            double tol = 1e-6;

            DebugLogger.Log($"ElementId={elementId.IntegerValue} dx={dx}, dy={dy}, width={elemWidth}, height={elemHeight}");

            if (Math.Abs(dx - elemWidth) < tol && Math.Abs(dy - elemHeight) < tol)
            {
                DebugLogger.Log("dx matches width, dy matches height, no rotation needed.");
                return false;
            }
            else if (Math.Abs(dx - elemHeight) < tol && Math.Abs(dy - elemWidth) < tol)
            {
                DebugLogger.Log("dx matches height, dy matches width, rotation needed.");
                // Rotate 90 degrees about Z axis at placementPoint
                Line axis = Line.CreateBound(placementPoint, placementPoint + XYZ.BasisZ);
                ElementTransformUtils.RotateElement(doc, sleeveInstance.Id, axis, Math.PI / 2);
                return true;
            }
            else
            {
                DebugLogger.Log("Ambiguous bounding box/parameter match, no rotation applied.");
                return false;
            }
        }
    }
}

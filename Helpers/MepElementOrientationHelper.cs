using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Mechanical;
using System;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Helpers
{
    /// <summary>
    /// Helper class to determine MEP element width orientation using bounding box analysis
    /// </summary>
    public static class MepElementOrientationHelper
    {
        /// <summary>
        /// Determines if a cable tray's width runs along X or Y world axis using bounding box analysis
        /// </summary>
        public static (string orientation, XYZ widthDirection) GetCableTrayWidthOrientation(CableTray cableTray)
        {
            try
            {
                DebugLogger.Log($"[MepElementOrientationHelper] === CABLETRAY ORIENTATION ANALYSIS ===");
                
                // Get the cable tray's width parameter
                double widthParam = cableTray.get_Parameter(BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM)?.AsDouble() ?? 0;
                double widthParamMM = UnitUtils.ConvertFromInternalUnits(widthParam, UnitTypeId.Millimeters);
                DebugLogger.Log($"[MepElementOrientationHelper] CableTray Width Parameter: {widthParamMM:F1}mm");
                
                // Get the cable tray's bounding box
                BoundingBoxXYZ bbox = cableTray.get_BoundingBox(null);
                if (bbox == null)
                {
                    DebugLogger.Log($"[MepElementOrientationHelper] ERROR: Could not get bounding box for cable tray");
                    return ("Y-ORIENTED", XYZ.BasisY); // Default fallback
                }
                
                // Calculate bounding box dimensions
                double bboxWidth = bbox.Max.X - bbox.Min.X;  // X dimension
                double bboxHeight = bbox.Max.Y - bbox.Min.Y; // Y dimension  
                double bboxDepth = bbox.Max.Z - bbox.Min.Z;  // Z dimension
                
                double bboxWidthMM = UnitUtils.ConvertFromInternalUnits(bboxWidth, UnitTypeId.Millimeters);
                double bboxHeightMM = UnitUtils.ConvertFromInternalUnits(bboxHeight, UnitTypeId.Millimeters);
                double bboxDepthMM = UnitUtils.ConvertFromInternalUnits(bboxDepth, UnitTypeId.Millimeters);
                
                DebugLogger.Log($"[MepElementOrientationHelper] BoundingBox Dimensions:");
                DebugLogger.Log($"  - X dimension: {bboxWidthMM:F1}mm");
                DebugLogger.Log($"  - Y dimension: {bboxHeightMM:F1}mm");
                DebugLogger.Log($"  - Z dimension: {bboxDepthMM:F1}mm");
                
                // Find which bounding box dimension matches the width parameter (within tolerance)
                double tolerance = UnitUtils.ConvertToInternalUnits(10.0, UnitTypeId.Millimeters); // 10mm tolerance
                
                bool matchesXDimension = Math.Abs(widthParam - bboxWidth) < tolerance;
                bool matchesYDimension = Math.Abs(widthParam - bboxHeight) < tolerance;
                bool matchesZDimension = Math.Abs(widthParam - bboxDepth) < tolerance;
                
                DebugLogger.Log($"[MepElementOrientationHelper] Width Parameter Match Analysis:");
                DebugLogger.Log($"  - Matches X dimension: {matchesXDimension} (diff: {Math.Abs(widthParamMM - bboxWidthMM):F1}mm)");
                DebugLogger.Log($"  - Matches Y dimension: {matchesYDimension} (diff: {Math.Abs(widthParamMM - bboxHeightMM):F1}mm)");
                DebugLogger.Log($"  - Matches Z dimension: {matchesZDimension} (diff: {Math.Abs(widthParamMM - bboxDepthMM):F1}mm)");
                
                // Determine orientation based on which dimension the width parameter matches
                if (matchesXDimension)
                {
                    DebugLogger.Log($"[MepElementOrientationHelper] Width parameter matches X dimension → X-ORIENTED");
                    return ("X-ORIENTED", XYZ.BasisX);
                }
                else if (matchesYDimension)
                {
                    DebugLogger.Log($"[MepElementOrientationHelper] Width parameter matches Y dimension → Y-ORIENTED");
                    return ("Y-ORIENTED", XYZ.BasisY);
                }
                else if (matchesZDimension)
                {
                    DebugLogger.Log($"[MepElementOrientationHelper] Width parameter matches Z dimension → Z-ORIENTED (treating as Y-ORIENTED)");
                    return ("Y-ORIENTED", XYZ.BasisY); // Treat Z-oriented as Y-oriented for rotation
                }
                else
                {
                    DebugLogger.Log($"[MepElementOrientationHelper] WARNING: Width parameter doesn't match any dimension clearly. Using fallback Y-ORIENTED.");
                    return ("Y-ORIENTED", XYZ.BasisY); // Default fallback
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[MepElementOrientationHelper] ERROR in GetCableTrayWidthOrientation: {ex.Message}");
                return ("Y-ORIENTED", XYZ.BasisY); // Default fallback
            }
        }
        
        /// <summary>
        /// Determines if a duct's width runs along X or Y world axis using bounding box analysis
        /// </summary>
        public static (string orientation, XYZ widthDirection) GetDuctWidthOrientation(Duct duct)
        {
            try
            {
                DebugLogger.Log($"[MepElementOrientationHelper] === DUCT ORIENTATION ANALYSIS ===");
                
                // Get the duct's width parameter
                double widthParam = duct.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM)?.AsDouble() ?? 0;
                double widthParamMM = UnitUtils.ConvertFromInternalUnits(widthParam, UnitTypeId.Millimeters);
                DebugLogger.Log($"[MepElementOrientationHelper] Duct Width Parameter: {widthParamMM:F1}mm");
                
                // Get the duct's bounding box
                BoundingBoxXYZ bbox = duct.get_BoundingBox(null);
                if (bbox == null)
                {
                    DebugLogger.Log($"[MepElementOrientationHelper] ERROR: Could not get bounding box for duct");
                    return ("Y-ORIENTED", XYZ.BasisY); // Default fallback
                }
                
                // Calculate bounding box dimensions
                double bboxWidth = bbox.Max.X - bbox.Min.X;  // X dimension
                double bboxHeight = bbox.Max.Y - bbox.Min.Y; // Y dimension  
                double bboxDepth = bbox.Max.Z - bbox.Min.Z;  // Z dimension
                
                double bboxWidthMM = UnitUtils.ConvertFromInternalUnits(bboxWidth, UnitTypeId.Millimeters);
                double bboxHeightMM = UnitUtils.ConvertFromInternalUnits(bboxHeight, UnitTypeId.Millimeters);
                double bboxDepthMM = UnitUtils.ConvertFromInternalUnits(bboxDepth, UnitTypeId.Millimeters);
                
                DebugLogger.Log($"[MepElementOrientationHelper] BoundingBox Dimensions:");
                DebugLogger.Log($"  - X dimension: {bboxWidthMM:F1}mm");
                DebugLogger.Log($"  - Y dimension: {bboxHeightMM:F1}mm");
                DebugLogger.Log($"  - Z dimension: {bboxDepthMM:F1}mm");
                
                // Find which bounding box dimension matches the width parameter (within tolerance)
                double tolerance = UnitUtils.ConvertToInternalUnits(10.0, UnitTypeId.Millimeters); // 10mm tolerance
                
                bool matchesXDimension = Math.Abs(widthParam - bboxWidth) < tolerance;
                bool matchesYDimension = Math.Abs(widthParam - bboxHeight) < tolerance;
                bool matchesZDimension = Math.Abs(widthParam - bboxDepth) < tolerance;
                
                DebugLogger.Log($"[MepElementOrientationHelper] Width Parameter Match Analysis:");
                DebugLogger.Log($"  - Matches X dimension: {matchesXDimension} (diff: {Math.Abs(widthParamMM - bboxWidthMM):F1}mm)");
                DebugLogger.Log($"  - Matches Y dimension: {matchesYDimension} (diff: {Math.Abs(widthParamMM - bboxHeightMM):F1}mm)");
                DebugLogger.Log($"  - Matches Z dimension: {matchesZDimension} (diff: {Math.Abs(widthParamMM - bboxDepthMM):F1}mm)");
                
                // Determine orientation based on which dimension the width parameter matches
                if (matchesXDimension)
                {
                    DebugLogger.Log($"[MepElementOrientationHelper] Width parameter matches X dimension → X-ORIENTED");
                    return ("X-ORIENTED", XYZ.BasisX);
                }
                else if (matchesYDimension)
                {
                    DebugLogger.Log($"[MepElementOrientationHelper] Width parameter matches Y dimension → Y-ORIENTED");
                    return ("Y-ORIENTED", XYZ.BasisY);
                }
                else if (matchesZDimension)
                {
                    DebugLogger.Log($"[MepElementOrientationHelper] Width parameter matches Z dimension → Z-ORIENTED (treating as Y-ORIENTED)");
                    return ("Y-ORIENTED", XYZ.BasisY); // Treat Z-oriented as Y-oriented for rotation
                }
                else
                {
                    DebugLogger.Log($"[MepElementOrientationHelper] WARNING: Width parameter doesn't match any dimension clearly. Using fallback Y-ORIENTED.");
                    return ("Y-ORIENTED", XYZ.BasisY); // Default fallback
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[MepElementOrientationHelper] ERROR in GetDuctWidthOrientation: {ex.Message}");
                return ("Y-ORIENTED", XYZ.BasisY); // Default fallback
            }
        }
    }
}

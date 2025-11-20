using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Validation
{
    /// <summary>
    /// Validates clash zones using the 3-point validation strategy:
    /// 1. MEP element exists
    /// 2. Structural element exists
    /// 3. Elements still intersect
    /// </summary>
    public class ThreePointValidator : IValidationStrategy
    {
        /// <summary>
        /// Validates a clash zone and returns the validation result.
        /// Uses optimized fast-path: checks element existence first, only does expensive geometry if needed.
        /// </summary>
        public ValidationResult Validate(ClashZone clashZone, Document document)
        {
            if (clashZone == null)
            {
                return ValidationResult.Invalid("ClashZone is null");
            }

            if (document == null)
            {
                return ValidationResult.Invalid("Document is null");
            }

            try
            {
                // ✅ OPTIMIZATION: Fast-path check - verify element IDs from XML match current state
                // Point 1: Quick check if MEP Element exists (O(1) in host, O(L) in links - but fast lookup)
                var mepElement = ElementRetrievalService.GetElementFromDocumentOrLinked(document, clashZone.MepElementId, enableLogging: false);
                if (mepElement == null)
                {
                    return ValidationResult.Invalid($"MEP element {clashZone.MepElementId} not found");
                }

                // Point 2: Quick check if Structural Element exists (O(1) in host, O(L) in links - but fast lookup)
                var structuralElement = ElementRetrievalService.GetElementFromDocumentOrLinked(document, clashZone.StructuralElementId, enableLogging: false);
                if (structuralElement == null)
                {
                    return ValidationResult.Invalid($"Structural element {clashZone.StructuralElementId} not found");
                }

                // ✅ FAST-PATH OPTIMIZATION: Quick bounding box check before expensive geometry intersection
                // If intersection point from XML is within element bounding boxes, likely still valid
                var xmlIntersectionPoint = new XYZ(
                    clashZone.IntersectionPointX,
                    clashZone.IntersectionPointY,
                    clashZone.IntersectionPointZ
                );

                // Quick check: Is XML intersection point near both elements' bounding boxes?
                if (IsPointNearElementBoundingBox(xmlIntersectionPoint, mepElement) &&
                    IsPointNearElementBoundingBox(xmlIntersectionPoint, structuralElement))
                {
                    // ✅ FAST PATH: Elements exist and intersection point is likely still valid
                    // Skip expensive geometry calculation - assume XML data is correct
                    // Only do full validation if user requests or if point is far from bounding boxes
                    return ValidationResult.Valid(); // XML data matches - skip expensive check
                }

                // ✅ FALLBACK: If fast-path fails, do full geometry intersection (expensive but accurate)
                // This happens if intersection point moved significantly or bounding box check fails
                var newIntersectionPoint = CalculateIntersectionPoint(mepElement, structuralElement);
                if (newIntersectionPoint == null)
                {
                    return ValidationResult.Invalid("Elements no longer intersect");
                }

                // Check if intersection point needs updating
                var distance = xmlIntersectionPoint.DistanceTo(newIntersectionPoint);

                if (distance > 0.001) // If points differ by more than 1mm
                {
                // Calculate active document coordinates
                var activeDocPoint = CoordinateTransformService.TransformToActiveDocumentCoordinates(newIntersectionPoint, mepElement.Document);
                    
                    return ValidationResult.ValidWithUpdate(
                        newIntersectionPoint,
                        activeDocPoint,
                        distance
                    );
                }
                else
                {
                    // All valid, no update needed
                    return ValidationResult.Valid();
                }
            }
            catch (Exception ex)
            {
                return ValidationResult.Invalid($"Exception during validation: {ex.Message}");
            }
        }

        /// <summary>
        /// Quick check if a point is near an element's bounding box (fast approximation).
        /// Returns true if point is within tolerance of element's bounding box.
        /// </summary>
        private bool IsPointNearElementBoundingBox(XYZ point, Element element)
        {
            try
            {
                var bbox = element.get_BoundingBox(null);
                if (bbox == null)
                {
                    return false; // Can't verify without bounding box - require full check
                }

                // Expand bounding box by tolerance (1mm = ~0.003ft)
                const double tolerance = 0.003; // 1mm tolerance
                var expandedMin = new XYZ(
                    bbox.Min.X - tolerance,
                    bbox.Min.Y - tolerance,
                    bbox.Min.Z - tolerance
                );
                var expandedMax = new XYZ(
                    bbox.Max.X + tolerance,
                    bbox.Max.Y + tolerance,
                    bbox.Max.Z + tolerance
                );

                // Check if point is within expanded bounding box
                return point.X >= expandedMin.X && point.X <= expandedMax.X &&
                       point.Y >= expandedMin.Y && point.Y <= expandedMax.Y &&
                       point.Z >= expandedMin.Z && point.Z <= expandedMax.Z;
            }
            catch
            {
                // If bounding box check fails, require full validation
                return false;
            }
        }

        // ✅ OOP REFACTORING: Removed duplicate GetElementFromDocumentOrLinked - now uses ElementRetrievalService.GetElementFromDocumentOrLinked()

        /// <summary>
        /// Calculates the intersection point between MEP and structural elements.
        /// </summary>
        private XYZ CalculateIntersectionPoint(Element mepElement, Element structuralElement)
        {
            try
            {
                // Get geometry from both elements
                var mepGeometry = mepElement.get_Geometry(Helpers.GeometryOptionsFactory.CreateIntersectionOptions());
                var structuralGeometry = structuralElement.get_Geometry(Helpers.GeometryOptionsFactory.CreateIntersectionOptions());

                if (mepGeometry == null || structuralGeometry == null)
                {
                    return null;
                }

                // Find intersection points
                var intersectionPoints = new List<XYZ>();

                // Get MEP element line
                Line? line = null;
                foreach (GeometryObject geo in mepGeometry)
                {
                    if (geo is Curve curve)
                    {
                        line = curve as Line;
                        if (line != null) break;
                    }
                }

                if (line == null)
                {
                    return null;
                }

                // Get structural element solid and find intersections
                foreach (GeometryObject structGeo in structuralGeometry)
                {
                    if (structGeo is Solid structSolid)
                    {
                        foreach (Face face in structSolid.Faces)
                        {
                            if (face == null) continue;
                            
                            IntersectionResultArray? ira;
                            var res = face.Intersect(line, out ira);
                            
                            if (res == SetComparisonResult.Overlap && ira != null)
                            {
                                foreach (Autodesk.Revit.DB.IntersectionResult ir in ira)
                                {
                                    intersectionPoints.Add(GetIntersectionPointFromRevitResult(ir));
                                }
                                
                                if (intersectionPoints.Count > 0)
                                {
                                    // Return first intersection point
                                    return intersectionPoints[0];
                                }
                            }
                        }
                    }
                }

                return null;
            }
            catch (Exception ex)
            {
                DebugLogger.Warning($"[ThreePointValidator] Failed to calculate intersection point: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Helper method to get intersection point from Revit's IntersectionResult (handles both Point and XYZPoint properties)
        /// </summary>
        private static XYZ GetIntersectionPointFromRevitResult(Autodesk.Revit.DB.IntersectionResult ir)
        {
            try
            {
                // Try Point property first (Revit 2024+)
                var pointProperty = ir.GetType().GetProperty("Point");
                if (pointProperty != null)
                    return (XYZ)pointProperty.GetValue(ir);
            }
            catch { }
            
            try
            {
                // Try XYZPoint property (Revit 2020-2023)
                var xyzPointProperty = ir.GetType().GetProperty("XYZPoint");
                if (xyzPointProperty != null)
                    return (XYZ)xyzPointProperty.GetValue(ir);
            }
            catch { }
            
            throw new InvalidOperationException("Unable to get intersection point from IntersectionResult");
        }

        // ✅ OOP REFACTORING: Removed duplicate TransformToActiveDocumentCoordinates - now uses CoordinateTransformService.TransformToActiveDocumentCoordinates()
    }
}


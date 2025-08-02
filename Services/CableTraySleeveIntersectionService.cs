using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    public static class CableTraySleeveIntersectionService
    {
        /// <summary>
        /// Performs direct solid intersection check between a CableTray and structural elements,
        /// returning bounding boxes of intersecting structural elements and intersection points.
        /// </summary>
        public static List<(Element structuralElement, BoundingBoxXYZ bbox, XYZ intersectionPoint)> FindDirectStructuralIntersectionBoundingBoxesVisibleOnly(
            CableTray cableTray, List<(Element element, Transform? linkTransform)> structuralElements, Line cableTrayCenterline)
        {
            var intersections = new List<(Element, BoundingBoxXYZ, XYZ)>();

            if (cableTrayCenterline == null)
            {
                DebugLogger.Log($"[CableTraySleeveIntersectionService] Cable tray centerline is null for CableTray ID={cableTray.Id.Value}");
                return intersections;
            }

            foreach (var (structuralElement, linkTransform) in structuralElements)
            {
                try
                {
                    bool isLinkedElement = linkTransform != null;
                    var structuralOptions = new Options();
                    var structuralGeometry = structuralElement.get_Geometry(structuralOptions);
                    Solid? structuralSolid = null;

                    foreach (var geomObj in structuralGeometry)
                    {
                        if (geomObj is Solid solid && solid.Volume > 0)
                        {
                            structuralSolid = solid;
                            break;
                        }
                        else if (geomObj is GeometryInstance instance)
                        {
                            foreach (var instObj in instance.GetInstanceGeometry())
                            {
                                if (instObj is Solid instSolid && instSolid.Volume > 0)
                                {
                                    structuralSolid = instSolid;
                                    break;
                                }
                            }
                            if (structuralSolid != null) break;
                        }
                    }

                    if (structuralSolid == null)
                    {
                        DebugLogger.Log($"[CableTraySleeveIntersectionService] No solid geometry found for structural element ID={structuralElement.Id.Value}");
                        continue;
                    }

                    // Apply linkTransform if this is a linked element
                    if (isLinkedElement && linkTransform != null)
                    {
                        structuralSolid = SolidUtils.CreateTransformed(structuralSolid, linkTransform);
                    }

                    // Check intersection using face.Intersect(line, out ira)
                    var intersectionPoints = new List<XYZ>();
                    foreach (Face face in structuralSolid.Faces)
                    {
                        IntersectionResultArray? ira = null;
                        SetComparisonResult res = face.Intersect(cableTrayCenterline, out ira);
                        if (res == SetComparisonResult.Overlap && ira != null)
                        {
                            foreach (IntersectionResult ir in ira)
                            {
                                intersectionPoints.Add(ir.XYZPoint);
                            }
                        }
                    }

                    if (intersectionPoints.Count > 0)
                    {
                        // If two or more intersection points, use the midpoint between the two furthest apart (entry/exit)
                        XYZ intersectionPoint;
                        if (intersectionPoints.Count >= 2)
                        {
                            // Find the two points with the maximum distance between them
                            double maxDist = double.MinValue;
                            XYZ? ptA = null; XYZ? ptB = null;
                            for (int i = 0; i < intersectionPoints.Count - 1; i++)
                            {
                                for (int j = i + 1; j < intersectionPoints.Count; j++)
                                {
                                    double dist = intersectionPoints[i].DistanceTo(intersectionPoints[j]);
                                    if (dist > maxDist)
                                    {
                                        maxDist = dist;
                                        ptA = intersectionPoints[i];
                                        ptB = intersectionPoints[j];
                                    }
                                }
                            }
                            intersectionPoint = new XYZ((ptA.X + ptB.X) / 2, (ptA.Y + ptB.Y) / 2, (ptA.Z + ptB.Z) / 2);
                        }
                        else
                        {
                            // Only one intersection point, use as is
                            intersectionPoint = intersectionPoints[0];
                        }

                        intersections.Add((structuralElement, structuralSolid.GetBoundingBox(), intersectionPoint));
                    }
                }
                catch (Exception ex)
                {
                    DebugLogger.Log($"[CableTraySleeveIntersectionService] Error testing structural element ID={structuralElement.Id.Value}: {ex.Message}");
                }
            }

            return intersections;
        }
                /// <summary>
        /// Helper method to check if a point is within a bounding box
        /// </summary>
        private static bool IsPointInBoundingBox(XYZ point, BoundingBoxXYZ bbox)
        {
            return point.X >= bbox.Min.X && point.X <= bbox.Max.X &&
                   point.Y >= bbox.Min.Y && point.Y <= bbox.Max.Y &&
                   point.Z >= bbox.Min.Z && point.Z <= bbox.Max.Z;
        }
    }
        
}
using System;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Geometry
{
    /// <summary>
    /// ✅ SOLID COMPLIANCE (SRP): Interface for sleeve corner calculation service.
    /// </summary>
    public interface ISleeveCornerCalculationService
    {
        (XYZ corner1, XYZ corner2, XYZ corner3, XYZ corner4)? CalculateCorners(XYZ placementPoint, double width, double height, double rotationAngleRad);
        (XYZ corner1, XYZ corner2, XYZ corner3, XYZ corner4)? CalculateCornersFromZone(ClashZone zone, double width, double height);
    }

    /// <summary>
    /// ✅ SOLID COMPLIANCE (SRP): Service responsible for calculating sleeve corner coordinates.
    /// Pure mathematical operations - no database or Revit API dependencies.
    /// Can be called in parallel safely.
    /// </summary>
    public class SleeveCornerCalculationService : ISleeveCornerCalculationService
    {
        /// <summary>
        /// ✅ SRP: Calculates 4 corner coordinates in WORLD space from placement point, dimensions, and rotation.
        /// Pure math operation - thread-safe and can be parallelized.
        /// </summary>
        /// <param name="placementPoint">Placement point of the sleeve (center)</param>
        /// <param name="width">Width of the sleeve (in internal units, feet)</param>
        /// <param name="height">Height of the sleeve (in internal units, feet)</param>
        /// <param name="rotationAngleRad">Rotation angle in radians (around Z-axis)</param>
        /// <returns>Tuple of 4 corner coordinates (corner1, corner2, corner3, corner4) or null if invalid</returns>
        public (XYZ corner1, XYZ corner2, XYZ corner3, XYZ corner4)? CalculateCorners(
            XYZ placementPoint, double width, double height, double rotationAngleRad)
        {
            try
            {
                if (placementPoint == null)
                    return null;

                if (width <= 0 || height <= 0)
                    return null;

                // ✅ Calculate corner offsets in local coordinate system (before rotation)
                double halfWidth = width / 2.0;
                double halfHeight = height / 2.0;
                // Note: Z coordinate will be set to placementPoint.Z for all corners (2D opening)

                // ✅ CORNER ORDER (per methodology line 872): 1=Bottom-left, 2=Bottom-right, 3=Top-left, 4=Top-right
                // Local corner offsets (before rotation) - Z=0 in local space, will be translated to placementPoint.Z
                var localCorners = new[]
                {
                    new XYZ(-halfWidth, -halfHeight, 0),  // Corner 1: Bottom-left
                    new XYZ(halfWidth, -halfHeight, 0),    // Corner 2: Bottom-right
                    new XYZ(-halfWidth, halfHeight, 0),    // Corner 3: Top-left
                    new XYZ(halfWidth, halfHeight, 0)      // Corner 4: Top-right
                };

                // ✅ Apply rotation around Z-axis (if rotation angle is non-zero)
                var worldCorners = new XYZ[4];
                if (Math.Abs(rotationAngleRad) > 1e-6)
                {
                    // Rotation matrix for Z-axis rotation
                    double cos = Math.Cos(rotationAngleRad);
                    double sin = Math.Sin(rotationAngleRad);

                    for (int i = 0; i < 4; i++)
                    {
                        var local = localCorners[i];
                        // Rotate around Z-axis
                        double rotatedX = local.X * cos - local.Y * sin;
                        double rotatedY = local.X * sin + local.Y * cos;
                        // ✅ FIX: Set Z based on corner height (bottom vs top)
                        // Corners 0,1 are bottom (Y < 0), Corners 2,3 are top (Y > 0)
                        double cornerZ = placementPoint.Z + local.Y; // Y represents height offset in local space
                        // Translate to world coordinates
                        worldCorners[i] = new XYZ(
                            placementPoint.X + rotatedX,
                            placementPoint.Y + rotatedY,
                            cornerZ
                        );
                    }
                }
                else
                {
                    // No rotation - just translate to world coordinates
                    for (int i = 0; i < 4; i++)
                    {
                        var local = localCorners[i];
                        // ✅ FIX: Set Z based on corner height (bottom vs top)
                        // Corners 0,1 are bottom (Y < 0), Corners 2,3 are top (Y > 0)
                        double cornerZ = placementPoint.Z + local.Y; // Y represents height offset in local space
                        worldCorners[i] = new XYZ(
                            placementPoint.X + local.X,
                            placementPoint.Y + local.Y,
                            cornerZ
                        );
                    }
                }

                return (worldCorners[0], worldCorners[1], worldCorners[2], worldCorners[3]);
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Warning($"[SleeveCornerCalculationService] Error calculating corners: {ex.Message}");
                }
                return null;
            }
        }

        /// <summary>
        /// ✅ SRP: Calculates corners from a ClashZone (convenience method).
        /// </summary>
        public (XYZ corner1, XYZ corner2, XYZ corner3, XYZ corner4)? CalculateCornersFromZone(
            ClashZone zone, double width, double height)
        {
            if (zone == null || zone.SleevePlacementPoint == null)
                return null;

            return CalculateCorners(
                zone.SleevePlacementPoint,
                width,
                height,
                zone.MepElementRotationAngle);
        }

        /// <summary>
        /// ✅ REFACTORED: Calculates corners directly from the placed Revit FamilyInstance geometry.
        /// This ensures the stored corners match the actual physical element in the model,
        /// regardless of any calculation drift or parameter mismatches.
        /// REQUIRES MAIN THREAD ACCESS (Revit API).
        /// </summary>
        public (XYZ corner1, XYZ corner2, XYZ corner3, XYZ corner4)? CalculateCornersFromInstance(FamilyInstance sleeve)
        {
            try
            {
                if (sleeve == null || !sleeve.IsValidObject) return null;

                // 1. Get Geometry (ComputeReferences = true to ensure accuracy)
                Options opt = new Options { ComputeReferences = true, DetailLevel = ViewDetailLevel.Fine };
                GeometryElement geomElem = sleeve.get_Geometry(opt);

                if (geomElem == null) return null;

                Solid solid = null;

                foreach (GeometryObject obj in geomElem)
                {
                    if (obj is Solid s && s.Volume > 0)
                    {
                        solid = s;
                        break; // Use the first valid solid
                    }
                    else if (obj is GeometryInstance gi)
                    {
                        // Handle geometry inside instances (nested families)
                        foreach (GeometryObject obj2 in gi.SymbolGeometry)
                        {
                            if (obj2 is Solid s2 && s2.Volume > 0)
                            {
                                // We need to transform this solid to world coordinates
                                solid = SolidUtils.CreateTransformed(s2, gi.Transform);
                                break;
                            }
                        }
                    }
                    if (solid != null) break;
                }

                if (solid == null)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Warning($"[SleeveCornerCalculationService] ⚠️ No solid with volume found for sleeve {sleeve.Id}");
                    return null;
                }

                // ... (existing comments)

                Transform trf = sleeve.GetTransform();
                
                BoundingBoxXYZ localBox = null;
                foreach (GeometryObject obj in geomElem)
                {
                     if (obj is GeometryInstance gi)
                     {
                         // Symbol geometry is defined in local space!
                         localBox = GetGeometryBoundingBox(gi.SymbolGeometry);
                         if (localBox != null && !DeploymentConfiguration.DeploymentMode)
                         {
                             // LOG LOCALS
                             double w = localBox.Max.X - localBox.Min.X;
                             double h = localBox.Max.Y - localBox.Min.Y;
                             double d = localBox.Max.Z - localBox.Min.Z;
                             DebugLogger.Info($"[SleeveCornerCalculationService] 📏 Sleeve {sleeve.Id} Geometry Bounds (Local): " +
                                 $"X=[{localBox.Min.X:F4}, {localBox.Max.X:F4}] ({w:F4}), " +
                                 $"Y=[{localBox.Min.Y:F4}, {localBox.Max.Y:F4}] ({h:F4}), " +
                                 $"Z=[{localBox.Min.Z:F4}, {localBox.Max.Z:F4}] ({d:F4})");
                         }
                         break; // Assuming one main instance
                     }
                }
                
                if (localBox == null)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                         DebugLogger.Warning($"[SleeveCornerCalculationService] ⚠️ Failed to get local bounding box for sleeve {sleeve.Id}");
                    return null; 
                }
                
                // Now we have the exact local bounds of the geometry!
                // localBox.Min and localBox.Max give us the extent width/height/depth logic.
                // Corners in local space:
                // Z is usually depth (or Y depending on family).
                // Assuming standard MEP families (Up is Z, Facing is Y, Right is X).
                // Usually Width is X-axis, Height is Y-axis (or Z?).
                
                // Let's assume standard local bounds:
                XYZ p1_local = new XYZ(localBox.Min.X, localBox.Min.Y, localBox.Min.Z); // Min-Min
                XYZ p2_local = new XYZ(localBox.Max.X, localBox.Min.Y, localBox.Min.Z); // Max-Min
                XYZ p3_local = new XYZ(localBox.Min.X, localBox.Max.Y, localBox.Min.Z); // Min-Max
                XYZ p4_local = new XYZ(localBox.Max.X, localBox.Max.Y, localBox.Min.Z); // Max-Max
                
                // Transform to World
                XYZ c1 = trf.OfPoint(p1_local);
                XYZ c2 = trf.OfPoint(p2_local);
                XYZ c3 = trf.OfPoint(p3_local);
                XYZ c4 = trf.OfPoint(p4_local);
                
                return (c1, c2, c3, c4);
            }
            catch
            {
                return null;
            }
        }
        
        private BoundingBoxXYZ GetGeometryBoundingBox(GeometryElement geom)
        {
            double minX = double.MaxValue, minY = double.MaxValue, minZ = double.MaxValue;
            double maxX = double.MinValue, maxY = double.MinValue, maxZ = double.MinValue;
            bool found = false;
            
            foreach (GeometryObject obj in geom)
            {
                if (obj is Solid s && s.Volume > 0)
                {
                    BoundingBoxXYZ bbox = s.GetBoundingBox();
                    if (bbox != null)
                    {
                        minX = Math.Min(minX, bbox.Min.X); minY = Math.Min(minY, bbox.Min.Y); minZ = Math.Min(minZ, bbox.Min.Z);
                        maxX = Math.Max(maxX, bbox.Max.X); maxY = Math.Max(maxY, bbox.Max.Y); maxZ = Math.Max(maxZ, bbox.Max.Z);
                        found = true;
                    }
                }
            }
            
            if (!found) return null;
            
            return new BoundingBoxXYZ { Min = new XYZ(minX, minY, minZ), Max = new XYZ(maxX, maxY, maxZ) };
        }
    }
}


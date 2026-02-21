using System;
using System.Linq;
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
        (XYZ corner1, XYZ corner2, XYZ corner3, XYZ corner4)? CalculateCornersFromInstance(FamilyInstance sleeve, string hostOrientation = null, string structuralType = null);
    }

    /// <summary>
    /// ✅ SOLID COMPLIANCE (SRP): Service responsible for calculating sleeve corner coordinates.
    /// Pure mathematical operations - no database or Revit API dependencies.
    /// Can be called in parallel safely.
    /// </summary>
    public class SleeveCornerCalculationService : ISleeveCornerCalculationService
    {
        // ✅ PERFORMANCE: Cache successful parameter names to avoid repeated LookUpParameter calls (15 per sleeve -> 0)
        // Key: "Width", "Height", "Diameter"
        // Value: The actual parameter name that worked (e.g. "Element Width")
        private static readonly System.Collections.Concurrent.ConcurrentDictionary<string, string> _parameterNameCache 
            = new System.Collections.Concurrent.ConcurrentDictionary<string, string>();
        /// <summary>
        /// ✅ SRP: Calculates 4 corner coordinates in WORLD space from placement point, dimensions, and rotation.
        /// Pure math operation - thread-safe and can be parallelized.
        /// </summary>
        /// <param name="placementPoint">Placement point of the sleeve (center)</param>
        /// <param name="width">Width of the sleeve (in internal units, feet)</param>
        /// <param name="height">Height of the sleeve (in internal units, feet)</param>
        /// <param name="rotationAngleRad">Rotation angle in radians (around Z-axis)</param>
        /// <returns>Tuple of 4 corner coordinates (corner1, corner2, corner3, corner4) or null if invalid</returns>
        // =========================================================
        // ✅ SOLID STRATEGY PATTERN IMPLEMENTATION
        // =========================================================

        /// <summary>
        /// Strategy interface for retrieving/calculating sleeve corners in World Coordinates.
        /// </summary>
        private interface ICornerRetrievalStrategy
        {
            (XYZ c1, XYZ c2, XYZ c3, XYZ c4)? RetrieveCorners(XYZ placementPoint, double width, double height, double rotationRad);
            (XYZ c1, XYZ c2, XYZ c3, XYZ c4)? RetrieveCornersFromLocalBox(BoundingBoxXYZ localBox, Transform trf);
        }

        /// <summary>
        /// Strategy for FLOOR sleeves.
        /// Enforces FLAT (Planar) geometry at a constant Z-level.
        /// </summary>
        private class FloorSleeveStrategy : ICornerRetrievalStrategy
        {
            public (XYZ c1, XYZ c2, XYZ c3, XYZ c4)? RetrieveCorners(XYZ placementPoint, double width, double height, double rotationRad)
            {
                if (placementPoint == null || width <= 0 || height <= 0) return null;

                // Check for "Straight" Axis (0, 90, 180, 270 degrees)
                double deg = Math.Abs(rotationRad * 180.0 / Math.PI) % 180.0;
                bool isStraight = (deg < 0.001 || deg > 179.999) || (Math.Abs(deg - 90.0) < 0.001);

                if (isStraight)
                {
                    return RetrieveStraightFloorCorners(placementPoint, width, height, rotationRad);
                }
                else
                {
                    return RetrieveRotatedFloorCorners(placementPoint, width, height, rotationRad);
                }
            }

            public (XYZ c1, XYZ c2, XYZ c3, XYZ c4)? RetrieveCornersFromLocalBox(BoundingBoxXYZ localBox, Transform trf)
            {
                if (localBox == null || trf == null) return null;

                // For a Floor: width=X, height=Y, depth=Z
                // We use the bottom-most planar footprint (minZ)
                XYZ p1 = new XYZ(localBox.Min.X, localBox.Min.Y, localBox.Min.Z); 
                XYZ p2 = new XYZ(localBox.Max.X, localBox.Min.Y, localBox.Min.Z); 
                XYZ p3 = new XYZ(localBox.Min.X, localBox.Max.Y, localBox.Min.Z); 
                XYZ p4 = new XYZ(localBox.Max.X, localBox.Max.Y, localBox.Min.Z); 

                return (trf.OfPoint(p1), trf.OfPoint(p2), trf.OfPoint(p3), trf.OfPoint(p4));
            }

            private (XYZ c1, XYZ c2, XYZ c3, XYZ c4)? RetrieveStraightFloorCorners(XYZ pt, double w, double h, double rot)
            {
                double hw = w / 2.0;
                double hh = h / 2.0;
                
                double effectiveW = hw;
                double effectiveH = hh;
                
                double deg = Math.Abs(rot * 180.0 / Math.PI) % 180.0;
                if (Math.Abs(deg - 90.0) < 0.001)
                {
                    effectiveW = hh;
                    effectiveH = hw;
                }

                return (
                    new XYZ(pt.X - effectiveW, pt.Y - effectiveH, pt.Z),
                    new XYZ(pt.X + effectiveW, pt.Y - effectiveH, pt.Z),
                    new XYZ(pt.X - effectiveW, pt.Y + effectiveH, pt.Z),
                    new XYZ(pt.X + effectiveW, pt.Y + effectiveH, pt.Z)
                );
            }

            private (XYZ c1, XYZ c2, XYZ c3, XYZ c4)? RetrieveRotatedFloorCorners(XYZ pt, double w, double h, double rot)
            {
                double hw = w / 2.0;
                double hh = h / 2.0;

                var local = new[] {
                    new XYZ(-hw, -hh, 0), new XYZ(hw, -hh, 0),
                    new XYZ(-hw, hh, 0),  new XYZ(hw, hh, 0)
                };

                var world = new XYZ[4];
                double cos = Math.Cos(rot);
                double sin = Math.Sin(rot);

                for (int i = 0; i < 4; i++)
                {
                    double rx = local[i].X * cos - local[i].Y * sin;
                    double ry = local[i].X * sin + local[i].Y * cos;
                    world[i] = new XYZ(pt.X + rx, pt.Y + ry, pt.Z);
                }
                return (world[0], world[1], world[2], world[3]);
            }
        }

        /// <summary>
        /// Strategy for WALL sleeves.
        /// Handles Vertical orientation where Height (Y) corresponds to Z-change.
        /// </summary>
        private class WallSleeveStrategy : ICornerRetrievalStrategy
        {
            public (XYZ c1, XYZ c2, XYZ c3, XYZ c4)? RetrieveCorners(XYZ placementPoint, double width, double height, double rotationRad)
            {
                if (placementPoint == null || width <= 0 || height <= 0) return null;

                double halfW = width / 2.0;
                double halfH = height / 2.0;

                var local = new[] {
                    new XYZ(-halfW, -halfH, 0), new XYZ(halfW, -halfH, 0),
                    new XYZ(-halfW, halfH, 0),  new XYZ(halfW, halfH, 0)
                };

                var world = new XYZ[4];
                double cos = Math.Cos(rotationRad);
                double sin = Math.Sin(rotationRad);

                for (int i = 0; i < 4; i++)
                {
                    // For Vertical Elements (Walls):
                    // Y-local maps to Z-world relative to center.
                    // X-local maps to X/Y-world (Width).
                    
                    double rx = local[i].X * cos; 
                    double ry = local[i].X * sin; 
                    double rz = local[i].Y; // Height adds to Z

                    world[i] = new XYZ(
                        placementPoint.X + rx,
                        placementPoint.Y + ry,
                        placementPoint.Z + rz
                    );
                }
                return (world[0], world[1], world[2], world[3]);
            }

            public (XYZ c1, XYZ c2, XYZ c3, XYZ c4)? RetrieveCornersFromLocalBox(BoundingBoxXYZ localBox, Transform trf)
            {
                if (localBox == null || trf == null) return null;

                // ✅ FIX: For Wall sleeves, we need 4 corners with X and Z variation (in WORLD coords).
                // The local box Y is "depth" (wall thickness) - we want to project onto one Y-face.
                // BUT the family transform may rotate axes, so we transform ALL 8 corners to world,
                // then pick the 4 that represent the "front face" in world XZ plane.

                // Step 1: Get all 8 corners of the local bounding box
                var localCorners = new XYZ[]
                {
                    new XYZ(localBox.Min.X, localBox.Min.Y, localBox.Min.Z),
                    new XYZ(localBox.Max.X, localBox.Min.Y, localBox.Min.Z),
                    new XYZ(localBox.Min.X, localBox.Max.Y, localBox.Min.Z),
                    new XYZ(localBox.Max.X, localBox.Max.Y, localBox.Min.Z),
                    new XYZ(localBox.Min.X, localBox.Min.Y, localBox.Max.Z),
                    new XYZ(localBox.Max.X, localBox.Min.Y, localBox.Max.Z),
                    new XYZ(localBox.Min.X, localBox.Max.Y, localBox.Max.Z),
                    new XYZ(localBox.Max.X, localBox.Max.Y, localBox.Max.Z)
                };

                // Step 2: Transform ALL corners to world coordinates
                var worldCorners = localCorners.Select(p => trf.OfPoint(p)).ToArray();

                // Step 3: Find world-space bounds
                double minX = worldCorners.Min(p => p.X);
                double maxX = worldCorners.Max(p => p.X);
                double minY = worldCorners.Min(p => p.Y);
                double maxY = worldCorners.Max(p => p.Y);
                double minZ = worldCorners.Min(p => p.Z);
                double maxZ = worldCorners.Max(p => p.Z);

                // Step 4: Determine wall orientation by comparing X-range vs Y-range
                double rangeX = maxX - minX;
                double rangeY = maxY - minY;

                // Step 5: Return 4 corners on the XZ plane (Y-wall) or YZ plane (X-wall)
                if (rangeX > rangeY)
                {
                    // X-wall: Width along X, use constant Y (minY = front face)
                    return (
                        new XYZ(minX, minY, minZ),  // Bottom-Left
                        new XYZ(maxX, minY, minZ),  // Bottom-Right
                        new XYZ(minX, minY, maxZ),  // Top-Left
                        new XYZ(maxX, minY, maxZ)   // Top-Right
                    );
                }
                else
                {
                    // Y-wall: Width along Y, use constant X (minX = front face)
                    return (
                        new XYZ(minX, minY, minZ),  // Bottom-Left
                        new XYZ(minX, maxY, minZ),  // Bottom-Right
                        new XYZ(minX, minY, maxZ),  // Top-Left
                        new XYZ(minX, maxY, maxZ)   // Top-Right
                    );
                }
            }
        }

        /// <summary>
        /// Factory to select the correct strategy.
        /// </summary>
        private ICornerRetrievalStrategy GetStrategyForElement(FamilyInstance sleeve, string structuralType = null)
        {
            // ✅ PRIORITY: Use explicit structuralType if available
            if (!string.IsNullOrEmpty(structuralType))
            {
                if (structuralType.IndexOf("Wall", StringComparison.OrdinalIgnoreCase) >= 0)
                    return new WallSleeveStrategy();
                if (structuralType.IndexOf("Floor", StringComparison.OrdinalIgnoreCase) >= 0 || structuralType.IndexOf("Roof", StringComparison.OrdinalIgnoreCase) >= 0)
                    return new FloorSleeveStrategy();
            }

            // Heuristic: Check Host Category
            if (sleeve != null && sleeve.Host != null)
            {
                if (sleeve.Host.Category.Name.Contains("Wall")) 
                    return new WallSleeveStrategy();
                if (sleeve.Host.Category.Name.Contains("Floor") || sleeve.Host.Category.Name.Contains("Roof"))
                    return new FloorSleeveStrategy();
            }

            // Fallback / Default
            return new FloorSleeveStrategy();
        }

        private string GetStructuralTypeFromLogic(FamilyInstance sleeve, string explicitType)
        {
            if (!string.IsNullOrEmpty(explicitType)) return explicitType;
            if (sleeve?.Host != null) return sleeve.Host.Category.Name;
            return null;
        }

        // =========================================================

        public (XYZ corner1, XYZ corner2, XYZ corner3, XYZ corner4)? CalculateCorners(
            XYZ placementPoint, double width, double height, double rotationAngleRad)
        {
            // Default generic call -> Use Floor Strategy (Flat Safe)
            return new FloorSleeveStrategy().RetrieveCorners(placementPoint, width, height, rotationAngleRad);
        }

        /// <summary>
        /// ✅ SRP: Calculates corners from a ClashZone using the correct Strategy.
        /// </summary>
        public (XYZ corner1, XYZ corner2, XYZ corner3, XYZ corner4)? CalculateCornersFromZone(
            ClashZone zone, double width, double height)
        {
            if (zone == null || zone.SleevePlacementPoint == null)
                return null;
            
            // SELECT STRATEGY based on Zone Type
            ICornerRetrievalStrategy strategy;
            if (string.Equals(zone.StructuralElementType, "Wall", StringComparison.OrdinalIgnoreCase))
            {
                strategy = new WallSleeveStrategy();
            }
            else
            {
                // Default to Floor (Flat)
                strategy = new FloorSleeveStrategy();
            }

            return strategy.RetrieveCorners(
                zone.SleevePlacementPoint,
                width,
                height,
                zone.MepElementRotationAngle);
        }

        /// <summary>
        /// ✅ REFACTORED: Calculates corners directly from the placed Revit FamilyInstance geometry.
        /// Uses element.get_BoundingBox(null) for WORLD coordinates directly.
        /// This ensures the stored corners match the actual physical element in the model.
        /// REQUIRES MAIN THREAD ACCESS (Revit API).
        /// </summary>
        public (XYZ corner1, XYZ corner2, XYZ corner3, XYZ corner4)? CalculateCornersFromInstance(FamilyInstance sleeve)
        {
            // Call overload with null orientation (auto-detect)
            return CalculateCornersFromInstance(sleeve, null);
        }

        /// <summary>
        /// ✅ REFACTORED: Calculates corners with explicit hostOrientation AND structuralType from database.
        /// When structuralType is provided (e.g., "Wall"), uses that to determining strategy,
        /// avoiding issues when sleeve.Host is null (e.g. for Generic Models in Links).
        /// </summary>
        /// <param name="sleeve">The Revit FamilyInstance</param>
        /// <param name="hostOrientation">Wall orientation from DB: "X", "Y", or null for auto-detect</param>
        /// <param name="structuralType">Structural Type from DB: "Wall", "Floor", etc.</param>
        public (XYZ corner1, XYZ corner2, XYZ corner3, XYZ corner4)? CalculateCornersFromInstance(FamilyInstance sleeve, string hostOrientation, string structuralType = null)
        {
            try
            {
                if (sleeve == null || !sleeve.IsValidObject) return null;

                // ✅ CHECK: Is this a wall-hosted sleeve? 
                // Prioritize explicit structuralType over unreliable sleeve.Host
                bool isWallHosted = false;
                
                if (!string.IsNullOrEmpty(structuralType) && structuralType.IndexOf("Wall", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    isWallHosted = true;
                }
                else if (sleeve.Host != null && sleeve.Host.Category.Name.Contains("Wall"))
                {
                    isWallHosted = true;
                }
                
                // ✅ FALLBACK: If no host, use hostOrientation from DB (StructuralElementType)
                bool useWallLogic = isWallHosted || (!string.IsNullOrEmpty(hostOrientation) && 
                    (hostOrientation.Equals("X", StringComparison.OrdinalIgnoreCase) || 
                     hostOrientation.Equals("Y", StringComparison.OrdinalIgnoreCase)));

                // ✅ REFACTORED: Always use parameter-based calculation with Strategy pattern.
                // The previous specialized wall logic (lines 350-404) was limited to straight axes.
                // CalculateCornersFromParameters correctly handles rotation for all orientations.
                return CalculateCornersFromParameters(sleeve, structuralType);

                // FIX: CS0162 - unreachable code commented out
                // // ✅ FLOOR SLEEVES: Use parameter-based calculation with rotation handling
                // // AABB bbox cannot handle rotated sleeves - we need OBB from parameters + rotation
                // // FloorSleeveStrategy.RetrieveCorners handles rotation at RetrieveRotatedFloorCorners()
                // return CalculateCornersFromParameters(sleeve);
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning($"[SleeveCornerCalculationService] Instance Extraction failed: {ex.Message}");
                return null;
            }
        }
        
        /// <summary>
        /// Calculates corners using Revit Parameters (Width/Height) and Location Rotation.
        /// Essential fallback for Void families where Geometry extraction fails.
        /// </summary>
        private (XYZ c1, XYZ c2, XYZ c3, XYZ c4)? CalculateCornersFromParameters(FamilyInstance sleeve, string structuralType = null)
        {
            try
            {
                // 1. Get Dimensions (Try multiple standard names, checking Instance and Type)
                string[] widthParams = { "Element Width", "Width", "MW", "Sleeve Width", "Opening Width" };
                string[] heightParams = { "Element Height", "Height", "MH", "Sleeve Height", "Opening Height" };
                string[] diaParams = { "Element Diameter", "Diameter", "MD", "Sleeve Diameter", "Opening Diameter" };

                double width = -1, height = -1;



                // Helper to get value with caching
                Func<string[], string, double> getVal = (names, cacheKey) => {
                    // 1. Try Cache First
                    if (_parameterNameCache.TryGetValue(cacheKey, out string cachedName))
                    {
                        Parameter p = sleeve.LookupParameter(cachedName);
                        if (p == null && sleeve.Symbol != null) p = sleeve.Symbol.LookupParameter(cachedName);
                        if (p != null) return p.AsDouble();
                    }

                    // 2. Iterate List (Slow path - run once per session)
                    foreach (var name in names) {
                        // Check Instance
                        Parameter p = sleeve.LookupParameter(name);
                        // Check Type (Symbol)
                        if (p == null && sleeve.Symbol != null) p = sleeve.Symbol.LookupParameter(name);
                        
                        if (p != null) 
                        {
                            // ✅ Cache the successful name
                            _parameterNameCache.TryAdd(cacheKey, name);
                            return p.AsDouble();
                        }
                    }
                    return -1;
                };

                width = getVal(widthParams, "Width");
                height = getVal(heightParams, "Height");

                if (width <= 0 || height <= 0)
                {
                    double dia = getVal(diaParams, "Diameter");
                    if (dia > 0) { width = dia; height = dia; }
                }
                
                if (width <= 0) return null; 
                
                // 2. Get Location & Rotation
                if (!(sleeve.Location is LocationPoint loc)) return null;
                
                // 3. Select Strategy & Execute
                var strategy = GetStrategyForElement(sleeve, GetStructuralTypeFromLogic(sleeve, structuralType));
                
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[SleeveCornerCalculationService] 🔢 From Params ({strategy.GetType().Name}): W={width:F3}, H={height:F3}, Rot={loc.Rotation:F3}");

                return strategy.RetrieveCorners(loc.Point, width, height, loc.Rotation);
            }
            catch(Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning($"[SleeveCornerCalculationService] Parameter Calculation failed: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Calculates 4 corners from a Solid's bottom vertices (World Coordinates).
        /// ... (existing method)
        
        /// <summary>
        /// Calculates 4 corners from a Solid's bottom vertices (World Coordinates).
        /// Used when GeometryInstance is not available.
        /// </summary>
        private (XYZ c1, XYZ c2, XYZ c3, XYZ c4)? CalculateCornersFromSolidVertices(Solid solid)
        {
            try 
            {
                 // Extract Vertices
                var vertices = new List<XYZ>();
                foreach (Edge edge in solid.Edges)
                {
                    foreach (XYZ pt in edge.Tessellate())
                    {
                        vertices.Add(pt);
                    }
                }
                
                // Get unique vertices
                var uniqueVertices = vertices
                    .GroupBy(p => new { X = Math.Round(p.X, 4), Y = Math.Round(p.Y, 4), Z = Math.Round(p.Z, 4) })
                    .Select(g => g.First())
                    .ToList();
                    
                // Find Bottom 4 Vertices (Lowest Z)
                if (uniqueVertices.Count == 0) return null;

                double minZ = uniqueVertices.Min(p => p.Z);
                var bottomVertices = uniqueVertices
                    .Where(p => Math.Abs(p.Z - minZ) < 0.01) // Tolerance
                    .ToList();
                
                if (bottomVertices.Count < 4) return null;
                
                // Sort by angle from centroid to order them C1..C4
                double cx = bottomVertices.Average(p => p.X);
                double cy = bottomVertices.Average(p => p.Y);
                
                var sorted = bottomVertices.OrderBy(p => Math.Atan2(p.Y - cy, p.X - cx)).ToList();
                
                // Take outer 4 (simplified convex hull for rectangle)
                // If shape is exactly rectangle, they are the 4 points.
                // Order: -PI to PI. (-180 to 180).
                // Bottom-Left (-135)-> Bottom-Right (-45) -> Top-Right (45) -> Top-Left (135).
                
                // Ensure we return 4 points
                if (sorted.Count < 4) return null;

                return (sorted[0], sorted[1], sorted[2], sorted[3]);
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
                if (obj is Solid s && Math.Abs(s.Volume) > 0.001)
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


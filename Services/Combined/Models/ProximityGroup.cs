using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Combined.Models
{
    /// <summary>
    /// Represents a group of sleeves that are in proximity to each other
    /// and should be combined into a single combined sleeve.
    /// 
    /// A proximity group contains sleeves from different categories
    /// (e.g., Pipes + Ducts) that are within the proximity threshold.
    /// </summary>
    public class ProximityGroup
    {
        /// <summary>
        /// List of sleeves in this proximity group
        /// </summary>
        public List<UnifiedSleeve> Sleeves { get; set; } = new List<UnifiedSleeve>();
        
        /// <summary>
        /// Gets the distinct categories represented in this group
        /// </summary>
        public List<string> Categories => Sleeves
            .Select(s => s.Category)
            .Distinct()
            .OrderBy(c => c)
            .ToList();
        
        /// <summary>
        /// Gets the count of sleeves in this group
        /// </summary>
        public int Count => Sleeves.Count;
        
        /// <summary>
        /// Adds a sleeve to this proximity group
        /// </summary>
        public void Add(UnifiedSleeve sleeve)
        {
            if (sleeve == null)
                throw new ArgumentNullException(nameof(sleeve));
            
            Sleeves.Add(sleeve);
        }
        
        /// <summary>
        /// Adds multiple sleeves to this proximity group
        /// </summary>
        public void AddRange(IEnumerable<UnifiedSleeve> sleeves)
        {
            if (sleeves == null)
                throw new ArgumentNullException(nameof(sleeves));
            
            Sleeves.AddRange(sleeves);
        }
        
        /// <summary>
        /// Calculates the combined bounding box that encompasses all sleeves in the group
        /// </summary>
        public BoundingBoxXYZ CalculateCombinedBoundingBox()
        {
            if (Sleeves.Count == 0)
                throw new InvalidOperationException("Cannot calculate bounding box for empty proximity group");
            
            // Collect all corner points from all sleeves
            var allCorners = Sleeves
                .Where(s => s.Corners != null && s.Corners.Count > 0)
                .SelectMany(s => s.Corners)
                .ToList();
            
            if (allCorners.Count == 0)
            {
                // Fallback: use bounding boxes if corners not available
                allCorners = Sleeves
                    .Where(s => s.BoundingBox != null)
                    .SelectMany(s => new[] { s.BoundingBox.Min, s.BoundingBox.Max })
                    .ToList();
            }
            
            if (allCorners.Count == 0)
                throw new InvalidOperationException("No geometry data available for bounding box calculation");
            
            var minX = allCorners.Min(c => c.X);
            var minY = allCorners.Min(c => c.Y);
            var minZ = allCorners.Min(c => c.Z);
            var maxX = allCorners.Max(c => c.X);
            var maxY = allCorners.Max(c => c.Y);
            var maxZ = allCorners.Max(c => c.Z);
            
            return new BoundingBoxXYZ
            {
                Min = new XYZ(minX, minY, minZ),
                Max = new XYZ(maxX, maxY, maxZ)
            };
        }
        
        /// <summary>
        /// Calculates the combined dimensions (width, height, depth) for the group
        /// </summary>
        public (double width, double height, double depth) CalculateCombinedDimensions()
        {
            string hostType = GetHostType();
            string orientation = GetHostOrientation();
            var referenceSleeve = Sleeves.First();

            // Floor (or no host): use axis-aligned XY only (rotation = 0).
            // Framing (Y or X): no rotation — use axis-aligned extents so Y framing gets width = X span, height = Z span.
            bool isFloor = string.Equals(hostType, "Floor", StringComparison.OrdinalIgnoreCase)
                || string.IsNullOrWhiteSpace(hostType);
            bool isFraming = (hostType ?? string.Empty).IndexOf("Framing", StringComparison.OrdinalIgnoreCase) >= 0;
            double rotationRad = (isFloor || isFraming) ? 0.0 : (referenceSleeve.RotationAngleDeg * (Math.PI / 180.0));

            // 1. Define Projection Vectors (Rotated Axes; for Floor these are just X and Y)
            XYZ vecU = new XYZ(Math.Cos(rotationRad), Math.Sin(rotationRad), 0);
            XYZ vecV = new XYZ(-Math.Sin(rotationRad), Math.Cos(rotationRad), 0);
            XYZ vecZ = XYZ.BasisZ;

            // 2. Project all corners — diagnostic: log per-sleeve geometry so we can see what causes rejection
            for (int i = 0; i < Sleeves.Count; i++)
            {
                var s = Sleeves[i];
                bool hasCorners = s.Corners != null && s.Corners.Count > 0;
                bool hasBbox = s.BoundingBox != null;
                DebugLogger.Info($"[CombinedDimensions] Sleeve[{i}] Id={s?.Id} Type={s?.Type} Corners={(hasCorners ? s.Corners.Count.ToString() : "null/0")} BoundingBox={(hasBbox ? "yes" : "null")}");
            }
            var allCorners = Sleeves
                .Where(s => s.Corners != null && s.Corners.Count > 0)
                .SelectMany(s => s.Corners)
                .ToList();

            if (allCorners.Count == 0) // Fallback
            {
                 DebugLogger.Info($"[CombinedDimensions] No corners from Sleeves.Corners; using BoundingBox fallback");
                 allCorners = Sleeves.Where(s => s.BoundingBox != null)
                     .SelectMany(s => new[] { 
                         s.BoundingBox.Min, 
                         s.BoundingBox.Max,
                         new XYZ(s.BoundingBox.Max.X, s.BoundingBox.Min.Y, s.BoundingBox.Min.Z),
                         new XYZ(s.BoundingBox.Min.X, s.BoundingBox.Max.Y, s.BoundingBox.Min.Z)
                     }).ToList();
            }
            DebugLogger.Info($"[CombinedDimensions] allCorners.Count={allCorners.Count} (from {(allCorners.Count > 0 && Sleeves.Any(s => s.Corners != null && s.Corners.Count > 0) ? "Corners" : "BBox")})");

            double minU = double.MaxValue, maxU = double.MinValue;
            double minV = double.MaxValue, maxV = double.MinValue;
            double minZ = double.MaxValue, maxZ = double.MinValue;

            foreach (var pt in allCorners)
            {
                double u = pt.DotProduct(vecU);
                double v = pt.DotProduct(vecV);
                double z = pt.DotProduct(vecZ);

                if (u < minU) minU = u; if (u > maxU) maxU = u;
                if (v < minV) minV = v; if (v > maxV) maxV = v;
                if (z < minZ) minZ = z; if (z > maxZ) maxZ = z;
            }

            // 3. Determine Dimensions based on Host Type & Orientation
            double rangeU = maxU - minU; // Rotated X-Range
            double rangeV = maxV - minV; // Rotated Y-Range
            double rangeZ = maxZ - minZ; // Z-Range (Vertical)

            double calculatedWidth, calculatedHeight, calculatedDepth;

            // Depth = structural thickness for all host types (Floor, Wall, Framing). Then constituent ClusterDepth.
            double? dbThickness = null;
            foreach (var s in Sleeves)
            {
                if (s.SourceData is ClashZone cz && cz.StructuralElementThickness > 0.001)
                {
                    dbThickness = cz.StructuralElementThickness;
                    break;
                }
                if (s.SourceData is ClusterSleeveData csd && csd.ClusterDepth > 0.001)
                {
                    dbThickness = csd.ClusterDepth;
                    break;
                }
            }
            double defaultDepth = GetConstituentThicknessFallback();

            if (isFloor)
            {
                // Floor (or no host): Width & height from corner XY extents only (U/V). Depth = structural/constituent thickness only.
                calculatedWidth = rangeU;
                calculatedHeight = rangeV;
                calculatedDepth = dbThickness ?? defaultDepth; // Never use rangeZ — depth is slab/sleeve thickness only
            }
            else if (isFraming)
            {
                // Framing: axis-aligned (rotationRad=0). U=X, V=Y, Z=Z.
                // Y framing: width = Y extent (rangeV), height = Z (rangeZ), depth = X extent (rangeU) or thickness.
                // X framing: width = X extent (rangeU), height = Z (rangeZ), depth = Y extent (rangeV) or thickness.
                calculatedHeight = rangeZ;
                if (string.Equals(orientation, "Y", StringComparison.OrdinalIgnoreCase))
                {
                    calculatedWidth = rangeV;  // Y extent
                    calculatedDepth = dbThickness ?? rangeU; // X extent or thickness
                }
                else
                {
                    calculatedWidth = rangeU;  // X extent
                    calculatedDepth = dbThickness ?? rangeV; // Y extent or thickness
                }
            }
            else
            {
                // Wall: Use Orientation to swap U/V for Width vs Depth. Z is always Height.
                calculatedHeight = rangeZ;
                if (string.Equals(orientation, "Y", StringComparison.OrdinalIgnoreCase))
                {
                    calculatedWidth = rangeV;
                    calculatedDepth = dbThickness ?? rangeU;
                }
                else
                {
                    calculatedWidth = rangeU;
                    calculatedDepth = dbThickness ?? rangeV;
                }
            }

            // ✅ FIX: If width or height is 0 (or near-zero) from corner/bbox extent, use constituent fallback so combined sleeve never gets W=0
            const double minExtentFt = 0.01; // ~3mm
            var (fallbackW, fallbackH) = GetConstituentWidthHeightFallback();
            if (calculatedWidth < minExtentFt && fallbackW > minExtentFt)
                calculatedWidth = fallbackW;
            if (calculatedHeight < minExtentFt && fallbackH > minExtentFt)
                calculatedHeight = fallbackH;

            DebugLogger.Info($"[CombinedDimensions] result W={calculatedWidth:F3} H={calculatedHeight:F3} D={calculatedDepth:F3} ft");
            return (calculatedWidth, calculatedHeight, calculatedDepth);
        }

        /// <summary>
        /// Gets fallback width and height from constituents (max of SleeveWidth/SleeveHeight or ClusterWidth/ClusterHeight).
        /// Used when corner/bbox extent gives 0 or near-zero so combined sleeve never gets W=0 or H=0.
        /// </summary>
        private (double width, double height) GetConstituentWidthHeightFallback()
        {
            double maxW = 0, maxH = 0;
            foreach (var s in Sleeves)
            {
                if (s.SourceData is ClashZone cz)
                {
                    if (cz.SleeveWidth > maxW) maxW = cz.SleeveWidth;
                    if (cz.SleeveHeight > maxH) maxH = cz.SleeveHeight;
                }
                else if (s.SourceData is ClusterSleeveData csd)
                {
                    if (csd.ClusterWidth > maxW) maxW = csd.ClusterWidth;
                    if (csd.ClusterHeight > maxH) maxH = csd.ClusterHeight;
                }
            }
            return (maxW, maxH);
        }
        
        /// <summary>
        /// Calculates the rotation angle in radians for the combined sleeve.
        /// Floor: always 0 degrees (no rotation). Wall/Framing: use host orientation (Y → 0°, X → 90°).
        /// </summary>
        public double CalculateCombinedRotation()
        {
            if (string.Equals(GetHostType(), "Floor", StringComparison.OrdinalIgnoreCase))
                return 0.0; // Floor: no rotation for combined sleeve — always 0 degrees

            // Wall / Framing: use host orientation
            var orientation = GetHostOrientation();
            if (string.Equals(orientation, "Y", StringComparison.OrdinalIgnoreCase))
                return 0.0; // Y-Wall (North-South): native alignment
            return Math.PI / 2.0; // X-Wall (East-West): rotate 90°
        }

        /// <summary>
        /// Calculates the combined placement point.
        /// Floor: mid X, mid Y from bbox; Z from constituent(s) so it stays same as individual (mid-thickness). Otherwise: center of combined bbox.
        /// </summary>
        public XYZ CalculateCombinedPlacementPoint()
        {
            var bbox = CalculateCombinedBoundingBox();
            var ht = GetHostType();
            bool isFloor = string.Equals(ht, "Floor", StringComparison.OrdinalIgnoreCase) || string.IsNullOrWhiteSpace(ht);
            if (isFloor)
            {
                double midX = (bbox.Min.X + bbox.Max.X) / 2.0;
                double midY = (bbox.Min.Y + bbox.Max.Y) / 2.0;
                // Z from constituent(s) — do not change from individual placement (mid-thickness)
                double floorZ = GetConstituentZForFloor();
                if (Math.Abs(floorZ) < 1e-9)
                    floorZ = bbox.Min.Z;
                return new XYZ(midX, midY, floorZ);
            }
            return (bbox.Min + bbox.Max) / 2.0;
        }

        /// <summary>Z for floor combined = constituent value (average of placement Z). Keeps Z same as individual.</summary>
        private double GetConstituentZForFloor()
        {
            var zList = new List<double>();
            foreach (var s in Sleeves)
            {
                if (s.PlacementPoint != null && Math.Abs(s.PlacementPoint.Z) > 1e-9)
                    zList.Add(s.PlacementPoint.Z);
                else if (s.SourceData is ClashZone cz && Math.Abs(cz.SleevePlacementPointZ) > 1e-9)
                    zList.Add(cz.SleevePlacementPointZ);
                else if (s.SourceData is ClusterSleeveData csd && Math.Abs(csd.PlacementZ) > 1e-9)
                    zList.Add(csd.PlacementZ);
            }
            return zList.Count > 0 ? zList.Average() : 0;
        }
        
        /// <summary>
        /// Gets fallback depth: structural thickness (Floor/Wall/Framing) or constituent ClusterDepth.
        /// </summary>
        private double GetConstituentThicknessFallback()
        {
            foreach (var s in Sleeves)
            {
                if (s.SourceData is ClashZone cz && cz.StructuralElementThickness > 0.001)
                    return cz.StructuralElementThickness;
                if (s.SourceData is ClusterSleeveData csd && csd.ClusterDepth > 0.001)
                    return csd.ClusterDepth;
            }
            return 0.1; // ~30mm default only when no constituent has thickness
        }

        /// <summary>
        /// Reference sleeve for host/orientation: prefer Framing so mixed groups (e.g. Duct in Wall + Cluster in Y Framing) use Y framing → 0° rotation.
        /// </summary>
        private UnifiedSleeve GetReferenceSleeveForHost()
        {
            var framing = Sleeves.FirstOrDefault(s => (s?.HostType ?? string.Empty).IndexOf("Framing", StringComparison.OrdinalIgnoreCase) >= 0);
            return framing ?? Sleeves.FirstOrDefault();
        }

        /// <summary>
        /// Gets the host type; prefers a Framing sleeve when present so Y framing gets 0° rotation.
        /// </summary>
        public string GetHostType()
        {
            return GetReferenceSleeveForHost()?.HostType;
        }
        
        /// <summary>
        /// Gets the host orientation; prefers a Framing sleeve when present. Y framing / Y wall → no rotation (0°).
        /// </summary>
        public string GetHostOrientation()
        {
            var refSleeve = GetReferenceSleeveForHost();
            if (refSleeve == null) return null;
            var hostType = refSleeve.HostType ?? string.Empty;
            var ht = hostType.Trim();
            if (string.Equals(ht, "Floor", StringComparison.OrdinalIgnoreCase))
                return null; // Floor has no host orientation
            if (!string.Equals(ht, "Wall", StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(ht, "Framing", StringComparison.OrdinalIgnoreCase) &&
                ht.IndexOf("Framing", StringComparison.OrdinalIgnoreCase) < 0)
                return null; // Only Wall and Framing have orientation
            return refSleeve.HostOrientation;
        }
        
        /// <summary>
        /// Checks if this group contains sleeves from multiple categories (cross-category)
        /// </summary>
        public bool IsCrossCategory()
        {
            return Categories.Count > 1;
        }
        
        /// <summary>
        /// Gets a summary string describing this proximity group
        /// </summary>
        public string GetSummary()
        {
            var categorySummary = string.Join("+", Categories);
            var typeSummary = Sleeves.GroupBy(s => s.Type)
                .Select(g => $"{g.Count()} {g.Key}")
                .ToList();
            
            return $"{categorySummary} ({string.Join(", ", typeSummary)})";
        }
        
        /// <summary>
        /// Validates that this proximity group is valid for combined sleeve creation
        /// </summary>
        public bool IsValid()
        {
            // Must have at least 2 sleeves
            if (Sleeves.Count < 2)
            {
                DebugLogger.Info($"[CombinedReject] REJECT: Sleeves.Count={Sleeves.Count} (need >= 2)");
                return false;
            }
            
            // Must be cross-category (different categories)
            if (!IsCrossCategory())
            {
                DebugLogger.Info($"[CombinedReject] REJECT: not cross-category (categories: {string.Join(", ", Sleeves.Select(s => s?.Category ?? "?"))})");
                return false;
            }
            
            // All sleeves must have geometry data
            var noGeometry = Sleeves.Where(s => s.BoundingBox == null && (s.Corners == null || s.Corners.Count == 0)).ToList();
            if (noGeometry.Any())
            {
                foreach (var s in noGeometry)
                    DebugLogger.Info($"[CombinedReject] REJECT: sleeve Id={s?.Id} Type={s?.Type} has no geometry (BoundingBox=null, Corners=null or Count=0)");
                return false;
            }
            
            return true;
        }
    }
}

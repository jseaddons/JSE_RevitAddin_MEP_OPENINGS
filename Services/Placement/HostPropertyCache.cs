using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Utils;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Placement
{
    /// <summary>
    /// Caches host element properties (walls, floors, framing) to ensure consistent
    /// depth, centerline, and orientation calculations for all sleeves on the same host.
    /// Fixes "half in/out of wall" issues caused by inconsistent property queries.
    /// </summary>
    public class HostPropertyCache
    {
        private readonly Document _doc;
        private readonly Dictionary<ElementId, HostProperties> _cache;

        public HostPropertyCache(Document doc)
        {
            _doc = doc;
            _cache = new Dictionary<ElementId, HostProperties>();
        }

        /// <summary>
        /// Gets cached host properties or computes them if not cached.
        /// </summary>
        public HostProperties GetOrCompute(Element host)
        {
            if (host == null)
                return null;

            if (_cache.ContainsKey(host.Id))
                return _cache[host.Id];

            var properties = ComputeHostProperties(host);
            _cache[host.Id] = properties;
            return properties;
        }

        /// <summary>
        /// Clears the cache. Call this when starting a new placement batch.
        /// </summary>
        public void Clear()
        {
            _cache.Clear();
        }

        private HostProperties ComputeHostProperties(Element host)
        {
            var props = new HostProperties
            {
                HostId = host.Id,
                HostCategory = host.Category?.Name ?? "Unknown"
            };

            // Determine host type and compute properties
            if (host is Wall wall)
            {
                ComputeWallProperties(wall, props);
            }
            else if (host is Floor floor)
            {
                ComputeFloorProperties(floor, props);
            }
            else if (host.Category?.Id.IntegerValue == (int)BuiltInCategory.OST_StructuralFraming)
            {
                ComputeFramingProperties(host, props);
            }
            else if (host.Category?.Id.IntegerValue == (int)BuiltInCategory.OST_StructuralColumns)
            {
                ComputeColumnProperties(host, props);
            }

            return props;
        }

        private void ComputeWallProperties(Wall wall, HostProperties props)
        {
            // Get wall thickness
            props.Thickness = wall.Width;

            // Get wall location curve
            LocationCurve locCurve = wall.Location as LocationCurve;
            if (locCurve != null)
            {
                Curve curve = locCurve.Curve;
                XYZ start = curve.GetEndPoint(0);
                XYZ end = curve.GetEndPoint(1);

                // Calculate wall direction and normal
                XYZ direction = (end - start).Normalize();
                XYZ normal = new XYZ(-direction.Y, direction.X, 0).Normalize();

                props.Normal = normal;
                props.Direction = direction;

                // Determine if wall is vertical (X or Y aligned)
                double angleToX = Math.Abs(direction.AngleTo(XYZ.BasisX));
                double angleToY = Math.Abs(direction.AngleTo(XYZ.BasisY));

                props.IsVertical = true;
                props.IsXAligned = angleToX < 0.1 || angleToX > (Math.PI - 0.1); // ~0 or ~180 degrees
                props.IsYAligned = angleToY < 0.1 || angleToY > (Math.PI - 0.1);

                // Calculate centerline (midpoint of wall curve)
                XYZ midpoint = (start + end) / 2.0;
                props.CenterlineX = midpoint.X;
                props.CenterlineY = midpoint.Y;
                props.CenterlineZ = midpoint.Z;
            }
        }

        private void ComputeFloorProperties(Floor floor, HostProperties props)
        {
            // Get floor thickness
            Parameter thicknessParam = floor.get_Parameter(BuiltInParameter.FLOOR_ATTR_THICKNESS_PARAM);
            if (thicknessParam != null && thicknessParam.StorageType == StorageType.Double)
            {
                props.Thickness = thicknessParam.AsDouble();
            }

            // Floor normal is typically Z-axis (pointing up)
            props.Normal = XYZ.BasisZ;
            props.IsVertical = false;

            // Get floor level elevation as centerline Z
            Level level = _doc.GetElement(floor.LevelId) as Level;
            if (level != null)
            {
                props.CenterlineZ = level.Elevation;
            }
        }

        private void ComputeFramingProperties(Element framing, HostProperties props)
        {
            // Get framing dimensions (typically beams)
            // Framing elements have width/height parameters
            Parameter widthParam = framing.LookupParameter("b") ?? framing.LookupParameter("Width");
            Parameter heightParam = framing.LookupParameter("h") ?? framing.LookupParameter("Height");

            if (widthParam != null && widthParam.StorageType == StorageType.Double)
            {
                props.Thickness = widthParam.AsDouble();
            }

            // Get framing location curve
            LocationCurve locCurve = framing.Location as LocationCurve;
            if (locCurve != null)
            {
                Curve curve = locCurve.Curve;
                XYZ start = curve.GetEndPoint(0);
                XYZ end = curve.GetEndPoint(1);

                XYZ direction = (end - start).Normalize();
                props.Direction = direction;

                // Calculate normal (perpendicular to beam direction)
                XYZ normal = new XYZ(-direction.Y, direction.X, 0).Normalize();
                props.Normal = normal;

                // Centerline is midpoint of beam
                XYZ midpoint = (start + end) / 2.0;
                props.CenterlineX = midpoint.X;
                props.CenterlineY = midpoint.Y;
                props.CenterlineZ = midpoint.Z;
            }
        }

        private void ComputeColumnProperties(Element column, HostProperties props)
        {
            // Get column dimensions
            Parameter widthParam = column.LookupParameter("b") ?? column.LookupParameter("Width");
            Parameter depthParam = column.LookupParameter("h") ?? column.LookupParameter("Depth");

            if (widthParam != null && widthParam.StorageType == StorageType.Double)
            {
                props.Thickness = widthParam.AsDouble();
            }

            // Column normal is typically horizontal (perpendicular to column axis)
            LocationPoint locPoint = column.Location as LocationPoint;
            if (locPoint != null)
            {
                XYZ point = locPoint.Point;
                props.CenterlineX = point.X;
                props.CenterlineY = point.Y;
                props.CenterlineZ = point.Z;
            }

            props.IsVertical = true;
        }
    }

    /// <summary>
    /// Cached properties for a host element (wall, floor, framing).
    /// </summary>
    public class HostProperties
    {
        public ElementId HostId { get; set; }
        public string HostCategory { get; set; }

        // Centerline coordinates
        public double CenterlineX { get; set; }
        public double CenterlineY { get; set; }
        public double CenterlineZ { get; set; }

        // Thickness (wall width, floor thickness, beam width, etc.)
        public double Thickness { get; set; }

        // Orientation
        public XYZ Normal { get; set; }
        public XYZ Direction { get; set; }
        public bool IsVertical { get; set; }
        public bool IsXAligned { get; set; }
        public bool IsYAligned { get; set; }
    }
}

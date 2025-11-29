using System;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.DamperDetection
{
    /// <summary>
    /// Detects MEP connector presence and direction for dampers.
    /// Follows Single Responsibility Principle - only responsible for connector detection.
    /// </summary>
    public class DamperConnectorDetector : IDamperConnectorDetector
    {
        /// <summary>
        /// Checks if a damper has at least one MEP connector.
        /// </summary>
        public bool HasMepConnector(FamilyInstance damper)
        {
            if (damper == null)
                return false;

            var cm = damper.MEPModel?.ConnectorManager;
            if (cm == null)
                return false;

            return cm.Connectors != null && cm.Connectors.Size > 0;
        }

        /// <summary>
        /// Detects if a damper has an MEP connector and returns its direction.
        /// </summary>
        /// <param name="damper">The damper FamilyInstance to check</param>
        /// <param name="useWorldCoordinates">If true, uses world coordinates (for non-standard dampers). If false, uses local coordinates (for standard dampers)</param>
        /// <param name="connector">Output parameter containing the detected connector, or null if not found</param>
        /// <returns>Connector side direction ("Left", "Right", "Top", "Bottom") or empty string if not detected</returns>
        public string DetectConnectorSide(FamilyInstance damper, bool useWorldCoordinates, out Connector connector)
        {
            connector = null;

            if (damper == null)
                return string.Empty;

            if (useWorldCoordinates)
            {
                return DetectConnectorSideWorld(damper, out connector);
            }
            else
            {
                return DetectConnectorSideLocal(damper, out connector);
            }
        }

        /// <summary>
        /// Gets the connector side using world coordinates (for non-standard dampers like MSFD).
        /// </summary>
        private string DetectConnectorSideWorld(FamilyInstance damper, out Connector connector)
        {
            connector = null;

            var cm = damper.MEPModel?.ConnectorManager;
            if (cm == null)
                return "Right";

            // Pick the connector with the largest absolute component
            Connector best = null;
            double maxAbs = 0;

            foreach (Connector c in cm.Connectors)
            {
                var dir = c.CoordinateSystem.BasisX;  // world direction
                double maxComp = Math.Max(Math.Max(Math.Abs(dir.X), Math.Abs(dir.Y)), Math.Abs(dir.Z));
                if (maxComp > maxAbs)
                {
                    maxAbs = maxComp;
                    best = c;
                }
            }

            if (best == null)
                return "Right";

            connector = best;
            var d = best.CoordinateSystem.BasisX;

            // Precedence: Z for top/bottom, then X for left/right
            if (Math.Abs(d.Z) >= 0.9)
                return d.Z > 0 ? "Top" : "Bottom";
            if (Math.Abs(d.X) >= 0.9)
                return d.X > 0 ? "Right" : "Left";

            return "Right";
        }

        /// <summary>
        /// Gets the connector side using local coordinates (for standard dampers).
        /// </summary>
        private string DetectConnectorSideLocal(FamilyInstance damper, out Connector connector)
        {
            connector = null;

            var cons = (damper.MEPModel as MechanicalFitting)?.ConnectorManager?.Connectors
                    ?? damper.MEPModel?.ConnectorManager?.Connectors;

            if (cons == null || cons.Size == 0)
                return "Unknown";

            // Pick the connector closest to origin
            connector = cons.Cast<Connector>()
                             .OrderBy(c => c.Origin.DistanceTo(damper.GetTransform().Origin))
                             .FirstOrDefault();

            if (connector == null)
                return "Unknown";

            var damperT = damper.GetTransform();
            var connDir = connector.CoordinateSystem.BasisX; // connector normal

            // Dot products with damper local axes
            double dotX = connDir.DotProduct(damperT.BasisX);
            double dotY = connDir.DotProduct(damperT.BasisY);
            double dotZ = connDir.DotProduct(damperT.BasisZ);

            const double tol = 0.9;

            if (Math.Abs(dotX) >= tol)
                return dotX > 0 ? "Right" : "Left";
            if (Math.Abs(dotZ) >= tol)
                return dotZ > 0 ? "Top" : "Bottom";

            return "Unknown";
        }
    }
}


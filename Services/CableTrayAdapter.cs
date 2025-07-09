using System;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;  // for CableTray

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Adapter for CableTray elements to provide common MEP interface.
    /// </summary>
    public class CableTrayAdapter : IMepElementAdapter
    {
        public Element Element { get; }
        private readonly CableTray tray;

        public CableTrayAdapter(CableTray cableTray)
        {
            tray = cableTray ?? throw new ArgumentNullException(nameof(cableTray));
            Element = cableTray;
        }

        public Curve GetCurve()
        {
            // Straight segments return as LocationCurve; use that for intersection rays
            return (tray.Location as LocationCurve)?.Curve;
        }

        public double GetDiameter()
        {
            return 0.0;
        }

        public double GetWidth()
        {
            return tray.Width;  // CableTray exposes Width directly
        }

        public double GetHeight()
        {
            return tray.Height; // CableTray exposes Height directly
        }

        public XYZ GetDirection()
        {
            var line = GetCurve() as Line;
            return line != null ? line.Direction : XYZ.BasisX;
        }

        public string ElementTypeName => "CableTray";
    }
}

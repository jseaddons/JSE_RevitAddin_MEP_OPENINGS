using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Helpers
{
    public interface IViewGeometryHelper
    {
        Outline GetWorldViewBounds(ViewPlan plan);
        bool IsPointInView(XYZ point, Outline worldBounds);
    }

    public class ViewGeometryHelper : IViewGeometryHelper
    {
        public Outline GetWorldViewBounds(ViewPlan plan)
        {
            if (plan == null || !plan.CropBoxActive || plan.CropBox == null)
                return null;

            BoundingBoxXYZ viewExtent = plan.CropBox;
            Transform viewTransform = viewExtent.Transform;
            
            XYZ bMin = viewExtent.Min;
            XYZ bMax = viewExtent.Max;
            
            var corners = new List<XYZ>
            {
                viewTransform.OfPoint(new XYZ(bMin.X, bMin.Y, bMin.Z)),
                viewTransform.OfPoint(new XYZ(bMax.X, bMin.Y, bMin.Z)),
                viewTransform.OfPoint(new XYZ(bMax.X, bMax.Y, bMin.Z)),
                viewTransform.OfPoint(new XYZ(bMin.X, bMax.Y, bMin.Z))
            };
            
            double worldMinX = corners.Min(c => c.X);
            double worldMinY = corners.Min(c => c.Y);
            double worldMaxX = corners.Max(c => c.X);
            double worldMaxY = corners.Max(c => c.Y);
            
            // Apply tolerance
            double tolerance = 0.01; 
            return new Outline(
                new XYZ(worldMinX - tolerance, worldMinY - tolerance, 0),
                new XYZ(worldMaxX + tolerance, worldMaxY + tolerance, 0)
            );
        }

        public bool IsPointInView(XYZ point, Outline worldBounds)
        {
            if (worldBounds == null) return true; // No bounds means all is in view

            return point.X >= worldBounds.MinimumPoint.X && point.X <= worldBounds.MaximumPoint.X &&
                   point.Y >= worldBounds.MinimumPoint.Y && point.Y <= worldBounds.MaximumPoint.Y;
        }
    }
}

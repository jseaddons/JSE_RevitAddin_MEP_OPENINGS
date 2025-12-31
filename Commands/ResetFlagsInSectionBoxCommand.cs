using System;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories;

namespace JSE_RevitAddin_MEP_OPENINGS.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class ResetFlagsInSectionBoxCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            var uiDoc = commandData.Application.ActiveUIDocument;
            var doc = uiDoc.Document;

            try
            {
                // 1. Validate Active View is 3D with Section Box
                var view3D = doc.ActiveView as View3D;
                if (view3D == null || !view3D.IsSectionBoxActive)
                {
                    TaskDialog.Show("Requirements", "Please go to a 3D View with an active Section Box to use this command.");
                    return Result.Cancelled;
                }

                // 2. Get Section Box Geometry (World Coordinates)
                var sectionBox = view3D.GetSectionBox();
                if (sectionBox == null)
                {
                    TaskDialog.Show("Error", "Could not retrieve Section Box geometry.");
                    return Result.Failed;
                }
                
                // Transform handling:
                // GetSectionBox() returns a BoundingBoxXYZ that might have a Transform. 
                // However, we need the World Coordinate bounds that encompass this box for the SQL query.
                // If it's rotated, the Min/Max of the BBox definition are in local coords.
                // We must transform the 8 corners to World space and find the World Min/Max.
                
                BoundingBoxXYZ worldBox = GetWorldBoundingBox(sectionBox);

                // 3. Reset Flags in DB
                int count = 0;
                // Use a transient context for this operation
                using (var context = new SleeveDbContext(doc, msg => System.Diagnostics.Debug.WriteLine(msg)))
                {
                    var repo = new ClashZoneRepository(context);
                    count = repo.ResetResolvedFlagsInSectionBox(worldBox);
                }

                // 4. Feedback
                TaskDialog.Show("Success", $"Reset flags for {count} clash zones within the selected Section Box.\n\nYou can now run 'Refresh Clash Zones' to re-process them.");

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }

        private BoundingBoxXYZ GetWorldBoundingBox(BoundingBoxXYZ localBox)
        {
            var t = localBox.Transform;
            
            // If identity, simple return
            if (t.IsIdentity) return localBox;

            // Construct 8 corners
            XYZ min = localBox.Min;
            XYZ max = localBox.Max;

            XYZ[] corners = new XYZ[8];
            corners[0] = t.OfPoint(new XYZ(min.X, min.Y, min.Z));
            corners[1] = t.OfPoint(new XYZ(max.X, min.Y, min.Z));
            corners[2] = t.OfPoint(new XYZ(min.X, max.Y, min.Z));
            corners[3] = t.OfPoint(new XYZ(max.X, max.Y, min.Z));
            corners[4] = t.OfPoint(new XYZ(min.X, min.Y, max.Z));
            corners[5] = t.OfPoint(new XYZ(max.X, min.Y, max.Z));
            corners[6] = t.OfPoint(new XYZ(min.X, max.Y, max.Z));
            corners[7] = t.OfPoint(new XYZ(max.X, max.Y, max.Z));

            double wMinX = double.MaxValue, wMinY = double.MaxValue, wMinZ = double.MaxValue;
            double wMaxX = double.MinValue, wMaxY = double.MinValue, wMaxZ = double.MinValue;

            foreach (var p in corners)
            {
                if (p.X < wMinX) wMinX = p.X;
                if (p.Y < wMinY) wMinY = p.Y;
                if (p.Z < wMinZ) wMinZ = p.Z;

                if (p.X > wMaxX) wMaxX = p.X;
                if (p.Y > wMaxY) wMaxY = p.Y;
                if (p.Z > wMaxZ) wMaxZ = p.Z;
            }

            return new BoundingBoxXYZ 
            { 
                Min = new XYZ(wMinX, wMinY, wMinZ), 
                Max = new XYZ(wMaxX, wMaxY, wMaxZ) 
            };
        }
    }
}

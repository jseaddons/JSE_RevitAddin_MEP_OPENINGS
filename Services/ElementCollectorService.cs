using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    public class ElementCollectorService
    {
        private readonly Action<string> _log;

        public ElementCollectorService(Action<string> log)
        {
            _log = log ?? throw new ArgumentNullException(nameof(log));
        }

        public void CollectElementsInSectionBox(Document hostDoc,
                                               View3D view3D,
                                               out List<Element> mepInSectionBox,
                                               out List<Element> structuralInSectionBox)
        {
            mepInSectionBox = new List<Element>();
            structuralInSectionBox = new List<Element>();

            // Get section box from view
            BoundingBoxXYZ sectionBox = view3D.GetSectionBox();
            
            // CRITICAL: Remove the view's transform to get model coordinates
            Transform viewTransform = view3D.CropBox.Transform;
            XYZ minModel = viewTransform.OfPoint(sectionBox.Min);
            XYZ maxModel = viewTransform.OfPoint(sectionBox.Max);
            
            // Ensure min/max are correct
            BoundingBoxXYZ modelBox = new BoundingBoxXYZ
            {
                Min = new XYZ(Math.Min(minModel.X, maxModel.X), 
                             Math.Min(minModel.Y, maxModel.Y), 
                             Math.Min(minModel.Z, maxModel.Z)),
                Max = new XYZ(Math.Max(minModel.X, maxModel.X), 
                             Math.Max(minModel.Y, maxModel.Y), 
                             Math.Max(minModel.Z, maxModel.Z))
            };

            _log($"Section box (view space): Min({sectionBox.Min.X:F3}, {sectionBox.Min.Y:F3}, {sectionBox.Min.Z:F3})");
            _log($"Section box (model space): Min({modelBox.Min.X:F3}, {modelBox.Min.Y:F3}, {modelBox.Min.Z:F3})  Max({modelBox.Max.X:F3}, {modelBox.Max.Y:F3}, {modelBox.Max.Z:F3})");

            Outline hostOutline = new Outline(modelBox.Min, modelBox.Max);

            // Collect host elements
            BuiltInCategory[] mepCats = {
                // Ducts
                BuiltInCategory.OST_DuctCurves,
                BuiltInCategory.OST_DuctFitting,
                BuiltInCategory.OST_DuctAccessory,
                BuiltInCategory.OST_DuctTerminal,
                // Pipes
                BuiltInCategory.OST_PipeCurves,
                BuiltInCategory.OST_PipeFitting,
                BuiltInCategory.OST_PipeAccessory,
                // Cable Management
                BuiltInCategory.OST_CableTray,
                BuiltInCategory.OST_CableTrayFitting,
                BuiltInCategory.OST_Conduit,
                BuiltInCategory.OST_ConduitFitting,
                // Equipment
                BuiltInCategory.OST_MechanicalEquipment,
                BuiltInCategory.OST_PlumbingFixtures,
                BuiltInCategory.OST_Sprinklers,
                BuiltInCategory.OST_ElectricalFixtures,
                BuiltInCategory.OST_ElectricalEquipment,
                BuiltInCategory.OST_LightingFixtures
            };
            
            foreach (var cat in mepCats)
            {
                var elements = new FilteredElementCollector(hostDoc)
                    .OfCategory(cat)
                    .WhereElementIsNotElementType()
                    .WherePasses(new BoundingBoxIntersectsFilter(hostOutline))
                    .ToElements();
                mepInSectionBox.AddRange(elements);
                if (elements.Count > 0)
                    _log($"Host Category {cat}: Found {elements.Count} elements");
            }

            BuiltInCategory[] structCats = {
                BuiltInCategory.OST_Walls, BuiltInCategory.OST_StructuralFraming,
                BuiltInCategory.OST_Floors
            };
            
            foreach (var cat in structCats)
            {
                var elements = new FilteredElementCollector(hostDoc)
                    .OfCategory(cat)
                    .WhereElementIsNotElementType()
                    .WherePasses(new BoundingBoxIntersectsFilter(hostOutline))
                    .ToElements();
                structuralInSectionBox.AddRange(elements);
                if (elements.Count > 0)
                    _log($"Host Category {cat}: Found {elements.Count} elements");
            }

            // Process linked files
            var links = new FilteredElementCollector(hostDoc)
                            .OfClass(typeof(RevitLinkInstance))
                            .Cast<RevitLinkInstance>()
                            .ToList();

            _log($"\n=== FOUND {links.Count} LINKED FILES ===");
            foreach (var link in links)
            {
                Document linkDoc = link.GetLinkDocument();
                string status = linkDoc != null ? "LOADED" : "NOT LOADED";
                _log($"  {link.Name} - {status}");
            }
            _log("");

            _log("=== PROCESSING LINKS ===");
            _log("");

            foreach (RevitLinkInstance link in links)
            {
                Document linkDoc = link.GetLinkDocument();
                if (linkDoc == null) 
                {
                    _log($"Link {link.Name}: Document not loaded");
                    continue;
                }

                _log($"Link: {link.Name}");
                
                // Get the transform from link to host
                Transform linkTransform = link.GetTotalTransform();
                _log($"  Link transform origin: ({linkTransform.Origin.X:F3}, {linkTransform.Origin.Y:F3}, {linkTransform.Origin.Z:F3})");
                
                // Transform section box to link's coordinate system using INVERSE
                Transform invTransform = linkTransform.Inverse;
                XYZ minLink = invTransform.OfPoint(modelBox.Min);
                XYZ maxLink = invTransform.OfPoint(modelBox.Max);
                
                BoundingBoxXYZ linkBox = new BoundingBoxXYZ
                {
                    Min = new XYZ(Math.Min(minLink.X, maxLink.X), 
                                 Math.Min(minLink.Y, maxLink.Y), 
                                 Math.Min(minLink.Z, maxLink.Z)),
                    Max = new XYZ(Math.Max(minLink.X, maxLink.X), 
                                 Math.Max(minLink.Y, maxLink.Y), 
                                 Math.Max(minLink.Z, maxLink.Z))
                };
                
                _log($"  Section in link coords: Min=({linkBox.Min.X:F1}, {linkBox.Min.Y:F1}, {linkBox.Min.Z:F1}) Max=({linkBox.Max.X:F1}, {linkBox.Max.Y:F1}, {linkBox.Max.Z:F1})");

                Outline linkOutline = new Outline(linkBox.Min, linkBox.Max);

                // Collect link elements
                foreach (var cat in mepCats)
                {
                    var elements = new FilteredElementCollector(linkDoc)
                        .OfCategory(cat)
                        .WhereElementIsNotElementType()
                        .WherePasses(new BoundingBoxIntersectsFilter(linkOutline))
                        .ToElements();
                    
                    if (elements.Count > 0)
                    {
                        mepInSectionBox.AddRange(elements);
                        _log($"  {cat}: {elements.Count} elements");
                        
                        // Log first element details
                        var first = elements.First();
                        var bbox = first.get_BoundingBox(null);
                        if (bbox != null)
                            _log($"    Example ID {first.Id}: ({bbox.Min.X:F1}, {bbox.Min.Y:F1}, {bbox.Min.Z:F1})");
                    }
                }

                foreach (var cat in structCats)
                {
                    var elements = new FilteredElementCollector(linkDoc)
                        .OfCategory(cat)
                        .WhereElementIsNotElementType()
                        .WherePasses(new BoundingBoxIntersectsFilter(linkOutline))
                        .ToElements();
                    
                    if (elements.Count > 0)
                    {
                        structuralInSectionBox.AddRange(elements);
                        _log($"  {cat}: {elements.Count} elements");
                    }
                }
            }

            _log($"\n=== FINAL RESULTS ===");
            _log($"Total MEP: {mepInSectionBox.Count}");
            _log($"Total Structural: {structuralInSectionBox.Count}");
        }
    }
}
using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Utils;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Calculation
{
    /// <summary>
    /// Phase 2 Service: Extracts geometry (corners) from PLACED individual sleeves
    /// and persists them to the Database for use in Phase 3 (Clustering).
    /// </summary>
    public class BatchSleeveCornerExtractor
    {
        private readonly IClashZoneRepository _repository;

        public BatchSleeveCornerExtractor(IClashZoneRepository repository)
        {
            _repository = repository ?? throw new ArgumentNullException(nameof(repository));
        }

        /// <summary>
        /// Iterates all placed sleeves in the document (that are known to DB),
        /// extracts their corner coordinates, and saves to DB.
        /// </summary>
        public void ExtractAndSaveCorners(Document doc)
        {
            SafeFileLogger.SafeAppendText("geometry_extraction.log", "Starting Batch Corner Extraction...\n");

            // 1. Get all known placed sleeves from DB
            // We need ClashZones that have a valid SleeveInstanceId
            // Optimization: Get distinct categories to query efficiently or just GetAll?
            // Let's assume we can get all zones with SleeveInstanceId > 0
            var categories = _repository.GetDistinctCategories();
            var updates = new List<(Guid Guid, double c1x, double c1y, double c1z, double c2x, double c2y, double c2z, double c3x, double c3y, double c3z, double c4x, double c4y, double c4z)>();

            foreach (var cat in categories)
            {
                var zones = _repository.GetClashZonesByCategory(cat);
                var placedZones = zones.Where(z => z.SleeveInstanceId > 0).ToList();

                foreach (var zone in placedZones)
                {
                   try
                   {
                        Element elem = doc.GetElement(new ElementId(zone.SleeveInstanceId));
                        if (elem == null || !(elem is FamilyInstance)) continue;

                        // Extract Corners
                        var corners = ExtractCornersFromSolid(elem, doc);
                        if (corners != null && corners.Count == 4)
                        {
                            // Flatten for DB
                            updates.Add((
                                zone.Id,
                                corners[0].X, corners[0].Y, corners[0].Z,
                                corners[1].X, corners[1].Y, corners[1].Z,
                                corners[2].X, corners[2].Y, corners[2].Z,
                                corners[3].X, corners[3].Y, corners[3].Z
                            ));
                        }
                   }
                   catch (Exception ex)
                   {
                       SafeFileLogger.SafeAppendText("geometry_extraction_errors.log", $"Error extracting {zone.SleeveInstanceId}: {ex.Message}\n");
                   }
                }
            }

            // 2. Batch Save
            if (updates.Any())
            {
                _repository.BatchUpdateSleeveCorners(updates);
                SafeFileLogger.SafeAppendText("geometry_extraction.log", $"Extracted and saved corners for {updates.Count} sleeves.\n");
            }
            else
            {
                SafeFileLogger.SafeAppendText("geometry_extraction.log", "No sleeves found/extracted.\n");
            }
        }

        private List<XYZ> ExtractCornersFromSolid(Element elem, Document doc)
        {
            // Get geometry
            Options opt = new Options { DetailLevel = ViewDetailLevel.Fine, ComputeReferences = true };
            GeometryElement geo = elem.get_Geometry(opt);
            
            Solid solid = null;

            if (geo != null)
            {
                foreach (GeometryObject obj in geo)
                {
                    if (obj is Solid s && s.Volume > 0)
                    {
                        solid = s;
                        break;
                    }
                    else if (obj is GeometryInstance gi)
                    {
                        foreach (GeometryObject obj2 in gi.GetSymbolGeometry())
                        {
                             if (obj2 is Solid s2 && s2.Volume > 0)
                             {
                                 // Transform to instance location
                                 solid = SolidUtils.CreateTransformed(s2, gi.Transform);
                                 break;
                             }
                        }
                    }
                    if (solid != null) break;
                }
            }

            if (solid == null) return null;

            // Get bottom face or just bounding box corners?
            // User requested "CORNERS". Typically for clustering we need the OBB (Oriented Bounding Box) or floor plan profile.
            // A robust way: getting the 4 bottom corners of the bounding box ALIGNED TO THE ELEMENT is hard if we only have the solid in World/Instance coordinates.
            // But we specifically placed these.
            // Actually, for clustering, we want the "Rotated Bounding Box" corners on the XY plane.
            
            // Logic:
            // 1. Get Oriented Bounding Box? No, native API doesn't give OBB easily.
            // 2. Calculate from Solid faces: Find the bottom-most major face (Norm Z ~ -1).
            //    Get its 4 vertices.
            
            return GetBottomFaceCorners(solid);
        }

        private List<XYZ> GetBottomFaceCorners(Solid solid)
        {
            Face bottomFace = null;
            double minZ = double.MaxValue;

            foreach (Face f in solid.Faces)
            {
                 if (f is PlanarFace pf)
                 {
                     if (pf.FaceNormal.IsAlmostEqualTo(XYZ.BasisZ.Negate(), 0.1)) // Pointing down
                     {
                         // Check elevation
                         if (pf.Origin.Z < minZ)
                         {
                             minZ = pf.Origin.Z;
                             bottomFace = pf;
                         }
                     }
                 }
            }

            if (bottomFace == null) return null;

            // Get vertices from outer loop
            var displayMesh = bottomFace.Triangulate(); 
            // Triangulate is one way, or EdgeLoops.
            // EdgeLoops is better for corners.
            
            List<XYZ> vertices = new List<XYZ>();
            foreach (EdgeArray loop in bottomFace.EdgeLoops)
            {
                foreach (Edge edge in loop)
                {
                     XYZ p = edge.AsCurve().GetEndPoint(0);
                     if (!vertices.Any(v => v.IsAlmostEqualTo(p))) vertices.Add(p);
                }
            }
            
            // Filter to 4 corners if it's rectangular
            if (vertices.Count > 4)
            {
                // Bounding box of vertices on that plane?
                // Just return simple bounding box of the face's vertices
                // This is safer if family has chamfers etc.
                double minX = vertices.Min(v => v.X);
                double maxX = vertices.Max(v => v.X);
                double minY = vertices.Min(v => v.Y);
                double maxY = vertices.Max(v => v.Y);
                double z = vertices.First().Z;
                
                return new List<XYZ>
                {
                    new XYZ(minX, minY, z),
                    new XYZ(maxX, minY, z),
                    new XYZ(maxX, maxY, z),
                    new XYZ(minX, maxY, z)
                };
            }
            
            return vertices.Take(4).ToList();
        }
    }
}

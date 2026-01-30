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
            SafeFileLogger.SafeAppendTextAlways("geometry_extraction.log", "Starting Batch Corner Extraction...\n");

            // 1. Get all known placed sleeves from DB EFFICIENTLY in a single query
            var placedZones = _repository.GetPlacedClashZones();
            
            if (placedZones == null || !placedZones.Any())
            {
                SafeFileLogger.SafeAppendText("geometry_extraction.log", "No placed sleeves found in database.\n");
                return;
            }

            ExtractAndSaveCornersInternal(doc, placedZones.ToList());
        }

        /// <summary>
        /// ✅ NEW: Extracts corners for specific zones by GUID and ElementId.
        /// This avoids SQLite WAL visibility issues where GetPlacedClashZones() returns 0 items.
        /// </summary>
        public int ExtractAndSaveCornersForZones(Document doc, IEnumerable<ClashZone> placedZones)
        {
            var zonesList = placedZones?.ToList() ?? new List<ClashZone>();
            SafeFileLogger.SafeAppendTextAlways("geometry_extraction.log", 
                $"Starting Batch Corner Extraction for {zonesList.Count} specific zones (using saved orientation)...\n");

            if (!zonesList.Any())
            {
                SafeFileLogger.SafeAppendTextAlways("geometry_extraction.log", "No zones provided for extraction.\n");
                return 0;
            }

            return ExtractAndSaveCornersInternal(doc, zonesList);
        }



        /// <summary>
        /// Internal method that does the actual extraction work
        /// </summary>
        private int ExtractAndSaveCornersInternal(Document doc, List<ClashZone> zones)
        {
            var updates = new List<(Guid Guid, double c1x, double c1y, double c1z, double c2x, double c2y, double c2z, double c3x, double c3y, double c3z, double c4x, double c4y, double c4z)>();

            foreach (var zone in zones)
            {
                try
                {
                    Element elem = doc.GetElement(new ElementId(zone.SleeveInstanceId));
                    if (elem == null || !(elem is FamilyInstance)) continue;

                    string hostOrientation = zone.HostOrientation;

                    // Extract Corners
                    var corners = ExtractCornersFromSolid(elem, doc, hostOrientation);
                    
                    // ✅ FALLBACK: If solid extraction fails, use BoundingBox corners
                    if (corners == null || corners.Count != 4)
                    {
                        var bbox = elem.get_BoundingBox(null);
                        if (bbox != null)
                        {
                            if (!string.IsNullOrEmpty(hostOrientation))
                            {
                                if (hostOrientation.Equals("X", StringComparison.OrdinalIgnoreCase))
                                {
                                    // X-Wall (Face is XZ plane) - Preserve Z-height
                                    double yMid = (bbox.Min.Y + bbox.Max.Y) / 2.0;
                                    corners = new List<XYZ>
                                    {
                                        new XYZ(bbox.Min.X, yMid, bbox.Min.Z),
                                        new XYZ(bbox.Max.X, yMid, bbox.Min.Z),
                                        new XYZ(bbox.Max.X, yMid, bbox.Max.Z),
                                        new XYZ(bbox.Min.X, yMid, bbox.Max.Z)
                                    };
                                }
                                else if (hostOrientation.Equals("Y", StringComparison.OrdinalIgnoreCase))
                                {
                                    // Y-Wall (Face is YZ plane) - Preserve Z-height
                                    double xMid = (bbox.Min.X + bbox.Max.X) / 2.0;
                                    corners = new List<XYZ>
                                    {
                                        new XYZ(xMid, bbox.Min.Y, bbox.Min.Z),
                                        new XYZ(xMid, bbox.Max.Y, bbox.Min.Z),
                                        new XYZ(xMid, bbox.Max.Y, bbox.Max.Z),
                                        new XYZ(xMid, bbox.Min.Y, bbox.Max.Z)
                                    };
                                }
                            }
                            
                            // Default Fallback (Floor or unknown) - Preserve Area
                            if (corners == null)
                            {
                                double z = bbox.Min.Z;
                                corners = new List<XYZ>
                                {
                                    new XYZ(bbox.Min.X, bbox.Min.Y, z),
                                    new XYZ(bbox.Max.X, bbox.Min.Y, z),
                                    new XYZ(bbox.Max.X, bbox.Max.Y, z),
                                    new XYZ(bbox.Min.X, bbox.Max.Y, z)
                                };
                            }
                            SafeFileLogger.SafeAppendTextAlways("geometry_extraction.log", $"[FALLBACK] Using BoundingBox corners (Orientation: {hostOrientation ?? "Floor"}) for sleeve {zone.SleeveInstanceId} (Solid extraction failed)\n");
                        }
                    }

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
                    SafeFileLogger.SafeAppendTextAlways("geometry_extraction_errors.log", $"Error extracting {zone.SleeveInstanceId}: {ex.Message}\n");
                }
            }

            // 2. Batch Save
            if (updates.Any())
            {
                _repository.BatchUpdateSleeveCorners(updates);
                SafeFileLogger.SafeAppendTextAlways("geometry_extraction.log", $"Extracted and saved corners for {updates.Count} sleeves.\n");
            }
            else
            {
                SafeFileLogger.SafeAppendTextAlways("geometry_extraction.log", "No sleeves found/extracted.\n");
            }

            return updates.Count;
        }

        private List<XYZ> ExtractCornersFromSolid(Element elem, Document doc, string hostOrientation)
        {
            // Get geometry - IncludeNonVisibleGeometry for VOIDS
            Options opt = new Options { 
                DetailLevel = ViewDetailLevel.Fine, 
                ComputeReferences = true, 
                IncludeNonVisibleObjects = true 
            };
            GeometryElement geo = elem.get_Geometry(opt);
            
            Solid solid = null;

            if (geo != null)
            {
                foreach (GeometryObject obj in geo)
                {
                    // REMOVED 's.Volume > 0' check to support Void-based families
                    if (obj is Solid s && s.Faces.Size > 0)
                    {
                        solid = s;
                        break;
                    }
                    else if (obj is GeometryInstance gi)
                    {
                        // Check symbol geometry (standard for family instances)
                        foreach (GeometryObject obj2 in gi.GetSymbolGeometry())
                        {
                             if (obj2 is Solid s2 && s2.Faces.Size > 0)
                             {
                                 solid = SolidUtils.CreateTransformed(s2, gi.Transform);
                                 break;
                             }
                        }
                        
                        // Fallback to instance geometry if symbol geometry lacks solids
                        if (solid == null)
                        {
                            foreach (GeometryObject obj2 in gi.GetInstanceGeometry())
                            {
                                if (obj2 is Solid s2 && s2.Faces.Size > 0)
                                {
                                    solid = s2;
                                    break;
                                }
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
            
            return GetOpeningProfileCorners(solid, hostOrientation);
        }

        private List<XYZ> GetOpeningProfileCorners(Solid solid, string hostOrientation)
        {
            Face profileFace = null;
            double maxArea = -1;

            // Determine target normal based on host orientation
            XYZ targetNormal = XYZ.BasisZ; // Default for floors
            bool isWall = false;

            if (!string.IsNullOrEmpty(hostOrientation))
            {
                if (hostOrientation.Equals("X", StringComparison.OrdinalIgnoreCase))
                {
                    targetNormal = XYZ.BasisX;
                    isWall = true;
                }
                else if (hostOrientation.Equals("Y", StringComparison.OrdinalIgnoreCase))
                {
                    targetNormal = XYZ.BasisY;
                    isWall = true;
                }
            }

            foreach (Face f in solid.Faces)
            {
                if (f is PlanarFace pf)
                {
                    // For walls, we want a face whose normal is parallel to the wall normal (X or Y)
                    // For floors, we want a face whose normal is vertical (Z)
                    bool isParallel = false;
                    if (isWall)
                    {
                        // Check if normal is parallel to X or Y (ignoring Z)
                        isParallel = Math.Abs(pf.FaceNormal.X) > 0.9 || Math.Abs(pf.FaceNormal.Y) > 0.9;
                    }
                    else
                    {
                        // Check if normal is vertical
                        isParallel = Math.Abs(pf.FaceNormal.Z) > 0.9;
                    }

                    if (isParallel)
                    {
                        if (f.Area > maxArea)
                        {
                            maxArea = f.Area;
                            profileFace = f;
                        }
                    }
                }
            }

            // Fallback: Just take the largest face if no parallel face found
            if (profileFace == null)
            {
                foreach (Face f in solid.Faces)
                {
                    if (f.Area > maxArea)
                    {
                        maxArea = f.Area;
                        profileFace = f;
                    }
                }
            }

            if (profileFace == null) return null;

            // Get vertices from outer loop
            List<XYZ> vertices = new List<XYZ>();
            foreach (EdgeArray loop in profileFace.EdgeLoops)
            {
                foreach (Edge edge in loop)
                {
                    XYZ p = edge.AsCurve().GetEndPoint(0);
                    if (!vertices.Any(v => v.IsAlmostEqualTo(p))) vertices.Add(p);
                }
            }
            
            // ✅ FIX: If we have exactly 4 vertices (Standard Rectangle), return them directly!
            // This preserves the ORIENTATION (OBB) instead of falling back to AABB.
            if (vertices.Count == 4)
            {
                double cx = vertices.Average(v => v.X);
                double cy = vertices.Average(v => v.Y);
                double cz = vertices.Average(v => v.Z);

                if (Math.Abs(vertices.Max(v => v.Z) - vertices.Min(v => v.Z)) < 0.01) // Horizontal face
                    return vertices.OrderBy(v => Math.Atan2(v.Y - cy, v.X - cx)).ToList();
                
                if (Math.Abs(vertices.Max(v => v.X) - vertices.Min(v => v.X)) < 0.01) // YZ plane (Vertical)
                    return vertices.OrderBy(v => Math.Atan2(v.Z - cz, v.Y - cy)).ToList();

                // XZ plane or slanted
                return vertices.OrderBy(v => Math.Atan2(v.Z - cz, v.X - cx)).ToList();
            }

            // If not exactly 4 vertices, calculate the bounding box of the face's vertices
            // This handles circular sleeves (1-2 vertices) and complex shapes
            double minX = vertices.Min(v => v.X);
            double maxX = vertices.Max(v => v.X);
            double minY = vertices.Min(v => v.Y);
            double maxY = vertices.Max(v => v.Y);
            double minZ = vertices.Min(v => v.Z);
            double maxZ = vertices.Max(v => v.Z);
            
            // For circular sleeves, we need more points to get an accurate bounding box
            // Triangulate the face to get a better representation of the boundary
            var mesh = profileFace.Triangulate();
            if (mesh != null)
            {
                foreach (XYZ p in mesh.Vertices)
                {
                    minX = Math.Min(minX, p.X);
                    maxX = Math.Max(maxX, p.X);
                    minY = Math.Min(minY, p.Y);
                    maxY = Math.Max(maxY, p.Y);
                    minZ = Math.Min(minZ, p.Z);
                    maxZ = Math.Max(maxZ, p.Z);
                }
            }

            // Return the 4 corners of the face's bounding box
            // Note: This is still a 2D face in 3D space. 
            // The ProximityHelper will extrude it into a 3D volume.
            
            // We need to decide which 4 corners to return. 
            // If it's a vertical face (wall), we return (minY, minZ) to (maxY, maxZ) etc.
            // For simplicity, we'll return 4 points that cover the min/max of the face.
            
            if (Math.Abs(maxZ - minZ) < 0.001) // Horizontal face
            {
                return new List<XYZ>
                {
                    new XYZ(minX, minY, minZ),
                    new XYZ(maxX, minY, minZ),
                    new XYZ(maxX, maxY, minZ),
                    new XYZ(minX, maxY, minZ)
                };
            }
            else if (Math.Abs(maxX - minX) < 0.001) // YZ face
            {
                return new List<XYZ>
                {
                    new XYZ(minX, minY, minZ),
                    new XYZ(minX, maxY, minZ),
                    new XYZ(minX, maxY, maxZ),
                    new XYZ(minX, minY, maxZ)
                };
            }
            else // XZ face or slanted
            {
                return new List<XYZ>
                {
                    new XYZ(minX, minY, minZ),
                    new XYZ(maxX, minY, minZ),
                    new XYZ(maxX, minY, maxZ),
                    new XYZ(minX, minY, maxZ)
                };
            }
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;

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
            SafeFileLogger.SafeAppendText("corner_extraction.log", $"[{DateTime.Now:HH:mm:ss}] [START] BATCH EXTRACTION (GetPlacedClashZones)...\n");

            // 1. Get all known placed sleeves from DB EFFICIENTLY in a single query
            var placedZones = _repository.GetPlacedClashZones();
            
            if (placedZones == null || !placedZones.Any())
            {
                SafeFileLogger.SafeAppendText("geometry_extraction.log", "No placed sleeves found in database.\n");
                SafeFileLogger.SafeAppendText("corner_extraction.log", $"[{DateTime.Now:HH:mm:ss}] [WARN] No placed sleeves found in database.\n");
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
            SafeFileLogger.SafeAppendText("corner_extraction.log",
                $"[{DateTime.Now:HH:mm:ss}] [START] BATCH EXTRACTION: Processing {zonesList.Count} zones...\n");

            if (!zonesList.Any())
            {
                SafeFileLogger.SafeAppendTextAlways("geometry_extraction.log", "No zones provided for extraction.\n");
                SafeFileLogger.SafeAppendText("corner_extraction.log", $"[{DateTime.Now:HH:mm:ss}] [WARN] No zones provided for extraction.\n");
                return 0;
            }

            return ExtractAndSaveCornersInternal(doc, zonesList);
        }

        /// <summary>
        /// After placement only: extracts corners from placement data (no pre-calculation, no Revit geometry API).
        /// Uses CalculatedSleeveWidth/Height, SleevePlacementPoint, and HostOrientation already set on each ClashZone.
        /// From the four corners, derives axis-aligned BoundingBoxMin/Max (min/max of corner X,Y,Z) and persists
        /// both SleeveCorner1-4 and SleeveBoundingBoxMin*/Max* to the database.
        /// </summary>
        public int ComputeAndSaveCornersFromPlacementData(List<ClashZone> zones)
        {
            if (zones == null || zones.Count == 0) return 0;

            var updates = new List<(Guid Guid, double c1x, double c1y, double c1z, double c2x, double c2y, double c2z, double c3x, double c3y, double c3z, double c4x, double c4y, double c4z)>();

            foreach (var zone in zones)
            {
                try
                {
                    double cx = zone.SleevePlacementPointX;
                    double cy = zone.SleevePlacementPointY;
                    double cz = zone.SleevePlacementPointZ;

                    // Use calculated dimensions (set during placement from planning DTO)
                    double halfW = zone.CalculatedSleeveWidth / 2.0;
                    double halfH = zone.CalculatedSleeveHeight / 2.0;

                    // For circular sleeves, use diameter for both dimensions
                    if (zone.CalculatedSleeveDiameter > 0 && halfW <= 0)
                    {
                        halfW = zone.CalculatedSleeveDiameter / 2.0;
                        halfH = zone.CalculatedSleeveDiameter / 2.0;
                    }

                    // Skip if no valid dimensions
                    if (halfW <= 0 && halfH <= 0) continue;

                    string hostOrientation = zone.HostOrientation;

                    double c1x, c1y, c1z, c2x, c2y, c2z, c3x, c3y, c3z, c4x, c4y, c4z;

                    if (!string.IsNullOrEmpty(hostOrientation) &&
                        hostOrientation.Equals("X", StringComparison.OrdinalIgnoreCase))
                    {
                        // X-wall: opening in XZ plane, Y constant
                        c1x = cx - halfW; c1y = cy; c1z = cz - halfH;
                        c2x = cx + halfW; c2y = cy; c2z = cz - halfH;
                        c3x = cx + halfW; c3y = cy; c3z = cz + halfH;
                        c4x = cx - halfW; c4y = cy; c4z = cz + halfH;
                    }
                    else if (!string.IsNullOrEmpty(hostOrientation) &&
                             hostOrientation.Equals("Y", StringComparison.OrdinalIgnoreCase))
                    {
                        // Y-wall: opening in YZ plane, X constant
                        c1x = cx; c1y = cy - halfW; c1z = cz - halfH;
                        c2x = cx; c2y = cy + halfW; c2z = cz - halfH;
                        c3x = cx; c3y = cy + halfW; c3z = cz + halfH;
                        c4x = cx; c4y = cy - halfW; c4z = cz + halfH;
                    }
                    else
                    {
                        // Floor/slab: opening in XY plane, Z constant
                        // Apply rotation if present
                        double angle = zone.MepElementRotationAngle;
                        if (Math.Abs(angle) > 1e-6)
                        {
                            double cos = Math.Cos(angle);
                            double sin = Math.Sin(angle);
                            // Rotated rectangle corners
                            c1x = cx + (-halfW * cos - (-halfH) * sin);
                            c1y = cy + (-halfW * sin + (-halfH) * cos);
                            c1z = cz;
                            c2x = cx + (halfW * cos - (-halfH) * sin);
                            c2y = cy + (halfW * sin + (-halfH) * cos);
                            c2z = cz;
                            c3x = cx + (halfW * cos - halfH * sin);
                            c3y = cy + (halfW * sin + halfH * cos);
                            c3z = cz;
                            c4x = cx + (-halfW * cos - halfH * sin);
                            c4y = cy + (-halfW * sin + halfH * cos);
                            c4z = cz;
                        }
                        else
                        {
                            c1x = cx - halfW; c1y = cy - halfH; c1z = cz;
                            c2x = cx + halfW; c2y = cy - halfH; c2z = cz;
                            c3x = cx + halfW; c3y = cy + halfH; c3z = cz;
                            c4x = cx - halfW; c4y = cy + halfH; c4z = cz;
                        }
                    }

                    // Queue corner update for DB
                    updates.Add((zone.Id, c1x, c1y, c1z, c2x, c2y, c2z, c3x, c3y, c3z, c4x, c4y, c4z));

                    // Also set on the in-memory object for immediate use by proximity checker
                    zone.SleeveCorner1X = c1x; zone.SleeveCorner1Y = c1y; zone.SleeveCorner1Z = c1z;
                    zone.SleeveCorner2X = c2x; zone.SleeveCorner2Y = c2y; zone.SleeveCorner2Z = c2z;
                    zone.SleeveCorner3X = c3x; zone.SleeveCorner3Y = c3y; zone.SleeveCorner3Z = c3z;
                    zone.SleeveCorner4X = c4x; zone.SleeveCorner4Y = c4y; zone.SleeveCorner4Z = c4z;

                    // Derive axis-aligned bbox from corners (no extra calculation — just min/max of corner coords) and persist.
                    double minX = Math.Min(Math.Min(c1x, c2x), Math.Min(c3x, c4x));
                    double maxX = Math.Max(Math.Max(c1x, c2x), Math.Max(c3x, c4x));
                    double minY = Math.Min(Math.Min(c1y, c2y), Math.Min(c3y, c4y));
                    double maxY = Math.Max(Math.Max(c1y, c2y), Math.Max(c3y, c4y));
                    double minZ = Math.Min(Math.Min(c1z, c2z), Math.Min(c3z, c4z));
                    double maxZ = Math.Max(Math.Max(c1z, c2z), Math.Max(c3z, c4z));

                    // Update in-memory ClashZone for immediate consumers
                    zone.SleeveBoundingBoxMinX = minX;
                    zone.SleeveBoundingBoxMinY = minY;
                    zone.SleeveBoundingBoxMinZ = minZ;
                    zone.SleeveBoundingBoxMaxX = maxX;
                    zone.SleeveBoundingBoxMaxY = maxY;
                    zone.SleeveBoundingBoxMaxZ = maxZ;

                    // Persist to database (also maintains R-tree index)
                    _repository.UpdateSleeveBoundingBoxes(zone.Id, minX, minY, minZ, maxX, maxY, maxZ);
                }
                catch (Exception ex)
                {
                    SafeFileLogger.SafeAppendText("corner_extraction.log",
                        $"[{DateTime.Now:HH:mm:ss}] [ERROR] Math corner computation failed for zone {zone.Id}: {ex.Message}\n");
                }
            }

            if (updates.Any())
            {
                _repository.BatchUpdateSleeveCorners(updates);
            }

            return updates.Count;
        }

        /// <summary>
        /// Extract corners from placed cluster sleeves in Revit and save to both ClusterSleeves and ClusterSleeves_v2 tables.
        /// </summary>
        public int ExtractAndSaveCornersForClusters(Document doc)
        {
            var clusters = _repository.GetPlacedClusterSleeves();
            if (clusters == null || !clusters.Any())
            {
                SafeFileLogger.SafeAppendText("corner_extraction.log", $"[{DateTime.Now:HH:mm:ss}] [CLUSTER] No placed cluster sleeves in DB.\n");
                return 0;
            }
            SafeFileLogger.SafeAppendText("corner_extraction.log", $"[{DateTime.Now:HH:mm:ss}] [CLUSTER] Extracting corners for {clusters.Count} cluster sleeves...\n");
            var updates = new List<(long ClusterInstanceId, double c1x, double c1y, double c1z, double c2x, double c2y, double c2z, double c3x, double c3y, double c3z, double c4x, double c4y, double c4z)>();
            foreach (var (clusterInstanceId, hostOrientation) in clusters)
            {
                try
                {
                    Element elem = doc.GetElement(ElementIdCompat.FromValue(clusterInstanceId));
                    if (elem == null || !(elem is FamilyInstance))
                    {
                        SafeFileLogger.SafeAppendText("corner_extraction.log", $"[{DateTime.Now:HH:mm:ss}] [CLUSTER] Cluster {clusterInstanceId} not found in Revit\n");
                        continue;
                    }

                    List<XYZ> corners = null;

                    // Same rule as individual sleeves: for HostOrientation X/Y (walls + framing),
                    // always use bbox-based WCS corners; do not depend on solid face selection.
                    if (!string.IsNullOrEmpty(hostOrientation) &&
                        (hostOrientation.Equals("X", StringComparison.OrdinalIgnoreCase) ||
                         hostOrientation.Equals("Y", StringComparison.OrdinalIgnoreCase)))
                    {
                        var bbox = elem.get_BoundingBox(null);
                        if (bbox != null)
                        {
                            if (hostOrientation.Equals("X", StringComparison.OrdinalIgnoreCase))
                            {
                                double yMid = (bbox.Min.Y + bbox.Max.Y) / 2.0;
                                corners = new List<XYZ>
                                {
                                    new XYZ(bbox.Min.X, yMid, bbox.Min.Z),
                                    new XYZ(bbox.Max.X, yMid, bbox.Min.Z),
                                    new XYZ(bbox.Max.X, yMid, bbox.Max.Z),
                                    new XYZ(bbox.Min.X, yMid, bbox.Max.Z)
                                };
                            }
                            else // "Y"
                            {
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
                        else
                        {
                            SafeFileLogger.SafeAppendText("corner_extraction.log",
                                $"[{DateTime.Now:HH:mm:ss}] [CLUSTER] ERROR: No bbox for cluster {clusterInstanceId} (HostOrientation={hostOrientation})\n");
                        }
                    }
                    else
                    {
                        corners = ExtractCornersFromSolid(elem, doc, hostOrientation);
                        if (corners == null || corners.Count != 4)
                        {
                            var bbox = elem.get_BoundingBox(null);
                            if (bbox != null)
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
                        }
                    }

                    if (corners != null && corners.Count == 4)
                    {
                        updates.Add((clusterInstanceId,
                            corners[0].X, corners[0].Y, corners[0].Z,
                            corners[1].X, corners[1].Y, corners[1].Z,
                            corners[2].X, corners[2].Y, corners[2].Z,
                            corners[3].X, corners[3].Y, corners[3].Z));
                    }
                }
                catch (Exception ex)
                {
                    SafeFileLogger.SafeAppendText("corner_extraction.log", $"[{DateTime.Now:HH:mm:ss}] [CLUSTER] Error cluster {clusterInstanceId}: {ex.Message}\n");
                }
            }
            if (updates.Any())
            {
                _repository.BatchUpdateClusterSleeveCorners(updates);
                SafeFileLogger.SafeAppendText("corner_extraction.log", $"[{DateTime.Now:HH:mm:ss}] [CLUSTER] Saved corners for {updates.Count} cluster sleeves (ClusterSleeves + ClusterSleeves_v2).\n");
            }
            return updates.Count;
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
                    Element elem = doc.GetElement(ElementIdCompat.FromValue(zone.SleeveInstanceId));
                    if (elem == null || !(elem is FamilyInstance))
                    {
                        SafeFileLogger.SafeAppendText("corner_extraction.log", $"[{DateTime.Now:HH:mm:ss}] [WARN] Sleeve {zone.SleeveInstanceId} not found in Revit\n");
                        continue;
                    }

                    string hostOrientation = zone.HostOrientation;

                    List<XYZ> corners = null;

                    // ✅ SINGLE SOURCE OF TRUTH FOR WALL/FRAMING:
                    // If HostOrientation is X/Y, always compute corners from the element's bounding box in WCS,
                    // just like walls, and do NOT rely on solid face selection (which can pick horizontal faces).
                    if (!string.IsNullOrEmpty(hostOrientation) &&
                        (hostOrientation.Equals("X", StringComparison.OrdinalIgnoreCase) ||
                         hostOrientation.Equals("Y", StringComparison.OrdinalIgnoreCase)))
                    {
                        var bbox = elem.get_BoundingBox(null);
                        if (bbox != null)
                        {
                            if (hostOrientation.Equals("X", StringComparison.OrdinalIgnoreCase))
                            {
                                // X-wall / X-framing: opening in XZ plane → Y constant, X/Z vary
                                double yMid = (bbox.Min.Y + bbox.Max.Y) / 2.0;
                                corners = new List<XYZ>
                                {
                                    new XYZ(bbox.Min.X, yMid, bbox.Min.Z),
                                    new XYZ(bbox.Max.X, yMid, bbox.Min.Z),
                                    new XYZ(bbox.Max.X, yMid, bbox.Max.Z),
                                    new XYZ(bbox.Min.X, yMid, bbox.Max.Z)
                                };
                            }
                            else // hostOrientation == "Y"
                            {
                                // Y-wall / Y-framing: opening in YZ plane → X constant, Y/Z vary
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
                        else
                        {
                            SafeFileLogger.SafeAppendText("corner_extraction.log", $"[{DateTime.Now:HH:mm:ss}] [ERROR] Failed to extract corners for sleeve {zone.SleeveInstanceId}\n");
                        }
                    }
                    else
                    {
                        // Floor / unknown: keep existing solid-based logic + XY fallback
                        corners = ExtractCornersFromSolid(elem, doc, hostOrientation);

                        if (corners == null || corners.Count != 4)
                        {
                            var bbox = elem.get_BoundingBox(null);
                            if (bbox != null)
                            {
                                double z = bbox.Min.Z;
                                corners = new List<XYZ>
                                {
                                    new XYZ(bbox.Min.X, bbox.Min.Y, z),
                                    new XYZ(bbox.Max.X, bbox.Min.Y, z),
                                    new XYZ(bbox.Max.X, bbox.Max.Y, z),
                                    new XYZ(bbox.Min.X, bbox.Max.Y, z)
                                };

                                SafeFileLogger.SafeAppendTextAlways("geometry_extraction.log",
                                    $"[FALLBACK] Using BoundingBox corners (Orientation: {hostOrientation ?? "Floor"}) for sleeve {zone.SleeveInstanceId} (Solid extraction failed)\n");
                                SafeFileLogger.SafeAppendText("corner_extraction.log",
                                    $"[{DateTime.Now:HH:mm:ss}] [WARN] Sleeve {zone.SleeveInstanceId}: Solid extraction failed, used BoundingBox fallback\n");
                            }
                            else
                            {
                                SafeFileLogger.SafeAppendText("corner_extraction.log",
                                    $"[{DateTime.Now:HH:mm:ss}] [ERROR] Failed to extract corners for sleeve {zone.SleeveInstanceId}\n");
                            }
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

                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            SafeFileLogger.SafeAppendText("corner_extraction.log", 
                                $"[{DateTime.Now:HH:mm:ss}] [DEBUG] Successfully extracted corners for sleeve {zone.SleeveInstanceId} (Zone: {zone.Id})\n");
                        }
                    }
                }
                catch (Exception ex)
                {
                    SafeFileLogger.SafeAppendTextAlways("geometry_extraction_errors.log", $"Error extracting {zone.SleeveInstanceId}: {ex.Message}\n");
                    SafeFileLogger.SafeAppendText("corner_extraction.log", $"[{DateTime.Now:HH:mm:ss}] [ERROR] Exception for sleeve {zone.SleeveInstanceId}: {ex.Message}\n");
                }
            }

            // 2. Batch Save
            int failedCount = zones.Count - updates.Count;
            if (updates.Any())
            {
                _repository.BatchUpdateSleeveCorners(updates);
                SafeFileLogger.SafeAppendTextAlways("geometry_extraction.log", $"Extracted and saved corners for {updates.Count} sleeves.\n");
                SafeFileLogger.SafeAppendText("corner_extraction.log",
                    $"[{DateTime.Now:HH:mm:ss}] [SUCCESS] BATCH EXTRACTION FINISHED: {updates.Count} extracted, {failedCount} failed\n");
            }
            else
            {
                SafeFileLogger.SafeAppendTextAlways("geometry_extraction.log", "No sleeves found/extracted.\n");
                SafeFileLogger.SafeAppendText("corner_extraction.log",
                    $"[{DateTime.Now:HH:mm:ss}] [SUCCESS] BATCH EXTRACTION FINISHED: 0 extracted, {zones.Count} failed\n");
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

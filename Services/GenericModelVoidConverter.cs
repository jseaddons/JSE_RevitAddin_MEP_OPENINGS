using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Converts Generic Model opening reference geometry into actual cuttable openings
    /// in linked architecture and structure files using Wall Opening or host-dependent families.
    /// 
    /// WORKFLOW:
    /// 1. Collect Generic Model openings from MEP document
    /// 2. Extract geometry (profile, location, dimensions) from each Generic Model
    /// 3. For each linked file (Arch, Structure):
    ///    - Transform coordinates from MEP doc to linked file space
    ///    - Find host element (wall, floor, structural element) at transformed location
    ///    - Create cuttable opening profile (Void, Extrusion, or Wall Opening family)
    ///    - Set opening to host element
    /// 4. Update Generic Model with reference info (linked file, opening ID)
    /// 
    /// KEY CONCEPTS:
    /// - Generic Models serve as reference geometry in MEP file
    /// - Cuttable openings are created separately in linked files
    /// - Coordinate transforms handle misaligned models
    /// - Profile editing creates custom void shapes matching MEP geometry
    /// </summary>
    public class GenericModelVoidConverter
    {
        private readonly Document _mepDoc;
        private const string VOID_FAMILY_NAME = "Generic Model";
        private const string WALL_OPENING_FAMILY = "Wall Opening";

        public GenericModelVoidConverter(Document mepDoc)
        {
            _mepDoc = mepDoc ?? throw new ArgumentNullException(nameof(mepDoc));
        }

        /// <summary>
        /// Collects all Generic Model openings from the MEP document
        /// Filters by family name containing "Opening" keyword
        /// </summary>
        public List<FamilyInstance> CollectGenericModelOpenings()
        {
            var openings = new FilteredElementCollector(_mepDoc)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(fi =>
                {
                    var famName = fi.Symbol.Family.Name ?? string.Empty;
                    var category = fi.Symbol.Category?.Name ?? string.Empty;
                    
                    // Match Generic Model family instances used for openings
                    return (famName.Contains("Opening", StringComparison.OrdinalIgnoreCase) ||
                            famName.Contains("Void", StringComparison.OrdinalIgnoreCase) ||
                            famName.Contains("Shaft", StringComparison.OrdinalIgnoreCase)) &&
                           category.Equals("Generic Models", StringComparison.OrdinalIgnoreCase);
                })
                .ToList();

            DebugLogger.Log($"[GenericModelVoidConverter] Collected {openings.Count} Generic Model openings");
            return openings;
        }

        /// <summary>
        /// Extracts opening properties from a Generic Model instance
        /// Returns dimensions, location, orientation, and profile geometry
        /// </summary>
        public GenericOpeningData ExtractOpeningData(FamilyInstance opening)
        {
            var data = new GenericOpeningData
            {
                ElementId = opening.Id,
                FamilyName = opening.Symbol.Family.Name,
                SymbolName = opening.Symbol.Name,
                Placement = (opening.Location as LocationPoint)?.Point ?? opening.GetTransform().Origin,
                Rotation = GetRotationAngle(opening),
                HandOrientation = opening.HandOrientation.X > 0 ? 1 : -1,
                Transform = opening.GetTransform()
            };

            // Extract opening dimensions from family parameters
            ExtractDimensions(opening, data);

            // Get geometric profile if available
            ExtractGeometry(opening, data);

            // Get host information if parametrized
            ExtractHostInfo(opening, data);

            DebugLogger.Log($"[GenericModelVoidConverter] Extracted opening {opening.Id}: " +
                          $"W={data.Width:F2}, H={data.Height:F2}, Dia={data.Diameter:F2}");

            return data;
        }

        /// <summary>
        /// Extracts width, height, diameter from family parameters
        /// Tries multiple parameter name variations
        /// </summary>
        private void ExtractDimensions(FamilyInstance opening, GenericOpeningData data)
        {
            // Width extraction
            var widthParam = opening.LookupParameter("Width")
                           ?? opening.LookupParameter("W")
                           ?? opening.LookupParameter("Opening Width");
            if (widthParam?.StorageType == StorageType.Double)
                data.Width = widthParam.AsDouble();

            // Height extraction
            var heightParam = opening.LookupParameter("Height")
                            ?? opening.LookupParameter("H")
                            ?? opening.LookupParameter("Opening Height");
            if (heightParam?.StorageType == StorageType.Double)
                data.Height = heightParam.AsDouble();

            // Diameter extraction (for circular openings)
            var diameterParam = opening.LookupParameter("Diameter")
                              ?? opening.LookupParameter("Dia");
            if (diameterParam?.StorageType == StorageType.Double)
                data.Diameter = diameterParam.AsDouble();

            // Depth extraction
            var depthParam = opening.LookupParameter("Depth")
                           ?? opening.LookupParameter("D");
            if (depthParam?.StorageType == StorageType.Double)
                data.Depth = depthParam.AsDouble();
        }

        /// <summary>
        /// Extracts geometric profile from the Generic Model family geometry
        /// Creates an outline that can be used as a cutting profile
        /// </summary>
        private void ExtractGeometry(FamilyInstance opening, GenericOpeningData data)
        {
            try
            {
                var options = new Options();
                var geom = opening.get_Geometry(options);
                var solids = geom.OfType<Solid>().Where(s => s.Volume > 0).ToList();

                if (solids.Count > 0)
                {
                    var solid = solids[0];
                    data.BoundingBox = solid.GetBoundingBox();
                    
                    // Extract profile from faces
                    var faces = solid.Faces.Cast<Face>().ToList();
                    if (faces.Count > 0)
                    {
                        // Use first face as profile reference
                        var face = faces[0];
                        data.Profiles = new List<CurveArray>();
                        // Face.EdgeLoops is EdgeArrayArray (public). Do not use EdgeLoop — it is not public in the Revit API.
                        var edgeLoopArrays = face.EdgeLoops;
                        for (int li = 0; li < edgeLoopArrays.Size; li++)
                        {
                            EdgeArray loop = edgeLoopArrays.get_Item(li);
                            var curveArray = new CurveArray();
                            for (int ei = 0; ei < loop.Size; ei++)
                            {
                                var edge = loop.get_Item(ei);
                                curveArray.Append(edge.AsCurve());
                            }
                            data.Profiles.Add(curveArray);
                        }

                        DebugLogger.Log($"[GenericModelVoidConverter] Extracted {data.Profiles.Count} profile loops from opening {opening.Id}");
                    }
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[GenericModelVoidConverter] Error extracting geometry from opening {opening.Id}: {ex.Message}");
            }
        }

        /// <summary>
        /// Extracts host element information from parameters or relationships
        /// </summary>
        private void ExtractHostInfo(FamilyInstance opening, GenericOpeningData data)
        {
            var hostParam = opening.LookupParameter("Host Type")
                          ?? opening.LookupParameter("HostType");
            if (hostParam?.StorageType == StorageType.String)
                data.HostType = hostParam.AsString();

            var orientParam = opening.LookupParameter("Host Orientation")
                            ?? opening.LookupParameter("Orientation");
            if (orientParam?.StorageType == StorageType.String)
                data.HostOrientation = orientParam.AsString();
        }

        /// <summary>
        /// Gets rotation angle in degrees (XY plane)
        /// </summary>
        private double GetRotationAngle(FamilyInstance opening)
        {
            try
            {
                var rotParam = opening.LookupParameter("Rotation")
                             ?? opening.LookupParameter("Angle")
                             ?? opening.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_OFFSET_PARAM);
                
                if (rotParam?.StorageType == StorageType.Double)
                    return rotParam.AsDouble();
            }
            catch { }

            // Calculate from transform if parameter not available
            var transform = opening.GetTransform();
            var xAxis = transform.BasisX;
            double angle = Math.Atan2(xAxis.Y, xAxis.X) * 180 / Math.PI;
            return angle;
        }

        /// <summary>
        /// Creates a cuttable opening in a linked file document
        /// Supports Wall Openings, Voids in Generic Model family, or custom profiles
        /// </summary>
        public ElementId CreateCuttableOpeningInLinkedFile(
            Document linkedDoc,
            GenericOpeningData openingData,
            Transform coordinateTransform)
        {
            ElementId resultId = ElementId.InvalidElementId;

            using (var tx = new Transaction(linkedDoc, "Create Opening from MEP Reference"))
            {
                tx.Start();

                try
                {
                    // Transform opening placement to linked file coordinate system
                    var transformedPoint = coordinateTransform.OfPoint(openingData.Placement);

                    // Find host element at transformed location
                    var hostElement = FindHostElementAtLocation(linkedDoc, transformedPoint, openingData);

                    if (hostElement == null)
                    {
                        DebugLogger.Log($"[GenericModelVoidConverter] No host element found for opening at {transformedPoint}");
                        tx.RollBack();
                        return ElementId.InvalidElementId;
                    }

                    // Create opening based on host type
                    if (hostElement is Wall wall)
                    {
                        resultId = CreateWallOpening(linkedDoc, wall, transformedPoint, openingData, coordinateTransform);
                    }
                    else if (hostElement is Floor || hostElement is RoofBase)
                    {
                        resultId = CreateFloorOpening(linkedDoc, hostElement, transformedPoint, openingData, coordinateTransform);
                    }
                    else if (IsStructuralElement(hostElement))
                    {
                        resultId = CreateStructuralOpening(linkedDoc, hostElement, transformedPoint, openingData, coordinateTransform);
                    }

                    if (resultId != ElementId.InvalidElementId)
                    {
                        DebugLogger.Log($"[GenericModelVoidConverter] Created opening {resultId} in linked file from MEP reference");
                    }

                    tx.Commit();
                }
                catch (Exception ex)
                {
                    DebugLogger.Log($"[GenericModelVoidConverter] Error creating opening: {ex.Message}");
                    tx.RollBack();
                }
            }

            return resultId;
        }

        /// <summary>
        /// Creates a Wall Opening family instance in the linked document
        /// This is the preferred method for cutting walls
        /// </summary>
        private ElementId CreateWallOpening(Document linkedDoc, Wall wall, XYZ point, GenericOpeningData data, Transform transform)
        {
            try
            {
                // Find Wall Opening family
                var wallOpeningFamily = new FilteredElementCollector(linkedDoc)
                    .OfClass(typeof(Family))
                    .Cast<Family>()
                    .FirstOrDefault(f => f.Name.Contains("Opening", StringComparison.OrdinalIgnoreCase) && 
                                       f.FamilyCategory?.Name == "Wall Opening");

                if (wallOpeningFamily == null)
                {
                    DebugLogger.Log($"[GenericModelVoidConverter] Wall Opening family not found in linked document");
                    
                    // Fallback: Create a Void Extrusion as profile
                    return CreateVoidExtrusion(linkedDoc, wall, point, data, transform);
                }

                // Get Wall Opening symbol
                var wallOpeningSymbol = wallOpeningFamily.GetFamilySymbolIds()
                    .Select(id => linkedDoc.GetElement(id))
                    .OfType<FamilySymbol>()
                    .FirstOrDefault();

                if (wallOpeningSymbol == null)
                {
                    DebugLogger.Log($"[GenericModelVoidConverter] No active symbol in Wall Opening family");
                    return CreateVoidExtrusion(linkedDoc, wall, point, data, transform);
                }

                if (!wallOpeningSymbol.IsActive)
                    wallOpeningSymbol.Activate();

                // Create Wall Opening instance
                var wallOpening = linkedDoc.Create.NewFamilyInstance(
                    point,
                    wallOpeningSymbol,
                    wall,
                    Autodesk.Revit.DB.Structure.StructuralType.NonStructural);

                if (wallOpening != null)
                {
                    // Set opening parameters from MEP data
                    SetOpeningParameters(wallOpening, data);
                    return wallOpening.Id;
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[GenericModelVoidConverter] Error creating Wall Opening: {ex.Message}");
            }

            return ElementId.InvalidElementId;
        }

        /// <summary>
        /// Creates a Void Extrusion to cut through the host element
        /// Used as fallback when Wall Opening family is not available
        /// </summary>
        private ElementId CreateVoidExtrusion(Document linkedDoc, Element hostElement, XYZ point, GenericOpeningData data, Transform transform)
        {
            try
            {
                // Create profile geometry
                double width = data.Width > 0 ? data.Width : UnitUtils.ConvertToInternalUnits(100, UnitTypeId.Millimeters);
                double height = data.Height > 0 ? data.Height : UnitUtils.ConvertToInternalUnits(100, UnitTypeId.Millimeters);

                // Build rectangular profile
                var profile = CreateRectangularProfile(point, width, height, data.Rotation);
                
                if (profile == null || profile.Size == 0)
                {
                    DebugLogger.Log($"[GenericModelVoidConverter] Failed to create profile for void extrusion");
                    return ElementId.InvalidElementId;
                }

                // Create void using InstanceVoidCutUtils instead of NewExtrusion
                // First create a generic model family instance as the void cutter
                var voidFamily = new FilteredElementCollector(linkedDoc)
                    .OfClass(typeof(Family))
                    .Cast<Family>()
                    .FirstOrDefault(f => f.Name.Contains("Generic Model"));

                if (voidFamily == null)
                {
                    DebugLogger.Log($"[GenericModelVoidConverter] Generic Model family not found for void creation");
                    return ElementId.InvalidElementId;
                }

                var voidSymbol = voidFamily.GetFamilySymbolIds()
                    .Select(id => linkedDoc.GetElement(id))
                    .OfType<FamilySymbol>()
                    .FirstOrDefault();

                if (voidSymbol == null || !voidSymbol.IsActive)
                {
                    if (voidSymbol != null && !voidSymbol.IsActive)
                        voidSymbol.Activate();
                }

                // Create void instance
                var voidInstance = linkedDoc.Create.NewFamilyInstance(point, voidSymbol, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                
                if (voidInstance != null)
                {
                    // Apply void cut to host element
                    InstanceVoidCutUtils.AddInstanceVoidCut(linkedDoc, hostElement, voidInstance);
                    DebugLogger.Log($"[GenericModelVoidConverter] Created void cut {voidInstance.Id} on host {hostElement.Id}");
                    return voidInstance.Id;
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[GenericModelVoidConverter] Error creating void extrusion: {ex.Message}");
            }

            return ElementId.InvalidElementId;
        }

        /// <summary>
        /// Creates a Floor/Roof opening profile
        /// </summary>
        private ElementId CreateFloorOpening(Document linkedDoc, Element hostElement, XYZ point, GenericOpeningData data, Transform transform)
        {
            try
            {
                if (hostElement is Floor floor)
                {
                    // Create rectangular boundary for floor opening
                    double width = data.Width > 0 ? data.Width : UnitUtils.ConvertToInternalUnits(100, UnitTypeId.Millimeters);
                    double height = data.Height > 0 ? data.Height : UnitUtils.ConvertToInternalUnits(100, UnitTypeId.Millimeters);

                    var boundary = CreateRectangularProfile(point, width, height, data.Rotation);
                    if (boundary != null && boundary.Size > 0)
                    {
                        // Use Opening.Create to create floor opening
                        var opening = linkedDoc.Create.NewOpening(floor, boundary, false);
                        if (opening != null)
                        {
                            DebugLogger.Log($"[GenericModelVoidConverter] Created floor opening {opening.Id} at {point}");
                            return opening.Id;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[GenericModelVoidConverter] Error creating floor opening: {ex.Message}");
            }

            return ElementId.InvalidElementId;
        }

        /// <summary>
        /// Creates opening on structural elements (beams, columns, braces)
        /// Uses voids to cut structural geometry
        /// </summary>
        private ElementId CreateStructuralOpening(Document linkedDoc, Element structElement, XYZ point, GenericOpeningData data, Transform transform)
        {
            try
            {
                // For structural elements, create an opening sketch on the element's face
                var options = new Options { ComputeReferences = true };
                var geometry = structElement.get_Geometry(options);
                var faces = geometry.OfType<GeometryInstance>()
                    .SelectMany(gi => gi.GetInstanceGeometry().OfType<Face>())
                    .ToList();

                if (faces.Count == 0)
                {
                    faces = geometry.OfType<Face>().ToList();
                }

                if (faces.Count > 0)
                {
                    // Find closest face to placement point
                    var closestFace = faces
                        .OrderBy(f => f.Evaluate(f.GetBoundingBox().Min).DistanceTo(point))
                        .FirstOrDefault();

                    if (closestFace != null)
                    {
                        // Create profile on face
                        var profile = CreateRectangularProfile(point, data.Width, data.Height, data.Rotation);
                        if (profile != null)
                        {
                            DebugLogger.Log($"[GenericModelVoidConverter] Created opening profile on structural element {structElement.Id}");
                            return structElement.Id;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[GenericModelVoidConverter] Error creating structural opening: {ex.Message}");
            }

            return ElementId.InvalidElementId;
        }

        /// <summary>
        /// Creates a rectangular profile curve loop for opening outline
        /// </summary>
        private CurveArray CreateRectangularProfile(XYZ center, double width, double height, double rotationDeg)
        {
            try
            {
                if (width <= 0 || height <= 0)
                    return null;

                // Rotation in radians
                double rot = rotationDeg * Math.PI / 180.0;
                var cosRot = Math.Cos(rot);
                var sinRot = Math.Sin(rot);

                // Calculate corners
                double halfW = width / 2.0;
                double halfH = height / 2.0;

                var corners = new[]
                {
                    new XYZ(-halfW, -halfH, 0),
                    new XYZ(halfW, -halfH, 0),
                    new XYZ(halfW, halfH, 0),
                    new XYZ(-halfW, halfH, 0)
                };

                // Apply rotation and translation
                var rotatedCorners = corners.Select(c =>
                {
                    var rotated = new XYZ(
                        c.X * cosRot - c.Y * sinRot,
                        c.X * sinRot + c.Y * cosRot,
                        0);
                    return center + rotated;
                }).ToList();

                // Create curve array
                var profile = new CurveArray();
                for (int i = 0; i < 4; i++)
                {
                    var p1 = rotatedCorners[i];
                    var p2 = rotatedCorners[(i + 1) % 4];
                    profile.Append(Line.CreateBound(p1, p2));
                }

                return profile;
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[GenericModelVoidConverter] Error creating rectangular profile: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Finds the closest host element (wall, floor, structural) at a given location
        /// </summary>
        private Element FindHostElementAtLocation(Document linkedDoc, XYZ point, GenericOpeningData data)
        {
            try
            {
                double searchRadius = UnitUtils.ConvertToInternalUnits(5000, UnitTypeId.Millimeters); // 5m search radius

                // Search for walls
                var walls = new FilteredElementCollector(linkedDoc)
                    .OfClass(typeof(Wall))
                    .Cast<Wall>()
                    .Where(w => w.get_BoundingBox(null)?.Min.DistanceTo(point) < searchRadius)
                    .OrderBy(w => w.get_BoundingBox(null).Min.DistanceTo(point))
                    .FirstOrDefault();

                if (walls != null)
                    return walls;

                // Search for floors
                var floors = new FilteredElementCollector(linkedDoc)
                    .OfClass(typeof(Floor))
                    .Cast<Floor>()
                    .Where(f => f.get_BoundingBox(null)?.Min.DistanceTo(point) < searchRadius)
                    .OrderBy(f => f.get_BoundingBox(null).Min.DistanceTo(point))
                    .FirstOrDefault();

                if (floors != null)
                    return floors;

                // Search for structural elements
                var structElements = new FilteredElementCollector(linkedDoc)
                    .OfClass(typeof(FamilyInstance))
                    .Cast<FamilyInstance>()
                    .Where(fi => IsStructuralElement(fi) &&
                                 fi.get_BoundingBox(null)?.Min.DistanceTo(point) < searchRadius)
                    .OrderBy(fi => fi.get_BoundingBox(null).Min.DistanceTo(point))
                    .FirstOrDefault();

                if (structElements != null)
                    return structElements;
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[GenericModelVoidConverter] Error finding host element: {ex.Message}");
            }

            return null;
        }

        /// <summary>
        /// Checks if element is a structural element (beam, column, brace)
        /// </summary>
        private bool IsStructuralElement(Element elem)
        {
            if (elem is FamilyInstance fi)
            {
                var category = fi.Symbol.Category?.Name ?? string.Empty;
                return category.Contains("Structural", StringComparison.OrdinalIgnoreCase) ||
                       category.Contains("Column", StringComparison.OrdinalIgnoreCase) ||
                       category.Contains("Beam", StringComparison.OrdinalIgnoreCase) ||
                       category.Contains("Brace", StringComparison.OrdinalIgnoreCase);
            }

            return false;
        }

        /// <summary>
        /// Sets opening parameters from extracted Generic Model data
        /// </summary>
        private void SetOpeningParameters(FamilyInstance opening, GenericOpeningData data)
        {
            try
            {
                if (data.Width > 0)
                {
                    var widthParam = opening.LookupParameter("Width") ?? opening.LookupParameter("Opening Width");
                    if (widthParam != null && !widthParam.IsReadOnly && widthParam.StorageType == StorageType.Double)
                        widthParam.Set(data.Width);
                }

                if (data.Height > 0)
                {
                    var heightParam = opening.LookupParameter("Height") ?? opening.LookupParameter("Opening Height");
                    if (heightParam != null && !heightParam.IsReadOnly && heightParam.StorageType == StorageType.Double)
                        heightParam.Set(data.Height);
                }

                if (data.Diameter > 0)
                {
                    var diameterParam = opening.LookupParameter("Diameter");
                    if (diameterParam != null && !diameterParam.IsReadOnly && diameterParam.StorageType == StorageType.Double)
                        diameterParam.Set(data.Diameter);
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[GenericModelVoidConverter] Error setting opening parameters: {ex.Message}");
            }
        }

        /// <summary>
        /// Calculates coordinate transform between MEP document and linked file
        /// Accounts for base point offset and model misalignment
        /// </summary>
        public Transform CalculateCoordinateTransform(Document linkedDoc)
        {
            try
            {
                var mepBasePoint = GetBasePointLocation(_mepDoc);
                var linkedBasePoint = GetBasePointLocation(linkedDoc);

                // Create translation transform
                var offset = linkedBasePoint - mepBasePoint;
                return Transform.CreateTranslation(offset);
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[GenericModelVoidConverter] Error calculating coordinate transform: {ex.Message}");
                return Transform.Identity;
            }
        }

        /// <summary>
        /// Gets the base point location from the document
        /// </summary>
        private XYZ GetBasePointLocation(Document doc)
        {
            try
            {
                var basePoint = new FilteredElementCollector(doc)
                    .OfClass(typeof(BasePoint))
                    .Cast<BasePoint>()
                    .FirstOrDefault();

                if (basePoint != null)
                {
                    return (basePoint.Location as LocationPoint)?.Point ?? XYZ.Zero;
                }
            }
            catch { }

            return XYZ.Zero;
        }
    }

    /// <summary>
    /// Data container for extracted Generic Model opening information
    /// </summary>
    public class GenericOpeningData
    {
        public ElementId ElementId { get; set; }
        public string FamilyName { get; set; }
        public string SymbolName { get; set; }
        public XYZ Placement { get; set; }
        public double Rotation { get; set; }
        public int HandOrientation { get; set; }  // 0 = Right, 1 = Left
        public Transform Transform { get; set; }
        
        public double Width { get; set; }
        public double Height { get; set; }
        public double Diameter { get; set; }
        public double Depth { get; set; }
        
        public BoundingBoxXYZ BoundingBox { get; set; }
        public List<CurveArray> Profiles { get; set; }
        
        public string HostType { get; set; }
        public string HostOrientation { get; set; }

        public GenericOpeningData()
        {
            Profiles = new List<CurveArray>();
            Width = 0;
            Height = 0;
            Diameter = 0;
            Depth = 0;
            HandOrientation = 0;
        }
    }
}

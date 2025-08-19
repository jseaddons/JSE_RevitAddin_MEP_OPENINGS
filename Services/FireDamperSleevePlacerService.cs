using System;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Structure;




namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Unified service for placing fire damper/duct accessory sleeves.
    /// Supports both production and debug modes.
    /// </summary>
    public class FireDamperSleevePlacerService
    {
        private readonly Document _doc;
        private readonly bool _enableDebugLogging;

        public FireDamperSleevePlacerService(Document doc, bool enableDebugLogging = false)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
            _enableDebugLogging = enableDebugLogging;
        }

        private void Log(string message)
        {
            if (_enableDebugLogging)
            {
                DamperLogger.Log(message);
            }
        }

        /// <summary>
        /// Gets the connector side using world coordinates (for MSFD dampers).
        /// </summary>
        public static string GetConnectorSideWorld(FamilyInstance damper, out Connector? connector)
        {
            connector = null;

            var cm = damper.MEPModel?.ConnectorManager;
            if (cm == null) return "Right";

            // Pick the connector with the largest absolute component
            Connector? best = null;
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

            if (best == null) return "Right";
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
        private string GetConnectorSideLocal(FamilyInstance damper, out Connector? connector)
        {
            connector = null;
            var cons = (damper.MEPModel as MechanicalFitting)?.ConnectorManager?.Connectors
                    ?? damper.MEPModel?.ConnectorManager?.Connectors;

            if (cons == null || cons.Size == 0)
            {
                Log("[GetConnectorSideLocal] No connectors found. Returning 'Unknown'.");
                return "Unknown";
            }

            // Debug logging for connectors
            if (_enableDebugLogging)
            {
                foreach (Connector c in cons)
                {
                    Log($"[GetConnectorSideLocal] Connector at {c.Origin}, BasisX={c.CoordinateSystem.BasisX}, BasisY={c.CoordinateSystem.BasisY}, BasisZ={c.CoordinateSystem.BasisZ}");
                }
            }

            // Pick the connector closest to origin
            connector = cons.Cast<Connector>()
                             .OrderBy(c => c.Origin.DistanceTo(damper.GetTransform().Origin))
                             .FirstOrDefault();

            if (connector == null)
            {
                Log("[GetConnectorSideLocal] Closest connector is null. Returning 'Unknown'.");
                return "Unknown";
            }

            var damperT = damper.GetTransform();
            var connDir = connector.CoordinateSystem.BasisX; // connector normal

            // Dot products with damper local axes
            double dotX = connDir.DotProduct(damperT.BasisX);
            double dotY = connDir.DotProduct(damperT.BasisY);
            double dotZ = connDir.DotProduct(damperT.BasisZ);

            const double tol = 0.9;

            if (Math.Abs(dotX) >= tol) return dotX > 0 ? "Right" : "Left";
            if (Math.Abs(dotZ) >= tol) return dotZ > 0 ? "Top" : "Bottom";
            
            Log("[GetConnectorSideLocal] Could not determine side, returning 'Unknown'.");
            return "Unknown";
        }

        /// <summary>
        /// Determines if a wall is oriented along the Y-axis.
        /// Y-axis wall: wall runs parallel to Y-axis, so its normal is in X direction.
        /// </summary>
        private static bool IsYAxisWall(Wall wall)
        {
            XYZ wallOrientation = wall.Orientation;
            return Math.Abs(wallOrientation.Y) < 0.1 && Math.Abs(wallOrientation.X) > 0.9;
        }

        /// <summary>
        /// Gets the correct world offset direction based on connector side and wall orientation.
        /// </summary>
        private XYZ GetOffsetDirection(string side, bool isYAxisWall)
        {
            if (isYAxisWall)
            {
                // Y-axis wall: Right/Left are in Y direction, Top/Bottom in Z direction
                return side switch
                {
                    "Right" => XYZ.BasisY,   // Positive Y direction
                    "Left" => -XYZ.BasisY,   // Negative Y direction
                    "Top" => XYZ.BasisZ,     // Positive Z direction
                    "Bottom" => -XYZ.BasisZ, // Negative Z direction
                    _ => XYZ.BasisX          // Fallback
                };
            }
            else
            {
                // X-axis wall: Right/Left are in X direction, Top/Bottom in Z direction
                return side switch
                {
                    "Right" => XYZ.BasisX,   // Positive X direction
                    "Left" => -XYZ.BasisX,   // Negative X direction
                    "Top" => XYZ.BasisZ,     // Positive Z direction
                    "Bottom" => -XYZ.BasisZ, // Negative Z direction
                    _ => XYZ.BasisX          // Fallback
                };
            }
        }

        /// <summary>
        /// Places a sleeve for a fire damper accessory.
        /// </summary>
        /// <param name="accessory">The fire damper element</param>
        /// <param name="sleeveSymbol">The sleeve family symbol to place</param>
        /// <param name="linkedWall">The wall to place the sleeve in</param>
        /// <returns>True if sleeve was placed successfully, false if skipped or failed</returns>
        public bool PlaceFireDamperSleeve(
            FamilyInstance accessory,
            FamilySymbol sleeveSymbol,
            Wall linkedWall,
            Transform? accessoryTransform = null,
            Transform? wallTransform = null)
        {
            Log("FireDamperSleevePlacerService: Starting placement.");

            try
            {
                // Get bounding boxes for the damper and the wall
                var damperBBox = accessory.get_BoundingBox(null);
                var wallBBox = linkedWall.get_BoundingBox(null);

                // If caller provided transforms (for linked instances), transform bounding boxes into host space
                if (damperBBox != null && accessoryTransform != null)
                    damperBBox = UtilityClass.TransformBoundingBox(damperBBox, accessoryTransform);
                if (wallBBox != null && wallTransform != null)
                    wallBBox = UtilityClass.TransformBoundingBox(wallBBox, wallTransform);

                if (damperBBox == null || wallBBox == null)
                {
                    DebugLogger.Info($"[PlaceFireDamperSleeve] Bounding box unavailable. damperId={accessory?.Id.IntegerValue ?? -1}, damperBBoxNull={(damperBBox==null)}, wallId={linkedWall?.Id.IntegerValue ?? -1}, wallBBoxNull={(wallBBox==null)}");
                    Log("Bounding box unavailable for damper or wall. Skipping placement.");
                    return false;
                }

                // Check if bounding boxes intersect
                if (!BoundingBoxesIntersect(damperBBox, wallBBox))
                {
                    DebugLogger.Info($"[PlaceFireDamperSleeve] BBox intersection test failed for damperId={accessory.Id.IntegerValue}, wallId={linkedWall.Id.IntegerValue}. damperMin={damperBBox.Min}, damperMax={damperBBox.Max}, wallMin={wallBBox.Min}, wallMax={wallBBox.Max}");
                    Log("Damper bounding box does not intersect wall bounding box. Skipping placement.");
                    return false;
                }

                // Retrieve damper dimensions using correct parameter names
                var widthParam = accessory.LookupParameter("Damper Width");
                var heightParam = accessory.LookupParameter("Damper Height");

                if (widthParam == null || heightParam == null)
                {
                    DebugLogger.Info($"[PlaceFireDamperSleeve] Missing damper params for damperId={accessory.Id.IntegerValue}: widthParamNull={(widthParam==null)}, heightParamNull={(heightParam==null)}");
                    Log("Damper parameters 'Damper Width' or 'Damper Height' not found. Skipping placement.");
                    return false;
                }

                double damperWidth = widthParam.AsDouble();
                double damperHeight = heightParam.AsDouble();

                Log($"Retrieved damper dimensions: Width={damperWidth}, Height={damperHeight}");
                DebugLogger.Info($"[PlaceFireDamperSleeve] Damper dims internal: damperId={accessory.Id.IntegerValue}, Width={damperWidth}, Height={damperHeight}");
                Log($"Damper dimensions (mm): Width={UnitUtils.ConvertFromInternalUnits(damperWidth, UnitTypeId.Millimeters)}, Height={UnitUtils.ConvertFromInternalUnits(damperHeight, UnitTypeId.Millimeters)}");

                // Get symbol type name and determine damper type (null-safe)
                string familyTypeName = accessory.Symbol?.Name ?? "<unknown type>";
                Log($"Damper family type name: {familyTypeName}");

                // Clearance values
                double clearance50 = UnitUtils.ConvertToInternalUnits(50.0, UnitTypeId.Millimeters);
                double clearance100 = UnitUtils.ConvertToInternalUnits(100.0, UnitTypeId.Millimeters);

                // Determine damper type (MSFD or Standard)
                string typeNameUpper = familyTypeName?.Trim().ToUpperInvariant() ?? "";
                bool isMSFD = typeNameUpper.Contains("MSFD");
                Log($"Damper type name: '{familyTypeName}', isMSFD: {isMSFD}");

                // Get connector side based on damper type
                Connector? conn;
                string side;
                if (isMSFD)
                {
                    side = GetConnectorSideWorld(accessory, out conn);
                    Log($"[DEBUG] Used GetConnectorSideWorld for MSFD. Detected connector side: {side}");
                    
                    // Debug: Log the actual connector direction for investigation
                    if (_enableDebugLogging && conn != null)
                    {
                        var connDir = conn.CoordinateSystem.BasisX;
                        Log($"[DEBUG] Connector BasisX direction: ({connDir.X:F3}, {connDir.Y:F3}, {connDir.Z:F3})");
                        Log($"[DEBUG] Damper ID: {accessory.Id.IntegerValue}");
                    }
                }
                else
                {
                    side = GetConnectorSideLocal(accessory, out conn);
                    Log($"[DEBUG] Used GetConnectorSideLocal for non-MSFD. Detected connector side: {side}");
                }

                // Set clearances for each side
                double left = clearance50, right = clearance50, top = clearance50, bottom = clearance50;
                if (isMSFD)
                {
                    // If connector side is unknown, fallback to Right
                    string effectiveSide = side;
                    if (side == "Unknown")
                    {
                        effectiveSide = "Right";
                        Log("[MSFD CLEARANCE] Connector side unknown, defaulting to 'Right' for clearance.");
                    }
                    
                    // Apply 100mm clearance to the connector side, 50mm elsewhere
                    switch (effectiveSide)
                    {
                        case "Left": left = clearance100; break;
                        case "Right": right = clearance100; break;
                        case "Top": top = clearance100; break;
                        case "Bottom": bottom = clearance100; break;
                    }
                    
                    Log($"[MSFD CLEARANCE] Applied 100mm clearance to {effectiveSide}, 50mm elsewhere. left={UnitUtils.ConvertFromInternalUnits(left, UnitTypeId.Millimeters)}, right={UnitUtils.ConvertFromInternalUnits(right, UnitTypeId.Millimeters)}, top={UnitUtils.ConvertFromInternalUnits(top, UnitTypeId.Millimeters)}, bottom={UnitUtils.ConvertFromInternalUnits(bottom, UnitTypeId.Millimeters)}");
                }
                else
                {
                    Log($"[STANDARD CLEARANCE] 50mm all sides. left={UnitUtils.ConvertFromInternalUnits(left, UnitTypeId.Millimeters)}, right={UnitUtils.ConvertFromInternalUnits(right, UnitTypeId.Millimeters)}, top={UnitUtils.ConvertFromInternalUnits(top, UnitTypeId.Millimeters)}, bottom={UnitUtils.ConvertFromInternalUnits(bottom, UnitTypeId.Millimeters)}");
                }

                // Compute sleeve dimensions
                double sleeveWidth = damperWidth + left + right;
                double sleeveHeight = damperHeight + top + bottom;
                Log($"Sleeve dims (internal): Width={sleeveWidth}, Height={sleeveHeight}");
                Log($"Sleeve dims (mm): Width={UnitUtils.ConvertFromInternalUnits(sleeveWidth, UnitTypeId.Millimeters)}, Height={UnitUtils.ConvertFromInternalUnits(sleeveHeight, UnitTypeId.Millimeters)}");

                // Activate sleeve symbol if not already active
                if (!sleeveSymbol.IsActive)
                {
                    sleeveSymbol.Activate();
                    Log("Activated sleeve family symbol.");
                }

                // Retrieve a level for placement
                var level = new FilteredElementCollector(_doc)
                                .OfClass(typeof(Level))
                                .Cast<Level>()
                                .FirstOrDefault();
                if (level == null)
                {
                    DebugLogger.Info($"[PlaceFireDamperSleeve] No Level found for damperId={accessory.Id.IntegerValue}. Aborting placement.");
                    Log("No Level found; aborting placement.");
                    return false;
                }

                // Determine raw damper insertion point (MEP connector or bounding box center)
                XYZ rawCenter = (accessory.MEPModel?.ConnectorManager.Connectors.Cast<Connector>().FirstOrDefault()?.Origin)
                                ?? ((damperBBox.Min + damperBBox.Max) * 0.5);
                if (rawCenter == null) rawCenter = new XYZ(0, 0, 0);
                // Ensure a non-null center for subsequent non-nullable operations
                XYZ safeCenter = rawCenter ?? new XYZ(0, 0, 0);
                // If accessory is from a linked doc, transform the raw center into host space for projection
                if (accessoryTransform != null && rawCenter != null)
                    rawCenter = accessoryTransform.OfPoint(rawCenter);
                // Update safeCenter after any transform
                safeCenter = rawCenter ?? new XYZ(0, 0, 0);
                Log($"Raw damper center before projection: {safeCenter}");
                DebugLogger.Info($"[PlaceFireDamperSleeve] Raw center damperId={accessory.Id.IntegerValue}: {safeCenter}");

                // Project raw center onto wall plane (wall face intersection)
                XYZ wallFaceOrigin;
                try
                {
                    var lc = linkedWall.Location as LocationCurve;
                    XYZ? endPt = lc?.Curve?.GetEndPoint(0);
                    wallFaceOrigin = endPt ?? safeCenter;
                }
                catch
                {
                    wallFaceOrigin = safeCenter;
                }
                XYZ wallNormal = linkedWall.Orientation;  // wall face normal
                // If wall came from a linked doc, transform wall geometry into host space
                if (wallTransform != null)
                {
                    try { wallFaceOrigin = wallTransform.OfPoint(wallFaceOrigin); } catch { }
                    try { wallNormal = wallTransform.OfVector(wallNormal); } catch { }
                }
                double distance = (safeCenter - wallFaceOrigin).DotProduct(wallNormal);
                XYZ projectedCenter = safeCenter - wallNormal.Multiply(distance);
                Log($"Projected damper-wall intersection center: {projectedCenter}");
                DebugLogger.Info($"[PlaceFireDamperSleeve] Projected center damperId={accessory.Id.IntegerValue}: {projectedCenter}, wallNormal={wallNormal}");

                // Calculate offset for MSFD dampers (25mm toward connector direction)
                XYZ offset = XYZ.Zero;
                double offset25 = UnitUtils.ConvertToInternalUnits(25.0, UnitTypeId.Millimeters);
                if (isMSFD && conn != null)
                {
                    // Determine wall orientation and get correct offset direction
                    bool isYAxisWall = IsYAxisWall(linkedWall);
                    Log($"[OFFSET] Wall orientation: ({wallNormal.X:F3}, {wallNormal.Y:F3}, {wallNormal.Z:F3}), isYAxisWall: {isYAxisWall}");
                    
                    XYZ worldOffsetDir = GetOffsetDirection(side, isYAxisWall);
                    offset = worldOffsetDir * offset25;
                    Log($"[OFFSET] Side: {side}, wall type: {(isYAxisWall ? "Y-axis" : "X-axis")}, using world direction: ({worldOffsetDir.X:F3}, {worldOffsetDir.Y:F3}, {worldOffsetDir.Z:F3})");
                }
                
                XYZ sleeveCenter = projectedCenter + offset;
                Log($"[CONNECTOR OFFSET] Sleeve placement center after offset: {sleeveCenter}");

                // DUPLICATION SUPPRESSION: Check if a sleeve already exists at this location
                double sleeveCheckRadius = UnitUtils.ConvertToInternalUnits(100.0, UnitTypeId.Millimeters); // 100mm tolerance
                bool duplicateExists = false;
                try
                {
                    BoundingBoxXYZ? sectionBox = null;
                    if (_doc.ActiveView is View3D vb) sectionBox = JSE_RevitAddin_MEP_OPENINGS.Helpers.SectionBoxHelper.GetSectionBoxBounds(vb);
                    // Require same family for damper duplication checks so nearby duct sleeves do not block damper placement
                    duplicateExists = OpeningDuplicationChecker.IsAnySleeveAtLocationEnhanced(_doc, sleeveCenter, sleeveCheckRadius, clusterExpansion: 0.0, ignoreIds: null, hostType: "OpeningOnWall", sectionBox: sectionBox, requireSameFamily: true, familyName: accessory.Symbol?.Family?.Name);
                }
                catch (Exception ex)
                {
                    Log($"[ERROR] Duplication suppression check failed: {ex.Message}");
                }
                if (duplicateExists)
                {
                    DebugLogger.Info($"[PlaceFireDamperSleeve] Duplicate detected at center for damperId={accessory.Id.IntegerValue}, center={sleeveCenter}");
                    Log("Duplicate damper sleeve detected at this location. Skipping placement.");
                    return false;
                }

                // Place the sleeve at the computed center
                var instance = _doc.Create.NewFamilyInstance(
                    sleeveCenter,
                    sleeveSymbol,
                    level,
                    StructuralType.NonStructural);
                
                    
                // Ensure instance is not null before touching parameters
                if (instance == null)
                {
                    DebugLogger.Error($"[PlaceFireDamperSleeve] Created instance is null for damperId={accessory.Id.IntegerValue}");
                    Log("Failed to create sleeve instance; skipping placement.");
                    return false;
                }

                // Set the Level parameter for schedule consistency
                Parameter levelParam = instance.get_Parameter(BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM);
                if (levelParam != null && !levelParam.IsReadOnly)
                {
                    levelParam.Set(level.Id);
                }
                else
                {
                    var levelByName = instance.LookupParameter("Level");
                    if (levelByName != null && !levelByName.IsReadOnly)
                        levelByName.Set(level.Id);
                }
                Log("Created sleeve at computed center.");
                DebugLogger.Info($"[PlaceFireDamperSleeve] Created sleeve instanceId={(instance.Id.IntegerValue)} for damperId={accessory.Id.IntegerValue} at {sleeveCenter}");

                // Set sleeve dimensions and depth (null-safe since instance is non-null)
                instance.LookupParameter("Width")?.Set(sleeveWidth);
                instance.LookupParameter("Height")?.Set(sleeveHeight);
                try
                {
                    double wallWidth = 0.0;
                    try { wallWidth = linkedWall.Width; } catch { wallWidth = 0.0; }
                    instance.LookupParameter("Depth")?.Set(wallWidth);
                }
                catch { }
                Log("Set sleeve dimensions and depth.");

                // Rotate sleeve for Y-axis dampers only (simplified rotation logic)
                try
                {
                    var damperDir = accessory.GetTotalTransform().BasisX; // damper flow direction
                    bool isYAxisDamper = Math.Abs(damperDir.Y) > Math.Abs(damperDir.X);
                    
                    if (isYAxisDamper)
                    {
                        Log("[ROTATE] Y-axis damper detected, applying 90° rotation");
                        _doc.Regenerate(); // Force Revit to update the instance location
                        LocationPoint? loc = instance.Location as LocationPoint;
                        if (loc != null)
                        {
                            XYZ rotationPoint = loc.Point;
                            Line rotationAxis = Line.CreateBound(rotationPoint, rotationPoint.Add(XYZ.BasisZ));
                            double rotationAngle = Math.PI / 2; // 90 degrees
                            ElementTransformUtils.RotateElement(_doc, instance.Id, rotationAxis, rotationAngle);
                            Log($"[ROTATE] Rotated Y-axis damper sleeve 90° at {rotationPoint}");
                        }
                        else
                        {
                            Log("[ROTATE ERROR] Instance location is not a LocationPoint!");
                        }
                    }
                    else
                    {
                        Log("[ROTATE] X-axis damper - no rotation needed");
                    }
                }
                catch (Exception rex)
                {
                    Log($"[ERROR] Failed to rotate sleeve: {rex.Message}");
                }

                // Set the Schedule Level parameter
                Level? refLevel = JSE_RevitAddin_MEP_OPENINGS.Helpers.HostLevelHelper.GetHostReferenceLevel(_doc, accessory);
                if (refLevel != null)
                {
                    Parameter schedLevelParam = instance.LookupParameter("Schedule Level");
                    if (schedLevelParam != null && !schedLevelParam.IsReadOnly)
                    {
                        schedLevelParam.Set(refLevel.Id);
                        Log($"[FireDamperSleevePlacerService] Set Schedule Level to {refLevel.Name} for damper {accessory.Id.IntegerValue}");
                    }
                }

                Log("Set sleeve dimensions and completed placement.");
                DebugLogger.Info($"[PlaceFireDamperSleeve] Placement completed for damperId={accessory.Id.IntegerValue}");
                return true;
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[PlaceFireDamperSleeve] Exception: {ex.Message}\n{ex.StackTrace}");
                Log($"Exception in PlaceFireDamperSleeve: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Process a batch of dampers using the same workflow as the command.
        /// Returns (placedCount, skippedCount, errorCount).
        /// </summary>
        public (int placed, int skipped, int errors) ProcessDamperBatch(
            System.Collections.Generic.List<(FamilyInstance instance, Transform? transform)> damperTuples,
            System.Collections.Generic.List<(Wall wall, Transform? transform)> wallTuples,
            FamilySymbol openingSymbol,
            System.Action<string> log,
            bool enableDebugging = true,
            System.Collections.Generic.List<(Element element, Transform? transform)>? structuralElements = null)
        {
            int placedCount = 0;
            int skippedExisting = 0;
            int errors = 0;

            if (openingSymbol == null)
            {
                log?.Invoke("[ProcessDamperBatch] openingSymbol is null, aborting damper batch.");
                return (0, 0, 0);
            }

            // Enable damper-specific logging similar to duct logging when requested
            if (enableDebugging)
            {
                JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.SetDamperLogFile();
                JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.InitLogFile("dampersleeveplacer");
                DamperLogger.InitLogFile();
                // Write an immediate marker so file exists even if no placements occur
                DamperLogger.Log("[ProcessDamperBatch] Debugging enabled - initialized damper log.");
            }

            // Debug summary: list counts and a small sample of wall candidates for quick diagnosis
            try
            {
                var damperListSafe = damperTuples ?? new System.Collections.Generic.List<(FamilyInstance, Transform?)>();
                var wallListSafe = wallTuples ?? new System.Collections.Generic.List<(Wall, Transform?)>();
                DebugLogger.Log($"[ProcessDamperBatch] DEBUG START: damperCount={damperListSafe.Count}, wallCandidates={wallListSafe.Count}, structuralProvided={(structuralElements!=null ? structuralElements.Count.ToString() : "false")}");
                int sampleIndex = 0;
                foreach (var wtuple in wallListSafe.Take(10))
                {
                    var w = wtuple.wall;
                    var wtr = wtuple.transform;
                    BoundingBoxXYZ? bbox = null;
                    try { bbox = w?.get_BoundingBox(null); } catch { }
                    string bboxText = bbox != null ? $"Min={bbox.Min},Max={bbox.Max}" : "<no-bbox>";
                    DebugLogger.Log($"[ProcessDamperBatch] WALL SAMPLE {sampleIndex++}: id={(w?.Id.IntegerValue.ToString() ?? "<null>")}, transformProvided={(wtr!=null)}, bbox={bboxText}");
                }
            }
            catch { }

            // Activate symbol if necessary
            try
            {
                using (var tx = new Transaction(_doc, "Activate Damper Opening Symbol"))
                {
                    tx.Start();
                    if (!openingSymbol.IsActive) openingSymbol.Activate();
                    tx.Commit();
                }
            }
            catch { }

            double offset25 = UnitUtils.ConvertToInternalUnits(25.0, UnitTypeId.Millimeters);
            double sleeveCheckRadius = UnitUtils.ConvertToInternalUnits(100.0, UnitTypeId.Millimeters);

            var damperListSafeMain = damperTuples ?? new System.Collections.Generic.List<(FamilyInstance, Transform?)>();
            var wallListSafeMain = wallTuples ?? new System.Collections.Generic.List<(Wall, Transform?)>();
            foreach (var tuple in damperListSafeMain)
            {
                try
                {
                    var damper = tuple.instance;
                    var damperTransform = tuple.transform;

                    if (damper == null) { skippedExisting++; continue; }

                    // Get damper location
                    XYZ damperLoc = (damper.Location as LocationPoint)?.Point ?? damper.GetTransform().Origin;

                    // Determine connector side (prefer world for MSFD)
                    string side = GetConnectorSideWorld(damper, out var connector);

                    XYZ offsetVec = UtilityClass.OffsetVector4Way(side, damper.GetTotalTransform());
                    XYZ sleevePos = damperLoc + offsetVec * offset25;

                    // Transform if damper is from a link
                    if (damperTransform != null)
                        sleevePos = damperTransform.OfPoint(sleevePos);

                    // NOTE: Delay duplication suppression until after we identify a target wall so we can pass hostType.

                    // Find a wall that contains this sleevePos
                    Wall? targetWall = null;
                    foreach (var wtuple in wallListSafeMain)
                    {
                        var w = wtuple.wall;
                        var wtr = wtuple.transform;
                        if (w == null) continue;
                        var bbox = w.get_BoundingBox(null);
                        if (bbox == null) continue;
                        if (wtr != null) bbox = UtilityClass.TransformBoundingBox(bbox, wtr);
                        if (bbox == null) continue;
                        if (UtilityClass.PointInBoundingBox(sleevePos, bbox)) { targetWall = w; break; }
                    }

                    // If no wall matched via bbox search, try an intersection fallback using the MepIntersectionService
                    if (targetWall == null)
                    {
                        try
                        {
                            DebugLogger.Log($"[ProcessDamperBatch] No wall found by bbox for damper {damper.Id.IntegerValue}, attempting intersection fallback.");
                            // Construct a short test line through the sleevePos in world Z to probe nearby walls
                            var p1 = sleevePos + new Autodesk.Revit.DB.XYZ(0, 0, -5.0);
                            var p2 = sleevePos + new Autodesk.Revit.DB.XYZ(0, 0, 5.0);
                            var testLine = Line.CreateBound(p1, p2);
                            // Use the structural elements collected from section box helper if possible
                            var structural = structuralElements ?? MepIntersectionService.CollectStructuralElementsForDirectIntersectionVisibleOnly(_doc, log ?? (_ => { }));
                            var intersects = MepIntersectionService.FindIntersections(testLine, null, structural, log ?? (_ => { }));
                            if (intersects != null && intersects.Count > 0)
                            {
                                targetWall = intersects[0].Item1 as Wall;
                                if (targetWall != null)
                                {
                                    DebugLogger.Log($"[ProcessDamperBatch] Intersection fallback found wall {targetWall.Id.IntegerValue} for damper {damper.Id.IntegerValue}.");
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            log?.Invoke($"[ProcessDamperBatch] Intersection fallback failed: {ex.Message}");
                        }
                    }

                    if (targetWall == null)
                    {
                        skippedExisting++;
                        continue;
                    }

                    // Now that we have a target wall, perform duplication suppression using wall-based host type
                    bool duplicateExistsForWall = false;
                    try
                    {
                        // Use hostType filter to only consider wall clusters/sleeves
                        BoundingBoxXYZ? sectionBox = null;
                        if (_doc.ActiveView is View3D vb2) sectionBox = JSE_RevitAddin_MEP_OPENINGS.Helpers.SectionBoxHelper.GetSectionBoxBounds(vb2);
                        // Require same family for damper batch-level duplicate check as well
                        duplicateExistsForWall = OpeningDuplicationChecker.IsAnySleeveAtLocationEnhanced(_doc, sleevePos, sleeveCheckRadius, clusterExpansion: 0.0, ignoreIds: null, hostType: "OpeningOnWall", sectionBox: sectionBox, requireSameFamily: true, familyName: damper.Symbol?.Family?.Name);
                    }
                    catch { }
                    if (duplicateExistsForWall)
                    {
                        skippedExisting++;
                        continue;
                    }

                    // Delegate to single-point placer which contains the projection/clearance/duplication logic
                    // Find transforms for accessory and wall from the original tuples
                    Transform? accTr = damperTransform;
                    Transform? wallTr = null;
                    // Try to find the wall tuple that matched to get its transform
                    var matching = wallListSafeMain.FirstOrDefault(wt => wt.wall != null && wt.wall.Id == targetWall.Id);
                    if (matching.wall != null)
                        wallTr = matching.transform;

                    var placed = PlaceFireDamperSleeve(damper, openingSymbol, targetWall, accTr, wallTr);
                    if (placed) placedCount++; else skippedExisting++;
                }
                catch (System.Exception ex)
                {
                    errors++;
                    log?.Invoke($"[ProcessDamperBatch] ERROR processing damper: {ex.Message}");
                }
            }

            return (placedCount, skippedExisting, errors);
        }

        private bool BoundingBoxesIntersect(BoundingBoxXYZ bbox1, BoundingBoxXYZ bbox2)
        {
            return !(bbox1.Max.X < bbox2.Min.X || bbox1.Min.X > bbox2.Max.X ||
                     bbox1.Max.Y < bbox2.Min.Y || bbox1.Min.Y > bbox2.Max.Y ||
                     bbox1.Max.Z < bbox2.Min.Z || bbox1.Min.Z > bbox2.Max.Z);
        }
    }
}

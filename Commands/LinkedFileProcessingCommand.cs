using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB.Plumbing; // For Pipe
using Autodesk.Revit.DB.Mechanical; // For Duct
using Autodesk.Revit.DB.Electrical; // For CableTray
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;
using System.Collections.Generic;
using System.Linq;
using System;

namespace JSE_RevitAddin_MEP_OPENINGS.Commands
{
    /// <summary>
    /// Task 2: Exclusive Command Class for Linked File Processing
    /// This command handles sleeve placement for MEP elements intersecting with structural elements in linked files.
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class LinkedFileProcessingCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            DebugLogger.InitCustomLogFileOverwrite("LinkedFileProcessing"); // Ensure log file is overwritten
            DebugLogger.Info("Starting LinkedFileProcessingCommand execution for linked file sleeve placement.");

            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;

            try
            {
                // Step 1: Collect linked file instances
                var linkedInstances = CollectLinkedFileInstances(doc);
                if (linkedInstances.Count == 0)
                {
                    TaskDialog.Show("Info", "No linked files found in the current document.");
                    return Result.Succeeded;
                }

                DebugLogger.Info($"Found {linkedInstances.Count} linked file instances.");

                // Step 2: Process each linked file for MEP-structural intersections
                var allIntersections = new List<LinkedIntersectionResult>();
                foreach (var linkInstance in linkedInstances)
                {
                    var linkedDoc = linkInstance.GetLinkDocument();
                    if (linkedDoc == null)
                    {
                        DebugLogger.Debug($"Skipping linked instance {linkInstance.Id.IntegerValue} - document not accessible.");
                        continue;
                    }

                    DebugLogger.Info($"Processing linked file: {linkedDoc.Title}");

                    // Collect MEP elements from linked document
                    var linkedMEPElements = CollectMEPElementsFromLinkedFile(linkedDoc);
                    DebugLogger.Info($"Found {linkedMEPElements.Count} MEP elements in linked file {linkedDoc.Title}");

                    // Collect structural elements from active document for intersection testing
                    var structuralElements = CollectStructuralElementsFromActiveDocument(doc);
                    DebugLogger.Info($"Found {structuralElements.Count} structural elements in active document for intersection testing");

                    // Detect intersections between linked MEP elements and active structural elements
                    var intersections = DetectIntersectionsInLinkedFile(doc, linkInstance, linkedMEPElements, structuralElements);
                    allIntersections.AddRange(intersections);

                    DebugLogger.Info($"Detected {intersections.Count} intersections in linked file {linkedDoc.Title}");
                }

                DebugLogger.Info($"Total intersections detected across all linked files: {allIntersections.Count}");

                if (allIntersections.Count == 0)
                {
                    TaskDialog.Show("Info", "No intersections detected between linked MEP elements and structural elements.");
                    return Result.Succeeded;
                }

                // Step 3: Activate required family symbols for sleeve placement
                var familySymbol = ActivateSleeveFamily(doc);
                if (familySymbol == null)
                {
                    TaskDialog.Show("Error", "No suitable sleeve family found. Please ensure the required families are loaded.");
                    return Result.Failed;
                }

                // Step 4: Place sleeves for detected intersections
                using (var tx = new Transaction(doc, "Place Sleeves for Linked File Intersections"))
                {
                    tx.Start();

                    int placedSleeves = 0;
                    foreach (var intersection in allIntersections)
                    {
                        try
                        {
                            PlaceSleeveForLinkedIntersection(doc, intersection, familySymbol);
                            placedSleeves++;
                            DebugLogger.Info($"Placed sleeve for linked MEP element {intersection.LinkedMEPElementId} " +
                                           $"intersecting with structural element {intersection.StructuralElement.Id.IntegerValue}");
                        }
                        catch (Exception ex)
                        {
                            DebugLogger.Error($"Failed to place sleeve for intersection: {ex.Message}");
                        }
                    }

                    tx.Commit();
                    DebugLogger.Info($"Successfully placed {placedSleeves} sleeves for linked file intersections.");
                }

                TaskDialog.Show("Success", $"LinkedFileProcessingCommand completed successfully. Placed {allIntersections.Count} sleeves for linked file intersections.");
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"Error during LinkedFileProcessingCommand execution: {ex.Message}");
                message = ex.Message;
                return Result.Failed;
            }
        }

        /// <summary>
        /// Collect all RevitLinkInstance elements from the active document
        /// </summary>
        private List<RevitLinkInstance> CollectLinkedFileInstances(Document doc)
        {
            try
            {
                var linkedInstances = new FilteredElementCollector(doc)
                    .OfClass(typeof(RevitLinkInstance))
                    .Cast<RevitLinkInstance>()
                    .Where(li => li.GetLinkDocument() != null) // Only include accessible linked documents
                    .ToList();

                foreach (var linkInstance in linkedInstances)
                {
                    var linkedDoc = linkInstance.GetLinkDocument();
                    DebugLogger.Debug($"Found linked file: {linkedDoc?.Title ?? "Unknown"} (ID: {linkInstance.Id.IntegerValue})");
                }

                return linkedInstances;
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"Error collecting linked file instances: {ex.Message}");
                return new List<RevitLinkInstance>();
            }
        }

        /// <summary>
        /// Collect MEP elements from a linked document
        /// Reuses logic from MEPElementCollector but adapted for linked files
        /// </summary>
        private List<Element> CollectMEPElementsFromLinkedFile(Document linkedDoc)
        {
            try
            {
                var mepElements = new List<Element>();

                // Collect Pipes
                var pipes = new FilteredElementCollector(linkedDoc)
                    .OfClass(typeof(Pipe))
                    .WhereElementIsNotElementType()
                    .ToList();
                mepElements.AddRange(pipes);
                DebugLogger.Debug($"Collected {pipes.Count} pipes from linked file {linkedDoc.Title}");

                // Collect Ducts
                var ducts = new FilteredElementCollector(linkedDoc)
                    .OfClass(typeof(Duct))
                    .WhereElementIsNotElementType()
                    .ToList();
                mepElements.AddRange(ducts);
                DebugLogger.Debug($"Collected {ducts.Count} ducts from linked file {linkedDoc.Title}");

                // Collect Cable Trays
                var cableTrays = new FilteredElementCollector(linkedDoc)
                    .OfClass(typeof(CableTray))
                    .WhereElementIsNotElementType()
                    .ToList();
                mepElements.AddRange(cableTrays);
                DebugLogger.Debug($"Collected {cableTrays.Count} cable trays from linked file {linkedDoc.Title}");

                return mepElements;
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"Error collecting MEP elements from linked file {linkedDoc.Title}: {ex.Message}");
                return new List<Element>();
            }
        }

        /// <summary>
        /// Collect structural elements from the active document for intersection testing
        /// </summary>
        private List<Element> CollectStructuralElementsFromActiveDocument(Document doc)
        {
            try
            {
                var structuralElements = new List<Element>();

                // Collect Floors
                var floors = new FilteredElementCollector(doc)
                    .OfClass(typeof(Floor))
                    .WhereElementIsNotElementType()
                    .ToList();
                structuralElements.AddRange(floors);
                DebugLogger.Debug($"Collected {floors.Count} floors from active document");

                // Collect Structural Framing
                var framingElements = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilyInstance))
                    .WhereElementIsNotElementType()
                    .Where(e => e.Category?.Id?.IntegerValue == (int)BuiltInCategory.OST_StructuralFraming)
                    .ToList();
                structuralElements.AddRange(framingElements);
                DebugLogger.Debug($"Collected {framingElements.Count} structural framing elements from active document");

                return structuralElements;
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"Error collecting structural elements from active document: {ex.Message}");
                return new List<Element>();
            }
        }

        /// <summary>
        /// Detect intersections between linked MEP elements and active document structural elements
        /// Reuses helper methods from ReferenceIntersectorHelper
        /// </summary>
        private List<LinkedIntersectionResult> DetectIntersectionsInLinkedFile(Document doc, RevitLinkInstance linkInstance, 
            List<Element> linkedMEPElements, List<Element> structuralElements)
        {
            var intersections = new List<LinkedIntersectionResult>();

            try
            {
                // Get a 3D view for ReferenceIntersector
                var view3D = new FilteredElementCollector(doc)
                    .OfClass(typeof(View3D))
                    .Cast<View3D>()
                    .FirstOrDefault(v => !v.IsTemplate);

                if (view3D == null)
                {
                    DebugLogger.Error("No 3D view found for intersection detection");
                    return intersections;
                }

                // Create filter for structural elements
                ElementFilter structuralFilter = new LogicalOrFilter(
                    new ElementCategoryFilter(BuiltInCategory.OST_Floors),
                    new ElementCategoryFilter(BuiltInCategory.OST_StructuralFraming));

                // Create ReferenceIntersector for detecting intersections
                var refIntersector = new ReferenceIntersector(structuralFilter, FindReferenceTarget.Face, view3D)
                {
                    FindReferencesInRevitLinks = false // Focus on active document structural elements
                };

                foreach (var mepElement in linkedMEPElements)
                {
                    var locCurve = mepElement.Location as LocationCurve;
                    if (locCurve?.Curve is Line line)
                    {
                        // Sample points along the MEP element for intersection testing
                        var sampleFractions = new[] { 0.0, 0.25, 0.5, 0.75, 1.0 };
                        
                        foreach (double t in sampleFractions)
                        {
                            var samplePt = line.Evaluate(t, true);
                            var direction = line.Direction.Normalize();
                            
                            // Test multiple directions for comprehensive detection
                            var directions = new[]
                            {
                                direction,                      // Forward along element
                                direction.Negate(),             // Backward along element
                                XYZ.BasisZ,                     // Up (for floors above)
                                XYZ.BasisZ.Negate(),            // Down (for floors below)
                                new XYZ(-direction.Y, direction.X, 0).Normalize(), // Perpendicular 1
                                new XYZ(direction.Y, -direction.X, 0).Normalize()  // Perpendicular 2
                            };

                            foreach (var testDirection in directions)
                            {
                                var hits = refIntersector.Find(samplePt, testDirection);
                                if (hits != null && hits.Count > 0)
                                {
                                    foreach (var hit in hits)
                                    {
                                        if (hit.Proximity > 0.5) continue; // Skip distant hits

                                        var structuralElement = doc.GetElement(hit.GetReference().ElementId);
                                        // Use GlobalPoint for face hits (ReferenceIntersector with FindReferenceTarget.Face)
                                        XYZ intersectionPoint = hit.GetReference().GlobalPoint;
                                        if (structuralElement != null && 
                                            (structuralElement.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Floors ||
                                             structuralElement.Category.Id.IntegerValue == (int)BuiltInCategory.OST_StructuralFraming))
                                        {
                                            // Create intersection result
                                            var intersection = new LinkedIntersectionResult
                                            {
                                                LinkInstance = linkInstance,
                                                LinkedMEPElement = mepElement,
                                                LinkedMEPElementId = mepElement.Id.IntegerValue,
                                                StructuralElement = structuralElement,
                                                IntersectionPoint = intersectionPoint,
                                                Proximity = hit.Proximity
                                            };

                                            // Check for duplicates
                                            bool isDuplicate = intersections.Any(i => 
                                                i.LinkedMEPElementId == intersection.LinkedMEPElementId &&
                                                i.StructuralElement.Id.IntegerValue == intersection.StructuralElement.Id.IntegerValue);

                                            if (!isDuplicate)
                                            {
                                                intersections.Add(intersection);
                                                DebugLogger.Debug($"Detected intersection: Linked {mepElement.Category.Name} {mepElement.Id.IntegerValue} " +
                                                                $"with {structuralElement.Category.Name} {structuralElement.Id.IntegerValue} " +
                                                                $"at proximity {hit.Proximity:F3}");
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                return intersections;
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"Error detecting intersections in linked file: {ex.Message}");
                return intersections;
            }
        }

        /// <summary>
        /// Activate suitable sleeve family for placement
        /// Reuses existing service classes for symbol activation
        /// </summary>
        private FamilySymbol ActivateSleeveFamily(Document doc)
        {
            try
            {
                var allFamilySymbols = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilySymbol))
                    .Cast<FamilySymbol>()
                    .ToList();

                // Find suitable sleeve family (following existing patterns)
                var familySymbol = allFamilySymbols
                    .FirstOrDefault(sym => sym.Family.Name.Contains("OnWall", StringComparison.OrdinalIgnoreCase) ||
                                           sym.Family.Name.Contains("PipeOpeningOnSlab", StringComparison.OrdinalIgnoreCase) ||
                                           sym.Family.Name.Contains("DuctOpeningOnSlab", StringComparison.OrdinalIgnoreCase));

                if (familySymbol != null && !familySymbol.IsActive)
                {
                    using (var txActivate = new Transaction(doc, "Activate Linked File Sleeve Symbol"))
                    {
                        txActivate.Start();
                        familySymbol.Activate();
                        txActivate.Commit();
                    }
                }

                return familySymbol;
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"Error activating sleeve family: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Place sleeve for a detected linked file intersection
        /// </summary>
        private void PlaceSleeveForLinkedIntersection(Document doc, LinkedIntersectionResult intersection, FamilySymbol familySymbol)
        {
            try
            {
                // Create family instance at intersection point
                var familyInstance = doc.Create.NewFamilyInstance(
                    intersection.IntersectionPoint,
                    familySymbol,
                    Autodesk.Revit.DB.Structure.StructuralType.NonStructural);

                // Set sleeve dimensions following the exact logic from working commands
                SetLinkedSleeveDimensions(familyInstance, intersection.LinkedMEPElement, intersection.StructuralElement);

                // Log placement details
                DebugLogger.Info($"Placed sleeve family {familySymbol.Family.Name} at {intersection.IntersectionPoint} " +
                               $"for linked MEP element {intersection.LinkedMEPElementId} " +
                               $"intersecting structural element {intersection.StructuralElement.Id.IntegerValue}");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"Error placing sleeve for linked intersection: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Set sleeve dimensions for linked file intersections following working command patterns
        /// </summary>
        private void SetLinkedSleeveDimensions(FamilyInstance sleeveInstance, Element linkedMEPElement, Element structuralElement)
        {
            try
            {
                Document doc = sleeveInstance.Document;
                
                // Get structural element thickness for depth parameter
                double structuralThickness = GetStructuralElementThickness(structuralElement);
                
                if (linkedMEPElement.Category.Id.IntegerValue == (int)BuiltInCategory.OST_PipeCurves)
                {
                    // PIPE LOGIC: Following PipeSleevePlacer pattern exactly
                    var pipe = linkedMEPElement as Pipe;
                    if (pipe != null)
                    {
                        // Get pipe diameter
                        double pipeDiameter = 0.0;
                        var diameterParam = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER);
                        if (diameterParam != null && diameterParam.HasValue)
                        {
                            pipeDiameter = diameterParam.AsDouble();
                        }
                        
                        // Check for insulation (exact logic from PipeSleeveCommand)
                        double clearancePerSideMM = 50.0; // default: 50mm per side for non-insulated
                        var insulationThicknessParam = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_INSULATION_THICKNESS);
                        if (insulationThicknessParam != null && insulationThicknessParam.HasValue && insulationThicknessParam.AsDouble() > 0)
                        {
                            clearancePerSideMM = 25.0; // 25mm per side for insulated
                        }
                        
                        // Calculate final diameter with clearance (following working command pattern)
                        double totalClearance = UnitUtils.ConvertToInternalUnits(clearancePerSideMM * 2.0, UnitTypeId.Millimeters);
                        double finalDiameter = pipeDiameter + totalClearance;
                        
                        // Set parameters
                        SetParameterSafely(sleeveInstance, "Diameter", finalDiameter, linkedMEPElement.Id.IntegerValue);
                        SetParameterSafely(sleeveInstance, "Depth", structuralThickness, linkedMEPElement.Id.IntegerValue);
                        
                        DebugLogger.Info($"LINKED PIPE SLEEVE DIMENSIONS: Pipe {linkedMEPElement.Id.IntegerValue} - " +
                                       $"Original Diameter: {UnitUtils.ConvertFromInternalUnits(pipeDiameter, UnitTypeId.Millimeters):F0}mm, " +
                                       $"Clearance: {clearancePerSideMM}mm/side, " +
                                       $"Final Diameter: {UnitUtils.ConvertFromInternalUnits(finalDiameter, UnitTypeId.Millimeters):F0}mm, " +
                                       $"Depth: {UnitUtils.ConvertFromInternalUnits(structuralThickness, UnitTypeId.Millimeters):F0}mm");
                    }
                }
                else if (linkedMEPElement.Category.Id.IntegerValue == (int)BuiltInCategory.OST_DuctCurves)
                {
                    // DUCT LOGIC: Following DuctSleevePlacer pattern exactly
                    var duct = linkedMEPElement as Duct;
                    if (duct != null)
                    {
                        // Get duct dimensions
                        double ductWidth = 0.0, ductHeight = 0.0;
                        var widthParam = duct.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM);
                        var heightParam = duct.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM);
                        
                        if (widthParam != null && widthParam.HasValue)
                            ductWidth = widthParam.AsDouble();
                        if (heightParam != null && heightParam.HasValue)
                            ductHeight = heightParam.AsDouble();
                        
                        // Add 50mm clearance on all sides (following working command pattern)
                        double clearanceMM = 50.0;
                        double clearance = UnitUtils.ConvertToInternalUnits(clearanceMM, UnitTypeId.Millimeters);
                        double finalWidth = ductWidth + (clearance * 2);
                        double finalHeight = ductHeight + (clearance * 2);
                        
                        // Set parameters
                        SetParameterSafely(sleeveInstance, "Width", finalWidth, linkedMEPElement.Id.IntegerValue);
                        SetParameterSafely(sleeveInstance, "Height", finalHeight, linkedMEPElement.Id.IntegerValue);
                        SetParameterSafely(sleeveInstance, "Depth", structuralThickness, linkedMEPElement.Id.IntegerValue);
                        
                        DebugLogger.Info($"LINKED DUCT SLEEVE DIMENSIONS: Duct {linkedMEPElement.Id.IntegerValue} - " +
                                       $"Original: {UnitUtils.ConvertFromInternalUnits(ductWidth, UnitTypeId.Millimeters):F0}mm x {UnitUtils.ConvertFromInternalUnits(ductHeight, UnitTypeId.Millimeters):F0}mm, " +
                                       $"Clearance: {clearanceMM}mm/side, " +
                                       $"Final: {UnitUtils.ConvertFromInternalUnits(finalWidth, UnitTypeId.Millimeters):F0}mm x {UnitUtils.ConvertFromInternalUnits(finalHeight, UnitTypeId.Millimeters):F0}mm, " +
                                       $"Depth: {UnitUtils.ConvertFromInternalUnits(structuralThickness, UnitTypeId.Millimeters):F0}mm");
                    }
                }
                else if (linkedMEPElement.Category.Id.IntegerValue == (int)BuiltInCategory.OST_CableTray)
                {
                    // CABLE TRAY LOGIC: Following CableTraySleevePlacer pattern exactly
                    var cableTray = linkedMEPElement as CableTray;
                    if (cableTray != null)
                    {
                        // Get cable tray dimensions
                        double trayWidth = 0.0, trayHeight = 0.0;
                        var widthParam = cableTray.get_Parameter(BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM);
                        var heightParam = cableTray.get_Parameter(BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM);
                        
                        if (widthParam != null && widthParam.HasValue)
                            trayWidth = widthParam.AsDouble();
                        if (heightParam != null && heightParam.HasValue)
                            trayHeight = heightParam.AsDouble();
                        
                        // Add 50mm clearance on all sides (following working command pattern)
                        double clearanceMM = 50.0;
                        double clearance = UnitUtils.ConvertToInternalUnits(clearanceMM, UnitTypeId.Millimeters);
                        double finalWidth = trayWidth + (clearance * 2);
                        double finalHeight = trayHeight + (clearance * 2);
                        
                        // Set parameters
                        SetParameterSafely(sleeveInstance, "Width", finalWidth, linkedMEPElement.Id.IntegerValue);
                        SetParameterSafely(sleeveInstance, "Height", finalHeight, linkedMEPElement.Id.IntegerValue);
                        SetParameterSafely(sleeveInstance, "Depth", structuralThickness, linkedMEPElement.Id.IntegerValue);
                        
                        DebugLogger.Info($"LINKED CABLE TRAY SLEEVE DIMENSIONS: Cable Tray {linkedMEPElement.Id.IntegerValue} - " +
                                       $"Original: {UnitUtils.ConvertFromInternalUnits(trayWidth, UnitTypeId.Millimeters):F0}mm x {UnitUtils.ConvertFromInternalUnits(trayHeight, UnitTypeId.Millimeters):F0}mm, " +
                                       $"Clearance: {clearanceMM}mm/side, " +
                                       $"Final: {UnitUtils.ConvertFromInternalUnits(finalWidth, UnitTypeId.Millimeters):F0}mm x {UnitUtils.ConvertFromInternalUnits(finalHeight, UnitTypeId.Millimeters):F0}mm, " +
                                       $"Depth: {UnitUtils.ConvertFromInternalUnits(structuralThickness, UnitTypeId.Millimeters):F0}mm");
                    }
                }
                
                // Regenerate document after setting parameters
                doc.Regenerate();
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"Error setting linked sleeve dimensions for MEP {linkedMEPElement.Id.IntegerValue}: {ex.Message}");
            }
        }
        
        /// <summary>
        /// Get structural element thickness for depth parameter
        /// </summary>
        private double GetStructuralElementThickness(Element structuralElement)
        {
            try
            {
                if (structuralElement.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Floors)
                {
                    // For floors, get thickness parameter
                    var floor = structuralElement as Floor;
                    if (floor != null)
                    {
                        var thicknessParam = floor.get_Parameter(BuiltInParameter.FLOOR_ATTR_THICKNESS_PARAM);
                        if (thicknessParam != null && thicknessParam.HasValue)
                        {
                            return thicknessParam.AsDouble();
                        }
                    }
                    // Fallback for floors: use default thickness
                    return UnitUtils.ConvertToInternalUnits(300.0, UnitTypeId.Millimeters); // 300mm default
                }
                else if (structuralElement.Category.Id.IntegerValue == (int)BuiltInCategory.OST_StructuralFraming)
                {
                    // For structural framing, try to get the width/depth
                    var familyInstance = structuralElement as FamilyInstance;
                    if (familyInstance != null)
                    {
                        // Try to get structural dimensions
                        var widthParam = familyInstance.LookupParameter("b") ?? familyInstance.LookupParameter("Width");
                        var depthParam = familyInstance.LookupParameter("h") ?? familyInstance.LookupParameter("Depth");
                        
                        if (widthParam != null && widthParam.HasValue)
                        {
                            return widthParam.AsDouble();
                        }
                        else if (depthParam != null && depthParam.HasValue)
                        {
                            return depthParam.AsDouble();
                        }
                    }
                    // Fallback for structural framing: use default depth
                    return UnitUtils.ConvertToInternalUnits(400.0, UnitTypeId.Millimeters); // 400mm default
                }
                
                // Default fallback
                return UnitUtils.ConvertToInternalUnits(300.0, UnitTypeId.Millimeters); // 300mm default
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"Error getting structural element thickness: {ex.Message}");
                return UnitUtils.ConvertToInternalUnits(300.0, UnitTypeId.Millimeters); // 300mm default
            }
        }
        
        /// <summary>
        /// Safely set parameter value (following working command pattern)
        /// </summary>
        private void SetParameterSafely(FamilyInstance instance, string paramName, double value, int elementId)
        {
            try
            {
                var param = instance.LookupParameter(paramName);
                if (param != null && !param.IsReadOnly)
                {
                    param.Set(value);
                    DebugLogger.Debug($"Set {paramName} = {UnitUtils.ConvertFromInternalUnits(value, UnitTypeId.Millimeters):F1}mm for linked element {elementId}");
                }
                else
                {
                    DebugLogger.Warning($"Parameter '{paramName}' not found or read-only for linked element {elementId}");
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"Failed to set parameter '{paramName}' for linked element {elementId}: {ex.Message}");
            }
        }
    }

    /// <summary>
    /// Data structure to represent an intersection between a linked MEP element and an active structural element
    /// </summary>
    public class LinkedIntersectionResult
    {
        public RevitLinkInstance LinkInstance { get; set; }
        public Element LinkedMEPElement { get; set; }
        public int LinkedMEPElementId { get; set; }
        public Element StructuralElement { get; set; }
        public XYZ IntersectionPoint { get; set; }
        public double Proximity { get; set; }
    }
}

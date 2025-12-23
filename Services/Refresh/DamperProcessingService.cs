using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Services.DamperDetection;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;
using JSE_RevitAddin_MEP_OPENINGS.Data;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Refresh
{
    /// <summary>
    /// ✅ SOLID SRP: Service responsible ONLY for processing dampers separately from MepIntersectionService.
    /// Single Responsibility: Collect dampers, find walls, create ClashZones, capture parameters.
    /// 
    /// This service processes dampers outside the normal intersection detection flow:
    /// 1. Collects dampers using DamperCollectionService (excludes VCD/VOLUME)
    /// 2. Finds which walls each damper intersects/contacts
    /// 3. Creates ClashZone objects with MEP and Host parameters
    /// 4. Returns ClashZones ready for database persistence
    /// 
    /// SOLID Principles:
    /// ✅ SRP: Single Responsibility = process dampers only
    /// ✅ OCP: Open/Closed = can be extended via interface inheritance
    /// ✅ LSP: Liskov Substitution = can be replaced with alternative implementations
    /// ✅ ISP: Interface Segregation = focused interface for damper processing
    /// ✅ DIP: Dependency Inversion = depends on abstractions (IDamperCollectionService, IParameterSnapshotService)
    /// </summary>
    public class DamperProcessingService
    {
        private readonly Document _document;
        private readonly Action<string> _logger;
        private readonly ClashZoneStorage _clashZoneStorage;
        private readonly IDamperTypeDetector _damperTypeDetector;
        private readonly IDamperConnectorDetector _connectorDetector;
        private readonly ParameterSnapshotService _parameterSnapshotService;
        private readonly PerformanceMonitor? _performanceMonitor;
        private readonly GuidManager? _guidManager; // ✅ CRITICAL: For deterministic GUID generation (3-point validation)

        public DamperProcessingService(
            Document document,
            Action<string> logger,
            ClashZoneStorage clashZoneStorage)
            : this(
                document,
                logger,
                clashZoneStorage,
                new DamperTypeDetector(),
                new DamperConnectorDetector(),
                new ParameterSnapshotService(),
                null,
                null)
        {
        }

        public DamperProcessingService(
            Document document,
            Action<string> logger,
            ClashZoneStorage clashZoneStorage,
            IDamperTypeDetector? damperTypeDetector,
            IDamperConnectorDetector? connectorDetector,
            ParameterSnapshotService? parameterSnapshotService,
            PerformanceMonitor? performanceMonitor = null,
            GuidManager? guidManager = null)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _clashZoneStorage = clashZoneStorage ?? throw new ArgumentNullException(nameof(clashZoneStorage));
            // ✅ FIX: Handle null values by creating defaults (SOLID - dependency injection with fallback)
            _damperTypeDetector = damperTypeDetector ?? new DamperTypeDetector();
            _connectorDetector = connectorDetector ?? new DamperConnectorDetector();
            _parameterSnapshotService = parameterSnapshotService ?? new ParameterSnapshotService();
            _performanceMonitor = performanceMonitor;
            _guidManager = guidManager ?? new GuidManager(document); // ✅ CRITICAL: Create GuidManager for deterministic GUID generation
        }

        /// <summary>
        /// ✅ MAIN ENTRY POINT: Process all dampers and create ClashZones.
        /// ✅ PERFORMANCE MONITORING: Tracks time for each operation using PerformanceMonitor (if provided).
        /// </summary>
        /// <param name="selectedMepCategories">MEP categories to process (should include "Duct Accessories")</param>
        /// <param name="selectedHostCategories">Host categories to check (should include "Walls")</param>
        /// <param name="selectedReferenceFiles">Reference files where MEP elements (dampers) are located. If null/empty, only active document is searched.</param>
        /// <param name="sectionBox">Optional section box to limit search area</param>
        /// <param name="existingClashZones">Optional list of existing clash zones to match against (prevents duplicate GUID creation)</param>
        /// <returns>List of ClashZones created from dampers</returns>
        public List<ClashZone> ProcessDampers(
            List<string> selectedMepCategories,
            List<string> selectedHostCategories,
            List<string>? selectedReferenceFiles = null,
            BoundingBoxXYZ? sectionBox = null,
            List<ClashZone>? existingClashZones = null)
        {
            _logger("[DamperProcessing] Starting separate damper processing...");

            var results = new List<ClashZone>();

            // ✅ STEP 1: Check if Duct Accessories category is selected
            if (!selectedMepCategories.Any(c => string.Equals(c, "Duct Accessories", StringComparison.OrdinalIgnoreCase)))
            {
                _logger("[DamperProcessing] Duct Accessories category not selected - skipping damper processing");
                return results;
            }

            // ✅ PERFORMANCE MONITORING: Track total damper processing time
            using (var totalTrackerRaw = _performanceMonitor?.TrackOperation("Damper Processing (Total)"))
                {
                    // ✅ STEP 2: Collect dampers from selected reference files only
                    // This matches the pattern in IntersectionProcessor.CollectMepElementsWithFilters
                    List<(Element damper, Transform? transform, XYZ placementPoint)> collectedDampers;
                    using (var collectTrackerRaw = _performanceMonitor?.TrackOperation("Collect Dampers"))
                    {
                        collectedDampers = CollectDampersFromSelectedFiles(selectedReferenceFiles, sectionBox);
                        // ✅ FIX: Cast to OperationTracker to access SetItemCount (OperationTracker is internal but accessible in same assembly)
                        if (collectTrackerRaw is PerformanceMonitor.OperationTracker collectTracker)
                        {
                            collectTracker.SetItemCount(collectedDampers.Count);
                        }
                    }
                
                _logger($"[DamperProcessing] ✅ Collected {collectedDampers.Count} dampers from selected reference files (VCD/VOLUME excluded)");

                if (collectedDampers.Count == 0)
                {
                    _logger("[DamperProcessing] No dampers found - skipping processing");
                    return results;
                }

                // ✅ STEP 3: Collect walls (host elements)
                List<(Element wall, Transform? transform)> walls;
                using (var wallsTrackerRaw = _performanceMonitor?.TrackOperation("Collect Walls"))
                {
                    walls = CollectWalls(selectedHostCategories, sectionBox);
                    // ✅ FIX: Cast to OperationTracker to access SetItemCount
                    if (wallsTrackerRaw is PerformanceMonitor.OperationTracker wallsTracker)
                    {
                        wallsTracker.SetItemCount(walls.Count);
                    }
                }
                _logger($"[DamperProcessing] ✅ Collected {walls.Count} walls for intersection checking");

                if (walls.Count == 0)
                {
                    _logger("[DamperProcessing] ⚠️ No walls found - cannot create ClashZones for dampers");
                    return results;
                }

                // ✅ STEP 4: For each damper, find intersecting wall and create ClashZone
                int processedCount = 0;
                int createdCount = 0;
                int skippedCount = 0;

                using (var processTrackerRaw = _performanceMonitor?.TrackOperation("Process Dampers (Find Intersections + Create ClashZones)"))
                {
                    foreach (var (damper, transform, placementPoint) in collectedDampers)
                    {
                        try
                        {
                            processedCount++;
                            
                            // ✅ DIAGNOSTIC: Log damper details to identify duplicates
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                string damperDocTitle = damper.Document?.Title ?? "NULL";
                                _logger($"[DamperProcessing] [{processedCount}/{collectedDampers.Count}] Processing damper {damper.Id} from document '{damperDocTitle}' at {placementPoint}");
                            }
                            
                            // ✅ PERFORMANCE: Track individual damper processing (only if diagnostic mode enabled)
                            if (!DeploymentConfiguration.DeploymentMode && OptimizationFlags.UseDiagnosticMode)
                            {
                                _logger($"[DamperProcessing] [{processedCount}/{collectedDampers.Count}] Processing damper {damper.Id} at {placementPoint}");
                            }

                            // Find wall that intersects with this damper
                            Element? intersectingWall;
                            Transform? wallTransform = null;
                            using (var findWallTrackerRaw = _performanceMonitor?.TrackOperation("Find Intersecting Wall"))
                            {
                                intersectingWall = FindIntersectingWall(damper, placementPoint, walls, transform);
                                
                                // ✅ CRITICAL FIX: Find the wall's transform from the walls list
                                // This is needed to transform the wall's bounding box in CreateClashZoneForDamper
                                if (intersectingWall != null)
                                {
                                    var wallMatch = walls.FirstOrDefault(w => w.Item1.Id == intersectingWall.Id);
                                    if (wallMatch.Item1 != null)
                                    {
                                        wallTransform = wallMatch.Item2;
                                    }
                                }
                                
                                // ✅ FIX: Cast to OperationTracker to access SetItemCount
                                if (findWallTrackerRaw is PerformanceMonitor.OperationTracker findWallTracker)
                                {
                                    findWallTracker.SetItemCount(1);
                                }
                            }
                            
                            if (intersectingWall == null)
                            {
                                skippedCount++;
                                if (!DeploymentConfiguration.DeploymentMode && OptimizationFlags.UseDiagnosticMode)
                                {
                                    _logger($"[DamperProcessing] ⚠️ Damper {damper.Id} has no intersecting wall - skipping");
                                }
                                continue;
                            }

                            // Create ClashZone for this damper-wall pair
                            ClashZone? clashZone;
                            using (var createTrackerRaw = _performanceMonitor?.TrackOperation("Create ClashZone"))
                            {
                                // ✅ CRITICAL FIX: Check if zone already exists for this MEP+Host pair before creating new one
                                // This prevents duplicate zones with new GUIDs and preserves existing flags
                                clashZone = FindOrCreateClashZoneForDamper(damper, intersectingWall, placementPoint, transform, wallTransform, existingClashZones, results);
                                // ✅ FIX: Cast to OperationTracker to access SetItemCount
                                if (createTrackerRaw is PerformanceMonitor.OperationTracker createTracker)
                                {
                                    createTracker.SetItemCount(1);
                                }
                            }
                            
                            if (clashZone != null)
                            {
                                // ✅ CRITICAL FIX: Check for duplicate zones by Host+Point (not just MEP+Host)
                                // Multiple dampers at the same location should share the same sleeve
                                int structuralIdValue = intersectingWall.Id.IntegerValue;
                                double tolerance = 0.1; // 0.1ft = ~30mm (same as GUID tolerance)
                                
                                var duplicate = results.FirstOrDefault(r => 
                                {
                                    if (r == null) return false;
                                    int rStructuralId = r.StructuralElementId?.IntegerValue ?? r.StructuralElementIdValue;
                                    if (rStructuralId != structuralIdValue) return false;
                                    
                                    // Check if placement points are within tolerance
                                    double dx = Math.Abs((r.SleevePlacementPoint?.X ?? r.SleevePlacementPointX) - placementPoint.X);
                                    double dy = Math.Abs((r.SleevePlacementPoint?.Y ?? r.SleevePlacementPointY) - placementPoint.Y);
                                    double dz = Math.Abs((r.SleevePlacementPoint?.Z ?? r.SleevePlacementPointZ) - placementPoint.Z);
                                    
                                    return dx < tolerance && dy < tolerance && dz < tolerance;
                                });
                                
                                if (duplicate != null)
                                {
                                    if (!DeploymentConfiguration.DeploymentMode)
                                    {
                                        _logger($"[DamperProcessing] ✅ REUSING existing ClashZone {duplicate.Id} for damper {damper.Id} with wall {intersectingWall.Id} at same location (Host+Point match)");
                                    }
                                    // Skip adding duplicate zone - reuse existing one
                                    continue;
                                }
                                
                                results.Add(clashZone);
                                createdCount++;
                                if (!DeploymentConfiguration.DeploymentMode && OptimizationFlags.UseDiagnosticMode)
                                {
                                    _logger($"[DamperProcessing] ✅ Created ClashZone {clashZone.Id} for damper {damper.Id} with wall {intersectingWall.Id}");
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            _logger($"[DamperProcessing] ⚠️ Error processing damper {damper.Id}: {ex.Message}");
                            // ✅ SAFETY: Log error using SafeFileLogger (respects deployment mode)
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                SafeFileLogger.SafeAppendText("damper_processing_errors.log",
                                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] Error processing damper {damper.Id}: {ex.Message}\nStackTrace: {ex.StackTrace}\n");
                            }
                        }
                    }
                    
                    // ✅ FIX: Cast to OperationTracker to access SetItemCount
                    if (processTrackerRaw is PerformanceMonitor.OperationTracker processTracker)
                    {
                        processTracker.SetItemCount(processedCount);
                    }
                }

                // ✅ FIX: Cast to OperationTracker to access SetItemCount
                if (totalTrackerRaw is PerformanceMonitor.OperationTracker totalTracker)
                {
                    totalTracker.SetItemCount(results.Count);
                }
                
                // ✅ PERFORMANCE LOGGING: Log summary using SafeFileLogger (respects deployment mode)
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    SafeFileLogger.SafeAppendText("damper_processing.log",
                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [SUMMARY] Processed: {processedCount} dampers, Created: {createdCount} ClashZones, Skipped: {skippedCount} (no intersecting wall)\n");
                }
                
                _logger($"[DamperProcessing] ✅ Processing complete: {createdCount} ClashZones created from {processedCount} dampers ({skippedCount} skipped)");
            }

            return results;
        }

        /// <summary>
        /// ✅ Collect dampers from selected reference files only (respects UI selection).
        /// Matches the pattern in IntersectionProcessor.CollectMepElementsWithFilters:
        /// - If no reference files selected → collect from active document only
        /// - If reference files selected → collect from those linked documents only
        /// </summary>
        private List<(Element damper, Transform? transform, XYZ placementPoint)> CollectDampersFromSelectedFiles(
            List<string>? selectedReferenceFiles,
            BoundingBoxXYZ? sectionBox)
        {
            var results = new List<(Element, Transform?, XYZ)>();
            var damperCollectionService = new DamperCollectionService(_damperTypeDetector, _connectorDetector);

            // ✅ STEP 1: Always collect from ACTIVE document first (if no reference files selected, or if active is selected)
            bool shouldCollectFromActive = selectedReferenceFiles == null || 
                                          selectedReferenceFiles.Count == 0 ||
                                          selectedReferenceFiles.Any(f => string.Equals(f, "Active Document", StringComparison.OrdinalIgnoreCase));

            if (shouldCollectFromActive)
            {
                _logger("[DamperProcessing] Collecting dampers from ACTIVE document...");
                var activeDampersList = CollectDampersFromDocument(_document, null, sectionBox);
                
                results.AddRange(activeDampersList);
                _logger($"[DamperProcessing] ✅ Collected {activeDampersList.Count} dampers from ACTIVE document");
            }

            // ✅ STEP 2: Collect from selected linked documents (if any)
            if (selectedReferenceFiles != null && selectedReferenceFiles.Count > 0)
            {
                // Get all link instances
                var allLinkInstances = new FilteredElementCollector(_document)
                    .OfClass(typeof(RevitLinkInstance))
                    .Cast<RevitLinkInstance>()
                    .Where(link => link.GetLinkDocument() != null)
                    .ToList();

                _logger($"[DamperProcessing] Found {allLinkInstances.Count} total link instances in document");

                // Match selected reference files to link instances
                var matchedLinks = new List<RevitLinkInstance>();
                foreach (var selectedFile in selectedReferenceFiles)
                {
                    if (string.Equals(selectedFile, "Active Document", StringComparison.OrdinalIgnoreCase))
                        continue; // Already handled above

                    foreach (var link in allLinkInstances)
                    {
                        var linkDoc = link.GetLinkDocument();
                        if (linkDoc == null) continue;

                        // Match by name, title, or path (same logic as IntersectionProcessor)
                        var linkName = link.Name ?? "";
                        var linkTitle = linkDoc.Title ?? "";
                        var linkPath = System.IO.Path.GetFileNameWithoutExtension(linkDoc.PathName ?? "");

                        bool matches = string.Equals(selectedFile, linkName, StringComparison.OrdinalIgnoreCase) ||
                                      string.Equals(selectedFile, linkTitle, StringComparison.OrdinalIgnoreCase) ||
                                      string.Equals(selectedFile, linkPath, StringComparison.OrdinalIgnoreCase);

                        if (matches && !matchedLinks.Contains(link))
                        {
                            matchedLinks.Add(link);
                            _logger($"[DamperProcessing] ✅ Matched linked file: Name='{link.Name}', Title='{linkDoc.Title}', PathName='{linkPath}'");
                            break;
                        }
                    }
                }

                _logger($"[DamperProcessing] Matched {matchedLinks.Count} linked files for damper collection");

                // Collect dampers from matched linked documents
                foreach (var link in matchedLinks)
                {
                    var linkDoc = link.GetLinkDocument();
                    if (linkDoc == null) continue;

                    var linkTransform = link.GetTotalTransform();
                    _logger($"[DamperProcessing] Collecting dampers from linked document: {linkDoc.Title}");

                    var linkDampersList = CollectDampersFromDocument(linkDoc, linkTransform, sectionBox);
                    
                    results.AddRange(linkDampersList);
                    _logger($"[DamperProcessing] ✅ Collected {linkDampersList.Count} dampers from linked document: {linkDoc.Title}");
                }
            }
            else
            {
                _logger("[DamperProcessing] ⚠️ No reference files selected - only collecting from active document");
            }

            _logger($"[DamperProcessing] ========== DAMPER COLLECTION COMPLETE: {results.Count} total dampers ==========");
            return results;
        }

        /// <summary>
        /// ✅ Collect dampers from a single document (host or linked).
        /// ✅ 28-FEATURE OPTIMIZATION: Applies section box filter at FilteredElementCollector level (like other MEP elements).
        /// Filters by Duct Accessories category and Standard/Motorized damper types.
        /// </summary>
        private List<(Element, Transform?, XYZ)> CollectDampersFromDocument(
            Document document,
            Transform? linkTransform,
            BoundingBoxXYZ? sectionBox)
        {
            var results = new List<(Element, Transform?, XYZ)>();
            
            try
            {
                // ✅ STEP 1: Build collector with category filter
                var collector = new FilteredElementCollector(document)
                    .OfCategory(BuiltInCategory.OST_DuctAccessory)
                    .OfClass(typeof(FamilyInstance));

                // ✅ 28-FEATURE OPTIMIZATION: Apply section box filter at collector level (before loading elements)
                // This prevents loading 611 dampers into memory when only 7 are needed
                if (sectionBox != null)
                {
                    // Transform section box to document coordinates if from linked document
                    Outline outline;
                    if (linkTransform != null)
                    {
                        // Transform section box min/max to linked document coordinates
                        var inverseTransform = linkTransform.Inverse;
                        var docMin = inverseTransform.OfPoint(sectionBox.Min);
                        var docMax = inverseTransform.OfPoint(sectionBox.Max);
                        outline = new Outline(docMin, docMax);
                    }
                    else
                    {
                        outline = new Outline(sectionBox.Min, sectionBox.Max);
                    }

                    // ✅ CRITICAL: Use BoundingBoxIntersectsFilter - handles PARTIAL intersections correctly
                    // This means dampers that are PARTIALLY within section box are included (as required)
                    collector = collector.WherePasses(new BoundingBoxIntersectsFilter(outline));
                    
                    _logger($"[DamperProcessing] ✅ Applied section box filter at collector level for document: {document.Title}");
                }

                // ✅ STEP 2: Load elements and filter by damper type (Standard/Motorized, exclude VCD/VOLUME)
                var allDuctAccessories = collector.Cast<FamilyInstance>().ToList();
                _logger($"[DamperProcessing] Found {allDuctAccessories.Count} total Duct Accessories in document: {document.Title} (before damper filter)");
                
                var ductAccessories = allDuctAccessories
                    .Where(fi => ShouldProcessDamper(fi))
                    .ToList();

                _logger($"[DamperProcessing] ✅ Filtered to {ductAccessories.Count} Standard/Motorized dampers in document: {document.Title} (after damper type filter)");

                // ✅ STEP 3: Calculate placement points for each damper (section box already filtered)
                foreach (var damper in ductAccessories)
                {
                    // ✅ STEP 3a: Detect damper type
                    string familyTypeName = damper.Symbol?.Name ?? "";
                    string damperType = _damperTypeDetector.DetectDamperType(familyTypeName);
                    bool isStandard = _damperTypeDetector.IsStandardDamper(damperType);

                    XYZ placementPoint;

                    // ✅ STEP 3b: Calculate placement point based on damper type
                    if (isStandard)
                    {
                        // ✅ STANDARD DAMPERS: Use bounding box centroid (DamperPlacementStrategy approach)
                        var bbox = damper.get_BoundingBox(null);
                        if (bbox == null)
                        {
                            _logger($"[DamperProcessing] ⚠️ Damper {damper.Id} has no bounding box - skipping");
                            continue;
                        }

                        // Calculate centroid of bounding box
                        placementPoint = new XYZ(
                            (bbox.Min.X + bbox.Max.X) / 2.0,
                            (bbox.Min.Y + bbox.Max.Y) / 2.0,
                            (bbox.Min.Z + bbox.Max.Z) / 2.0
                        );
                    }
                    else
                    {
                        // ✅ NON-STANDARD DAMPERS: Use connector detector to find placement point
                        // Use world coordinates for non-standard dampers (they may be rotated/flipped)
                        string connectorSide = _connectorDetector.DetectConnectorSide(
                            damper,
                            useWorldCoordinates: true, // Non-standard dampers use world coordinates
                            out Connector connector,
                            wallOrientation: null); // No wall orientation available at collection time

                        if (connector != null && connector.Origin != null)
                        {
                            placementPoint = connector.Origin;
                        }
                        else
                        {
                            // Fallback to bounding box centroid
                            var bbox = damper.get_BoundingBox(null);
                            if (bbox == null)
                            {
                                _logger($"[DamperProcessing] ⚠️ Damper {damper.Id} has no bounding box - skipping");
                                continue;
                            }

                            // Calculate centroid of bounding box
                            placementPoint = new XYZ(
                                (bbox.Min.X + bbox.Max.X) / 2.0,
                                (bbox.Min.Y + bbox.Max.Y) / 2.0,
                                (bbox.Min.Z + bbox.Max.Z) / 2.0
                            );
                        }
                    }

                    // ✅ STEP 4: Transform placement point to host coordinates if from linked document
                    if (linkTransform != null)
                    {
                        placementPoint = linkTransform.OfPoint(placementPoint);
                    }

                    results.Add((damper, linkTransform, placementPoint));
                }
            }
            catch (Exception ex)
            {
                _logger($"[DamperProcessing] ⚠️ Error collecting dampers from document {document.Title}: {ex.Message}");
            }

            return results;
        }

        /// <summary>
        /// ✅ SIMPLIFIED: Check if damper should be processed.
        /// Only processes duct accessories where family name contains "Damper" (case-insensitive).
        /// Excludes VCD/VOLUME families (not in walls).
        /// </summary>
        private bool ShouldProcessDamper(FamilyInstance damper)
        {
            var familyName = damper.Symbol?.Family?.Name ?? "";
            var typeName = damper.Symbol?.Name ?? "";

            // ✅ SIMPLIFIED: Only check if family name contains "Damper" (case-insensitive)
            // All other duct accessory types (filters, coils, etc.) are excluded
            if (familyName.IndexOf("Damper", StringComparison.OrdinalIgnoreCase) < 0)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    _logger($"[DamperProcessing] ⏭️ SKIP: Damper {damper.Id} - FamilyName='{familyName}' does not contain 'Damper'");
                }
                return false;
            }

            // ✅ CRITICAL: Exclude VCD/VOLUME families (not in walls)
            if (familyName.IndexOf("VCD", StringComparison.OrdinalIgnoreCase) >= 0 ||
                familyName.IndexOf("VOLUME", StringComparison.OrdinalIgnoreCase) >= 0 ||
                typeName.IndexOf("VCD", StringComparison.OrdinalIgnoreCase) >= 0 ||
                typeName.IndexOf("VOLUME", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    _logger($"[DamperProcessing] ⏭️ SKIP: Damper {damper.Id} - FamilyName='{familyName}', TypeName='{typeName}' contains VCD/VOLUME");
                }
                return false;
            }

            // ✅ SIMPLIFIED: Process ALL dampers (no complex type checking)
            if (!DeploymentConfiguration.DeploymentMode)
            {
                _logger($"[DamperProcessing] ✅ PROCESS: Damper {damper.Id} - FamilyName='{familyName}', TypeName='{typeName}'");
            }

            return true;
        }

        /// <summary>
        /// ✅ Collect walls from documents (host and linked).
        /// </summary>
        private List<(Element wall, Transform? transform)> CollectWalls(
            List<string> selectedHostCategories,
            BoundingBoxXYZ? sectionBox)
        {
            var walls = new List<(Element, Transform?)>();

            if (!selectedHostCategories.Any(c => string.Equals(c, "Walls", StringComparison.OrdinalIgnoreCase)))
            {
                return walls;
            }

            // Collect from host document
            var hostWalls = new FilteredElementCollector(_document)
                .OfCategory(BuiltInCategory.OST_Walls)
                .WhereElementIsNotElementType()
                .ToList();

            foreach (var wall in hostWalls)
            {
                // Apply section box filter if provided
                if (sectionBox != null)
                {
                    var wallBbox = wall.get_BoundingBox(null);
                    if (wallBbox != null && !BoundingBoxService.BoundingBoxesIntersect(
                        sectionBox.Min, sectionBox.Max, wallBbox.Min, wallBbox.Max))
                    {
                        continue;
                    }
                }
                walls.Add((wall, null));
            }

            // Collect from linked documents
            var links = new FilteredElementCollector(_document)
                .OfClass(typeof(RevitLinkInstance))
                .Cast<RevitLinkInstance>()
                .ToList();

            foreach (var link in links)
            {
                var linkDoc = link.GetLinkDocument();
                if (linkDoc == null) continue;

                var linkTransform = link.GetTotalTransform();
                var linkedWalls = new FilteredElementCollector(linkDoc)
                    .OfCategory(BuiltInCategory.OST_Walls)
                    .WhereElementIsNotElementType()
                    .ToList();

                foreach (var wall in linkedWalls)
                {
                    // Apply section box filter if provided (transform coordinates)
                    if (sectionBox != null)
                    {
                        var wallBbox = wall.get_BoundingBox(null);
                        if (wallBbox != null)
                        {
                            // Transform wall bbox to host coordinates
                            var transformedMin = linkTransform.OfPoint(wallBbox.Min);
                            var transformedMax = linkTransform.OfPoint(wallBbox.Max);
                            var transformedBbox = new BoundingBoxXYZ
                            {
                                Min = new XYZ(Math.Min(transformedMin.X, transformedMax.X), Math.Min(transformedMin.Y, transformedMax.Y), Math.Min(transformedMin.Z, transformedMax.Z)),
                                Max = new XYZ(Math.Max(transformedMin.X, transformedMax.X), Math.Max(transformedMin.Y, transformedMax.Y), Math.Max(transformedMin.Z, transformedMax.Z))
                            };

                            if (!BoundingBoxService.BoundingBoxesIntersect(
                                sectionBox.Min, sectionBox.Max, transformedBbox.Min, transformedBbox.Max))
                            {
                                continue;
                            }
                        }
                    }
                    walls.Add((wall, linkTransform));
                }
            }

            return walls;
        }

        /// <summary>
        /// ✅ Find the wall that intersects with a damper using bounding box overlap.
        /// ✅ CRITICAL: Prioritizes walls where the damper's placement point is INSIDE the wall's bounding box.
        /// ✅ CRITICAL: Also checks if damper has a Host property pointing to a wall (most reliable).
        /// </summary>
        private Element? FindIntersectingWall(
            Element damper,
            XYZ damperPlacementPoint,
            List<(Element wall, Transform? transform)> walls,
            Transform? damperTransform)
        {
            var damperBbox = damper.get_BoundingBox(null);
            if (damperBbox == null) return null;

            // Transform damper bbox to host coordinates if needed
            BoundingBoxXYZ hostDamperBbox = damperBbox;
            if (damperTransform != null)
            {
                var transformed = BoundingBoxService.TransformBoundingBox(damperBbox, damperTransform);
                if (transformed != null)
                {
                    hostDamperBbox = transformed;
                }
            }

            // ✅ PRIORITY 1: Check if damper has a Host property pointing to a wall (most reliable)
            if (damper is FamilyInstance damperInstance)
            {
                try
                {
                    var hostElement = damperInstance.Host;
                    if (hostElement != null && hostElement is Wall hostWall)
                    {
                        // Check if this host wall is in our walls list
                        var hostWallMatch = walls.FirstOrDefault(w => 
                            w.Item1.Id == hostWall.Id || 
                            (w.Item1 is Wall ww && ww.UniqueId == hostWall.UniqueId));
                        
                        if (hostWallMatch.Item1 != null)
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                _logger($"[DamperProcessing] ✅ Found host wall {hostWall.Id} for damper {damper.Id} (from Host property)");
                            }
                            return hostWallMatch.Item1;
                        }
                        else
                        {
                            // Host wall found but not in our list - add it if it's a valid wall
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                _logger($"[DamperProcessing] ⚠️ Damper {damper.Id} has host wall {hostWall.Id} but it's not in the walls list - will check other walls");
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        _logger($"[DamperProcessing] ⚠️ Error checking damper Host property: {ex.Message}");
                    }
                }
            }

            const double tolerance = 0.5; // 6 inches tolerance
            var expandedMin = new XYZ(
                hostDamperBbox.Min.X - tolerance,
                hostDamperBbox.Min.Y - tolerance,
                hostDamperBbox.Min.Z - tolerance);
            var expandedMax = new XYZ(
                hostDamperBbox.Max.X + tolerance,
                hostDamperBbox.Max.Y + tolerance,
                hostDamperBbox.Max.Z + tolerance);

            // ✅ PRIORITY 2: Find walls where placement point is INSIDE the wall's bounding box (most accurate)
            Element? bestWall = null;
            bool bestWallHasPointInside = false;
            
            // Check each wall for intersection
            int checkedWalls = 0;
            foreach (var (wall, wallTransform) in walls)
            {
                try
                {
                    checkedWalls++;
                    var wallBbox = wall.get_BoundingBox(null);
                    if (wallBbox == null) continue;

                    // Transform wall bbox to host coordinates if needed
                    if (wallTransform != null)
                    {
                        var transformed = BoundingBoxService.TransformBoundingBox(wallBbox, wallTransform);
                        if (transformed != null)
                        {
                            wallBbox = transformed;
                        }
                    }

                    // ✅ CRITICAL: Check if placement point is INSIDE wall's bounding box (with small tolerance)
                    bool isPointInsideWall = 
                        damperPlacementPoint.X >= wallBbox.Min.X - tolerance &&
                        damperPlacementPoint.X <= wallBbox.Max.X + tolerance &&
                        damperPlacementPoint.Y >= wallBbox.Min.Y - tolerance &&
                        damperPlacementPoint.Y <= wallBbox.Max.Y + tolerance &&
                        damperPlacementPoint.Z >= wallBbox.Min.Z - tolerance &&
                        damperPlacementPoint.Z <= wallBbox.Max.Z + tolerance;

                    // Check bounding box intersection
                    bool bboxesIntersect = BoundingBoxService.BoundingBoxesIntersect(
                        expandedMin, expandedMax, wallBbox.Min, wallBbox.Max);

                    if (isPointInsideWall)
                    {
                        // ✅ PRIORITY: Placement point is inside this wall - this is the host wall
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            _logger($"[DamperProcessing] ✅ Found HOST wall {wall.Id} for damper {damper.Id} - placement point is INSIDE wall bbox");
                        }
                        return wall; // Return immediately - this is the correct host wall
                    }
                    else if (bboxesIntersect && !bestWallHasPointInside)
                    {
                        // Check if damper is completely contained within wall (buried inside)
                        bool isDamperContainedInWall = 
                            hostDamperBbox.Min.X >= wallBbox.Min.X &&
                            hostDamperBbox.Min.Y >= wallBbox.Min.Y &&
                            hostDamperBbox.Min.Z >= wallBbox.Min.Z &&
                            hostDamperBbox.Max.X <= wallBbox.Max.X &&
                            hostDamperBbox.Max.Y <= wallBbox.Max.Y &&
                            hostDamperBbox.Max.Z <= wallBbox.Max.Z;

                        if (isDamperContainedInWall)
                        {
                            // Damper is contained in wall - this is likely the host wall
                            bestWall = wall;
                            bestWallHasPointInside = false; // Point not inside, but damper is contained
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                _logger($"[DamperProcessing] ✅ Found intersecting wall {wall.Id} for damper {damper.Id} (damper contained in wall)");
                            }
                        }
                        else if (bestWall == null)
                        {
                            // Fallback: Just intersecting (not ideal, but better than nothing)
                            bestWall = wall;
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                _logger($"[DamperProcessing] ⚠️ Found intersecting wall {wall.Id} for damper {damper.Id} (bboxes intersect but point not inside)");
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    _logger($"[DamperProcessing] ⚠️ Error checking wall {wall.Id}: {ex.Message}");
                }
            }

            // Return best wall found (prioritizes walls with point inside, then contained, then just intersecting)
            if (bestWall != null)
            {
                return bestWall;
            }

            // ✅ DIAGNOSTIC: Log why no wall was found
            if (checkedWalls > 0)
            {
                _logger($"[DamperProcessing] ⚠️ Damper {damper.Id} at ({damperPlacementPoint.X:F2}, {damperPlacementPoint.Y:F2}, {damperPlacementPoint.Z:F2}) checked {checkedWalls} walls but none intersect (tolerance={0.5 * 304.8:F1}mm)");
            }
            else
            {
                _logger($"[DamperProcessing] ⚠️ Damper {damper.Id} at ({damperPlacementPoint.X:F2}, {damperPlacementPoint.Y:F2}, {damperPlacementPoint.Z:F2}) - no walls available to check");
            }

            return null;
        }

        /// <summary>
        /// ✅ Create a ClashZone for a damper-wall pair.
        /// Captures MEP and Host parameters using ParameterSnapshotService.
        /// </summary>
        /// <summary>
        /// ✅ CRITICAL FIX: Find existing zone or create new one for damper.
        /// Matches by MEP element ID + Structural element ID to prevent duplicate GUID creation.
        /// </summary>
        private ClashZone? FindOrCreateClashZoneForDamper(
            Element damper,
            Element wall,
            XYZ placementPoint,
            Transform? damperTransform,
            Transform? wallTransform = null,
            List<ClashZone>? existingClashZones = null,
            List<ClashZone>? currentResults = null)
        {
            int mepIdValue = damper.Id.IntegerValue;
            int structuralIdValue = wall.Id.IntegerValue;
            double tolerance = 0.1; // 0.1ft = ~30mm (same as GUID tolerance)
            
            // ✅ STEP 1: Check if zone already exists for this MEP+Host pair (from previous refresh)
            if (existingClashZones != null && existingClashZones.Count > 0)
            {
                var existingZone = existingClashZones.FirstOrDefault(cz =>
                {
                    if (cz == null) return false;
                    int czMepId = cz.MepElementId?.IntegerValue ?? cz.MepElementIdValue;
                    int czStructuralId = cz.StructuralElementId?.IntegerValue ?? cz.StructuralElementIdValue;
                    return czMepId == mepIdValue && czStructuralId == structuralIdValue;
                });
                
                if (existingZone != null)
                {
                    // ✅ REUSE EXISTING ZONE: Update properties but preserve GUID and flags
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        _logger($"[DamperProcessing] ✅ REUSING existing ClashZone {existingZone.Id} for damper {damper.Id} with wall {wall.Id} (preserving GUID and flags)");
                    }
                    
                    // Update zone properties with fresh data from damper processing
                    var updatedZone = CreateClashZoneForDamper(damper, wall, placementPoint, damperTransform, wallTransform);
                    if (updatedZone != null)
                    {
                        // Preserve existing GUID and flags
                        updatedZone.Id = existingZone.Id;
                        updatedZone.IsResolved = existingZone.IsResolved;
                        updatedZone.IsClusterResolved = existingZone.IsClusterResolved;
                        updatedZone.SleeveInstanceId = existingZone.SleeveInstanceId;
                        updatedZone.ClusterSleeveInstanceId = existingZone.ClusterSleeveInstanceId;
                        updatedZone.ReadyForPlacement = existingZone.ReadyForPlacement;
                        
                        return updatedZone;
                    }
                }
            }
            
            // ✅ STEP 2: Check if zone already exists in current refresh results by Host+Point
            // Multiple dampers at the same location should share the same sleeve
            if (currentResults != null && currentResults.Count > 0)
            {
                var existingZoneByLocation = currentResults.FirstOrDefault(cz =>
                {
                    if (cz == null) return false;
                    int czStructuralId = cz.StructuralElementId?.IntegerValue ?? cz.StructuralElementIdValue;
                    if (czStructuralId != structuralIdValue) return false;
                    
                    // Check if placement points are within tolerance
                    double dx = Math.Abs((cz.SleevePlacementPoint?.X ?? cz.SleevePlacementPointX) - placementPoint.X);
                    double dy = Math.Abs((cz.SleevePlacementPoint?.Y ?? cz.SleevePlacementPointY) - placementPoint.Y);
                    double dz = Math.Abs((cz.SleevePlacementPoint?.Z ?? cz.SleevePlacementPointZ) - placementPoint.Z);
                    
                    return dx < tolerance && dy < tolerance && dz < tolerance;
                });
                
                if (existingZoneByLocation != null)
                {
                    // ✅ REUSE EXISTING ZONE: Use the same GUID for dampers at the same location
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        _logger($"[DamperProcessing] ✅ REUSING ClashZone {existingZoneByLocation.Id} for damper {damper.Id} with wall {wall.Id} at same location (Host+Point match within {tolerance * 304.8:F1}mm tolerance)");
                    }
                    
                    // Create new zone but reuse the GUID from existing zone
                    var newZone = CreateClashZoneForDamper(damper, wall, placementPoint, damperTransform, wallTransform);
                    if (newZone != null)
                    {
                        // Reuse GUID from existing zone at same location
                        newZone.Id = existingZoneByLocation.Id;
                        // Preserve flags from existing zone if it has a sleeve
                        if (existingZoneByLocation.SleeveInstanceId > 0)
                        {
                            newZone.IsResolved = existingZoneByLocation.IsResolved;
                            newZone.IsClusterResolved = existingZoneByLocation.IsClusterResolved;
                            newZone.SleeveInstanceId = existingZoneByLocation.SleeveInstanceId;
                            newZone.ClusterSleeveInstanceId = existingZoneByLocation.ClusterSleeveInstanceId;
                            newZone.ReadyForPlacement = existingZoneByLocation.ReadyForPlacement;
                        }
                        
                        return newZone;
                    }
                }
            }
            
            // ✅ STEP 3: No existing zone found - create new one
            return CreateClashZoneForDamper(damper, wall, placementPoint, damperTransform, wallTransform);
        }
        
        private ClashZone? CreateClashZoneForDamper(
            Element damper,
            Element wall,
            XYZ placementPoint,
            Transform? damperTransform,
            Transform? wallTransform = null)
        {
            try
            {
                // Get damper bounding box
                var damperBbox = damper.get_BoundingBox(null);
                if (damperBbox == null)
                {
                    _logger($"[DamperProcessing] ⚠️ Damper {damper.Id} has no bounding box - cannot create ClashZone");
                    return null;
                }

                // Transform damper bbox to host coordinates if needed
                if (damperTransform != null)
                {
                    var transformed = BoundingBoxService.TransformBoundingBox(damperBbox, damperTransform);
                    if (transformed != null)
                    {
                        damperBbox = transformed;
                    }
                }

                // Get wall bounding box
                var wallBbox = wall.get_BoundingBox(null);
                if (wallBbox == null)
                {
                    _logger($"[DamperProcessing] ⚠️ Wall {wall.Id} has no bounding box - cannot create ClashZone");
                    return null;
                }

                // ✅ CRITICAL FIX: Transform wall bbox to host coordinates if wall is from linked document
                if (wallTransform != null)
                {
                    var transformed = BoundingBoxService.TransformBoundingBox(wallBbox, wallTransform);
                    if (transformed != null)
                    {
                        wallBbox = transformed;
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            _logger($"[DamperProcessing] ✅ Transformed wall {wall.Id} bounding box to host coordinates");
                        }
                    }
                }

                // Calculate intersection bounding box
                var intersectionMin = new XYZ(
                    Math.Max(damperBbox.Min.X, wallBbox.Min.X),
                    Math.Max(damperBbox.Min.Y, wallBbox.Min.Y),
                    Math.Max(damperBbox.Min.Z, wallBbox.Min.Z));
                var intersectionMax = new XYZ(
                    Math.Min(damperBbox.Max.X, wallBbox.Max.X),
                    Math.Min(damperBbox.Max.Y, wallBbox.Max.Y),
                    Math.Min(damperBbox.Max.Z, wallBbox.Max.Z));

                // If damper is contained, use damper's bounding box
                bool isDamperContainedInWall = 
                    damperBbox.Min.X >= wallBbox.Min.X &&
                    damperBbox.Min.Y >= wallBbox.Min.Y &&
                    damperBbox.Min.Z >= wallBbox.Min.Z &&
                    damperBbox.Max.X <= wallBbox.Max.X &&
                    damperBbox.Max.Y <= wallBbox.Max.Y &&
                    damperBbox.Max.Z <= wallBbox.Max.Z;

                if (isDamperContainedInWall)
                {
                    intersectionMin = damperBbox.Min;
                    intersectionMax = damperBbox.Max;
                }
                else if (intersectionMin.X > intersectionMax.X ||
                         intersectionMin.Y > intersectionMax.Y ||
                         intersectionMin.Z > intersectionMax.Z)
                {
                    // ✅ CRITICAL FIX: If raw bboxes don't overlap (wall found via tolerance in FindIntersectingWall),
                    // use damper's bbox as intersection since the damper is what needs the sleeve
                    // This handles cases where wall is within tolerance but bboxes don't actually overlap
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        _logger($"[DamperProcessing] ⚠️ Raw bboxes don't overlap for damper {damper.Id} and wall {wall.Id} (wall found via tolerance) - using damper bbox as intersection");
                    }
                    intersectionMin = damperBbox.Min;
                    intersectionMax = damperBbox.Max;
                }

                var intersectionBbox = new BoundingBoxXYZ
                {
                    Min = intersectionMin,
                    Max = intersectionMax
                };

                // ✅ CRITICAL FIX: Calculate intersection point
                // Priority 1: Use placementPoint (already transformed to host coordinates)
                // Priority 2: Use intersection bbox center (already in host coordinates after transformation)
                // ✅ IMPORTANT: placementPoint is already transformed to host coordinates in CollectDampersFromDocument
                XYZ intersectionPoint = placementPoint ?? BoundingBoxService.GetBoundingBoxCenter(intersectionBbox);
                
                // ✅ DIAGNOSTIC: Log placement point to verify coordinate system
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    _logger($"[DamperProcessing] 🔍 CreateClashZoneForDamper: placementPoint={placementPoint?.ToString() ?? "NULL"}, intersectionPoint={intersectionPoint}, damperTransform={(damperTransform != null ? "Present" : "NULL")}, wallTransform={(wallTransform != null ? "Present" : "NULL")}");
                    _logger($"[DamperProcessing] 🔍 BoundingBoxes: damperBbox Min={damperBbox.Min}, Max={damperBbox.Max}, wallBbox Min={wallBbox.Min}, Max={wallBbox.Max}");
                    _logger($"[DamperProcessing] 🔍 IntersectionBbox: Min={intersectionBbox.Min}, Max={intersectionBbox.Max}, Center={BoundingBoxService.GetBoundingBoxCenter(intersectionBbox)}");
                }

                // ✅ STEP 1: Capture MEP parameters with smart level parameter logic
                // Priority: Reference Level > other level params > restrict to Reference Level only
                var mepParameters = CaptureMepParametersWithSmartLevel(damper);
                
                // ✅ STEP 2: Capture Host parameters (restricted to Level and Fire Rating only)
                var hostParameters = CaptureHostParametersRestricted(wall);

                // ✅ STEP 3: Extract damper dimensions (Width and Height) - CRITICAL for placement sizing
                // Use same logic as DamperPlacementStrategy.GetMepElementSize
                double damperWidth = 0.0;
                double damperHeight = 0.0;
                
                var damperInstance = damper as FamilyInstance;
                if (damperInstance != null)
                {
                    // Get size from damper-specific parameters (not generic Width/Height)
                    // Fire dampers use "Damper Width" and "Damper Height" parameters
                    var widthParam = damperInstance.LookupParameter("Damper Width") ?? 
                                    damperInstance.LookupParameter("Width") ?? 
                                    damperInstance.LookupParameter("width");
                    var heightParam = damperInstance.LookupParameter("Damper Height") ?? 
                                     damperInstance.LookupParameter("Height") ?? 
                                     damperInstance.LookupParameter("height");
                    
                    damperWidth = widthParam?.AsDouble() ?? 0.0;
                    damperHeight = heightParam?.AsDouble() ?? 0.0;
                    
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        _logger($"[DamperProcessing] ✅ Extracted damper dimensions: Width={damperWidth * 304.8:F1}mm ({damperWidth:F6}ft), Height={damperHeight * 304.8:F1}mm ({damperHeight:F6}ft) from parameters '{widthParam?.Definition.Name ?? "NULL"}' and '{heightParam?.Definition.Name ?? "NULL"}'");
                    }
                }
                
                // ✅ STEP 4: Create ClashZone using ClashZoneService_Legacy.CreateClashZone
                // We'll use a simplified approach to create ClashZone directly
                // ✅ CRITICAL: Set document keys for validation (IsValidClashZone requires at least one non-empty key)
                var damperDoc = damper.Document;
                var wallDoc = wall.Document;
                var hostDoc = _document; // Host document (where walls are typically located)
                
                // ✅ CRITICAL: Get StructuralElementType for filtering (required for placement)
                // ✅ SOLID: Use helper method to determine structural element type
                string structuralElementType = GetStructuralElementType(wall);
                
                // ✅ X-WALL/Y-WALL ROTATION: Get host orientation for rotation logic (same as other MEP elements)
                // This is critical for SleeveRotationService to determine correct rotation angle
                string hostOrientation = WallDirectionService.GetHostOrientation(wall);
                
                // ✅ X-WALL/Y-WALL ROTATION: Get damper orientation (MEP element orientation) for rotation logic
                // Dampers use transform BasisX as flow direction (similar to ducts/pipes)
                XYZ mepElementOrientation = null;
                string mepElementOrientationDirection = string.Empty;
                if (damperInstance != null)
                {
                    try
                    {
                        var transform = damperInstance.GetTotalTransform();
                        mepElementOrientation = transform.BasisX; // Use BasisX as flow direction
                        
                        // ✅ CRITICAL: Calculate MEP orientation direction (X or Y) for rotation logic
                        // This is used by SleeveRotationService to determine if sleeve needs 90° rotation
                        // Logic: If MEP flows primarily in X direction → "X", if primarily in Y → "Y"
                        double absX = Math.Abs(mepElementOrientation.X);
                        double absY = Math.Abs(mepElementOrientation.Y);
                        
                        if (absX > absY)
                        {
                            mepElementOrientationDirection = "X";
                        }
                        else if (absY > absX)
                        {
                            mepElementOrientationDirection = "Y";
                        }
                        // If equal or both near zero, leave empty (will use fallback in rotation service)
                        
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            _logger($"[DamperProcessing] 🔍 Damper Orientation: Vector=({mepElementOrientation.X:F3}, {mepElementOrientation.Y:F3}, {mepElementOrientation.Z:F3}), Direction='{mepElementOrientationDirection}'");
                        }
                    }
                    catch (Exception ex)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            _logger($"[DamperProcessing] ⚠️ Error getting damper orientation: {ex.Message}");
                        }
                    }
                }
                
                // ✅ X-WALL/Y-WALL ROTATION: Get structural element normal for caching
                XYZ structuralElementNormal = WallDirectionService.GetStructuralElementNormal(wall);
                
                // ✅ DEPTH PARAMETER: Get wall width and framing thickness (same as other MEP elements)
                // ✅ FLOW VERIFICATION: This mimics the exact same flow as ClashZoneService_Legacy:
                //    1. ClashZoneService_Legacy: GetWallThickness(structuralElement) → wall.Width
                //    2. ClashZoneService_Legacy: WallThickness = GetWallThickness(structuralElement) on ClashZone
                //    3. ClashZoneService_Legacy: ClashZone saved to DB via BulkInsertClashZones
                //    4. Placement: SleeveParameterService.SetDepthParameter uses GetThickness() → reads zone.WallThickness from DB
                //    DAMPERS: Same flow - wall.Width → WallThickness on ClashZone → saved to DB → read during placement
                // These are used by SleeveParameterService.SetDepthParameter to set "Wall Width" or "Depth" parameter
                double wallThickness = 0.0;
                double framingThickness = 0.0;
                double structuralElementThickness = 0.0;
                
                if (wall is Wall wallElement)
                {
                    // ✅ WALL: Get wall.Width property (used for "Wall Width" parameter on sleeve)
                    // ✅ SAME AS OTHER CATEGORIES: ClashZoneService_Legacy.GetWallThickness() also uses wall.Width
                    // ✅ CRITICAL FIX: If wall is from linked file, retrieve it from the linked document to ensure correct width
                    // This fixes the issue where wall.Width returns 100mm instead of 200mm for linked walls
                    wallThickness = wallElement.Width;
                    
                    // ✅ CRITICAL FIX: If wall is from linked file and width seems incorrect, try retrieving from linked document
                    if (wall.Document != _document && wallThickness > 0)
                    {
                        try
                        {
                            // Get the linked document
                            var linkDoc = wall.Document;
                            if (linkDoc != null)
                            {
                                // Try to get the wall element directly from the linked document
                                var wallFromLinkDoc = linkDoc.GetElement(wall.Id);
                                if (wallFromLinkDoc is Wall linkedWall && linkedWall.Width > 0)
                                {
                                    // Use the width from the linked document's wall element
                                    wallThickness = linkedWall.Width;
                                    if (!DeploymentConfiguration.DeploymentMode)
                                    {
                                        _logger($"[DamperProcessing] ✅ Retrieved wall width from linked document: {wallThickness * 304.8:F1}mm (was {wallElement.Width * 304.8:F1}mm from original element)");
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                _logger($"[DamperProcessing] ⚠️ Error retrieving wall from linked document: {ex.Message} - using original width");
                            }
                        }
                    }
                    
                    structuralElementThickness = wallThickness; // Also set as general thickness
                    
                    // ✅ DIAGNOSTIC: Log wall document and element ID to verify we're reading from correct wall
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        string wallDocTitle = wall.Document?.Title ?? "NULL";
                        string wallDocPath = wall.Document?.PathName ?? "NULL";
                        bool isLinked = wall.Document != _document;
                        
                        _logger($"[DamperProcessing] ✅ Wall Thickness: {wallThickness * 304.8:F1}mm (from wall.Width), WallId={wallElement.Id.IntegerValue}, WallDoc={wallDocTitle}, IsLinked={isLinked}");
                        SafeFileLogger.SafeAppendText("damper_depth_debug.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] [DamperProcessing] ✅ Wall ElementId={wallElement.Id.IntegerValue}, " +
                            $"WallDocTitle='{wallDocTitle}', WallDocPath='{wallDocPath}', IsLinked={isLinked}, " +
                            $"Wall.Width={wallThickness * 304.8:F1}mm ({wallThickness:F9}ft), " +
                            $"WallThickness={wallThickness * 304.8:F1}mm, " +
                            $"StructuralElementThickness={structuralElementThickness * 304.8:F1}mm\n");
                    }
                }
                else if (wall is FamilyInstance framingInstance && 
                         framingInstance.Category?.Id?.IntegerValue == (int)BuiltInCategory.OST_StructuralFraming)
                {
                    // ✅ FRAMING: Get type parameter 'b' (breadth/depth) - same logic as ClashZoneService
                    try
                    {
                        var typeId = framingInstance.GetTypeId();
                        var typeElem = wallDoc?.GetElement(typeId);
                        
                        if (typeElem != null)
                        {
                            // Try multiple parameter names for breadth/depth (same as ClashZoneService)
                            string[] possibleNames = { "b", "B", "Breadth", "Depth", "Width", "Height", "d", "D" };
                            
                            foreach (var paramName in possibleNames)
                            {
                                try
                                {
                                    var param = typeElem.LookupParameter(paramName);
                                    if (param != null && !param.IsReadOnly)
                                    {
                                        double bVal = param.AsDouble();
                                        if (bVal > 0)
                                        {
                                            framingThickness = bVal;
                                            structuralElementThickness = bVal; // Also set as general thickness
                                            
                                            if (!DeploymentConfiguration.DeploymentMode)
                                            {
                                                _logger($"[DamperProcessing] ✅ Framing Thickness: {framingThickness * 304.8:F1}mm (from type parameter '{paramName}')");
                                            }
                                            break;
                                        }
                                    }
                                }
                                catch { continue; }
                            }
                            
                            if (framingThickness <= 0 && !DeploymentConfiguration.DeploymentMode)
                            {
                                _logger($"[DamperProcessing] ⚠️ Could not find framing thickness parameter 'b' for framing {wall.Id} - depth will use fallback");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            _logger($"[DamperProcessing] ⚠️ Error getting framing thickness: {ex.Message}");
                        }
                    }
                }
                else if (wall is Floor floorElement)
                {
                    // ✅ FLOOR: Get floor thickness parameter
                    try
                    {
                        structuralElementThickness = floorElement.get_Parameter(BuiltInParameter.FLOOR_ATTR_THICKNESS_PARAM)?.AsDouble() ?? 0.0;
                        if (!DeploymentConfiguration.DeploymentMode && structuralElementThickness > 0)
                        {
                            _logger($"[DamperProcessing] ✅ Floor Thickness: {structuralElementThickness * 304.8:F1}mm");
                        }
                    }
                    catch (Exception ex)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            _logger($"[DamperProcessing] ⚠️ Error getting floor thickness: {ex.Message}");
                        }
                    }
                }
                
                // ✅ REFERENCE LEVEL: Extract MEP element's Reference Level name AND elevation
                // ✅ CRITICAL: This is used to set Schedule Level parameter on sleeve
                // ✅ CRITICAL: Level elevation is used to calculate Elevation from Level and Bottom of Opening
                // Without this, Bottom of Opening will calculate from 0 level instead of the correct Reference Level
                string mepElementLevelName = string.Empty;
                double mepElementLevelElevation = 0.0;
                
                // ✅ PRIORITY 1: Try to get from MEP element directly using HostLevelHelper (most reliable)
                try
                {
                    var refLevel = JSE_RevitAddin_MEP_OPENINGS.Helpers.HostLevelHelper.GetHostReferenceLevel(damper.Document, damper);
                    if (refLevel != null)
                    {
                        mepElementLevelName = refLevel.Name;
                        mepElementLevelElevation = refLevel.Elevation;
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            _logger($"[DamperProcessing] ✅ Extracted MEP Element Level: Name='{mepElementLevelName}', Elevation={mepElementLevelElevation * 304.8:F1}mm from HostLevelHelper for damper {damper.Id}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        _logger($"[DamperProcessing] ⚠️ Error getting Reference Level from HostLevelHelper: {ex.Message}");
                    }
                }
                
                // ✅ PRIORITY 2: If not found via HostLevelHelper, try from captured MEP parameters
                if (string.IsNullOrWhiteSpace(mepElementLevelName) && mepParameters != null && mepParameters.Count > 0)
                {
                    // Look for Reference Level or Level parameter in captured parameters
                    var refLevelParam = mepParameters.FirstOrDefault(kv => 
                        string.Equals(kv.Key, "Reference Level", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(kv.Key, "Level", StringComparison.OrdinalIgnoreCase));
                    
                    if (refLevelParam.Key != null && !string.IsNullOrWhiteSpace(refLevelParam.Value))
                    {
                        mepElementLevelName = refLevelParam.Value;
                        
                        // ✅ CRITICAL: Get elevation from level name
                        try
                        {
                            var levelByName = new FilteredElementCollector(damper.Document)
                                .OfClass(typeof(Level))
                                .Cast<Level>()
                                .FirstOrDefault(l => string.Equals(l.Name, mepElementLevelName, StringComparison.OrdinalIgnoreCase));
                            
                            if (levelByName != null)
                            {
                                mepElementLevelElevation = levelByName.Elevation;
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    _logger($"[DamperProcessing] ✅ Extracted MEP Element Level: Name='{mepElementLevelName}', Elevation={mepElementLevelElevation * 304.8:F1}mm from captured parameters for damper {damper.Id}");
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                _logger($"[DamperProcessing] ⚠️ Error getting elevation for level '{mepElementLevelName}': {ex.Message}");
                            }
                        }
                    }
                }
                
                // ✅ PRIORITY 3: If still not found, try direct parameter lookup
                if (string.IsNullOrWhiteSpace(mepElementLevelName))
                {
                    var referenceLevelWhitelist = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "Reference Level", "Level" };
                    var referenceLevelParams = _parameterSnapshotService.CaptureParams(damper, referenceLevelWhitelist);
                    var refLevelParam = referenceLevelParams.FirstOrDefault(kv => !string.IsNullOrWhiteSpace(kv.Value));
                    
                    if (refLevelParam.Key != null && !string.IsNullOrWhiteSpace(refLevelParam.Value))
                    {
                        mepElementLevelName = refLevelParam.Value;
                        
                        // ✅ CRITICAL: Get elevation from level name
                        try
                        {
                            var levelByName = new FilteredElementCollector(damper.Document)
                                .OfClass(typeof(Level))
                                .Cast<Level>()
                                .FirstOrDefault(l => string.Equals(l.Name, mepElementLevelName, StringComparison.OrdinalIgnoreCase));
                            
                            if (levelByName != null)
                            {
                                mepElementLevelElevation = levelByName.Elevation;
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    _logger($"[DamperProcessing] ✅ Extracted MEP Element Level: Name='{mepElementLevelName}', Elevation={mepElementLevelElevation * 304.8:F1}mm from direct parameter lookup for damper {damper.Id}");
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                _logger($"[DamperProcessing] ⚠️ Error getting elevation for level '{mepElementLevelName}': {ex.Message}");
                            }
                        }
                    }
                }
                
                // ✅ CRITICAL: Log warning if level name or elevation is not found (this will cause Bottom of Opening to use 0 level)
                if (string.IsNullOrWhiteSpace(mepElementLevelName) && !DeploymentConfiguration.DeploymentMode)
                {
                    _logger($"[DamperProcessing] ⚠️⚠️⚠️ WARNING: Could not extract MEP Element Level Name for damper {damper.Id} - Schedule Level may not be set correctly, Bottom of Opening will use 0 level!");
                }
                if (mepElementLevelElevation == 0.0 && !string.IsNullOrWhiteSpace(mepElementLevelName) && !DeploymentConfiguration.DeploymentMode)
                {
                    _logger($"[DamperProcessing] ⚠️⚠️⚠️ WARNING: Could not extract MEP Element Level Elevation for damper {damper.Id}, level '{mepElementLevelName}' - Elevation from Level and Bottom of Opening may use 0 elevation!");
                }
                
                // ✅ WALL CENTERLINE POINT: Calculate and save during refresh (enables multi-threaded placement)
                // ✅ CRITICAL: Pre-calculate wall centerline point when wall element is available
                // This avoids Revit API calls during placement, enabling multi-threading
                XYZ wallCenterlinePoint = intersectionPoint; // Default to intersection point if calculation fails
                try
                {
                    // Only calculate for walls and framing (floors don't need centerline adjustment)
                    if (wall is Wall || 
                        (wall is FamilyInstance framingInstance && 
                         framingInstance.Category?.Id?.IntegerValue == (int)BuiltInCategory.OST_StructuralFraming))
                    {
                        // ✅ DAMPER-SPECIFIC: Calculate wall centerline point using SIMPLE BBOX METHOD (no ray tracing)
                        // This avoids the projection/ray tracing issue that finds wall face instead of centerline
                        // Use simple bbox center method ONLY for dampers (other MEP elements use original method)
                        if (wall is Wall hostWallForCenterline)
                        {
                            wallCenterlinePoint = JSE_RevitAddin_MEP_OPENINGS.Helpers.WallCenterlineHelper.GetWallCenterlinePointFromBbox(
                                hostWallForCenterline, 
                                intersectionPoint, 
                                _document);
                        }
                        else
                        {
                            // For framing, use the element centerline method
                            wallCenterlinePoint = JSE_RevitAddin_MEP_OPENINGS.Helpers.WallCenterlineHelper.GetElementCenterlinePoint(
                                wall, 
                                intersectionPoint, 
                                _document);
                        }
                        
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            string hostType = wall is Wall ? "wall" : "framing";
                            _logger($"[DamperProcessing] ✅ Calculated Wall/Framing Centerline Point: ({wallCenterlinePoint.X:F3}, {wallCenterlinePoint.Y:F3}, {wallCenterlinePoint.Z:F3}) for damper {damper.Id}, {hostType} {wall.Id}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        _logger($"[DamperProcessing] ⚠️ Error calculating wall centerline point: {ex.Message}, using intersection point as fallback");
                    }
                    wallCenterlinePoint = intersectionPoint; // Fallback to intersection point
                }
                
                // ✅ DIAGNOSTIC: Log orientation values to verify they're being set correctly
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    _logger($"[DamperProcessing] 🔍 StructuralElementType for wall {wall.Id}: '{structuralElementType}' (Wall type: {wall.GetType().Name}, Category: {wall.Category?.Name ?? "NULL"})");
                    _logger($"[DamperProcessing] 🔍 HostOrientation: '{hostOrientation}', MepElementOrientationDirection: '{mepElementOrientationDirection}'");
                }
                
                // ✅ DAMPER-SPECIFIC: Calculate unique "Sleeve Placement Point" for deterministic GUID
                // Problem: Wall centerline point is too broad - one wall can host many dampers, causing duplicate GUIDs
                // Solution: Create a point that is:
                //   1. Unique per damper (uses damper's X/Y/Z coordinates)
                //   2. Aligned with wall centerline (stable, represents where sleeve will be placed)
                //   3. Represents actual sleeve placement location
                // This ensures each damper on the same wall gets a unique GUID
                XYZ sleevePlacementPoint = intersectionPoint; // Default fallback
                try
                {
                    if (wall is Wall) // Only for walls, not framing
                    {
                        if (hostOrientation == "X")
                        {
                            // X-wall: Wall runs along X axis, use wall centerline Y coordinate, keep damper X and Z
                            sleevePlacementPoint = new XYZ(placementPoint.X, wallCenterlinePoint.Y, placementPoint.Z);
                        }
                        else if (hostOrientation == "Y")
                        {
                            // Y-wall: Wall runs along Y axis, use wall centerline X coordinate, keep damper Y and Z
                            sleevePlacementPoint = new XYZ(wallCenterlinePoint.X, placementPoint.Y, placementPoint.Z);
                        }
                        else
                        {
                            // Unknown orientation - use full wall centerline point as fallback
                            sleevePlacementPoint = wallCenterlinePoint;
                        }
                        
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            _logger($"[DamperProcessing] ✅ Calculated Sleeve Placement Point: ({sleevePlacementPoint.X:F3}, {sleevePlacementPoint.Y:F3}, {sleevePlacementPoint.Z:F3}) for damper {damper.Id}, wall {wall.Id}, orientation '{hostOrientation}'");
                            _logger($"[DamperProcessing]   - Damper Placement Point: ({placementPoint.X:F3}, {placementPoint.Y:F3}, {placementPoint.Z:F3})");
                            _logger($"[DamperProcessing]   - Wall Centerline Point: ({wallCenterlinePoint.X:F3}, {wallCenterlinePoint.Y:F3}, {wallCenterlinePoint.Z:F3})");
                        }
                    }
                    else
                    {
                        // For framing or other host types, use damper placement point
                        sleevePlacementPoint = placementPoint;
                    }
                }
                catch (Exception ex)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        _logger($"[DamperProcessing] ⚠️ Error calculating sleeve placement point: {ex.Message}, using intersection point as fallback");
                    }
                    sleevePlacementPoint = intersectionPoint; // Fallback
                }
                
                // ✅ CRITICAL: Extract type name and family name for storage in ClashZone (avoids linked file access during placement)
                // These are used by DamperPlacementStrategy for branching logic (Standard vs non-standard with Motorized)
                // ✅ CRITICAL: Always extract these FIRST, before connector detection, so they're always available
                string damperTypeName = "";
                string damperFamilyName = "";
                
                if (damperInstance != null)
                {
                    // Extract type name and family name (ALWAYS do this, regardless of connector detection)
                    damperTypeName = damperInstance.Symbol?.Name ?? "";
                    damperFamilyName = damperInstance.Symbol?.Family?.Name ?? "";
                    
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        _logger($"[DamperProcessing] ✅ Extracted damper type/family: TypeName='{damperTypeName}', FamilyName='{damperFamilyName}' for damper {damper.Id}");
                    }
                }
                else
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        _logger($"[DamperProcessing] ⚠️ WARNING: damperInstance is null for damper {damper.Id} - cannot extract TypeName/FamilyName");
                    }
                }
                
                // ✅ SIMPLIFIED: Check if damper has MEP connectors
                // If has connectors, detect side for asymmetric clearance; if no connectors, use symmetric clearance
                bool hasMepConnector = false;
                string damperConnectorSide = string.Empty;

                if (damperInstance != null)
                {
                    try
                    {
                        // Check if damper has MEP connectors
                        string connectorSide = _connectorDetector.DetectConnectorSide(
                            damperInstance,
                            useWorldCoordinates: true, // Use world coordinates for dampers
                            out Connector connector,
                            wallOrientation: hostOrientation);

                        if (connector != null && !string.IsNullOrEmpty(connectorSide))
                        {
                            hasMepConnector = true;
                            damperConnectorSide = connectorSide;

                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                _logger($"[DamperProcessing] ✅ Damper {damper.Id} has MEP connector: HasMepConnector=true, DamperConnectorSide='{connectorSide}' - will use asymmetric clearance");
                            }
                        }
                        else
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                _logger($"[DamperProcessing] ℹ️ Damper {damper.Id} has no MEP connector - will use symmetric clearance");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            _logger($"[DamperProcessing] ⚠️ Error detecting connector for damper {damper.Id}: {ex.Message} - will use symmetric clearance");
                        }
                    }
                }
                
                // ✅ CRITICAL: Generate deterministic GUID using 3-point validation (MEP ID + Host ID + Placement Point)
                // ✅ DAMPER-SPECIFIC: For dampers, use placementPoint (damper centroid) for GUID generation
                // PlacementPoint is stable - only changes if damper moves, perfect for GUID generation
                // This matches non-damper logic which uses IntersectionPoint (also stable)
                int mepId = damper.Id.IntegerValue;
                int hostId = wall.Id.IntegerValue;
                Guid deterministicGuid;
                
                // ✅ USE PLACEMENT POINT (damper centroid) for GUID - stable, only changes if damper moves
                XYZ guidPoint = placementPoint; // Use damper centroid for deterministic GUID (stable)
                
                if (_guidManager != null && 
                    mepId > 0 && hostId > 0 &&
                    Math.Abs(guidPoint.X) > 1e-9 &&
                    Math.Abs(guidPoint.Y) > 1e-9 &&
                    Math.Abs(guidPoint.Z) > 1e-9)
                {
                    // ✅ DAMPER-SPECIFIC: Use tolerance (0.1ft = ~30mm) same as non-dampers
                    // PlacementPoint (centroid) is stable, so same tolerance as IntersectionPoint
                    double damperTolerance = 0.1; // 0.1ft = ~30mm (same as ducts/pipes)
                    
                    // ✅ 3-POINT VALIDATION: Use MEP ID + Host ID + Placement Point (centroid) to generate deterministic GUID
                    deterministicGuid = _guidManager.GetOrCreateDeterministicGuidDatabaseFirst(
                        mepId, 
                        hostId, 
                        guidPoint.X, 
                        guidPoint.Y, 
                        guidPoint.Z,
                        tolerance: damperTolerance); // Larger tolerance for dampers
                    
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        _logger($"[DamperProcessing] ✅ Generated deterministic GUID {deterministicGuid} for MEP={mepId}, Host={hostId}, PlacementPoint=({guidPoint.X:F3},{guidPoint.Y:F3},{guidPoint.Z:F3}), Tolerance={damperTolerance * 304.8:F1}mm (using damper centroid, stable)");
                    }
                }
                else
                {
                    // Fallback to random GUID if 3-point validation fails
                    deterministicGuid = Guid.NewGuid();
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        _logger($"[DamperProcessing] ⚠️ Using random GUID (3-point validation failed): MEP={mepId}, Host={hostId}, PlacementPoint=({guidPoint.X:F3},{guidPoint.Y:F3},{guidPoint.Z:F3})");
                    }
                }
                
                var clashZone = new ClashZone
                {
                    Id = deterministicGuid, // ✅ CRITICAL: Use deterministic GUID instead of random GUID
                    MepElementId = damper.Id,
                    MepElementUniqueId = damper.UniqueId,
                    StructuralElementId = wall.Id,
                    StructuralElementType = structuralElementType, // ✅ CRITICAL: Set for host type filtering
                    MepElementCategory = "Duct Accessories",
                    // ✅ CRITICAL FIX: Set damper dimensions for placement sizing
                    // These are used by ParallelSleevePlacementPlanner to calculate sleeve size
                    MepElementWidth = damperWidth,  // Already in Revit internal units (feet)
                    MepElementHeight = damperHeight, // Already in Revit internal units (feet)
                    // ✅ CRITICAL: Store type name and family name to avoid linked file access during placement
                    // These are used by DamperPlacementStrategy for branching logic (Standard vs non-standard with Motorized)
                    MepElementTypeName = damperTypeName,
                    MepElementFamilyName = damperFamilyName,
                    // ✅ X-WALL/Y-WALL ROTATION: Set orientation properties for rotation logic (same as other MEP elements)
                    HostOrientation = hostOrientation, // "X" or "Y" for walls/framing, empty for floors
                    MepElementOrientation = mepElementOrientation, // Damper flow direction vector
                    MepElementOrientationDirection = mepElementOrientationDirection, // "X" or "Y" for rotation logic
                    StructuralElementNormal = structuralElementNormal, // Wall normal for caching
                    // ✅ DEPTH PARAMETER: Set thickness properties for depth calculation (same as other MEP elements)
                    // These are used by SleeveParameterService.SetDepthParameter to set "Wall Width" or "Depth" parameter
                    WallThickness = wallThickness, // Wall.Width (for walls) - used for "Wall Width" parameter
                    FramingThickness = framingThickness, // Type parameter 'b' (for framing) - used for "Depth" parameter
                    StructuralElementThickness = structuralElementThickness, // General thickness (fallback for all host types)
                    MepElementLevelName = mepElementLevelName, // ✅ CRITICAL: Set MEP element's Reference Level name for Schedule Level mapping
                    MepElementLevelElevation = mepElementLevelElevation, // ✅ CRITICAL: Set MEP element's Reference Level elevation for Elevation from Level and Bottom of Opening calculation
                    IntersectionPoint = intersectionPoint, // ✅ DAMPER: This is the centroid of bounding box (MEP element center)
                    IntersectionPointX = intersectionPoint.X,
                    IntersectionPointY = intersectionPoint.Y,
                    IntersectionPointZ = intersectionPoint.Z,
                    WallCenterlinePoint = wallCenterlinePoint, // ✅ CRITICAL: Pre-calculated wall centerline point (enables multi-threaded placement)
                    WallCenterlinePointX = wallCenterlinePoint.X,
                    WallCenterlinePointY = wallCenterlinePoint.Y,
                    WallCenterlinePointZ = wallCenterlinePoint.Z,
                    MepParameterValues = mepParameters?.Select(kv => new SerializableKeyValue { Key = kv.Key, Value = kv.Value }).ToList() ?? new List<SerializableKeyValue>(),
                    HostParameterValues = hostParameters?.Select(kv => new SerializableKeyValue { Key = kv.Key, Value = kv.Value }).ToList() ?? new List<SerializableKeyValue>(),
                    // ✅ DAMPER-SPECIFIC: Use calculated sleeve placement point (unique per damper, aligned with wall centerline)
                    // This is the point where the sleeve will actually be placed, calculated from damper placement point and wall centerline
                    SleevePlacementPoint = sleevePlacementPoint,
                    SleevePlacementPointX = sleevePlacementPoint.X,
                    SleevePlacementPointY = sleevePlacementPoint.Y,
                    SleevePlacementPointZ = sleevePlacementPoint.Z,
                    ClashBoundingBox = intersectionBbox,
                    IsResolved = false,
                    ReadyForPlacement = true,
                    // ✅ CRITICAL: Set connector detection results for MEP connector side clearance logic
                    // These are used by DamperPlacementStrategy to determine if asymmetric clearance should be applied
                    HasMepConnector = hasMepConnector,
                    DamperConnectorSide = damperConnectorSide,
                    // ✅ CRITICAL: Set document keys for validation (IsValidClashZone requires at least one non-empty key)
                    SourceDocKey = damperDoc?.Title ?? damperDoc?.PathName ?? string.Empty,
                    HostDocKey = wallDoc?.Title ?? wallDoc?.PathName ?? string.Empty,
                    DocumentPath = hostDoc?.PathName ?? string.Empty,
                    StructuralElementDocumentTitle = wallDoc?.Title ?? string.Empty,
                    
                    // ✅ CRITICAL FIX: Set Size parameter value for database column
                    // This was missing, causing the 'Size' column in DB to be empty even if parameter was captured in JSON
                    MepElementSizeParameterValue = GetSizeParameterValue(mepParameters)
                };

                _logger($"[DamperProcessing] ✅ Created ClashZone {clashZone.Id} with {mepParameters?.Count ?? 0} MEP params and {hostParameters?.Count ?? 0} Host params, StructuralElementType='{clashZone.StructuralElementType}', MepElementWidth={clashZone.MepElementWidth * 304.8:F1}mm, MepElementHeight={clashZone.MepElementHeight * 304.8:F1}mm, TypeName='{clashZone.MepElementTypeName}', FamilyName='{clashZone.MepElementFamilyName}', HasMepConnector={clashZone.HasMepConnector}, DamperConnectorSide='{clashZone.DamperConnectorSide}'");

                // ✅ CRITICAL DIAGNOSTIC: Verify StructuralElementType is set before returning
                if (string.IsNullOrWhiteSpace(clashZone.StructuralElementType))
                {
                    _logger($"[DamperProcessing] ⚠️⚠️⚠️ WARNING: ClashZone {clashZone.Id} has EMPTY StructuralElementType! Wall={wall.Id}, WallType={wall.GetType().Name}");
                }

                return clashZone;
            }
            catch (Exception ex)
            {
                _logger($"[DamperProcessing] ⚠️ Error creating ClashZone for damper {damper.Id} and wall {wall.Id}: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// ✅ Capture MEP parameters with smart level parameter logic using ParameterSnapshotService:
        /// 1. Prefer "Reference Level"
        /// 2. If "Reference Level" not found, capture any level-related parameters
        /// 3. If no level-related parameters exist, restrict to "Reference Level" only
        /// </summary>
        private Dictionary<string, string> CaptureMepParametersWithSmartLevel(Element mepElement)
        {
            // Level-related parameter names (all variations)
            var levelParamNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "Reference Level",
                "Level",
                "Schedule Level",
                "Schedule of Level",
                "Reference Level Elevation"
            };

            // Essential MEP parameters (excluding level params)
            var essentialMepParams = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                // Size parameters (various naming conventions)
                "Size", "Diameter", "Width", "Height",
                "MEP Size", "Nominal Size", "Actual Size",
                "Duct Size", "Damper Size", "Outside Diameter",
                // System parameters
                "System Type", "System Name", "System Abbreviation",
                "Service Type", "System Classification",
                // Mark parameters
                "Mark"
            };

            // Build whitelist based on what's available
            var mepWhitelist = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            
            // First, check if "Reference Level" exists
            var referenceLevelWhitelist = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "Reference Level" };
            var referenceLevelParams = _parameterSnapshotService.CaptureParams(mepElement, referenceLevelWhitelist);
            bool hasReferenceLevel = referenceLevelParams.Any(kv => !string.IsNullOrWhiteSpace(kv.Value));

            if (hasReferenceLevel)
            {
                // Reference Level found - use it + essential params (no other level params)
                mepWhitelist.Add("Reference Level");
                foreach (var param in essentialMepParams)
                {
                    mepWhitelist.Add(param);
                }
                _logger($"[DamperProcessing] ✅ Found 'Reference Level' for MEP element {mepElement.Id} - using Reference Level + essential params");
            }
            else
            {
                // Reference Level not found - try other level params
                bool foundAnyLevelParam = false;
                foreach (var levelParamName in levelParamNames)
                {
                    if (levelParamName == "Reference Level") continue;
                    
                    var testWhitelist = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { levelParamName };
                    var testParams = _parameterSnapshotService.CaptureParams(mepElement, testWhitelist);
                    if (testParams.Any(kv => !string.IsNullOrWhiteSpace(kv.Value)))
                    {
                        mepWhitelist.Add(levelParamName);
                        foundAnyLevelParam = true;
                        _logger($"[DamperProcessing] ✅ Found '{levelParamName}' for MEP element {mepElement.Id} (Reference Level not found)");
                        break; // Use first available level param
                    }
                }

                if (!foundAnyLevelParam)
                {
                    _logger($"[DamperProcessing] ⚠️ No level-related parameters found for MEP element {mepElement.Id} - restricting to Reference Level only (empty)");
                    // Still add essential params even if no level param found
                }
                else
                {
                    // Add essential params
                    foreach (var param in essentialMepParams)
                    {
                        mepWhitelist.Add(param);
                    }
                }
            }

            // Use ParameterSnapshotService to capture parameters
            var capturedParams = _parameterSnapshotService.CaptureParams(mepElement, mepWhitelist);
            
            // ✅ DEBUG: Log what was captured
            if (!DeploymentConfiguration.DeploymentMode)
            {
                var capturedKeys = capturedParams.Where(kv => !string.IsNullOrWhiteSpace(kv.Value)).Select(kv => kv.Key).ToList();
                _logger($"[DamperProcessing] 📋 CAPTURED PARAMS for MEP {mepElement.Id}: [{string.Join(", ", capturedKeys)}]");
                
                // Check if Size is missing
                bool hasSize = capturedKeys.Any(k => k.Equals("Size", StringComparison.OrdinalIgnoreCase));
                if (!hasSize)
                {
                    _logger($"[DamperProcessing] ⚠️ SIZE NOT CAPTURED for MEP {mepElement.Id}! Whitelist contains Size: {mepWhitelist.Contains("Size")}");
                }
            }
            
            // Convert to Dictionary<string, string>
            var result = capturedParams
                .Where(kv => !string.IsNullOrWhiteSpace(kv.Value))
                .ToDictionary(kv => kv.Key, kv => kv.Value, StringComparer.OrdinalIgnoreCase);

            return result;
        }

        /// <summary>
        /// ✅ Helper to extract Size parameter value from dictionary using common keys
        /// </summary>
        private string GetSizeParameterValue(Dictionary<string, string> parameters)
        {
            if (parameters == null || parameters.Count == 0)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    _logger($"[DamperProcessing] ⚠️ GetSizeParameterValue: parameters is null or empty");
                return null;
            }

            // Prioritized list of keys to check for Size
            var sizeKeys = new[] 
            { 
                "Size", 
                "MEP Size", 
                "Diameter", 
                "Width", // Fallback for rectangular
                "Height", // Fallback for rectangular
                "Nominal Size", 
                "Duct Size", 
                "Damper Size" 
            };

            foreach (var key in sizeKeys)
            {
                if (parameters.TryGetValue(key, out string value) && !string.IsNullOrWhiteSpace(value))
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        _logger($"[DamperProcessing] ✅ GetSizeParameterValue: Found '{key}'='{value}'");
                    return value;
                }
            }
            
            // Log available keys if Size not found
            if (!DeploymentConfiguration.DeploymentMode)
            {
                var availableKeys = string.Join(", ", parameters.Keys.Take(10));
                _logger($"[DamperProcessing] ⚠️ GetSizeParameterValue: Size NOT FOUND in keys: [{availableKeys}]");
            }
            return null;
        }

        /// <summary>
        /// ✅ Capture Host parameters (restricted to Level and Fire Rating only) using ParameterSnapshotService.
        /// </summary>
        private Dictionary<string, string> CaptureHostParametersRestricted(Element hostElement)
        {
            // Only capture Level and Fire Rating
            var hostWhitelist = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "Level",
                "Fire Rating"
            };

            // Use ParameterSnapshotService to capture parameters
            var capturedParams = _parameterSnapshotService.CaptureParams(hostElement, hostWhitelist);
            
            // Convert to Dictionary<string, string> and log
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var kv in capturedParams)
            {
                if (!string.IsNullOrWhiteSpace(kv.Value))
                {
                    result[kv.Key] = kv.Value;
                    _logger($"[DamperProcessing] ✅ Captured '{kv.Key}'='{kv.Value}' for Host element {hostElement.Id}");
                }
            }

            return result;
        }

        /// <summary>
        /// ✅ Get StructuralElementType from element (matches ClashZoneService_Legacy logic)
        /// </summary>
        private string GetStructuralElementType(Element element)
        {
            // ✅ DIAGNOSTIC: Log element type for debugging
            if (!DeploymentConfiguration.DeploymentMode)
            {
                _logger($"[DamperProcessing] 🔍 GetStructuralElementType: Element={element.Id}, Type={element.GetType().Name}, Category={element.Category?.Name ?? "NULL"}, CategoryId={element.Category?.Id?.IntegerValue ?? -1}");
            }
            
            if (element is Wall)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    _logger($"[DamperProcessing] ✅ Element {element.Id} is Wall - returning 'Wall'");
                }
                return "Wall";
            }
            else if (element is Floor)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    _logger($"[DamperProcessing] ✅ Element {element.Id} is Floor - returning 'Floor'");
                }
                return "Floor";
            }
            else if (element is FamilyInstance famInst && 
                     famInst.Category?.Id?.IntegerValue == (int)BuiltInCategory.OST_StructuralFraming)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    _logger($"[DamperProcessing] ✅ Element {element.Id} is Structural Framing - returning 'Structural Framing'");
                }
                return "Structural Framing";
            }
            else
            {
                // ✅ CRITICAL: Check if it's a wall by category even if not Wall type (linked documents)
                if (element.Category?.Id?.IntegerValue == (int)BuiltInCategory.OST_Walls)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        _logger($"[DamperProcessing] ✅ Element {element.Id} has Wall category (but not Wall type) - returning 'Wall'");
                    }
                    return "Wall";
                }
                
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    _logger($"[DamperProcessing] ⚠️ Element {element.Id} is Unknown type - returning 'Unknown'");
                }
                return "Unknown";
            }
        }
    }
}

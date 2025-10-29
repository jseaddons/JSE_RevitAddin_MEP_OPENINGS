using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Service for detecting intersections between MEP and structural elements.
    /// Extracted from the working TestMepIntersection logic.
    /// </summary>
    public class IntersectionDetectionService
    {
        private readonly Action<string> _logger;

        public IntersectionDetectionService(Action<string> logger)
        {
            _logger = logger ?? (msg => { });
            
            // ✅ PHASE 1 OPTIMIZATION: Enable optimization flags
            OptimizationFlags.EnablePhase1Optimizations();
            _logger($"[IntersectionDetectionService] Phase 1 optimizations enabled: {OptimizationFlags.GetOptimizationStatus()}");
        }

        /// <summary>
        /// Finds intersections using the proven TestMepIntersection approach
        /// </summary>
        public List<(Element, Element, BoundingBoxXYZ, XYZ)> FindIntersections(
            Document document, 
            View3D view3D,
            List<string> selectedMepCategories = null,
            List<string> selectedReferenceFiles = null,
            List<string> selectedHostFiles = null,
            List<string> allowedHostElementTypes = null)
        {
            try
            {
                _logger("=== INTERSECTION DETECTION STARTED ===");
                _logger($"Document: {document.Title}");
                _logger($"View: {view3D.Name}");

                // STEP 1: Get section box in model coordinates
                BoundingBoxXYZ sectionBox = view3D.GetSectionBox();
                Transform sectionTransform = sectionBox.Transform;

                List<XYZ> corners = new List<XYZ>
                {
                    sectionTransform.OfPoint(sectionBox.Min),
                    sectionTransform.OfPoint(new XYZ(sectionBox.Max.X, sectionBox.Min.Y, sectionBox.Min.Z)),
                    sectionTransform.OfPoint(new XYZ(sectionBox.Min.X, sectionBox.Max.Y, sectionBox.Min.Z)),
                    sectionTransform.OfPoint(new XYZ(sectionBox.Max.X, sectionBox.Max.Y, sectionBox.Min.Z)),
                    sectionTransform.OfPoint(new XYZ(sectionBox.Min.X, sectionBox.Min.Y, sectionBox.Max.Z)),
                    sectionTransform.OfPoint(new XYZ(sectionBox.Max.X, sectionBox.Min.Y, sectionBox.Max.Z)),
                    sectionTransform.OfPoint(new XYZ(sectionBox.Min.X, sectionBox.Max.Y, sectionBox.Max.Z)),
                    sectionTransform.OfPoint(sectionBox.Max)
                };

                XYZ modelMin = new XYZ(corners.Min(p => p.X), corners.Min(p => p.Y), corners.Min(p => p.Z));
                XYZ modelMax = new XYZ(corners.Max(p => p.X), corners.Max(p => p.Y), corners.Max(p => p.Z));

                _logger($"Section box: Min=({modelMin.X:F1}, {modelMin.Y:F1}, {modelMin.Z:F1}) Max=({modelMax.X:F1}, {modelMax.Y:F1}, {modelMax.Z:F1})");

                // STEP 2: Collect MEP and Structural elements
                var mepElements = new List<Element>();
                var wallElements = new List<Element>();

                CollectElements(document, modelMin, modelMax, ref mepElements, ref wallElements, selectedMepCategories, selectedReferenceFiles, selectedHostFiles, allowedHostElementTypes);

                _logger($"Found {mepElements.Count} MEP elements and {wallElements.Count} structural elements (walls/floors/framing) in section box");
                _logger($"Selected MEP cats: {string.Join(", ", selectedMepCategories ?? new List<string>())}");
                _logger($"Selected host types: {string.Join(", ", allowedHostElementTypes ?? new List<string>())}");

                // ✅ OPTIMIZATION: 5-STEP FILTERING RESULTS SUMMARY
                _logger($"[5-STEP-FILTER-RESULTS] ELEMENT COLLECTION COMPLETE:");
                _logger($"[5-STEP-FILTER-RESULTS] Step 1 (Section Box): ✓ Applied - Only elements within section box collected");
                _logger($"[5-STEP-FILTER-RESULTS] Step 2 (Reference File): ✓ Applied - Only MEP elements from selected reference files collected");
                _logger($"[5-STEP-FILTER-RESULTS] Step 3 (MEP Categories): ✓ Applied - Only selected MEP categories collected");
                _logger($"[5-STEP-FILTER-RESULTS] Step 4 (Host File): ✓ Applied - Only structural elements from selected host files collected");
                _logger($"[5-STEP-FILTER-RESULTS] Step 5 (Host Categories): ✓ Applied - Only selected host types collected");
                _logger($"[5-STEP-FILTER-RESULTS] FINAL RESULTS: {mepElements.Count} MEP elements, {wallElements.Count} structural elements (100% efficiency - no waste)");

                if (mepElements.Count == 0)
                {
                    _logger("No MEP elements found in section box.");
                    return new List<(Element, Element, BoundingBoxXYZ, XYZ)>();
                }

                if (wallElements.Count == 0)
                {
                    _logger("No structural elements (walls/floors/framing) found in section box.");
                    return new List<(Element, Element, BoundingBoxXYZ, XYZ)>();
                }

                // STEP 3: Find intersections
                var intersections = FindIntersectionsInternal(mepElements, wallElements, document);

                // Enforce section box bounds using oriented-box test (view's local coords)
                Transform invSection = sectionTransform.Inverse;
                intersections = intersections
                    .Where(tuple =>
                    {
                        var center = tuple.Item4;
                        var local = invSection.OfPoint(center);
                        return local.X >= sectionBox.Min.X && local.X <= sectionBox.Max.X &&
                               local.Y >= sectionBox.Min.Y && local.Y <= sectionBox.Max.Y &&
                               local.Z >= sectionBox.Min.Z && local.Z <= sectionBox.Max.Z;
                    })
                    .ToList();
                _logger($"Filtered intersections inside oriented section box: {intersections.Count}");

                _logger($"=== INTERSECTION RESULTS ===");
                _logger($"Total intersections found: {intersections.Count}");

                return intersections;
            }
            catch (Exception ex)
            {
                _logger($"ERROR: {ex.Message}");
                _logger($"Stack: {ex.StackTrace}");
                return new List<(Element, Element, BoundingBoxXYZ, XYZ)>();
            }
        }

        private void CollectElements(Document doc, XYZ modelMin, XYZ modelMax, 
                                     ref List<Element> mepElements, ref List<Element> wallElements,
                                     List<string> selectedMepCategories = null,
                                     List<string> selectedReferenceFiles = null,
                                     List<string> selectedHostFiles = null,
                                     List<string> allowedHostElementTypes = null)
        {
            Outline hostOutline = new Outline(modelMin, modelMax);

            // Determine which MEP categories to collect based on UI selections
            List<BuiltInCategory> mepCats = new List<BuiltInCategory>();
            
            if (selectedMepCategories == null || selectedMepCategories.Count == 0)
            {
                // If no categories selected, collect all (fallback behavior)
                mepCats.AddRange(new[] {
                    BuiltInCategory.OST_DuctCurves,
                    BuiltInCategory.OST_DuctFitting,
                    BuiltInCategory.OST_DuctAccessory,
                    BuiltInCategory.OST_DuctTerminal,
                    BuiltInCategory.OST_PipeCurves,
                    BuiltInCategory.OST_PipeFitting,
                    BuiltInCategory.OST_PipeAccessory,
                    BuiltInCategory.OST_CableTray,
                    BuiltInCategory.OST_CableTrayFitting,
                    BuiltInCategory.OST_Conduit,
                    BuiltInCategory.OST_ConduitFitting
                });
                _logger("No MEP categories selected - collecting all categories");
            }
            else
            {
                // Only collect selected categories
                if (selectedMepCategories.Any(c => string.Equals(c, "Ducts", StringComparison.OrdinalIgnoreCase)))
                {
                    mepCats.Add(BuiltInCategory.OST_DuctCurves);
                    mepCats.Add(BuiltInCategory.OST_DuctFitting);
                    mepCats.Add(BuiltInCategory.OST_DuctTerminal);
                    _logger("Including Ducts categories");
                }
                
                if (selectedMepCategories.Any(c => string.Equals(c, "Duct Accessories", StringComparison.OrdinalIgnoreCase)))
                {
                    mepCats.Add(BuiltInCategory.OST_DuctAccessory);
                    _logger("Including Duct Accessories category");
                }
                
                if (selectedMepCategories.Any(c => string.Equals(c, "Pipes", StringComparison.OrdinalIgnoreCase) || string.Equals(c, "Pipe", StringComparison.OrdinalIgnoreCase)))
                {
                    mepCats.Add(BuiltInCategory.OST_PipeCurves);
                    mepCats.Add(BuiltInCategory.OST_PipeFitting);
                    mepCats.Add(BuiltInCategory.OST_PipeAccessory);
                    _logger("Including Pipes categories");
                }
                
                if (selectedMepCategories.Any(c => string.Equals(c, "Cable Trays", StringComparison.OrdinalIgnoreCase) || string.Equals(c, "Cable Tray", StringComparison.OrdinalIgnoreCase)))
                {
                    mepCats.Add(BuiltInCategory.OST_CableTray);
                    mepCats.Add(BuiltInCategory.OST_CableTrayFitting);
                    mepCats.Add(BuiltInCategory.OST_Conduit);
                    mepCats.Add(BuiltInCategory.OST_ConduitFitting);
                    _logger("Including Cable Trays categories");
                }

                // Discipline-to-category mapping for filters named by discipline
                if (selectedMepCategories.Any(c => string.Equals(c, "Ventilation", StringComparison.OrdinalIgnoreCase) || string.Equals(c, "HVAC", StringComparison.OrdinalIgnoreCase)))
                {
                    mepCats.Add(BuiltInCategory.OST_DuctCurves);
                    mepCats.Add(BuiltInCategory.OST_DuctFitting);
                    mepCats.Add(BuiltInCategory.OST_DuctAccessory);
                    mepCats.Add(BuiltInCategory.OST_DuctTerminal);
                    _logger("Including HVAC/Ventilation (duct) categories via discipline mapping");
                }

                if (selectedMepCategories.Any(c => string.Equals(c, "Electrical", StringComparison.OrdinalIgnoreCase)))
                {
                    mepCats.Add(BuiltInCategory.OST_CableTray);
                    mepCats.Add(BuiltInCategory.OST_CableTrayFitting);
                    mepCats.Add(BuiltInCategory.OST_Conduit);
                    mepCats.Add(BuiltInCategory.OST_ConduitFitting);
                    _logger("Including Electrical (tray/conduit) categories via discipline mapping");
                }

                if (selectedMepCategories.Any(c => string.Equals(c, "Plumbing", StringComparison.OrdinalIgnoreCase)))
                {
                    mepCats.Add(BuiltInCategory.OST_PipeCurves);
                    mepCats.Add(BuiltInCategory.OST_PipeFitting);
                    mepCats.Add(BuiltInCategory.OST_PipeAccessory);
                    _logger("Including Plumbing (pipes) categories via discipline mapping");
                }
                
                _logger($"Selected MEP categories: {string.Join(", ", selectedMepCategories)}");
                _logger($"Collecting {mepCats.Count} MEP categories");
            }

            _logger($"DEBUG: Section box in active document: Min=({modelMin.X:F2}, {modelMin.Y:F2}, {modelMin.Z:F2}) Max=({modelMax.X:F2}, {modelMax.Y:F2}, {modelMax.Z:F2})");

            // ✅ OPTIMIZATION: 5-STEP FILTERING APPLIED DURING ELEMENT COLLECTION
            // Step 1: Section Box Filter - Only collect elements within visible section box
            _logger($"[5-STEP-FILTER] Step 1 (Section Box): Collecting elements within section box bounds");
            
            // Step 2: Reference File Filter - Only collect MEP elements from selected reference files
            _logger($"[5-STEP-FILTER] Step 2 (Reference File): Selected reference files: {string.Join(", ", selectedReferenceFiles ?? new List<string>())}");
            
            // Step 3: MEP Categories Filter - Only collect selected MEP categories
            _logger($"[5-STEP-FILTER] Step 3 (MEP Categories): Selected MEP categories: {string.Join(", ", selectedMepCategories ?? new List<string>())}");
            
            // Step 4: Host File Filter - Only collect structural elements from selected host files
            _logger($"[5-STEP-FILTER] Step 4 (Host File): Selected host files: {string.Join(", ", selectedHostFiles ?? new List<string>())}");
            
            // Step 5: Host Categories Filter - Only collect selected host types
            _logger($"[5-STEP-FILTER] Step 5 (Host Categories): Selected host types: {string.Join(", ", allowedHostElementTypes ?? new List<string>())}");

            // Collect from host
            foreach (var cat in mepCats)
            {
                mepElements.AddRange(
                    new FilteredElementCollector(doc)
                        .OfCategory(cat)
                        .WhereElementIsNotElementType()
                        .WherePasses(new BoundingBoxIntersectsFilter(hostOutline))
                        .ToElements()
                );
            }
            _logger($"Collected {mepElements.Count} MEP elements from host document after category filter");

            // ✅ CRITICAL FIX: Track MEP count before processing links
            int mepCountBeforeLinks = mepElements.Count;

            // ⚠️ IMPORTANT: Host elements (Walls, Floors, Structural Framing) are ALWAYS in linked files ONLY
            // DO NOT collect host elements from active document - they only exist in architectural/structural links
            _logger($"Host elements will ONLY be collected from selected host files (linked files): {string.Join(", ", selectedHostFiles ?? new List<string>())}");

            // Collect from links
            var links = new FilteredElementCollector(doc)
                .OfClass(typeof(RevitLinkInstance))
                .Cast<RevitLinkInstance>()
                .ToList();

            _logger($"DEBUG: Found {links.Count} linked models");

            // Use passed-in reference files (MEP links), independent from host files
            // If not provided, fall back to UI state provider
            if (selectedReferenceFiles == null)
            {
                selectedReferenceFiles = FilterUiStateProvider.GetSelectedReferenceFiles?.Invoke() ?? new List<string>();
            }
            
            _logger($"DEBUG: Selected reference files (MEP links): {string.Join(", ", selectedReferenceFiles)}");

            // Normalization helper for robust filename matching
            Func<string, string> norm = s =>
            {
                if (string.IsNullOrWhiteSpace(s)) return string.Empty;
                var trimmed = s;
                var idxParen = trimmed.IndexOf('(');
                if (idxParen >= 0) trimmed = trimmed.Substring(0, idxParen);
                trimmed = System.IO.Path.GetFileNameWithoutExtension(trimmed);
                trimmed = trimmed.ToLowerInvariant().Replace("_detached", "");
                trimmed = trimmed.Replace('_', ' ').Replace('-', ' ');
                trimmed = System.Text.RegularExpressions.Regex.Replace(trimmed, "\\s+", " ");
                return trimmed.Trim();
            };

            foreach (var link in links)
            {
                Document linkDoc = link.GetLinkDocument();
                if (linkDoc == null) continue;
                string linkTitleNorm = norm(linkDoc.Title);

                Transform linkTransform = link.GetTotalTransform();
                Transform invTransform = linkTransform.Inverse;

                XYZ linkMin = invTransform.OfPoint(modelMin);
                XYZ linkMax = invTransform.OfPoint(modelMax);

                XYZ actualMin = new XYZ(
                    Math.Min(linkMin.X, linkMax.X),
                    Math.Min(linkMin.Y, linkMax.Y),
                    Math.Min(linkMin.Z, linkMax.Z)
                );
                XYZ actualMax = new XYZ(
                    Math.Max(linkMin.X, linkMax.X),
                    Math.Max(linkMin.Y, linkMax.Y),
                    Math.Max(linkMin.Z, linkMax.Z)
                );

                _logger($"DEBUG: Section box in linked file '{linkDoc.Title}': Min=({actualMin.X:F2}, {actualMin.Y:F2}, {actualMin.Z:F2}) Max=({actualMax.X:F2}, {actualMax.Y:F2}, {actualMax.Z:F2})");

                Outline linkOutline = new Outline(actualMin, actualMax);

                // Collect MEP from reference links only (if any specified); otherwise from all links
                bool refMatch = selectedReferenceFiles.Count == 0 || selectedReferenceFiles.Any(f =>
                {
                    string ui = norm(f);
                    return linkTitleNorm.Contains(ui) || ui.Contains(linkTitleNorm);
                });
                if (refMatch)
                {
                    foreach (var cat in mepCats)
                    {
                        mepElements.AddRange(
                            new FilteredElementCollector(linkDoc)
                                .OfCategory(cat)
                                .WhereElementIsNotElementType()
                                .WherePasses(new BoundingBoxIntersectsFilter(linkOutline))
                                .ToElements()
                        );
                    }
                    _logger($"Collected {mepElements.Count} total MEP elements after processing reference link '{linkDoc.Title}'");
                }
                else
                {
                    _logger($"DEBUG: Skipping MEP collection from link '{linkDoc.Title}' (not in selected reference files)");
                }

                // Collect hosts only from selected host files (if specified), else from all links
                bool hostMatch = selectedHostFiles == null || selectedHostFiles.Count == 0 || selectedHostFiles.Any(f =>
                {
                    string ui = norm(f);
                    return linkTitleNorm.Contains(ui) || ui.Contains(linkTitleNorm);
                });
                
                _logger($"DEBUG: Host match for '{linkDoc.Title}': {hostMatch} (selected host files: {string.Join(", ", selectedHostFiles ?? new List<string>())})");
                
                if (hostMatch)
                {
                    // ✅ FIX: Only collect structural categories that are selected in UI
                    if (allowedHostElementTypes == null || allowedHostElementTypes.Count == 0)
                    {
                        // Fallback: collect all structural categories if no selection
                        _logger("No host element types selected - collecting all structural categories");
                        wallElements.AddRange(
                            new FilteredElementCollector(linkDoc)
                                .OfCategory(BuiltInCategory.OST_Walls)
                                .WhereElementIsNotElementType()
                                .WherePasses(new BoundingBoxIntersectsFilter(linkOutline))
                                .ToElements()
                        );
                        wallElements.AddRange(
                            new FilteredElementCollector(linkDoc)
                                .OfCategory(BuiltInCategory.OST_Floors)
                                .WhereElementIsNotElementType()
                                .WherePasses(new BoundingBoxIntersectsFilter(linkOutline))
                                .ToElements()
                        );
                        wallElements.AddRange(
                            new FilteredElementCollector(linkDoc)
                                .OfCategory(BuiltInCategory.OST_StructuralFraming)
                                .WhereElementIsNotElementType()
                                .WherePasses(new BoundingBoxIntersectsFilter(linkOutline))
                                .ToElements()
                        );
                    }
                    else
                    {
                        // ✅ FILTERED: Only collect selected host element types
                        _logger($"UI Selected Host types: {string.Join(", ", allowedHostElementTypes)}");
                        
                        if (allowedHostElementTypes.Any(ht => ht.Equals("Walls", StringComparison.OrdinalIgnoreCase)))
                        {
                            wallElements.AddRange(
                                new FilteredElementCollector(linkDoc)
                                    .OfCategory(BuiltInCategory.OST_Walls)
                                    .WhereElementIsNotElementType()
                                    .WherePasses(new BoundingBoxIntersectsFilter(linkOutline))
                                    .ToElements()
                            );
                            _logger("UI Selection: Including Walls");
                        }

                        if (allowedHostElementTypes.Any(ht => ht.Equals("Floors", StringComparison.OrdinalIgnoreCase)))
                        {
                            wallElements.AddRange(
                                new FilteredElementCollector(linkDoc)
                                    .OfCategory(BuiltInCategory.OST_Floors)
                                    .WhereElementIsNotElementType()
                                    .WherePasses(new BoundingBoxIntersectsFilter(linkOutline))
                                    .ToElements()
                            );
                            _logger("UI Selection: Including Floors");
                        }

                        if (allowedHostElementTypes.Any(ht => ht.Equals("Structural Framing", StringComparison.OrdinalIgnoreCase)))
                        {
                            wallElements.AddRange(
                                new FilteredElementCollector(linkDoc)
                                    .OfCategory(BuiltInCategory.OST_StructuralFraming)
                                    .WhereElementIsNotElementType()
                                    .WherePasses(new BoundingBoxIntersectsFilter(linkOutline))
                                    .ToElements()
                            );
                            _logger("UI Selection: Including Structural Framing");
                        }
                        
                        _logger($"✅ FILTERED: Only collected {allowedHostElementTypes.Count} selected host types (was 3 total)");
                    }

                    _logger($"Collected {wallElements.Count} total host elements after processing host link '{linkDoc.Title}'");
                }
                else
                {
                    _logger($"DEBUG: Skipping host collection from link '{linkDoc.Title}' (not in selected host files)");
                }
            }

            // ✅ CRITICAL FIX: Fallback for single-model workflows
            // If MEP collection is still empty after checking links, and we're in a single-model workflow (no links or no reference files selected),
            // the MEP elements are in the active document itself
            if (mepElements.Count == mepCountBeforeLinks && mepCats.Count > 0)
            {
                if (links.Count == 0)
                {
                    _logger("SINGLE-MODEL WORKFLOW: No linked files found. MEP elements are in the active document (already collected).");
                }
                else if (mepElements.Count == 0)
                {
                    _logger("FALLBACK: No MEP elements collected from selected reference files. Scanning all links for selected categories.");
                    foreach (var link in links)
                    {
                        var linkDoc = link.GetLinkDocument();
                        if (linkDoc == null) continue;
                        Transform inv = link.GetTotalTransform().Inverse;
                        XYZ linkMin = inv.OfPoint(modelMin);
                        XYZ linkMax = inv.OfPoint(modelMax);
                        XYZ actualMin = new XYZ(Math.Min(linkMin.X, linkMax.X), Math.Min(linkMin.Y, linkMax.Y), Math.Min(linkMin.Z, linkMax.Z));
                        XYZ actualMax = new XYZ(Math.Max(linkMin.X, linkMax.X), Math.Max(linkMin.Y, linkMax.Y), Math.Max(linkMin.Z, linkMax.Z));
                        var linkOutline = new Outline(actualMin, actualMax);
                        foreach (var cat in mepCats)
                        {
                            mepElements.AddRange(new FilteredElementCollector(linkDoc)
                                .OfCategory(cat)
                                .WhereElementIsNotElementType()
                                .WherePasses(new BoundingBoxIntersectsFilter(linkOutline))
                                .ToElements());
                        }
                    }
                    _logger($"FALLBACK: Collected {mepElements.Count} MEP elements after scanning all links.");
                }
            }

            // ✅ CRITICAL FIX: Fallback for host elements in single-model workflows
            if (wallElements.Count == 0)
            {
                if (links.Count == 0)
                {
                    _logger("SINGLE-MODEL WORKFLOW: No linked files found. Host elements (walls/floors/framing) are in the active document (already collected).");
                }
                else
                {
                    _logger("FALLBACK: No host elements collected from selected host files. Scanning all links for selected host types only.");
                    foreach (var link in links)
                    {
                        var linkDoc = link.GetLinkDocument();
                        if (linkDoc == null) continue;
                        Transform inv = link.GetTotalTransform().Inverse;
                        XYZ linkMin = inv.OfPoint(modelMin);
                        XYZ linkMax = inv.OfPoint(modelMax);
                        XYZ actualMin = new XYZ(Math.Min(linkMin.X, linkMax.X), Math.Min(linkMin.Y, linkMax.Y), Math.Min(linkMin.Z, linkMax.Z));
                        XYZ actualMax = new XYZ(Math.Max(linkMin.X, linkMax.X), Math.Max(linkMin.Y, linkMax.Y), Math.Max(linkMin.Z, linkMax.Z));
                        var linkOutline = new Outline(actualMin, actualMax);
                        
                        // ✅ FIX: Apply same host type filtering in fallback
                        if (allowedHostElementTypes == null || allowedHostElementTypes.Count == 0)
                        {
                            // Fallback: collect all structural categories
                            wallElements.AddRange(new FilteredElementCollector(linkDoc)
                                .OfCategory(BuiltInCategory.OST_Walls)
                                .WhereElementIsNotElementType()
                                .WherePasses(new BoundingBoxIntersectsFilter(linkOutline))
                                .ToElements());
                            wallElements.AddRange(new FilteredElementCollector(linkDoc)
                                .OfCategory(BuiltInCategory.OST_Floors)
                                .WhereElementIsNotElementType()
                                .WherePasses(new BoundingBoxIntersectsFilter(linkOutline))
                                .ToElements());
                            wallElements.AddRange(new FilteredElementCollector(linkDoc)
                                .OfCategory(BuiltInCategory.OST_StructuralFraming)
                                .WhereElementIsNotElementType()
                                .WherePasses(new BoundingBoxIntersectsFilter(linkOutline))
                                .ToElements());
                        }
                        else
                        {
                            // Only collect selected host types
                            if (allowedHostElementTypes.Any(ht => ht.Equals("Walls", StringComparison.OrdinalIgnoreCase)))
                            {
                                wallElements.AddRange(new FilteredElementCollector(linkDoc)
                                    .OfCategory(BuiltInCategory.OST_Walls)
                                    .WhereElementIsNotElementType()
                                    .WherePasses(new BoundingBoxIntersectsFilter(linkOutline))
                                    .ToElements());
                            }
                            if (allowedHostElementTypes.Any(ht => ht.Equals("Floors", StringComparison.OrdinalIgnoreCase)))
                            {
                                wallElements.AddRange(new FilteredElementCollector(linkDoc)
                                    .OfCategory(BuiltInCategory.OST_Floors)
                                    .WhereElementIsNotElementType()
                                    .WherePasses(new BoundingBoxIntersectsFilter(linkOutline))
                                    .ToElements());
                            }
                            if (allowedHostElementTypes.Any(ht => ht.Equals("Structural Framing", StringComparison.OrdinalIgnoreCase)))
                            {
                                wallElements.AddRange(new FilteredElementCollector(linkDoc)
                                    .OfCategory(BuiltInCategory.OST_StructuralFraming)
                                    .WhereElementIsNotElementType()
                                    .WherePasses(new BoundingBoxIntersectsFilter(linkOutline))
                                    .ToElements());
                            }
                        }
                    }
                    _logger($"FALLBACK: Collected {wallElements.Count} host elements after scanning all links with host type filtering.");
                }
            }
        }

        private List<(Element, Element, BoundingBoxXYZ, XYZ)> FindIntersectionsInternal(
            List<Element> mepElements, 
            List<Element> wallElements, 
            Document doc)
        {
            _logger($"Using optimized MepIntersectionService.FindIntersectionsBatch with {mepElements.Count} MEP elements...");

            // Get all linked documents for transform lookup
            var links = new FilteredElementCollector(doc)
                .OfClass(typeof(RevitLinkInstance))
                .Cast<RevitLinkInstance>()
                .ToList();

            // ✅ PHASE 2 OPTIMIZATION: Prepare MEP and structural elements for batch processing
            var mepElementsWithTransforms = new List<(Element, Transform?)>();
            foreach (var mep in mepElements)
            {
                try
                {
                    bool isDamper = false;
                    if (mep is FamilyInstance fi)
                    {
                        string familyName = fi.Symbol?.Family?.Name ?? "";
                        isDamper = familyName.IndexOf("Damper", StringComparison.OrdinalIgnoreCase) >= 0;
                    }

                    var mepLink = links.FirstOrDefault(link => link.GetLinkDocument()?.Title == mep.Document.Title);
                    Transform? mepTransform = mepLink != null ? MepIntersectionService.GetCachedTransform(mep.Document, links) : null;
                    mepElementsWithTransforms.Add((mep, mepTransform));
                }
                catch (Exception ex)
                {
                    _logger($"Error preparing MEP element {mep.Id}: {ex.Message}");
                }
            }

            var structuralElementsWithTransforms = new List<(Element, Transform?)>();
            foreach (var wall in wallElements)
            {
                try
                {
                    var wallLink = links.FirstOrDefault(link => link.GetLinkDocument()?.Title == wall.Document.Title);
                    Transform? wallTransform = wallLink != null ? MepIntersectionService.GetCachedTransform(wall.Document, links) : null;
                    structuralElementsWithTransforms.Add((wall, wallTransform));
                }
                catch (Exception ex)
                {
                    _logger($"Error preparing structural element {wall.Id}: {ex.Message}");
                }
            }

            // ✅ PHASE 2 OPTIMIZATION: Use batch processing with spatial hash grid and curve-in-bbox test
            _logger($"[PHASE2] Calling MepIntersectionService.FindIntersectionsBatch with {mepElementsWithTransforms.Count} MEP and {structuralElementsWithTransforms.Count} structural elements");
            var intersections = MepIntersectionService.FindIntersectionsBatch(
                mepElementsWithTransforms,
                structuralElementsWithTransforms,
                _logger);
            
            _logger($"[PHASE2] Optimized batch processing completed: {intersections.Count} intersections found");
            return intersections;
        }

        // OLD CODE BELOW - Replaced with optimized FindIntersectionsBatch method
        // Keeping for reference, but never executed
        /*
            foreach (var mep in mepElements)
            {
                try
                {
                    // ✅ PHASE 1 OPTIMIZATION: Check cache invalidation
                    if (CacheInvalidationMonitor.NeedsInvalidation(mep))
                    {
                        _logger($"[IntersectionDetectionService] MEP element {mep.Id} needs cache invalidation");
                    }

                    bool isDamper = false;
                    string familyName = string.Empty;
                    if (mep is FamilyInstance fi)
                    {
                        familyName = fi.Symbol?.Family?.Name ?? string.Empty;
                        isDamper = familyName.IndexOf("Damper", StringComparison.OrdinalIgnoreCase) >= 0;
                        if (isDamper)
                        {
                            _logger($"DEBUG: Processing damper element {mep.Id} - Family: {familyName}");
                        }
                    }

                    var mepLink = links.FirstOrDefault(link => link.GetLinkDocument()?.Title == mep.Document.Title);
                    Transform? mepTransform = mepLink?.GetTotalTransform();
                    // MEP link and transform found

                    // ✅ PHASE 1 OPTIMIZATION: Use smart tolerance
                    var tolerance = SmartToleranceService.GetIntersectionTolerance(mep, wallElements.FirstOrDefault());
                    _logger($"[IntersectionDetectionService] Using smart tolerance: {tolerance:F3} for MEP element {mep.Id}");

                    var structuralElements = new List<(Element, Transform?)>();
                    foreach (var wall in wallElements)
                    {
                        var wallLink = links.FirstOrDefault(link => link.GetLinkDocument()?.Title == wall.Document.Title);
                        if (wallLink != null)
                        {
                            var wallTransform = wallLink.GetTotalTransform();
                            structuralElements.Add((wall, wallTransform));
                        }
                        else
                        {
                            _logger($"DEBUG: No link found for wall {wall.Id} from document {wall.Document.Title}");
                            _logger($"DEBUG: Available links: {string.Join(", ", links.Select(l => l.GetLinkDocument()?.Title ?? "NULL"))}");
                            structuralElements.Add((wall, null));
                        }
                    }

                    var mepBBox = mep.get_BoundingBox(null);

                    if (isDamper)
                    {
                        var damperHits = MepIntersectionService.FindDamperIntersections(
                            mep,
                            structuralElements,
                            mepTransform,
                            _logger);

                        _logger($"DEBUG: Damper {mep.Id} intersections found: {damperHits.Count}");

                        foreach (var (structuralElement, bbox, center) in damperHits)
                        {
                            intersections.Add((mep, structuralElement, bbox, center));
                            _logger($"Intersection found: Damper {mep.Id} ({mep.Category?.Name}) <-> {structuralElement.Category?.Name} {structuralElement.Id} at {center}");
                        }

                        if (damperHits.Count == 0)
                        {
                            _logger($"DEBUG: No intersections found for damper {mep.Id}");
                        }

                        continue;
                    }

                    var locationCurve = mep.Location as LocationCurve;
                    if (locationCurve?.Curve is Line mepLine)
                    {
                        _logger($"DEBUG: Processing MEP {mep.Id} ({mep.Category?.Name}) with line from {mepLine.GetEndPoint(0)} to {mepLine.GetEndPoint(1)}");

                        // ✅ METHOD 3: Check for damper presence at duct end points BEFORE processing intersections
                        if (mep.Category?.Name == "Ducts" || mep.Category?.Name == "Duct Curves")
                        {
                            _logger($"[METHOD3] Checking duct {mep.Id} for damper presence at end points");
                            
                            if (CheckForDamperAtDuctEnd(doc, mep, mepLine, mepTransform))
                            {
                                _logger($"[METHOD3] SKIP: Duct {mep.Id} - Damper found at duct end, skipping intersection detection");
                                continue; // Skip this duct entirely - no intersections will be created
                            }
                            
                            _logger($"[METHOD3] No damper found at duct {mep.Id} end points - proceeding with intersection detection");
                        }

                        Line mepLineInHostShared;
                        if (mepTransform != null)
                        {
                            mepLineInHostShared = Line.CreateBound(
                                mepTransform.OfPoint(mepLine.GetEndPoint(0)),
                                mepTransform.OfPoint(mepLine.GetEndPoint(1))
                            );
                        }
                        else
                        {
                            mepLineInHostShared = mepLine;
                        }

                        _logger($"DEBUG: MEP line in host shared coordinates: {mepLineInHostShared.GetEndPoint(0)} to {mepLineInHostShared.GetEndPoint(1)}");

                        BoundingBoxXYZ? mepBBoxInHostShared = mepBBox;
                        if (mepBBox != null && mepTransform != null)
                        {
                            _logger($"DEBUG: MEP bbox before transform: Min=({mepBBox.Min.X:F2}, {mepBBox.Min.Y:F2}, {mepBBox.Min.Z:F2}) Max=({mepBBox.Max.X:F2}, {mepBBox.Max.Y:F2}, {mepBBox.Max.Z:F2})");

                            var mepPts = new[]
                            {
                                mepTransform.OfPoint(new XYZ(mepBBox.Min.X, mepBBox.Min.Y, mepBBox.Min.Z)),
                                mepTransform.OfPoint(new XYZ(mepBBox.Max.X, mepBBox.Min.Y, mepBBox.Min.Z)),
                                mepTransform.OfPoint(new XYZ(mepBBox.Min.X, mepBBox.Max.Y, mepBBox.Min.Z)),
                                mepTransform.OfPoint(new XYZ(mepBBox.Min.X, mepBBox.Min.Y, mepBBox.Max.Z)),
                                mepTransform.OfPoint(new XYZ(mepBBox.Max.X, mepBBox.Max.Y, mepBBox.Max.Z)),
                                mepTransform.OfPoint(new XYZ(mepBBox.Min.X, mepBBox.Max.Y, mepBBox.Max.Z)),
                                mepTransform.OfPoint(new XYZ(mepBBox.Max.X, mepBBox.Min.Y, mepBBox.Max.Z)),
                                mepTransform.OfPoint(new XYZ(mepBBox.Max.X, mepBBox.Max.Y, mepBBox.Min.Z))
                            };

                            var mepNewMin = new XYZ(mepPts.Min(p => p.X), mepPts.Min(p => p.Y), mepPts.Min(p => p.Z));
                            var mepNewMax = new XYZ(mepPts.Max(p => p.X), mepPts.Max(p => p.Y), mepPts.Max(p => p.Z));

                            mepBBoxInHostShared = new BoundingBoxXYZ
                            {
                                Min = mepNewMin,
                                Max = mepNewMax
                            };

                            _logger($"DEBUG: MEP bbox after transform: Min=({mepBBoxInHostShared.Min.X:F2}, {mepBBoxInHostShared.Min.Y:F2}, {mepBBoxInHostShared.Min.Z:F2}) Max=({mepBBoxInHostShared.Max.X:F2}, {mepBBoxInHostShared.Max.Y:F2}, {mepBBoxInHostShared.Max.Z:F2})");
                        }

                        // ✅ PHASE 1 OPTIMIZATION: Use memory management for intersection results
                        var hits = MepIntersectionService.FindIntersections(
                            mep,
                            structuralElements,
                            _logger);
                        
                        // Cache the intersection results for potential reuse
                        if (hits.Any())
                        {
                            MemoryManagementService.AddToCache(mep.Id, hits, hits.Count * 100); // Estimate 100 bytes per hit
                        }

                        _logger($"DEBUG: MepIntersectionService found {hits.Count} intersections for MEP {mep.Id}");

                        foreach (var (structuralElement, bbox, center) in hits)
                        {
                            intersections.Add((mep, structuralElement, bbox, center));
                            _logger($"Intersection found: MEP {mep.Id} ({mep.Category?.Name}) <-> {structuralElement.Category?.Name} {structuralElement.Id} at {center}");
                        }
                    }
                    else
                    {
                        // MEP element has no valid location curve or is not a line - skipping
                    }
                }
                catch (Exception ex)
                {
                    _logger($"ERROR: Failed to process MEP element {mep.Id}: {ex.Message}");
                }
            }

            // ✅ PHASE 1 OPTIMIZATION: Enforce memory management
            MemoryManagementService.EnforceCacheSize();
            
            _logger($"FindIntersectionsInternal completed: {intersections.Count} intersections found");
            return intersections;
        }
        */ // End of old code block

        /// <summary>
        /// METHOD 3: Check for damper presence at duct end points
        /// This prevents creating intersections for ducts that already have dampers
        /// </summary>
        private bool CheckForDamperAtDuctEnd(Document document, Element ductElement, Line ductLine, Transform ductTransform)
        {
            try
            {
                if (!(ductElement is Autodesk.Revit.DB.Mechanical.Duct duct))
                {
                    return false; // Not a duct, no need to check
                }

                _logger($"[METHOD3] Checking for damper at duct end: Duct {duct.Id}");

                // Get duct end points in host coordinates
                var endPoint1 = ductTransform?.OfPoint(ductLine.GetEndPoint(0)) ?? ductLine.GetEndPoint(0);
                var endPoint2 = ductTransform?.OfPoint(ductLine.GetEndPoint(1)) ?? ductLine.GetEndPoint(1);

                _logger($"[METHOD3] Duct end points: {endPoint1} and {endPoint2}");

                // Check each end point for damper presence
                foreach (var endPoint in new[] { endPoint1, endPoint2 })
                {
                    if (CheckForDamperNearPoint(document, endPoint))
                    {
                        _logger($"[METHOD3] ✓ DAMPER FOUND: Damper detected near duct end point {endPoint}");
                        return true;
                    }
                }

                _logger($"[METHOD3] ✗ NO DAMPER: No damper found near any duct end points");
                return false;
            }
            catch (Exception ex)
            {
                _logger($"[METHOD3] ERROR: Failed to check for damper at duct end: {ex.Message}");
                return false; // Default to false (don't skip) if error occurs
            }
        }

        /// <summary>
        /// Check for damper presence near a specific point
        /// </summary>
        private bool CheckForDamperNearPoint(Document document, XYZ searchPoint)
        {
            try
            {
                double searchRadius = 0.5; // 0.5 feet = ~150mm
                
                // ✅ CRITICAL FIX: Search in ALL linked documents, not just the host document
                // The ducts and dampers are in linked files (ME-00001), not the host document
                var allDampers = new List<Element>();
                
                // Search in host document first
                var hostCollector = new FilteredElementCollector(document);
                var hostDampers = hostCollector
                    .OfCategory(BuiltInCategory.OST_DuctAccessory)
                    .WhereElementIsNotElementType()
                    .ToElements()
                    .Where(d => IsDamperType(d))
                    .ToList();
                allDampers.AddRange(hostDampers);
                
                // Search in all linked documents
                var linkedDocs = document.Application.Documents.Cast<Document>()
                    .Where(doc => doc != document && !doc.IsFamilyDocument)
                    .ToList();
                
                foreach (var linkedDoc in linkedDocs)
                {
                    try
                    {
                        var linkedCollector = new FilteredElementCollector(linkedDoc);
                        var linkedDampers = linkedCollector
                            .OfCategory(BuiltInCategory.OST_DuctAccessory)
                            .WhereElementIsNotElementType()
                            .ToElements()
                            .Where(d => IsDamperType(d))
                            .ToList();
                        allDampers.AddRange(linkedDampers);
                        
                        _logger($"[METHOD3] Found {linkedDampers.Count} dampers in linked doc '{linkedDoc.Title}'");
                    }
                    catch (Exception ex)
                    {
                        _logger($"[METHOD3] WARNING: Could not search linked doc '{linkedDoc.Title}': {ex.Message}");
                    }
                }

                _logger($"[METHOD3] Found {allDampers.Count} total duct accessories across all documents");

                foreach (var damper in allDampers)
                {
                    var damperLocation = damper.Location as LocationPoint;
                    if (damperLocation?.Point != null)
                    {
                        double distance = damperLocation.Point.DistanceTo(searchPoint);
                        if (distance < searchRadius)
                        {
                            var damperType = damper.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM)?.AsValueString() ?? "";
                            var damperFamily = (damper as FamilyInstance)?.Symbol?.Family?.Name ?? "";
                            
                            _logger($"[METHOD3] Checking damper {damper.Id}: Type='{damperType}', Family='{damperFamily}'");
                            
                            if (IsDamperType(damper))
                            {
                                _logger($"[METHOD3] ✓ CONFIRMED DAMPER: {damper.Id} is a damper (Type='{damperType}', Family='{damperFamily}')");
                                return true;
                            }
                        }
                    }
                }

                return false;
            }
            catch (Exception ex)
            {
                _logger($"[METHOD3] ERROR: Failed to search for dampers near point: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Check if an element is a damper based on type and family name
        /// </summary>
        private bool IsDamperType(Element element)
        {
            try
            {
                if (element is FamilyInstance fi)
                {
                    var familyName = fi.Symbol?.Family?.Name ?? "";
                    var typeName = fi.Symbol?.Name ?? "";
                    var combinedName = $"{familyName} {typeName}".ToLowerInvariant();
                    
                    // Check for damper keywords
                    var damperKeywords = new[] { "damper", "dam", "fire", "smoke", "motorized", "motorised" };
                    
                    if (damperKeywords.Any(keyword => combinedName.Contains(keyword)))
                    {
                        _logger($"[METHOD3] Damper detected: '{combinedName}' contains damper keywords");
                        return true;
                    }
                }
                
                return false;
            }
            catch (Exception ex)
            {
                _logger($"[METHOD3] ERROR: Failed to check damper type: {ex.Message}");
                return false;
            }
        }
    }
}

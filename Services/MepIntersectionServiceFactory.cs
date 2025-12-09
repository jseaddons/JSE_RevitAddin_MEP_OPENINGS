using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// ✅ NICE3POINT TEMPLATE PATTERN: Factory for creating version-specific MepIntersectionService implementations
    /// Uses runtime version detection to avoid static initialization issues
    /// </summary>
    public static class MepIntersectionServiceFactory
    {
        private static IMepIntersectionService? _instance;
        private static readonly object _lock = new object();

        /// <summary>
        /// Get the appropriate MepIntersectionService implementation for the current Revit version
        /// </summary>
        public static IMepIntersectionService GetInstance()
        {
            if (_instance == null)
            {
                lock (_lock)
                {
                    if (_instance == null)
                    {
                        // ✅ Use compile-time version detection (from VersionInfo)
                        // This avoids runtime reflection and static initialization issues
#if REVIT2024_OR_GREATER
                        _instance = new MepIntersectionService2024();
#else
                        _instance = new MepIntersectionService2023();
#endif
                    }
                }
            }
            return _instance;
        }

        /// <summary>
        /// Clear the cached instance (useful for testing or version switching)
        /// </summary>
        public static void ClearInstance()
        {
            lock (_lock)
            {
                _instance = null;
            }
        }
    }

    /// <summary>
    /// ✅ R2023 Implementation: Uses static caches for performance
    /// </summary>
#if !REVIT2024_OR_GREATER
    internal class MepIntersectionService2023 : IMepIntersectionService
    {
        // Use the existing static implementation via wrapper
        public Transform GetCachedTransform(Document doc, List<RevitLinkInstance> links, Action<string>? log = null)
        {
            // Delegate to static method - but this will only compile in R2023 builds
            return MepIntersectionService.GetCachedTransform(doc, links, log);
        }

        public List<(Element, Element, BoundingBoxXYZ, XYZ)> FindIntersectionsBatch(
            List<(Element, Transform?)> mepElements,
            List<(Element, Transform?)> structuralElements,
            Action<string> log,
            HashSet<(int mepId, int structuralId)>? knownValidPairs = null,
            bool skipKnownPairsGeometryCheck = false)
        {
            return MepIntersectionService.FindIntersectionsBatch(
                mepElements, structuralElements, log, knownValidPairs, skipKnownPairsGeometryCheck);
        }
    }
#endif

    /// <summary>
    /// ✅ R2024 Implementation: No static caches to avoid TypeInitializationException
    /// </summary>
#if REVIT2024_OR_GREATER
    internal class MepIntersectionService2024 : IMepIntersectionService
    {
        // Simple transform cache per instance (not static)
        private readonly Dictionary<Document, Transform> _transformCache = new Dictionary<Document, Transform>();

        public Transform GetCachedTransform(Document doc, List<RevitLinkInstance> links, Action<string>? log = null)
        {
            if (!_transformCache.ContainsKey(doc))
            {
                var link = links?.FirstOrDefault(l => l.GetLinkDocument()?.Title == doc.Title);
                if (link != null)
                {
                    _transformCache[doc] = link.GetTotalTransform();
                    log?.Invoke($"[TransformCache] Cached transform for document: {doc.Title}");
                }
                else
                {
                    _transformCache[doc] = Transform.Identity;
                    log?.Invoke($"[TransformCache] No link found for document: {doc.Title}, using Identity transform");
                }
            }
            return _transformCache[doc];
        }

        public List<(Element, Element, BoundingBoxXYZ, XYZ)> FindIntersectionsBatch(
            List<(Element, Transform?)> mepElements,
            List<(Element, Transform?)> structuralElements,
            Action<string> log,
            HashSet<(int mepId, int structuralId)>? knownValidPairs = null,
            bool skipKnownPairsGeometryCheck = false)
        {
            // ✅ R2024 FIX: Implement directly to avoid accessing static class
            // This prevents any static initialization issues
            var results = new List<(Element, Element, BoundingBoxXYZ, XYZ)>();
            
            try
            {
                log?.Invoke($"[R24-FIX] FindIntersectionsBatch called: {mepElements.Count} MEP + {structuralElements.Count} structural");
                log?.Invoke($"[R24-FIX] ✅ DAMPER DETECTION CODE IS ACTIVE - Starting element processing loop");
                
                int damperCount = 0;
                int nonDamperCount = 0;
                int processedCount = 0;
                
                foreach (var (mepElem, mepTransform) in mepElements)
                {
                    processedCount++;
                    try
                    {
                        // ✅ ALWAYS LOG: Log every element to see what's being processed
                        var categoryName = mepElem.Category?.Name ?? "NULL";
                        var categoryId = mepElem.Category?.Id?.IntegerValue ?? -1;
                        log?.Invoke($"[R24-FIX] [{processedCount}/{mepElements.Count}] Processing element {mepElem.Id} (Category={categoryName}, CategoryId={categoryId})");
                        
                        // ✅ CRITICAL FIX: Check for dampers FIRST, before trying to get line
                        // Dampers should use damper intersection logic, not line-based logic
                        bool isDamper = IsDamperElement(mepElem);
                        log?.Invoke($"[R24-FIX] [{processedCount}/{mepElements.Count}] Element {mepElem.Id}: IsDamperElement={isDamper}, Category={categoryName}, CategoryId={categoryId}, OST_DuctAccessory={(int)BuiltInCategory.OST_DuctAccessory}");
                        
                        if (isDamper)
                        {
                            damperCount++;
                            log?.Invoke($"[R24-FIX] [{processedCount}/{mepElements.Count}] ✅ Element {mepElem.Id} (Category={categoryName}) is a damper - using damper intersection logic (Damper #{damperCount})");
                            
                            var damperBBox = mepElem.get_BoundingBox(null);
                            if (damperBBox == null)
                            {
                                log?.Invoke($"[DamperDetection] ⚠️ Damper {mepElem.Id} has no bounding box - skipping");
                                continue;
                            }
                            
                            // Transform damper bbox if needed
                            if (mepTransform != null && !mepTransform.IsIdentity)
                            {
                                var min = mepTransform.OfPoint(damperBBox.Min);
                                var max = mepTransform.OfPoint(damperBBox.Max);
                                damperBBox = new BoundingBoxXYZ
                                {
                                    Min = new XYZ(Math.Min(min.X, max.X), Math.Min(min.Y, max.Y), Math.Min(min.Z, max.Z)),
                                    Max = new XYZ(Math.Max(min.X, max.X), Math.Max(min.Y, max.Y), Math.Max(min.Z, max.Z))
                                };
                            }
                            
                            // Use damper intersection logic
                            var damperResults = FindDamperIntersectionsForFactory(mepElem, damperBBox, structuralElements, mepTransform, log);
                            results.AddRange(damperResults);
                            continue;
                        }
                        
                        // Get MEP Line (Centerline) - for non-dampers
                        nonDamperCount++;
                        log?.Invoke($"[MEP-Processing] Element {mepElem.Id} is NOT a damper - using line-based logic (Non-Damper #{nonDamperCount})");
                        
                        var mepLine = GetElementLine(mepElem);
                        if (mepLine == null)
                        {
                            log?.Invoke($"[MEP-Processing] ⚠️ Element {mepElem.Id} has no line - skipping");
                            continue;
                        }
                        
                        // Transform MEP line if needed
                        if (mepTransform != null && !mepTransform.IsIdentity)
                        {
                            mepLine = mepLine.CreateTransformed(mepTransform) as Line;
                        }
                        
                        if (mepLine == null) continue;

                        var mepBBox = mepElem.get_BoundingBox(null);
                        
                        foreach (var (structElem, structTransform) in structuralElements)
                        {
                            try
                            {
                                // Spatial Filter (Bounding Box)
                                var structBBox = structElem.get_BoundingBox(null);
                                if (structBBox != null && mepBBox != null)
                                {
                                    // Transform structural bbox if needed
                                    if (structTransform != null && !structTransform.IsIdentity)
                                    {
                                        var min = structTransform.OfPoint(structBBox.Min);
                                        var max = structTransform.OfPoint(structBBox.Max);
                                        structBBox = new BoundingBoxXYZ
                                        {
                                            Min = new XYZ(Math.Min(min.X, max.X), Math.Min(min.Y, max.Y), Math.Min(min.Z, max.Z)),
                                            Max = new XYZ(Math.Max(min.X, max.X), Math.Max(min.Y, max.Y), Math.Max(min.Z, max.Z))
                                        };
                                    }
                                    
                                    // Simple overlap check
                                    if (!BoundingBoxesOverlap(mepBBox, structBBox)) continue;
                                }

                                // ✅ Use EfficientIntersectionService for solid intersection
                                var points = EfficientIntersectionService.PerformSolidIntersection(mepLine, structElem, structTransform);
                                
                                // Create Result
                                if (points.Count > 0)
                                {
                                    var bbox = CreateBoundingBox(points);
                                    if (bbox != null)
                                    {
                                        var center = (bbox.Min + bbox.Max) * 0.5;
                                        results.Add((mepElem, structElem, bbox, center));
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                // Silently continue - log only in debug mode
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        log?.Invoke($"[R24-FIX] MEP element {mepElem.Id} failed: {ex.Message}");
                    }
                }
                
                log?.Invoke($"[R24-FIX] ✅ Processing complete: {damperCount} dampers processed, {nonDamperCount} non-dampers processed");
                log?.Invoke($"[R24-FIX] Found {results.Count} intersections");
                return results;
            }
            catch (Exception ex)
            {
                log?.Invoke($"[R24-FIX] FATAL ERROR: {ex.Message}");
                return results;
            }
        }

        // Helper methods for R2024 implementation
        private static Line GetElementLine(Element element)
        {
            if (element.Location is LocationCurve lc && lc.Curve is Line line)
            {
                return line;
            }
            return null;
        }

        private static bool BoundingBoxesOverlap(BoundingBoxXYZ bb1, BoundingBoxXYZ bb2)
        {
            if (bb1 == null || bb2 == null) return false;
            return bb1.Min.X <= bb2.Max.X && bb1.Max.X >= bb2.Min.X &&
                   bb1.Min.Y <= bb2.Max.Y && bb1.Max.Y >= bb2.Min.Y &&
                   bb1.Min.Z <= bb2.Max.Z && bb1.Max.Z >= bb2.Min.Z;
        }

        /// <summary>
        /// ✅ Check if an element is a damper (Duct Accessory category, excluding VCD/VOLUME)
        /// </summary>
        private static bool IsDamperElement(Element element)
        {
            // ✅ CRITICAL FIX: Check category FIRST - Duct Accessories are dampers
            if (element?.Category?.Id?.IntegerValue == (int)BuiltInCategory.OST_DuctAccessory)
            {
                // ✅ EXCLUDE: Skip VCD and VOLUME dampers (not in walls)
                if (element is FamilyInstance fi && fi.Symbol != null)
                {
                    var familyName = fi.Symbol.Family?.Name ?? "";
                    var typeName = fi.Symbol.Name ?? "";
                    var combinedName = $"{familyName} {typeName}".ToUpperInvariant();
                    
                    if (combinedName.Contains("VCD") || combinedName.Contains("VOLUME"))
                        return false;
                }
                
                return true;
            }
            
            // ✅ FALLBACK: Check family name
            if (element is FamilyInstance fi2)
            {
                var familyName = fi2.Symbol?.Family?.Name ?? "";
                var typeName = fi2.Symbol?.Name ?? "";
                var combinedName = $"{familyName} {typeName}".ToUpperInvariant();
                
                if (combinedName.Contains("VCD") || combinedName.Contains("VOLUME"))
                    return false;
                
                return familyName.Contains("Damper", StringComparison.OrdinalIgnoreCase);
            }
            
            return false;
        }
        
        /// <summary>
        /// ✅ Find damper intersections using bounding box overlap and containment checks
        /// </summary>
        private static List<(Element, Element, BoundingBoxXYZ, XYZ)> FindDamperIntersectionsForFactory(
            Element damperElement,
            BoundingBoxXYZ damperBBox,
            List<(Element, Transform?)> structuralElements,
            Transform? damperTransform,
            Action<string>? log)
        {
            var results = new List<(Element, Element, BoundingBoxXYZ, XYZ)>();
            
            const double tolerance = 0.5; // 6 inches
            var expandedMin = new XYZ(
                damperBBox.Min.X - tolerance,
                damperBBox.Min.Y - tolerance,
                damperBBox.Min.Z - tolerance);
            var expandedMax = new XYZ(
                damperBBox.Max.X + tolerance,
                damperBBox.Max.Y + tolerance,
                damperBBox.Max.Z + tolerance);
            
            log?.Invoke($"[DamperIntersection] Processing damper {damperElement.Id} with bbox Min=({damperBBox.Min.X:F2}, {damperBBox.Min.Y:F2}, {damperBBox.Min.Z:F2}) Max=({damperBBox.Max.X:F2}, {damperBBox.Max.Y:F2}, {damperBBox.Max.Z:F2})");
            log?.Invoke($"[DamperIntersection] Testing against {structuralElements.Count} structural elements");
            
            foreach (var (structuralElement, structTransform) in structuralElements)
            {
                try
                {
                    var structBBox = structuralElement.get_BoundingBox(null);
                    if (structBBox == null) continue;
                    
                    // Transform structural bbox if needed
                    if (structTransform != null && !structTransform.IsIdentity)
                    {
                        var min = structTransform.OfPoint(structBBox.Min);
                        var max = structTransform.OfPoint(structBBox.Max);
                        structBBox = new BoundingBoxXYZ
                        {
                            Min = new XYZ(Math.Min(min.X, max.X), Math.Min(min.Y, max.Y), Math.Min(min.Z, max.Z)),
                            Max = new XYZ(Math.Max(min.X, max.X), Math.Max(min.Y, max.Y), Math.Max(min.Z, max.Z))
                        };
                    }
                    
                    // Check bounding box intersection
                    if (!BoundingBoxesOverlap(new BoundingBoxXYZ { Min = expandedMin, Max = expandedMax }, structBBox))
                    {
                        continue;
                    }
                    
                    log?.Invoke($"[DamperIntersection] ✅ Intersection candidate: damper {damperElement.Id} with structural {structuralElement.Id}");
                    
                    // Check if damper is completely contained within wall
                    bool isDamperContainedInWall = 
                        damperBBox.Min.X >= structBBox.Min.X &&
                        damperBBox.Min.Y >= structBBox.Min.Y &&
                        damperBBox.Min.Z >= structBBox.Min.Z &&
                        damperBBox.Max.X <= structBBox.Max.X &&
                        damperBBox.Max.Y <= structBBox.Max.Y &&
                        damperBBox.Max.Z <= structBBox.Max.Z;
                    
                    if (isDamperContainedInWall)
                    {
                        log?.Invoke($"[DamperIntersection] ✅ Damper {damperElement.Id} is COMPLETELY CONTAINED within structural {structuralElement.Id} (buried inside wall)");
                    }
                    
                    // Calculate intersection bounding box
                    var intersectionMin = new XYZ(
                        Math.Max(damperBBox.Min.X, structBBox.Min.X),
                        Math.Max(damperBBox.Min.Y, structBBox.Min.Y),
                        Math.Max(damperBBox.Min.Z, structBBox.Min.Z));
                    var intersectionMax = new XYZ(
                        Math.Min(damperBBox.Max.X, structBBox.Max.X),
                        Math.Min(damperBBox.Max.Y, structBBox.Max.Y),
                        Math.Min(damperBBox.Max.Z, structBBox.Max.Z));
                    
                    // If damper is contained, use damper's bounding box
                    if (isDamperContainedInWall)
                    {
                        intersectionMin = damperBBox.Min;
                        intersectionMax = damperBBox.Max;
                    }
                    else if (intersectionMin.X > intersectionMax.X ||
                             intersectionMin.Y > intersectionMax.Y ||
                             intersectionMin.Z > intersectionMax.Z)
                    {
                        continue; // No valid intersection
                    }
                    
                    var intersectionBBox = new BoundingBoxXYZ
                    {
                        Min = intersectionMin,
                        Max = intersectionMax
                    };
                    
                    // Calculate intersection point
                    XYZ intersectionPoint;
                    if (structuralElement is Wall wall)
                    {
                        // For walls, project damper center onto wall face
                        var damperCenter = (damperBBox.Min + damperBBox.Max) * 0.5;
                        intersectionPoint = damperCenter; // Simplified - could project onto wall face
                    }
                    else
                    {
                        // For floors/framing, use intersection bbox center
                        intersectionPoint = (intersectionBBox.Min + intersectionBBox.Max) * 0.5;
                    }
                    
                    results.Add((damperElement, structuralElement, intersectionBBox, intersectionPoint));
                    log?.Invoke($"[DamperIntersection] ✅ Added intersection: damper {damperElement.Id} with structural {structuralElement.Id} at ({intersectionPoint.X:F3}, {intersectionPoint.Y:F3}, {intersectionPoint.Z:F3})");
                }
                catch (Exception ex)
                {
                    log?.Invoke($"[DamperIntersection] ERROR: Failed to process damper intersection for element {structuralElement.Id}: {ex.Message}");
                }
            }
            
            log?.Invoke($"[DamperIntersection] ✅ Damper {damperElement.Id} found {results.Count} intersections");
            return results;
        }

        private static BoundingBoxXYZ? CreateBoundingBox(List<XYZ> points)
        {
            if (points == null || points.Count == 0) return null;
            
            double minX = double.MaxValue, minY = double.MaxValue, minZ = double.MaxValue;
            double maxX = double.MinValue, maxY = double.MinValue, maxZ = double.MinValue;
            
            foreach (var pt in points)
            {
                if (pt == null) continue;
                minX = Math.Min(minX, pt.X);
                minY = Math.Min(minY, pt.Y);
                minZ = Math.Min(minZ, pt.Z);
                maxX = Math.Max(maxX, pt.X);
                maxY = Math.Max(maxY, pt.Y);
                maxZ = Math.Max(maxZ, pt.Z);
            }
            
            if (minX == double.MaxValue) return null;
            
            return new BoundingBoxXYZ
            {
                Min = new XYZ(minX, minY, minZ),
                Max = new XYZ(maxX, maxY, maxZ)
            };
        }
    }
#endif
}


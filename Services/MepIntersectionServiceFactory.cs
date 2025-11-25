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
                
                foreach (var (mepElem, mepTransform) in mepElements)
                {
                    try
                    {
                        // Get MEP Line (Centerline)
                        var mepLine = GetElementLine(mepElem);
                        if (mepLine == null) continue;
                        
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


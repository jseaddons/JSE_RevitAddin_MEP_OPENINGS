using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// ✅ NICE3POINT TEMPLATE PATTERN: Interface-based approach to avoid static initialization issues
    /// This allows version-specific implementations without triggering TypeInitializationException
    /// </summary>
    public interface IMepIntersectionService
    {
        Transform GetCachedTransform(Document doc, List<RevitLinkInstance> links, Action<string>? log = null);
        
        List<(Element, Element, BoundingBoxXYZ, XYZ)> FindIntersectionsBatch(
            List<(Element, Transform?)> mepElements,
            List<(Element, Transform?)> structuralElements,
            Action<string> log,
            View3D? view3D = null,
            HashSet<(int mepId, int structuralId)>? knownValidPairs = null,
            bool skipKnownPairsGeometryCheck = false);
    }
}


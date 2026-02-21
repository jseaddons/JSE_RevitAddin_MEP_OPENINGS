using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Manages family loading and symbol retrieval for sleeve placement.
    /// Implements IFamilyManager to satisfy dependency injection requirements.
    /// </summary>
    public class FamilyManager : IFamilyManager
    {
        private readonly Document _doc;
        private readonly FamilyLoadingService _familyLoadingService;
        private readonly Dictionary<string, FamilySymbol> _symbolCache = new Dictionary<string, FamilySymbol>();

        public FamilyManager(Document doc)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
            _familyLoadingService = new FamilyLoadingService(doc);
        }

        /// <summary>
        /// Gets the appropriate family name for a clash zone based on host type and shape.
        /// Pure static method for use in planning phases.
        /// </summary>
        public static string SelectUniversalFamily(string hostType, string mepShape)
        {
            return SelectUniversalFamily(hostType, mepShape, hostOrientation: null);
        }

        /// <summary>
        /// Gets the appropriate family name for a clash zone based on host type, shape, and orientation.
        /// When hostOrientation is "X" for wall/framing hosts, returns pre-rotated "_X" family variant
        /// to eliminate per-sleeve rotation API calls (2-3 Revit API calls saved per sleeve).
        /// </summary>
        public static string SelectUniversalFamily(string hostType, string mepShape, string hostOrientation)
        {
            // Treat Structural Framing the same as Walls for opening families.
            // For both Walls and Structural Framing we want vertical-host openings (OnWall),
            // and only use slab openings for floors/slabs.
            bool isWallLikeHost =
                hostType.IndexOf("Wall", StringComparison.OrdinalIgnoreCase) >= 0 ||
                hostType.IndexOf("Framing", StringComparison.OrdinalIgnoreCase) >= 0;

            bool isRound = mepShape.IndexOf("Round", StringComparison.OrdinalIgnoreCase) >= 0 ||
                          mepShape.IndexOf("Circular", StringComparison.OrdinalIgnoreCase) >= 0;

            // ✅ PERF: For X-wall/framing, use pre-rotated "_X" family variant to skip runtime rotation
            bool isXOriented = isWallLikeHost &&
                string.Equals(hostOrientation?.Trim(), "X", StringComparison.OrdinalIgnoreCase);

            if (isWallLikeHost)
            {
                if (isXOriented)
                    return isRound ? "CircularOpeningOnWall_X" : "RectangularOpeningOnWall_X";
                else
                    return isRound ? "CircularOpeningOnWall" : "RectangularOpeningOnWall";
            }
            else // Slab/Floor
            {
                return isRound ? "CircularOpeningOnSlab" : "RectangularOpeningOnSlab";
            }
        }

        /// <summary>
        /// Interface implementation delegating to the static logic.
        /// </summary>
        public string GetFamilyNameForClashZone(string hostType, string mepShape)
        {
            return SelectUniversalFamily(hostType, mepShape, hostOrientation: null);
        }

        public string GetFamilyNameForClashZone(string hostType, string mepShape, string hostOrientation)
        {
            return SelectUniversalFamily(hostType, mepShape, hostOrientation);
        }

        /// <summary>
        /// Loads a family into the document if not already present.
        /// Delegates to FamilyLoadingService for resource loading.
        /// </summary>
        public bool LoadFamily(Document doc, string familyName)
        {
            if (_familyLoadingService.IsFamilyLoaded(familyName))
            {
                return true;
            }

            // Try to load all required families to ensure the requested one is available
            return _familyLoadingService.LoadAllRequiredFamilies() && _familyLoadingService.IsFamilyLoaded(familyName);
        }

        /// <summary>
        /// Pre-caches family symbols for all clash zones to optimize placement performance.
        /// </summary>
        public void PreCacheFamilySymbols(Document doc, List<ClashZone> clashZones)
        {
            if (clashZones == null || !clashZones.Any()) return;

            var requiredFamilies = new HashSet<string>();

            // Identify all required families
            foreach (var zone in clashZones)
            {
                string familyName = !string.IsNullOrEmpty(zone.SleeveFamilyName) 
                    ? zone.SleeveFamilyName 
                    : SelectUniversalFamily(zone.StructuralElementType, zone.SleeveDiameter > 0 && zone.SleeveWidth <= 0 ? "Round" : "Rectangular");
                
                if (!string.IsNullOrEmpty(familyName))
                {
                    requiredFamilies.Add(familyName);
                }
            }

            // Ensure they are loaded
            _familyLoadingService.LoadAllRequiredFamilies();

            // Cache symbols
            foreach (var familyName in requiredFamilies)
            {
                if (!_symbolCache.ContainsKey(familyName))
                {
                    var symbol = new FilteredElementCollector(_doc)
                        .OfClass(typeof(FamilySymbol))
                        .Cast<FamilySymbol>()
                        .FirstOrDefault(x => x.FamilyName == familyName);

                    if (symbol != null)
                    {
                        _symbolCache[familyName] = symbol;
                        if (!symbol.IsActive)
                        {
                            symbol.Activate();
                        }
                    }
                }
            }
        }
    }
}

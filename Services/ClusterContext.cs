using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Minimal context to carry document-scoped services and caches used by clustering
    /// and related helpers. This class is intentionally lightweight and non-invasive.
    /// </summary>
    public class ClusterContext
    {
        public Document Document { get; private set; }

        public FlagManager FlagManager { get; private set; }

        public FilterManagementService FilterService { get; private set; }

        // Caches used to avoid repeated API calls during clustering
        public Dictionary<ElementId, Element> MepElementCache { get; private set; }

        public Dictionary<FamilyInstance, BoundingBoxXYZ> BboxCache { get; private set; }

        public Dictionary<FamilyInstance, Dictionary<string, Parameter>> ParameterCache { get; private set; }

        // Cache of loaded clash zones from XML (or database). AddXmlLoadCache will safely merge lists.
        public List<ClashZone> LoadedClashZones { get; private set; }

        /// <summary>
        /// Create a minimal context for the given document. Caches are created but empty.
        /// FlagManager and FilterService are left null and can be initialized via InitializeServices.
        /// </summary>
        public ClusterContext(Document document)
        {
            Document = document ?? throw new ArgumentNullException(nameof(document));
            MepElementCache = new Dictionary<ElementId, Element>();
            BboxCache = new Dictionary<FamilyInstance, BoundingBoxXYZ>();
            ParameterCache = new Dictionary<FamilyInstance, Dictionary<string, Parameter>>();
            LoadedClashZones = new List<ClashZone>();
        }

        /// <summary>
        /// Constructor with explicit dependencies (optional).
        /// </summary>
        public ClusterContext(
            Document document,
            FlagManager flagManager,
            FilterManagementService filterService,
            Dictionary<ElementId, Element> mepElementCache = null,
            Dictionary<FamilyInstance, BoundingBoxXYZ> bboxCache = null,
            Dictionary<FamilyInstance, Dictionary<string, Parameter>> parameterCache = null,
            List<ClashZone> loadedClashZones = null)
        {
            Document = document ?? throw new ArgumentNullException(nameof(document));
            FlagManager = flagManager;
            FilterService = filterService;

            MepElementCache = mepElementCache ?? new Dictionary<ElementId, Element>();
            BboxCache = bboxCache ?? new Dictionary<FamilyInstance, BoundingBoxXYZ>();
            ParameterCache = parameterCache ?? new Dictionary<FamilyInstance, Dictionary<string, Parameter>>();
            LoadedClashZones = loadedClashZones ?? new List<ClashZone>();
        }

        /// <summary>
        /// Initialize services if they are not already provided. Safe: does not overwrite existing instances.
        /// </summary>
        public void InitializeServices(Action<string> log = null, Action<string> updateStatus = null)
        {
            if (FlagManager == null && Document != null)
                FlagManager = new FlagManager(Document);

            if (FilterService == null && Document != null)
                FilterService = new FilterManagementService(Document, log, updateStatus);
        }

        /// <summary>
        /// Ensure caches are initialized. Optionally clears existing content.
        /// </summary>
        public void InitializeCaches(bool clearExisting = false)
        {
            if (MepElementCache == null || clearExisting)
                MepElementCache = new Dictionary<ElementId, Element>();

            if (BboxCache == null || clearExisting)
                BboxCache = new Dictionary<FamilyInstance, BoundingBoxXYZ>();

            if (ParameterCache == null || clearExisting)
                ParameterCache = new Dictionary<FamilyInstance, Dictionary<string, Parameter>>();

            if (LoadedClashZones == null || clearExisting)
                LoadedClashZones = new List<ClashZone>();
        }

        /// <summary>
        /// Safe getter for FlagManager - will create one if missing and Document is available.
        /// </summary>
        public FlagManager GetFlagManagerSafe()
        {
            if (FlagManager == null && Document != null)
                FlagManager = new FlagManager(Document);
            return FlagManager;
        }

        /// <summary>
        /// Safe getter for FilterManagementService - will create one if missing and Document is available.
        /// </summary>
        public FilterManagementService GetFilterServiceSafe(Action<string> log = null, Action<string> updateStatus = null)
        {
            if (FilterService == null && Document != null)
                FilterService = new FilterManagementService(Document, log, updateStatus);
            return FilterService;
        }

        /// <summary>
        /// Add loaded clash zones from XML (or DB) into the context cache. Duplicates (by Id) are ignored.
        /// </summary>
        public void AddXmlLoadCache(IEnumerable<ClashZone> loadedClashZones)
        {
            if (loadedClashZones == null) return;

            if (LoadedClashZones == null)
                LoadedClashZones = new List<ClashZone>(loadedClashZones);
            else
            {
                var existingIds = new HashSet<Guid>(LoadedClashZones.Select(z => z.Id));
                foreach (var cz in loadedClashZones)
                {
                    if (cz == null) continue;
                    if (!existingIds.Contains(cz.Id))
                    {
                        LoadedClashZones.Add(cz);
                        existingIds.Add(cz.Id);
                    }
                }
            }
        }
    }
}

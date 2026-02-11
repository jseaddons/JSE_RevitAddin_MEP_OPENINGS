using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Refresh
{
    /// <summary>
    /// Holds all shared state and cached data for a refresh operation.
    /// Loaded once at start, reused throughout refresh, disposed at end.
    /// </summary>
    public class RefreshContext : IDisposable, IRefreshDocumentContext, IRefreshCacheContext, IRefreshSelectionContext, IRefreshSettingsContext
    {
        // Core document reference
        public Document Document { get; }
        public UIDocument UIDocument { get; private set; }
        
        // UI selections (immutable for this refresh)
        public List<string> SelectedFilterNames { get; }
        public List<string> SelectedMepCategories { get; }
        public List<string> SelectedReferenceFiles { get; }
        public List<string> SelectedHostFiles { get; }
        public List<string> SelectedHostTypes { get; }
        public Dictionary<string, double> ClearanceSettings { get; }
        
        // Settings
        public bool EnableThreePointValidation { get; set; }
        public bool IsDeploymentMode { get; set; }
        
        // Cached data (loaded once, reused)
        public XmlCache XmlCache { get; set; }
        public StringPool StringPool { get; set; }
        public GeometryCache GeometryCache { get; set; }
        
        // Working data (built during refresh)
        public List<ClashZone> ExistingClashZones { get; set; }
        public List<ClashZone> NewClashZones { get; set; }
        public List<ClashZone> AllClashZones { get; set; }
        public List<(Element, Element, BoundingBoxXYZ, XYZ)> CurrentIntersections { get; set; }
        
        // ✅ PATH 3: Track validated and invalidated zones separately for distinct placement flows
        public List<ClashZone> ValidatedZones { get; set; } = new List<ClashZone>();
        public List<ClashZone> InvalidatedZones { get; set; } = new List<ClashZone>();
        
        // Path strategy (determined after loading existing zones)
        public IRefreshPathStrategy PathStrategy { get; set; }
        
        // Performance tracking
        public PerformanceMonitor PerformanceMonitor { get; }
        public string RefreshLogName { get; }
        
        // Model change detection
        public DateTime LastModelModifiedTime { get; set; }
        public string LastDocumentPath { get; set; }
        
        public RefreshContext(
            Document document, 
            UIDocument uiDocument,
            List<string> selectedFilterNames,
            List<string> selectedMepCategories,
            List<string> selectedReferenceFiles,
            List<string> selectedHostFiles,
            List<string> selectedHostTypes,
            Dictionary<string, double> clearanceSettings,
            bool enableThreePointValidation)
        {
            Document = document ?? throw new ArgumentNullException(nameof(document));
            UIDocument = uiDocument ?? throw new ArgumentNullException(nameof(uiDocument));
            
            SelectedFilterNames = selectedFilterNames ?? new List<string>();
            SelectedMepCategories = selectedMepCategories ?? new List<string>();
            SelectedReferenceFiles = selectedReferenceFiles ?? new List<string>();
            SelectedHostFiles = selectedHostFiles ?? new List<string>();
            SelectedHostTypes = selectedHostTypes ?? new List<string>();
            ClearanceSettings = clearanceSettings ?? new Dictionary<string, double>();
            
            EnableThreePointValidation = enableThreePointValidation;
            IsDeploymentMode = DeploymentConfiguration.DeploymentMode;
            
            // Initialize caches
            StringPool = new StringPool();
            GeometryCache = new GeometryCache();
            
            // Performance tracking
            string timestamp = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss");
            RefreshLogName = $"Refresh_{timestamp}.log";
            PerformanceMonitor = new PerformanceMonitor(RefreshLogName);
            
            // Model state
            LastDocumentPath = document.PathName;
            if (!string.IsNullOrEmpty(LastDocumentPath) && System.IO.File.Exists(LastDocumentPath))
            {
                LastModelModifiedTime = System.IO.File.GetLastWriteTime(LastDocumentPath);
            }
        }
        
        public bool HasModelChanged()
        {
            if (string.IsNullOrEmpty(Document.PathName) || !System.IO.File.Exists(Document.PathName))
                return true;
            
            var currentPath = Document.PathName;
            var currentModified = System.IO.File.GetLastWriteTime(currentPath);
            
            bool changed = currentPath != LastDocumentPath || currentModified != LastModelModifiedTime;
            
            if (changed)
            {
                LastDocumentPath = currentPath;
                LastModelModifiedTime = currentModified;
            }
            
            return changed;
        }
        
        /// <summary>
        /// Reset session state (for reuse if needed)
        /// </summary>
        public void ResetSession()
        {
            ExistingClashZones?.Clear();
            NewClashZones?.Clear();
            AllClashZones?.Clear();
            CurrentIntersections?.Clear();
            ValidatedZones?.Clear();
            InvalidatedZones?.Clear();
            GeometryCache?.Clear();
            StringPool?.Clear();
        }
        
        public void Dispose()
        {
            // Clear large caches
            GeometryCache?.Clear();
            XmlCache?.Clear();
            
            // Generate final performance report
            PerformanceMonitor?.GenerateReport(AllClashZones?.Count ?? 0);
        }
    }
    
    /// <summary>
    /// String pool for interning - reduces memory by 99% for duplicate strings
    /// </summary>
    public class StringPool
    {
        private readonly Dictionary<string, string> _pool = new Dictionary<string, string>(StringComparer.Ordinal);
        
        public string Intern(string value)
        {
            if (value == null) return null;
            if (string.IsNullOrEmpty(value)) return string.Empty;
            
            if (_pool.TryGetValue(value, out var pooled))
                return pooled;
            
            _pool[value] = value;
            return value;
        }
        
        public int Count => _pool.Count;
        
        public void Clear() => _pool.Clear();
    }
    
    /// <summary>
    /// Geometry cache - don't clear between refreshes unless model changed
    /// </summary>
    public class GeometryCache
    {
        private readonly Dictionary<int, GeometryElement> _geometryCache = new Dictionary<int, GeometryElement>();
        private readonly Dictionary<int, XYZ> _transformCache = new Dictionary<int, XYZ>();
        
        public GeometryElement GetGeometry(Element element, Options options)
        {
            int key = element.Id.GetIntegerValue();
            
            if (_geometryCache.TryGetValue(key, out var cached))
                return cached;
            
            var geometry = element.get_Geometry(options);
            _geometryCache[key] = geometry;
            
            return geometry;
        }
        
        public void CacheTransform(int elementId, XYZ transform)
        {
            _transformCache[elementId] = transform;
        }
        
        public bool TryGetTransform(int elementId, out XYZ transform)
        {
            return _transformCache.TryGetValue(elementId, out transform);
        }
        
        public void Clear()
        {
            _geometryCache.Clear();
            _transformCache.Clear();
        }
        
        public int GeometryCount => _geometryCache.Count;
        public int TransformCount => _transformCache.Count;
    }
}
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined.Phase1And2.Services;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined.Phase1And2.Repository;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined.Phase1And2.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined.Phase1And2.Interfaces;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.BoundingBox;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Rotation;
using JSE_RevitAddin_MEP_OPENINGS.Services.Combined;
using JSE_RevitAddin_MEP_OPENINGS.Services.Combined.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Geometry;
using JSE_RevitAddin_MEP_OPENINGS.Data; // For GlobalData

// Explicit Alias for UI Types
using Visibility = System.Windows.Visibility;
using ICommand = System.Windows.Input.ICommand;
using CommandManager = System.Windows.Input.CommandManager;
using SelectionObjectType = Autodesk.Revit.UI.Selection.ObjectType;

namespace JSE_RevitAddin_MEP_OPENINGS.UI
{
    public class CombinedSleeveViewModel : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;
        
        private readonly Autodesk.Revit.UI.UIDocument _uiDocument;
        private readonly Document _document;
        
        // Visibility Flags
        private bool _isAutoMode = true;
        private string _statusMessage = "Ready";
        
        // Services - NEW Agent A/B Architecture
        private readonly ICombinedClusterDiscoveryService _discoveryService;
        private readonly ICrossCategoryProximityService _proximityService; // NEW: Agent B proximity detection
        private readonly CombinedSleevePlacementService _placementService; // NEW: Agent A placement

        private readonly IClashZoneRepository _repo;
        private readonly CrashSafeExecutor _executor; // Safety wrapper
        private readonly IManualClusterCalculationAdapter _manualCalculator;
        
        // UI Properties (Abbreviated for rewrite)
        public bool IsAutoMode
        {
            get => _isAutoMode;
            set
            {
                _isAutoMode = value;
                OnPropertyChanged(nameof(IsAutoMode));
                OnPropertyChanged(nameof(IsManualMode));
                OnPropertyChanged(nameof(AutoVisibility));
                OnPropertyChanged(nameof(ManualVisibility));
            }
        }
        public bool IsManualMode => !_isAutoMode;
        public Visibility AutoVisibility => _isAutoMode ? Visibility.Visible : Visibility.Collapsed;
        public Visibility ManualVisibility => !_isAutoMode ? Visibility.Visible : Visibility.Collapsed;

        public string StatusMessage
        {
            get => _statusMessage;
            set
            {
                _statusMessage = value;
                OnPropertyChanged(nameof(StatusMessage));
            }
        }

        public ObservableCollection<string> SelectedSleevesList { get; } = new ObservableCollection<string>();
        private List<ElementId> _selectedIds = new List<ElementId>();

        public ICommand RefreshCommand { get; }
        public ICommand CreateCombinedSleevesCommand { get; }
        public ICommand SelectCrossingCommand { get; }
        public ICommand SelectIndividuallyCommand { get; }
        public ICommand JoinSelectedCommand { get; }
        public bool CanJoin => _selectedIds.Count >= 2;

        public Action HideRequest { get; set; }
        public Action ShowRequest { get; set; }
        public Action CloseRequest { get; set; }
        
        private readonly Autodesk.Revit.UI.ExternalEvent _externalEvent;
        private readonly CombinedSleeveRequestHandler _requestHandler;

        public CombinedSleeveViewModel(
            Autodesk.Revit.UI.UIDocument uiDoc,
            Autodesk.Revit.UI.ExternalEvent externalEvent,
            CombinedSleeveRequestHandler requestHandler,
            IClashZoneRepository repo,
            ICombinedClusterDiscoveryService discoveryService,
            ICrossCategoryProximityService proximityService,
            CombinedSleevePlacementService placementService,
            IManualClusterCalculationAdapter manualCalculator)
        {
            _uiDocument = uiDoc;
            _document = uiDoc.Document;
            _externalEvent = externalEvent;
            _requestHandler = requestHandler;

            // Injected Dependencies
            _repo = repo ?? throw new ArgumentNullException(nameof(repo));
            _discoveryService = discoveryService ?? throw new ArgumentNullException(nameof(discoveryService));
            _proximityService = proximityService ?? throw new ArgumentNullException(nameof(proximityService));
            _placementService = placementService ?? throw new ArgumentNullException(nameof(placementService));
            _manualCalculator = manualCalculator ?? throw new ArgumentNullException(nameof(manualCalculator));

            _executor = new CrashSafeExecutor();

            // Commands
            RefreshCommand = new RelayCommand(Refresh);
            CreateCombinedSleevesCommand = new RelayCommand(CreateCombinedSleeves); // Auto flow

            SelectCrossingCommand = new RelayCommand(SelectByCrossing);
            SelectIndividuallyCommand = new RelayCommand(SelectIndividually);
            JoinSelectedCommand = new RelayCommand(JoinSelected, () => CanJoin);
        }

        // Category Selection Properties
        private bool _isDuctsSelected = false;
        public bool IsDuctsSelected
        {
            get => _isDuctsSelected;
            set { _isDuctsSelected = value; OnPropertyChanged(nameof(IsDuctsSelected)); }
        }

        private bool _isPipesSelected = false;
        public bool IsPipesSelected
        {
            get => _isPipesSelected;
            set { _isPipesSelected = value; OnPropertyChanged(nameof(IsPipesSelected)); }
        }

        private bool _isCableTraysSelected = false;
        public bool IsCableTraysSelected
        {
            get => _isCableTraysSelected;
            set { _isCableTraysSelected = value; OnPropertyChanged(nameof(IsCableTraysSelected)); }
        }

        private bool _isConduitsSelected = false;
        public bool IsConduitsSelected
        {
            get => _isConduitsSelected;
            set { _isConduitsSelected = value; OnPropertyChanged(nameof(IsConduitsSelected)); }
        }

        private bool _isDuctAccessoriesSelected = false;
        public bool IsDuctAccessoriesSelected
        {
            get => _isDuctAccessoriesSelected;
            set { _isDuctAccessoriesSelected = value; OnPropertyChanged(nameof(IsDuctAccessoriesSelected)); }
        }

        protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));

        private void Refresh()
        {
            StatusMessage = "Refreshed.";
        }

        private void CreateCombinedSleeves()
        {
            // Trigger the Auto Workflow
            // This usually runs CombinedSleeveManager.Execute()
            // Here we just delegate to it via External Event if strictly needed, 
            // OR if we are just calling logic:
            
            _executor.ExecuteWithTimeout(() =>
            {
                DebugLogger.SetCombinedSleeveLogFile();
                // ✅ USER REQUEST: Overwrite log instead of appending
                DebugLogger.InitLogFile("combinesleeveplacer");
                DebugLogger.SetServiceContext("CombinedSleeveAuto");
                DebugLogger.IsEnabled = true;
                StatusMessage = "Running Auto-Clustering...";

                try
                {
                    if (_requestHandler == null)
                    {
                        DebugLogger.Error("[AutoCluster] RequestHandler is null");
                        StatusMessage = "Error: RequestHandler not initialized";
                        return Autodesk.Revit.UI.Result.Failed;
                    }

                    _requestHandler.SetAction(async (uiapp) =>
                    {
                        DebugLogger.Info("[AutoCluster] Starting auto combined sleeve placement");
                        Document doc = uiapp.ActiveUIDocument.Document;
                        if (!doc.IsValidObject) return;

                        // Example: Use default categories and filter name (could be parameterized)
                        // CRITICAL FIX: Use user-selected categories instead of hardcoded defaults
                        var categories = new List<string>();
                        if (IsDuctsSelected) categories.Add(MepCategoryConstants.DUCTS);
                        if (IsPipesSelected) categories.Add(MepCategoryConstants.PIPES);
                        if (IsCableTraysSelected) categories.Add(MepCategoryConstants.CABLE_TRAYS);
                        if (IsConduitsSelected) categories.Add(MepCategoryConstants.CONDUITS);
                        if (IsDuctAccessoriesSelected) categories.Add(MepCategoryConstants.DUCT_ACCESSORIES);

                        DebugLogger.Info($"[AutoCluster] Selected Categories: [{string.Join(", ", categories)}]");

                        if (categories.Count == 0)
                        {
                            StatusMessage = "No categories selected.";
                            DebugLogger.Warning("[AutoCluster] No categories selected by user.");
                            return;
                        }

                        // CRITICAL FIX: Use wildcard filter to find sleeves across all filters
                        // "Combined" was likely an invalid filter name causing 0 results
                        var filterName = "*"; 

                        // ✅ SECTION BOX FILTERing: Pass section box if available
                        // CRITICAL FIX: Use SectionBoxHelper to get WORLD coordinates (with transform applied)
                        // Raw v3d.GetSectionBox() returns LOCAL coordinates which don't match element positions!
                        Autodesk.Revit.DB.BoundingBoxXYZ sectionBox = null;
                        if (doc.ActiveView is View3D v3d && v3d.IsSectionBoxActive)
                        {
                            sectionBox = Helpers.SectionBoxHelper.GetSectionBoxBounds(v3d);
                            DebugLogger.Info($"[AutoCluster] Using Section Box from View: {v3d.Name}");
                        }

                        // 1. Discover cluster sleeves (returns ClusterSleeveInfo)
                        var clusterSleeveInfos = _discoveryService.Discover(_uiDocument, filterName, categories, sectionBox);
                        DebugLogger.Info($"[AutoCluster] Discovered {clusterSleeveInfos.Count} cluster sleeves using categories: {string.Join(", ", categories)} and filter: {filterName}");
                        
                        if (clusterSleeveInfos.Count == 0)
                        {
                            StatusMessage = "No cluster sleeves found (Check selection/filters).";
                            return;
                        }

                        // 2. Convert to UnifiedSleeves for proximity detection
                        var unifiedSleeves = new List<UnifiedSleeve>();
                        foreach (var info in clusterSleeveInfos)
                        {
                            try
                            {
                                if (info.OriginalZone == null)
                                {
                                    DebugLogger.Warning($"[AutoCluster] Sleeve {info.SleeveInstanceId} has no OriginalZone. Skipping.");
                                    continue;
                                }

                                // ✅ CRITICAL FIX: Use static factory method to ensure BBox inflation (for flat floor sleeves)
                                // and uniform property mapping. 
                                var unified = UnifiedSleeve.FromClashZone(info.OriginalZone);
                                if (unified != null)
                                {
                                    unifiedSleeves.Add(unified);
                                }
                            }
                            catch (Exception ex)
                            {
                                DebugLogger.Warning($"[AutoCluster] Failed to convert sleeve {info.SleeveInstanceId}: {ex.Message}");
                            }
                        }

                        // 3. Detect proximity groups using NEW service
                        double proximityThreshold = 1.0; // 1 foot default
                        var proximityGroups = _proximityService.DetectProximityGroups(unifiedSleeves, proximityThreshold);
                        DebugLogger.Info($"[AutoCluster] Formed {proximityGroups.Count} proximity groups (threshold={proximityThreshold:F2} ft)");

                        if (proximityGroups.Count == 0)
                        {
                            StatusMessage = "No proximity groups found (sleeves too far apart).";
                            return;
                        }

                        // 4. Place combined sleeves using NEW placement service
                        // 4. Place combined sleeves using NEW placement service (Batch Method handles Transaction & DB Save)
                        try
                        {
                            var placedSleeves = _placementService.PlaceProximityGroups(proximityGroups, 0, 0);
                            StatusMessage = $"Auto-Clustered {placedSleeves.Count} combined sleeves.";
                            DebugLogger.Info($"[AutoCluster] Batch placement complete. Placed: {placedSleeves.Count}");
                        }
                        catch (Exception ex)
                        {
                            StatusMessage = "Auto-Cluster failed: " + ex.Message;
                            DebugLogger.Error("Auto-Cluster Failed: " + ex.ToString());
                        }
                    });

                    HideRequest?.Invoke();
                    if (System.Windows.Application.Current != null)
                    {
                        System.Windows.Application.Current.Dispatcher.BeginInvoke(new Action(() => {
                            try {
                                _externalEvent.Raise();
                            } catch (Exception ex) {
                                DebugLogger.Error("ExternalEvent.Raise() failed: " + ex.ToString());
                            }
                        }), System.Windows.Threading.DispatcherPriority.Background);
                    }
                    else
                    {
                        DebugLogger.Warning("[AutoCluster] Application.Current is null, raising directly");
                        _externalEvent.Raise();
                    }
                    return Autodesk.Revit.UI.Result.Succeeded;
                }
                catch (Exception ex)
                {
                    StatusMessage = "Auto-Cluster failed: " + ex.Message;
                    DebugLogger.Error("Auto-Cluster Failed: " + ex.ToString());
                    return Autodesk.Revit.UI.Result.Failed;
                }
            }, "Auto Cluster");
        }

        private void SelectByCrossing()
        {
            try
            {
                StatusMessage = "Select sleeves...";
                HideRequest?.Invoke(); // Hide before selection
                var refs = _uiDocument.Selection.PickObjects(SelectionObjectType.Element, new SleeveSelectionFilter(), "Select sleeves to join");
                ShowRequest?.Invoke(); // Show after selection
                
                UpdateSelection(refs);
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                StatusMessage = "Selection canceled.";
            }
            catch (Exception ex)
            {
                StatusMessage = "Selection error: " + ex.Message;
            }
        }

        private void SelectIndividually()
        {
            try
            {
                // Standard multi-selection is better than a custom loop that hides/shows UI
                StatusMessage = "Select sleeves (Finish when done)...";
                _selectedIds.Clear();
                SelectedSleevesList.Clear();
                
                // Use PickObjects to allow multiple selection in one go
                HideRequest?.Invoke(); // Hide before selection
                var refs = _uiDocument.Selection.PickObjects(SelectionObjectType.Element, new SleeveSelectionFilter(), "Select sleeves to join");
                ShowRequest?.Invoke(); // Show after selection
                
                UpdateSelection(refs);
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                StatusMessage = "Selection finished.";
            }
            catch (Exception ex)
            {
                StatusMessage = "Selection error: " + ex.Message;
            }
        }

        private void UpdateSelection(IList<Reference> refs)
        {
            _selectedIds.Clear();
            SelectedSleevesList.Clear();
            foreach (var r in refs)
            {
                _selectedIds.Add(r.ElementId);
                var elem = _document.GetElement(r.ElementId);
                SelectedSleevesList.Add($"{elem.Category.Name} - {elem.Id}");
            }
             OnPropertyChanged(nameof(CanJoin));
             StatusMessage = $"{_selectedIds.Count} sleeves selected.";
        }

        private void JoinSelected()
        {
            if (_selectedIds.Count < 2) return;

            // Refactored Manual Join: Use Auto-Cluster Logic (CombinedSleevePlacementService)
            // This ensures exact parity with Auto-Join workflow.
            
            _executor.ExecuteWithTimeout(() =>
            {
                // Configure Wrapper Logger
                DebugLogger.SetCombinedSleeveLogFile();
                // ✅ USER REQUEST: Overwrite log instead of appending
                DebugLogger.InitLogFile("combinesleeveplacer");
                DebugLogger.SetServiceContext("CombinedSleeveManual");
                DebugLogger.IsEnabled = true;
                
                DebugLogger.Info("Starting REFACTORED Manual Join sequence (Delegating to Auto-Cluster Service)...");

                try
                {
                    StatusMessage = "Joining...";
                    
                    if (_requestHandler == null)
                    {
                        DebugLogger.Error("[ManualJoin] RequestHandler is null");
                        StatusMessage = "Error: RequestHandler not initialized";
                        return Autodesk.Revit.UI.Result.Failed;
                    }

                    // Delegate Transaction to External Event
                    _requestHandler.SetAction((uiapp) =>
                    {
                        Document doc = uiapp.ActiveUIDocument.Document;
                        if (!doc.IsValidObject) return;

                        // 1. Construct ProximityGroup from Selected Sleeves
                        var group = new JSE_RevitAddin_MEP_OPENINGS.Services.Combined.Models.ProximityGroup();
                        
                        // We need to fetch the original sleeve data (ClashZone or ClusterSleeveData)
                        // The repository has methods for this.
                        var selectedIntIds = _selectedIds.Select(id => id.IntegerValue).ToList();
                        
                        var individualSleeves = _repo.GetClashZonesBySleeveIds(selectedIntIds);
                        // Note: GetClashZonesBySleeveIds might return multiple zones for same sleeve? Usually 1:1 for uncombined.
                        // Also fetch ClusterSleeves (if any selected sleeves are clusters)
                        
                        // We need to distinguish if a selected ID is a Cluster or Individual.
                        // Best way: Check element type or use Repository logic.
                        
                        // Let's iterate selected elements and construct UnifiedSleeves
                        var unifiedSleeves = new List<UnifiedSleeve>();
                        
                        foreach(var id in _selectedIds)
                        {
                            var elem = doc.GetElement(id);
                            if (elem == null) continue;
                            
                            // Check if it's a Cluster Sleeve
                            // We can check by parameter or family name, or try to find in DB. This is tricky without metadata.
                            // BUT ManualClusterCalculationAdapter logic used _repo.GetClashZonesBySleeveIds AND _repo.GetClusterSleevesByInstanceIds
                            
                            // Let's use the same approach to find source data
                            object sourceData = null;
                            SleeveType type = SleeveType.Individual;
                            
                            // Try Individual
                            var cz = individualSleeves.FirstOrDefault(z => z.SleeveInstanceId == id.IntegerValue);
                            if (cz != null)
                            {
                                sourceData = cz;
                                type = SleeveType.Individual;
                            }
                            else
                            {
                                // Try Cluster
                                // We need to access CombinedClusterRepository if possible, or use ClashZoneRepository's cluster method
                                // The adapter used _repo.GetClusterSleevesByInstanceIds? No, _repo is IClashZoneRepository
                                // Let's try casting _repo to ClashZoneRepository to access specific methods or use the interface
                                // if IClashZoneRepository has GetClusterSleevesByInstanceIds (it should).
                                
                                // Actually, IClashZoneRepository might not have it exposed.
                                // But we know it exists in ClashZoneRepository.
                                if (_repo is ClashZoneRepository concreteRepo)
                                {
                                     var clusters = concreteRepo.GetClusterSleevesByInstanceIds(new List<int> { id.IntegerValue });
                                     if(clusters != null && clusters.Count > 0)
                                     {
                                         sourceData = clusters[0];
                                         type = SleeveType.Cluster;
                                     }
                                }
                            }
                            
                            if (sourceData == null)
                            {
                                // If not found in DB, we can't fully support it in "Auto Logic" which relies on DB data.
                                // PROPOSAL: Create a lightweight fake source data? Or just skip?
                                // User wants "same as auto". Auto logic relies on DB.
                                // If manual selection involves sleeves not in DB, it should probably fail or warn.
                                DebugLogger.Warning($"[ManualJoin] Selected sleeve {id} not found in DB (Individual/Cluster tables). Skipping.");
                                continue;
                            }
                            
                            var unified = new UnifiedSleeve
                            {
                                Id = type == SleeveType.Individual ? $"I_{id.IntegerValue}" : $"C_{id.IntegerValue}",
                                Type = type,
                                Category = elem.Category?.Name ?? "Unknown",
                                BoundingBox = elem.get_BoundingBox(null),
                                SourceData = sourceData
                                // HostType/Orientation: usually extracted from SourceData. 
                                // UnifiedSleeve helper usually does this.
                                // We can use UnifiedSleeve.FromClashZone / FromClusterSleeve static methods!
                            };
                            
                            // Re-create unified using static helpers if possible
                            if (type == SleeveType.Individual && sourceData is Models.ClashZone z)
                            {
                                unified = UnifiedSleeve.FromClashZone(z);
                                
                                // ✅ CRITICAL FIX: Ensure Geometry is valid!
                                // If DB corners are 0,0,0 (not extracted yet), use Live Element BBox
                                bool cornersAreZero = unified.Corners.All(c => c.IsZeroLength());
                                if (cornersAreZero)
                                {
                                    DebugLogger.Warning($"[ManualJoin] Sleeve {id} has valid DB entry but 0,0,0 corners. Using Live BBox.");
                                    unified.Corners = null; // Force fallback to BBox in ProximityGroup logic
                                }
                            }
                            else if (type == SleeveType.Cluster && sourceData is Data.Repositories.ClusterSleeveData c)
                            {
                                unified = UnifiedSleeve.FromClusterSleeve(c);
                                
                                // ✅ CRITICAL FIX: Same check for Cluster Sleeves
                                bool cornersAreZero = unified.Corners.All(k => k.IsZeroLength());
                                if (cornersAreZero)
                                {
                                     DebugLogger.Warning($"[ManualJoin] Cluster Sleeve {id} has 0,0,0 corners in DB. Using Live BBox.");
                                     unified.Corners = null; 
                                }
                            }
                            else if (type == SleeveType.Cluster && sourceData is Models.ClusterSleeve cModel)
                            {
                                // ✅ FIX: Handle Model type returned by generic repository
                                unified = UnifiedSleeve.FromClusterSleeve(cModel);
                                
                                bool cornersAreZero = unified.Corners.All(k => k.IsZeroLength());
                                if (cornersAreZero)
                                {
                                     DebugLogger.Warning($"[ManualJoin] Cluster Sleeve (Model) {id} has 0,0,0 corners in DB. Using Live BBox.");
                                     unified.Corners = null; 
                                }
                            }
                                
                            // Always update BBox from live element to be safe (DB might be stale)
                            var liveBBox = elem.get_BoundingBox(null);
                            if (liveBBox != null)
                            {
                                unified.BoundingBox = liveBBox;
                                // Also update PlacementPoint if corners were zero
                                if (unified.Corners == null || unified.Corners.Count == 0 || unified.GetCenter().IsZeroLength())
                                {
                                    unified.PlacementPoint = (liveBBox.Min + liveBBox.Max) / 2.0;
                                }
                            }

                            unifiedSleeves.Add(unified);
                            group.Categories.Add(unified.Category);
                        }
                        
                        group.Sleeves = unifiedSleeves;
                        
                        DebugLogger.Info($"[ManualJoin] Constructed Proximity Group with {group.Sleeves.Count} sleeves.");
                        
                        if (group.Sleeves.Count < 2)
                        {
                            DebugLogger.Error("[ManualJoin] Not enough valid sleeves found in DB to form a group.");
                            // Fallback? No, user wants parity.
                            return;
                        }

                        // 2. Delegate to Placement Service
                        // Construct a list of 1 group
                        var groups = new List<JSE_RevitAddin_MEP_OPENINGS.Services.Combined.Models.ProximityGroup> { group };
                        
                        // Call PlaceProximityGroups (This handles Transaction, Placement, Flag Update, Cleanup)
                        // Use dummy comboId/filterId (0,0) as this is manual
                        var results = _placementService.PlaceProximityGroups(groups, 0, 0);
                        
                        DebugLogger.Info($"[ManualJoin] Service returned {results.Count} placed sleeves.");
                        
                        if (results.Count > 0)
                        {
                            StatusMessage = "Manual Join Successful.";
                        }
                        else
                        {
                             StatusMessage = "Manual Join Failed (Service returned 0 results).";
                        }
                        
                        // Clear Selection
                        _selectedIds.Clear();
                        SelectedSleevesList.Clear();
                        OnPropertyChanged(nameof(CanJoin));
                    });

                    // Ensure UI actions (Hide/Show)
                    HideRequest?.Invoke();
                    
                     if (System.Windows.Application.Current != null)
                    {
                        System.Windows.Application.Current.Dispatcher.BeginInvoke(new Action(() => {
                            try {
                                _externalEvent.Raise();
                            } catch (Exception ex) {
                                DebugLogger.Error("ExternalEvent.Raise() failed: " + ex.ToString());
                            }
                        }), System.Windows.Threading.DispatcherPriority.Background);
                    }
                    else
                    {
                        _externalEvent.Raise();
                    }
                    
                    return Autodesk.Revit.UI.Result.Succeeded;
                }
                catch (Exception ex)
                {
                    StatusMessage = "Manual Join Failed: " + ex.Message;
                    DebugLogger.Error("Manual Join Failed: " + ex.ToString());
                    return Autodesk.Revit.UI.Result.Failed;
                }
            }, "Manual Join");
        }








        
        /// <summary>
        /// Creates a combined sleeve from a candidate group (Auto Workflow)
        /// Uses CombinedSleevePlacementService for reliable placement and joining
        /// </summary>
        private void CreateCombinedSleeve(CombinedClusterCandidate candidate)
        {
            try
            {
                // Construct a temporary ProximityGroup to pass to the service
                // The service expects this structure to perform placement
                var group = new JSE_RevitAddin_MEP_OPENINGS.Services.Combined.Models.ProximityGroup
                {
                    // No Key property on ProximityGroup
                };
                
                // Add categories
                if (candidate.CategoriesInvolved != null)
                {
                    foreach(var cat in candidate.CategoriesInvolved) group.Categories.Add(cat);
                }
                
                // IMPORTANT: We need to set the pre-calculated geometry directly on the group
                // OR ensure the service uses the candidate's geometry if provided.
                // However, ProximityGroup calculates geometry from sleeves.
                // Since we don't have the original sleeves easily here (only IDs in candidate),
                // we'll use a slightly different overload or Strategy if possible.
                
                // ALTERNATIVE: Since we just want the Placement + Join logic, and we have the geometry:
                // We should expose a method in the service that accepts geometry directly OR
                // Update the service to handle this case.
                
                // Given constraints, I will use the service's PlaceSingleCombinedSleeveInRevit 
                // but I need to mock the group's geometric calculation methods OR
                // simpler: I'll duplicate the ROBUST PLACEMENT & JOIN logic right here for now
                // to avoid complex refactoring of the Service's input model which requires deep dependency changes.
                // The User wants AUTO JOIN FIXED.
                
                var bbox = candidate.CombinedBoundingBox;
                var width = candidate.CombinedWidth;
                var height = candidate.CombinedHeight;
                var placementPoint = candidate.CombinedBoundingBox.Min; // Using Min as placement point per previous logic
                
                // 1. Determine Family
                string familyName = "RectangularOpeningOnWall";
                if (candidate.HostType != null && (candidate.HostType.Contains("Floor") || candidate.HostType.Contains("Slab")))
                {
                    familyName = "RectangularOpeningOnSlab";
                }
                
                FamilySymbol symbol = new FilteredElementCollector(_document)
                    .OfClass(typeof(FamilySymbol))
                    .Cast<FamilySymbol>()
                    .FirstOrDefault(x => x.Name.Equals(familyName, StringComparison.OrdinalIgnoreCase) || 
                                         x.Family.Name.Equals(familyName, StringComparison.OrdinalIgnoreCase));
                                         
                if (symbol != null)
                {
                    if (!symbol.IsActive) symbol.Activate();
                    
                    // 3. Place Instance
                    var instance = _document.Create.NewFamilyInstance(placementPoint, symbol, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                    
                    if (instance != null)
                    {
                        // 4. Set Parameters
                        var pWidth = instance.LookupParameter("Width");
                        var pHeight = instance.LookupParameter("Height");
                        if (pWidth != null) pWidth.Set(width);
                        if (pHeight != null) pHeight.Set(height);
                        
                        var pComments = instance.LookupParameter("Comments");
                        if (pComments != null) pComments.Set($"Combined: {candidate.CombinedClusterInstanceId}");
                        
                        // 5. AUTO-JOIN (CRITICAL FIX)
                        try 
                        {
                            Element host = instance.Host;
                            // Search for host if null
                            if (host == null)
                            {
                                var potentialHosts = new FilteredElementCollector(_document)
                                    .OfClass(typeof(HostObject))
                                    .WherePasses(new BoundingBoxIntersectsFilter(new Outline(placementPoint - new XYZ(0.5, 0.5, 0.5), placementPoint + new XYZ(0.5, 0.5, 0.5))))
                                    .Cast<HostObject>()
                                    .ToList();

                                if (potentialHosts.Count > 0)
                                    host = potentialHosts.OrderBy(h => h.Location is LocationCurve lc ? lc.Curve.Distance(placementPoint) : 100).FirstOrDefault();
                            }
                            
                            if (host != null)
                            {
                                if (!JoinGeometryUtils.AreElementsJoined(_document, host, instance))
                                {
                                    JoinGeometryUtils.JoinGeometry(_document, host, instance);
                                }
                            }
                        }
                        catch { /* Ignore join errors */ }

                        // 6. Persistence (removed - using new placement service instead)
                        // _persistService.PersistCombinedCluster(candidate, instance.Id.IntegerValue);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error creating combined sleeve: {ex.Message}");
            }
        }
    }
    
    public class SleeveSelectionFilter : Autodesk.Revit.UI.Selection.ISelectionFilter
    {
        public bool AllowElement(Element elem)
        {
            if (elem is FamilyInstance fi)
            {
                return elem.Category.Name.Contains("Generic Models") || elem.Category.Name.Contains("Speciality Equipment") 
                    || elem.Category.Name.Contains("Plumbing Fixtures") || elem.Category.Name.Contains("Mechanical Equipment");
            }
            return false;
        }
        public bool AllowReference(Reference reference, XYZ position) => true;
    }

    public class RelayCommand : ICommand
    {
        private readonly Action _execute;
        private readonly Func<bool> _canExecute;
        public RelayCommand(Action execute, Func<bool> canExecute = null) { _execute = execute; _canExecute = canExecute; }
        public event EventHandler CanExecuteChanged { add { CommandManager.RequerySuggested += value; } remove { CommandManager.RequerySuggested -= value; } }
        public bool CanExecute(object parameter) => _canExecute == null || _canExecute();
        public void Execute(object parameter) => _execute();
    }
}

namespace JSE_RevitAddin_MEP_OPENINGS.UI
{
    // Helper extension for DetermineOrientationFromHost
    public static class CombinedSleeveViewModelHelper
    {
        public static void DetermineOrientationFromHost(ElementId masterId, XYZ basisX, XYZ basisY, Document doc, ref bool useXSpan)
        {
            var masterElement = doc.GetElement(masterId);
            Element host = (masterElement as FamilyInstance)?.Host;

            if (host is Wall w)
            {
                // Wall Orientation is the Normal vector.
                XYZ normal = w.Orientation;

                if (Math.Abs(normal.Y) > Math.Abs(normal.X))
                {
                    useXSpan = true; // Wall runs X
                    DebugLogger.Info($"[ManualJoin] Host is X-Wall (Normal: {normal}). Using Delta X.");
                }
                else
                {
                    useXSpan = false; // Wall runs Y
                    DebugLogger.Info($"[ManualJoin] Host is Y-Wall (Normal: {normal}). Using Delta Y.");
                }
            }
            else
            {
                // Floor or Non-Hosted: Use Master Sleeve Direction
                if (Math.Abs(basisX.X) > Math.Abs(basisX.Y))
                {
                    useXSpan = true;
                    DebugLogger.Info("[ManualJoin] Non-Wall Host. Master aligned to X. Using Delta X.");
                }
                else
                {
                    useXSpan = false;
                    DebugLogger.Info("[ManualJoin] Non-Wall Host. Master aligned to Y. Using Delta Y.");
                }
            }
        }
    }
}

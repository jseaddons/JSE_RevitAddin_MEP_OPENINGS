using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined.Phase1And2.Services;
using FormationService = JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined.Phase3And4.Services.CombinedClusterFormationService;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined.Phase3And4.Services;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined.Phase1And2.Repository;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined.Phase1And2.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined.Phase1And2.Interfaces;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined.Phase3And4.Interfaces;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.BoundingBox;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Rotation;
using JSE_RevitAddin_MEP_OPENINGS.Services.Combined;
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
        
        // Services
        // Services
        private readonly ICombinedClusterDiscoveryService _discoveryService;
        private readonly ICombinedClusterFormation _formationService; // Phase 3
        private readonly IParameterAggregatorService _paramService; // Phase 3
        private readonly ICombinedClusterPersistence _persistService; // Phase 4

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
            ICombinedClusterFormation formationService,
            IParameterAggregatorService paramService,
            ICombinedClusterPersistence persistService,
            IManualClusterCalculationAdapter manualCalculator)
        {
            _uiDocument = uiDoc;
            _document = uiDoc.Document;
            _externalEvent = externalEvent;
            _requestHandler = requestHandler;

            // Injected Dependencies
            _repo = repo ?? throw new ArgumentNullException(nameof(repo));
            _discoveryService = discoveryService ?? throw new ArgumentNullException(nameof(discoveryService));
            _formationService = formationService ?? throw new ArgumentNullException(nameof(formationService));
            _paramService = paramService ?? throw new ArgumentNullException(nameof(paramService));
            _persistService = persistService ?? throw new ArgumentNullException(nameof(persistService));
            _manualCalculator = manualCalculator ?? throw new ArgumentNullException(nameof(manualCalculator));

            _executor = new CrashSafeExecutor();
            
            // Commands
            RefreshCommand = new RelayCommand(Refresh);
            CreateCombinedSleevesCommand = new RelayCommand(CreateCombinedSleeves); // Auto flow
            
            SelectCrossingCommand = new RelayCommand(SelectByCrossing);
            SelectIndividuallyCommand = new RelayCommand(SelectIndividually);
            JoinSelectedCommand = new RelayCommand(JoinSelected);
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
                StatusMessage = "Running Auto-Clustering...";
                // Logic to call Manager?
                // CombinedSleeveManager manager = new CombinedSleeveManager(_document);
                // manager.Execute(true);
                StatusMessage = "Detailed Auto-Cluster UI not fully wired yet.";
                return Autodesk.Revit.UI.Result.Succeeded;
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
            
            _executor.ExecuteWithTimeout(() =>
            {
                // Configure Wrapper Logger (SafeFileLogger wrapped)
                DebugLogger.SetCombinedSleeveLogFile();
                DebugLogger.SetServiceContext("CombinedSleeveManual");
                
                // Force Enable Logging for Diagnostics
                DebugLogger.IsEnabled = true;
                
                DebugLogger.Info("Starting Manual Join sequence...");

                try
                {
                    // Log Build Timestamp to verify latest code
                    var assemblyLocation = System.Reflection.Assembly.GetExecutingAssembly().Location;
                    var buildTime = System.IO.File.GetLastWriteTime(assemblyLocation);
                    DebugLogger.Info($"[CombinedSleeveManual] Build Timestamp: {buildTime}");

                    DebugLogger.Info("[DEBUG] Step 1: About to set StatusMessage");
                    StatusMessage = "Joining...";

                    
                    DebugLogger.Info("[DEBUG] Step 2: Checking _requestHandler");
                    if (_requestHandler == null)
                    {
                        DebugLogger.Error("[ManualJoin] RequestHandler is null");
                        StatusMessage = "Error: RequestHandler not initialized";
                        return Autodesk.Revit.UI.Result.Failed;
                    }
                    DebugLogger.Info("[DEBUG] Step 3: About to call SetAction");
                    // Delegate Transaction to External Event
                    _requestHandler.SetAction((uiapp) =>
                    {
                        DebugLogger.Info("[DEBUG] Step 4: Inside SetAction");
                        Document doc = uiapp.ActiveUIDocument.Document;
                        if (!doc.IsValidObject) return;
                        
                        using (Transaction tx = new Transaction(doc, "Manual Join Combined Sleeve"))
                        {
                            tx.Start();
                            
                            // 1. Identify Master Sleeve (First one selected)
                            ElementId masterId = _selectedIds[0];
                            FamilyInstance masterSleeve = doc.GetElement(masterId) as FamilyInstance;
                            
                            if (masterSleeve == null) 
                            {
                                DebugLogger.Error("Master sleeve is null or not a FamilyInstance.");
                                return;
                            }

                            // 2. Setup Local Coordinate System from Master
                            Transform masterTransform = masterSleeve.GetTransform();
                            Transform inverseTransform = masterTransform.Inverse;
                            


                            // Legacy placeholders for Parameter Aggregation logic
                            double unionMinX = 0, unionMaxX = 0;
                            double unionMinZ = 0, unionMaxZ = 0;
                            List<dynamic> zones = new List<dynamic>(); // Placeholder list

                            var cats = new HashSet<string>();
                            var idsToDelete = new List<ElementId>();

                            DebugLogger.Info($"[ManualJoin] Master Sleeve {masterId} Transform Origin: {masterTransform.Origin}");

                            // 4. ADAPTER-BASED CALCULATION
                            // Using the new ManualClusterCalculationAdapter to ensure 1:1 match with Auto-Cluster logic.
                            // This delegates all geometry analysis (World Space, Rotation, etc.) to the "Golden" services.

                            var selectedElements = _selectedIds.Select(id => doc.GetElement(id)).Where(e => e != null).ToList();
                            
                            DebugLogger.Info($"[ManualJoin] Invoking Adapter for {selectedElements.Count} elements.");
                            
                            var joinResult = _manualCalculator.Calculate(selectedElements);

                            // Apply Results
                            // Apply Results
                            double newWidth = joinResult.WidthFeet;
                            double newHeight = joinResult.HeightFeet;
                            double newDepth = joinResult.DepthFeet;
                            double newRotation = joinResult.RotationAngleRad;
                            XYZ newCenter = joinResult.CenterPoint;
                            
                            // Initialize translation
                            XYZ worldTranslation = XYZ.Zero;

                            // Calculate World Translation
                            // We need to move the Master Sleeve from its current location to this optimal center.
                            
                            // Get Master Center (Try LocationPoint first, then BBox Center)
                            XYZ currentMasterCenter;
                            if (masterSleeve.Location is LocationPoint lp)
                            {
                                currentMasterCenter = lp.Point;
                            }
                            else
                            {
                                var bbox = masterSleeve.get_BoundingBox(null);
                                currentMasterCenter = (bbox.Min + bbox.Max) / 2.0;
                            }

                            // If newCenter is valid (and not origin due to some error), calculate shift.
                            // The Adapter returns world coordinates for the center.
                            if (!newCenter.IsZeroLength())
                            {
                                worldTranslation = newCenter - currentMasterCenter;
                            }
                            
                            DebugLogger.Info($"[ManualJoin] ADAPTER RESULT: Width={newWidth:F4}, Height={newHeight:F4}, Rot={newRotation:F4}, Shift={worldTranslation}");

                            // Log for comparison
                            double widthMM = newWidth * 304.8;
                            double heightMM = newHeight * 304.8;
                            DebugLogger.Info($"[ManualJoin] FINAL CALC: Width={newWidth:F4}ft ({widthMM:F1}mm), Height={newHeight:F4}ft ({heightMM:F1}mm)");

                            // 5. Update Master Parameters (Geometry)
                            var pWidth = masterSleeve.LookupParameter("Width");
                            var pHeight = masterSleeve.LookupParameter("Height");
                            
                            // Check parameter units - Assuming Internal Units (Feet)
                            if (pWidth != null) {
                                pWidth.Set(newWidth);
                                DebugLogger.Info($"[ManualJoin] Set Width Parameter to {newWidth}");
                            }
                            if (pHeight != null) {
                                pHeight.Set(newHeight);
                                DebugLogger.Info($"[ManualJoin] Set Height Parameter to {newHeight}");
                            }
                            
                            // ✅ SET ROTATION (Per Adapter Result)
                            // We need to rotate the element to match the calculated rotation.
                            // First, verify current rotation and rotate difference.
                            // Note: Rotating an Element is tricky. We often need to use ElementTransformUtils.RotateElement.
                            // Assuming typical sleeve is placed point-based with Rotation.
                            // We'll trust the Master's axis if rotation is near-identical, otherwise rotate.
                            
                            // Simple parameter set if it exists? Usually "Angle" or "Rotation" is read-only or doesn't exist.
                            // We must use ElementTransformUtils.RotateElement.
                            
                            // Get Current Rotation (Angle to X-Axis)
                             if (masterSleeve.Location is LocationPoint lpRot)
                             {
                                 // Calculate current rotation
                                 double currentRotation = lpRot.Rotation;
                                 double rotationDiff = newRotation - currentRotation;
                                 
                                 if (Math.Abs(rotationDiff) > 0.001) // Tolerance
                                 {
                                     Line axis = Line.CreateBound(lpRot.Point, lpRot.Point + XYZ.BasisZ);
                                     ElementTransformUtils.RotateElement(doc, masterSleeve.Id, axis, rotationDiff);
                                     DebugLogger.Info($"[ManualJoin] Rotated Master from {currentRotation * 180 / Math.PI:F1} to {newRotation * 180 / Math.PI:F1} degrees.");
                                 }
                             }

                            // 5b. TRANSFER PARAMETERS (Aggregation)
                            try
                            {
                                if (_paramService == null)
                                {
                                    DebugLogger.Warning("[ManualJoin] ParameterAggregatorService is null");
                                    var pComm = masterSleeve.LookupParameter("Comments");
                                    if (pComm != null) pComm.Set($"Manual Join: {string.Join(",", cats)}");
                                    // Skip the rest of parameter aggregation
                                }
                                else
                                {
                                    DebugLogger.Info("[ManualJoin] Starting Parameter Aggregation...");
                                    // Convert selected IDs to ClusterSleeveInfo (Transient) using Live Geometry
                                    // Old 'zones' variable logic replaced by transient creation
                                    var sleeveInfos = new List<ClusterSleeveInfo>();

                                    foreach (var id in _selectedIds)
                                    {
                                        var el = doc.GetElement(id);
                                        if (el == null) continue;

                                        // Create transient info from element
                                        // We don't have full ClashZone data here if it wasn't pre-fetched, but we do our best.
                                        var info = new ClusterSleeveInfo
                                        {
                                            SleeveInstanceId = id.IntegerValue,
                                            CategoryName = el.Category?.Name ?? "Unknown",
                                            // Approximate placement point (center of bbox)
                                            SleeveCenter = el.get_BoundingBox(null) is BoundingBoxXYZ b ? (b.Min + b.Max) / 2 : XYZ.Zero
                                        };
                                        sleeveInfos.Add(info);
                                    }

                                    // Create Dummy Candidate
                                    var candidate = new CombinedClusterCandidate(0, sleeveInfos);

                                    // Set Bounds (Using our new LOCAL bounds as proxy for structure)
                                    // Although these are Local, they represent the extent.
                                    // NOTE: Multi-Replace tool might need 'unionMinX' etc defined if we want to strictly match old logic, 
                                    // but here we pass the Calculated Dimensions which is what matters.

                                    // For Parameter Aggregation, the exact World Coords matter less than the *Sets* of parameters.
                                    // We populate dummy bounds to avoid null refs.
                                    candidate.CombinedBoundingBoxMinX = 0;
                                    candidate.CombinedBoundingBoxMinY = 0;
                                    candidate.CombinedBoundingBoxMinZ = 0;
                                    candidate.CombinedBoundingBoxMaxX = newWidth; // Use dimensions
                                    candidate.CombinedBoundingBoxMaxY = newDepth;
                                    candidate.CombinedBoundingBoxMaxZ = newHeight;

                                    // Generate Parameter Set (with null check)
                                    if (_paramService != null)
                                    {
                                        var paramSet = _paramService.CreateCombinedSleeveParameterSet(candidate, candidate.CombinedBoundingBox);

                                        // Apply to Master Sleeve
                                        foreach (var kvp in paramSet)
                                        {
                                            // FILTER: Do not overwrite Geometry parameters (Width, Height Set above. Depth/Length controlled by Host)
                                            // User Issue: "depth of sleeve is very big... depth of sleeve is always controlled by structural thickness"
                                            // The Aggregator calculates Depth from BoundingBox Y-Diff, which is wrong for Manual Join AABB logic.
                                            if (kvp.Key == "Width" || kvp.Key == "Height" || kvp.Key == "Depth" || kvp.Key == "Length")
                                            {
                                                continue;
                                            }

                                            var p = masterSleeve.LookupParameter(kvp.Key);
                                            if (p != null && !p.IsReadOnly)
                                            {
                                                if (kvp.Value is double d) p.Set(d);
                                                else if (kvp.Value is string s) p.Set(s);
                                                else if (kvp.Value is int i) p.Set(i);
                                            }
                                        }
                                        DebugLogger.Info($"[ManualJoin] Applied {paramSet.Count} aggregated parameters.");

                                        // 5c. SET NEW METADATA: "Combined Sleeve Instance ID" (Case Insensitive)
                                        // User request: "store in that as this si not cluster instance id but combined sleeve instance id"

                                        Parameter pCombinedId = masterSleeve.LookupParameter("Combined Sleeve Instance ID");
                                        if (pCombinedId == null)
                                        {
                                            // Try Case-Insensitive Search
                                            foreach (Parameter p in masterSleeve.Parameters)
                                            {
                                                if (string.Equals(p.Definition.Name, "Combined Sleeve Instance ID", StringComparison.OrdinalIgnoreCase))
                                                {
                                                    pCombinedId = p;
                                                    break;
                                                }
                                            }
                                        }

                                        if (pCombinedId != null && !pCombinedId.IsReadOnly)
                                        {
                                            bool result = false;
                                            if (pCombinedId.StorageType == StorageType.Integer)
                                            {
                                                result = pCombinedId.Set(masterId.IntegerValue);
                                            }
                                            else if (pCombinedId.StorageType == StorageType.String)
                                            {
                                                result = pCombinedId.Set(masterId.ToString());
                                            }
                                            else if (pCombinedId.StorageType == StorageType.Double)
                                            {
                                                result = pCombinedId.Set((double)masterId.IntegerValue);
                                            }

                                            if (result)
                                                DebugLogger.Info($"[ManualJoin] Set Combined Sleeve Instance ID: {masterId} (Type: {pCombinedId.StorageType})");
                                            else
                                                DebugLogger.Warning($"[ManualJoin] Failed to set Combined Sleeve Instance ID: {masterId} (Type: {pCombinedId.StorageType})");
                                        }
                                        else
                                        {
                                            DebugLogger.Warning("[ManualJoin] Parameter 'Combined Sleeve Instance ID' not found on family (Case Insensitive search).");
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                DebugLogger.Error($"[ManualJoin] Parameter Aggregation Failed: {ex.Message}");
                                // Fallback to simple comments if Aggregation fails?
                                var pComm = masterSleeve.LookupParameter("Comments");
                                if (pComm != null) pComm.Set($"Manual Join: {string.Join(",", cats)} (Agg Failed)");
                            }


                            // 6. Move Master to new Center
                            if (!worldTranslation.IsZeroLength())
                            {
                                ElementTransformUtils.MoveElement(doc, masterId, worldTranslation);
                            }

                            // 7. Delete other sleeves (Moved here to ensure safety)
                            if (idsToDelete.Count > 0)
                            {
                                doc.Delete(idsToDelete);
                            }

                            tx.Commit();
                            
                            DebugLogger.Info($"Manual Join Success. Kept {masterId}, deleted {idsToDelete.Count} sleeves.");
                            
                            _selectedIds.Clear();
                            SelectedSleevesList.Clear();
                            OnPropertyChanged(nameof(CanJoin));
                            StatusMessage = "Join complete.";
                        }
                    });
                    // Ensure the UI is closed before raising the external event
                    HideRequest?.Invoke(); // Hide/close the dialog first
                    // Use dispatcher to delay Raise until after UI is closed
                    // Ensure the UI is closed before raising the external event
                                   
                    // Check if WPF Application is available
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
                        // Fallback: Raise directly
                        DebugLogger.Warning("[ManualJoin] Application.Current is null, raising directly");
                        _externalEvent.Raise();
                    }
                    return Autodesk.Revit.UI.Result.Succeeded;
                }
                catch (Exception ex)
                {
                    StatusMessage = "Join failed: " + ex.Message;
                    DebugLogger.Error("Manual Join Failed: " + ex.ToString());
                    return Autodesk.Revit.UI.Result.Failed;
                }
            }, "Manual Join");
        }







        
        private void CreateCombinedSleeve(CombinedClusterCandidate candidate)
        {
             // Full logic implementation using services
             // Combined clustering creates openings; dampers placed by dedicated strategy
             string familyName = "RectangularOpeningOnWall"; // Can logic this
             
             // 1. Get Params
             var paramSet = _paramService.CreateCombinedSleeveParameterSet(candidate, candidate.CombinedBoundingBox);
             
             // 2. Load Symbol
             // Use FilteredElementCollector inside Orchestrator or here
             // Duplicated logic for now for command independence
             FamilySymbol symbol = new FilteredElementCollector(_document)
                    .OfClass(typeof(FamilySymbol))
                    .Cast<FamilySymbol>()
                    .FirstOrDefault(x => x.Name.Equals(familyName, StringComparison.OrdinalIgnoreCase) || 
                                         x.Family.Name.Equals(familyName, StringComparison.OrdinalIgnoreCase));
                                         
             if (symbol != null)
             {
                 if(!symbol.IsActive) symbol.Activate();
                 var instance = _document.Create.NewFamilyInstance(candidate.CombinedBoundingBox.Min, symbol, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                 
                 // Apply params
                 foreach(var kvp in paramSet)
                 {
                     var p = instance.LookupParameter(kvp.Key);
                     if(p != null)
                     {
                         if(kvp.Value is double d) p.Set(d);
                         else if(kvp.Value is string s) p.Set(s);
                         else if(kvp.Value is int i) p.Set(i);
                     }
                 }
                 
                 // 3. Persist
                 _persistService.PersistCombinedCluster(candidate, instance.Id.IntegerValue);
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
        public bool AllowReference(Reference reference, XYZ position) => false;
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

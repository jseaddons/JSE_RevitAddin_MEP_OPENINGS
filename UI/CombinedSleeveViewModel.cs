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
        private readonly CombinedClusterDiscoveryService _discoveryService;
        private readonly FormationService _formationService; // Phase 3
        private readonly ParameterAggregatorService _paramService; // Phase 3
        private readonly CombinedClusterPersistenceService _persistService; // Phase 4

        private readonly ClashZoneRepository _repo;
        private readonly CrashSafeExecutor _executor; // Safety wrapper
        
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

        public CombinedSleeveViewModel(Autodesk.Revit.UI.UIDocument uiDoc, Autodesk.Revit.UI.ExternalEvent externalEvent, CombinedSleeveRequestHandler requestHandler)
        {
            _uiDocument = uiDoc;
            _document = uiDoc.Document;
            _externalEvent = externalEvent;
            _requestHandler = requestHandler;

            _executor = new CrashSafeExecutor();
            
            // Initialize Services
            // Note: In real app, these should be injected or created via factory.
            // Simplified here for direct instantiation if context allows.
            
            // Creating Context for Repositories
            // Using shared context potentially
            // BUT here we create transient context or use one passed in?
            // Assuming simplified usage:
            try 
            {
                 // Using a dedicated context for VM operations if needed, but Manual Join uses Transaction context
                 var dbContext = new SleeveDbContext(_document); // Should be disposed?
                 _repo = new ClashZoneRepository(dbContext);
                 var combinedRepo = new CombinedClusterRepository(_repo);
                 
                 _discoveryService = new CombinedClusterDiscoveryService(combinedRepo);
                 _formationService = new FormationService(combinedRepo);
                 
                 _paramService = new ParameterAggregatorService();
                 _persistService = new CombinedClusterPersistenceService(_repo);
            }
            catch(Exception ex)
            {
                StatusMessage = "Service Init Error: " + ex.Message;
            }

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
                    StatusMessage = "Joining...";
                    
                    // Delegate Transaction to External Event
                    _requestHandler.SetAction((uiapp) =>
                    {
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
                            
                            // Initialize variables for outer scope (Fixing CS0103)
                            double newWidth = 0;
                            double newHeight = 0;
                            XYZ worldTranslation = XYZ.Zero;

                            // Legacy placeholders for Parameter Aggregation logic
                            double unionMinX = 0, unionMaxX = 0;
                            double unionMinZ = 0, unionMaxZ = 0;
                            List<dynamic> zones = new List<dynamic>(); // Placeholder list

                            var cats = new HashSet<string>();
                            var idsToDelete = new List<ElementId>();

                            DebugLogger.Info($"[ManualJoin] Master Sleeve {masterId} Transform Origin: {masterTransform.Origin}");

                            // 4. ROBUST "CLUSTER-MIMIC" LOGIC (Inverse Transform Strategy)
                            // This replicates the Auto-Cluster logic: "Rotate to Local -> Measure -> Rotate Back".
                            // We use MasterTransform.Inverse to put everything in the Master's aligned local space.
                            
                            // Bounds in Local Space
                            double localMinX = double.MaxValue, localMaxX = double.MinValue;
                            double localMinY = double.MaxValue, localMaxY = double.MinValue; // Depth/Thickness
                            double localMinZ = double.MaxValue, localMaxZ = double.MinValue; // Height

                            DebugLogger.Info($"[ManualJoin] Starting Cluster-Mimic Calculation using Master {masterId} Transform.");

                            foreach (var id in _selectedIds)
                            {
                                var el = doc.GetElement(id);
                                if(el == null) continue;
                                
                                BoundingBoxXYZ bb = el.get_BoundingBox(null);
                                if (bb == null) continue;

                                var corners = GetBoundingBoxCorners(bb);
                                foreach (var p in corners)
                                {
                                    // Transform World Point -> Master Local Space
                                    XYZ localP = inverseTransform.OfPoint(p);
                                    
                                    if (localP.X < localMinX) localMinX = localP.X;
                                    if (localP.X > localMaxX) localMaxX = localP.X;
                                    
                                    if (localP.Y < localMinY) localMinY = localP.Y;
                                    if (localP.Y > localMaxY) localMaxY = localP.Y;
                                    
                                    if (localP.Z < localMinZ) localMinZ = localP.Z;
                                    if (localP.Z > localMaxZ) localMaxZ = localP.Z;
                                }
                                
                                if (id != masterId) idsToDelete.Add(id);
                            }

                            // SIZING (Local Space)
                            // Width is simply the span along Local X
                            newWidth = localMaxX - localMinX;
                            
                            // Height is the span along Local Z
                            newHeight = localMaxZ - localMinZ;
                            
                            // Depth (Thickness) - We calculate it but usually preserve Master's Thickness logic 
                            // unless we want to resize thickness. Standard Sleeves usually fixed thickness by Host.
                            // But for Generic Models, we might want to update it.
                            // Let's log it for now.
                            double newDepth = localMaxY - localMinY;

                            // ROUNDING LOGIC (Nearest Millimeter)
                            double widthMM_calc = Math.Round(newWidth * 304.8);
                            double heightMM_calc = Math.Round(newHeight * 304.8);
                            
                            newWidth = widthMM_calc / 304.8;
                            newHeight = heightMM_calc / 304.8;

                            DebugLogger.Info($"[ManualJoin] CLUSTER-MIMIC CALC: Width={newWidth:F4}ft ({widthMM_calc}mm), Height={newHeight:F4}ft ({heightMM_calc}mm), Depth={newDepth:F4}ft");

                            // PLACEMENT (Local -> World)
                            // We find the center of the Union box in Local Space.
                            double localCenterX = (localMinX + localMaxX) / 2.0;
                            double localCenterZ = (localMinZ + localMaxZ) / 2.0;
                            // For Y (Thickness), we usually stay centered on the Master's plane (Y=0) if it's a Wall.
                            // However, if we are joining depth-wise (thick wall), we might want the new center.
                            // Mimicking Auto-Cluster: It calculates a "Theoretical Midpoint".
                            double localCenterY = (localMinY + localMaxY) / 2.0;
                            
                            // Construct Local Center Point
                            // Note: Using 'localCenterY' correctly centers it in the combined thickness.
                            // This handles off-center joins.
                            XYZ localCenterPoint = new XYZ(localCenterX, localCenterY, localCenterZ);
                            
                            // Transform Local Center -> World Space
                            XYZ worldCenterPoint = masterTransform.OfPoint(localCenterPoint);
                            
                            // Calculate Shift
                            worldTranslation = worldCenterPoint - masterTransform.Origin;
                            
                            DebugLogger.Info($"[ManualJoin] PLACEMENT: LocalCenter={localCenterPoint}, WorldShift={worldTranslation}");

                            // BYPASSING Previous Logic Blocks
                            // This replaces the entire "DB Loop" and "Orientation Checks".
                            // We trust the Master's coordinate system implicitly.

                            // Convert to MM for debug comparison
                            double widthMM = newWidth * 304.8;
                            double heightMM = newHeight * 304.8;
                            
                            // EMERGENCY LOG - TRACE 1
                            string tracePath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "JSE_ManualJoin_Trace.txt");
                            try {
                                System.IO.File.AppendAllText(tracePath, $"[{DateTime.Now}] ManualJoin Calc Start. Master:{masterId} ZonesFound:{zones?.Count ?? 0}\n");
                            } catch {}

                            DebugLogger.Info($"[ManualJoin] FINAL CALC: Width={newWidth:F4}ft ({widthMM:F1}mm), Height={newHeight:F4}ft ({heightMM:F1}mm)");
                            
                            // FORCE LOG regardless of settings
                            SafeFileLogger.SafeAppendTextAlways("combinesleeveplacer.log", $"[ManualJoin] FINAL CALC: Width={newWidth:F4}ft ({widthMM:F1}mm), Height={newHeight:F4}ft ({heightMM:F1}mm)");
                            
                            try {
                                System.IO.File.AppendAllText(tracePath, 
                                    $"[{DateTime.Now}] Size:{widthMM:F1}x{heightMM:F1}mm. WorldShift:{worldTranslation}\n");
                            } catch {}

                            DebugLogger.Info($"[ManualJoin] World Shift: {worldTranslation}");

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

                            // 5b. TRANSFER PARAMETERS (Aggregation)
                            try 
                            {
                                DebugLogger.Info("[ManualJoin] Starting Parameter Aggregation...");
                                // Convert selected IDs to ClusterSleeveInfo (Transient) using Live Geometry
                                // Old 'zones' variable logic replaced by transient creation
                                var sleeveInfos = new List<ClusterSleeveInfo>();
                                
                                foreach(var id in _selectedIds)
                                {
                                    var el = doc.GetElement(id);
                                    if(el == null) continue;
                                    
                                    // Create transient info from element
                                    // We don't have full ClashZone data here if it wasn't pre-fetched, but we do our best.
                                    var info = new ClusterSleeveInfo
                                    {
                                        SleeveInstanceId = id.IntegerValue,
                                        CategoryName = el.Category?.Name ?? "Unknown",
                                        // Approximate placement point (center of bbox)
                                        SleeveCenter = el.get_BoundingBox(null) is BoundingBoxXYZ b ? (b.Min + b.Max)/2 : XYZ.Zero
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

                                // Generate Parameter Set
                                var paramSet = _paramService.CreateCombinedSleeveParameterSet(candidate, candidate.CombinedBoundingBox);
                                
                                // Apply to Master Sleeve
                                foreach(var kvp in paramSet)
                                {
                                    // FILTER: Do not overwrite Geometry parameters (Width, Height Set above. Depth/Length controlled by Host)
                                    // User Issue: "depth of sleeve is very big... depth of sleeve is always controlled by structural thickness"
                                    // The Aggregator calculates Depth from BoundingBox Y-Diff, which is wrong for Manual Join AABB logic.
                                    if(kvp.Key == "Width" || kvp.Key == "Height" || kvp.Key == "Depth" || kvp.Key == "Length")
                                    {
                                        continue; 
                                    }

                                    var p = masterSleeve.LookupParameter(kvp.Key);
                                    if(p != null && !p.IsReadOnly)
                                    {
                                        if(kvp.Value is double d) p.Set(d);
                                        else if(kvp.Value is string s) p.Set(s);
                                        else if(kvp.Value is int i) p.Set(i);
                                    }
                                }
                                DebugLogger.Info($"[ManualJoin] Applied {paramSet.Count} aggregated parameters.");

                                // 5c. SET NEW METADATA: "Combined Sleeve Instance ID" (Case Insensitive)
                                // User request: "store in that as this si not cluster instance id but combined sleeve instance id"
                                
                                Parameter pCombinedId = masterSleeve.LookupParameter("Combined Sleeve Instance ID");
                                if (pCombinedId == null)
                                {
                                    // Try Case-Insensitive Search
                                    foreach(Parameter p in masterSleeve.Parameters)
                                    {
                                        if (string.Equals(p.Definition.Name, "Combined Sleeve Instance ID", StringComparison.OrdinalIgnoreCase))
                                        {
                                            pCombinedId = p;
                                            break;
                                        }
                                    }
                                }
                                
                                if(pCombinedId != null && !pCombinedId.IsReadOnly)
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
                                    
                                    if(result)
                                        DebugLogger.Info($"[ManualJoin] Set Combined Sleeve Instance ID: {masterId} (Type: {pCombinedId.StorageType})");
                                    else
                                        DebugLogger.Warning($"[ManualJoin] Failed to set Combined Sleeve Instance ID: {masterId} (Type: {pCombinedId.StorageType})");
                                }
                                else
                                {
                                    DebugLogger.Warning("[ManualJoin] Parameter 'Combined Sleeve Instance ID' not found on family (Case Insensitive search).");
                                }
                            }
                            catch (Exception ex)
                            {
                                DebugLogger.Error($"[ManualJoin] Parameter Aggregation Failed: {ex.Message}");
                                // Fallback to simple comments if Aggregation fails?
                                var pComm = masterSleeve.LookupParameter("Comments");
                                if(pComm != null) pComm.Set($"Manual Join: {string.Join(",", cats)} (Agg Failed)");
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
                    System.Windows.Application.Current.Dispatcher.BeginInvoke(new Action(() => {
                        try {
                            _externalEvent.Raise();
                        } catch (Exception ex) {
                            DebugLogger.Error("ExternalEvent.Raise() failed: " + ex.ToString());
                        }
                    }), System.Windows.Threading.DispatcherPriority.Background);
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

        private IEnumerable<XYZ> GetBoundingBoxCorners(BoundingBoxXYZ box)
        {
            yield return box.Min;
            yield return box.Max;
            yield return new XYZ(box.Max.X, box.Min.Y, box.Min.Z);
            yield return new XYZ(box.Min.X, box.Max.Y, box.Min.Z);
            yield return new XYZ(box.Min.X, box.Min.Y, box.Max.Z);
            yield return new XYZ(box.Max.X, box.Max.Y, box.Min.Z); // etc.. simplified needed?
            // Just all 8 permutations
            yield return new XYZ(box.Min.X, box.Max.Y, box.Max.Z);
            yield return new XYZ(box.Max.X, box.Min.Y, box.Max.Z);
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

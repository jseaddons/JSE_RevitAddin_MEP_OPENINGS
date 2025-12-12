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
        
        // External Event for Modeless Transaction Safety
        private readonly Autodesk.Revit.UI.ExternalEvent _externalEvent;
        private readonly CombinedSleeveRequestHandler _requestHandler;

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));

        // Window Control Actions
        public Action HideRequest { get; set; }
        public Action ShowRequest { get; set; }

        public CombinedSleeveViewModel(Autodesk.Revit.UI.UIDocument uiDocument)
        {
            _uiDocument = uiDocument;
            _document = uiDocument.Document;
            SelectedSleevesList = new ObservableCollection<string>();
            _selectedIds = new List<ElementId>();

            // Feature 28: Initialize Services with existing Infrastructure
            // Initialize DbContext with the current document
            var dbContext = new SleeveDbContext(_document); 
            _repo = new ClashZoneRepository(dbContext);
            _executor = new CrashSafeExecutor(); // Init safety executor 
            
            // Interfaces for new services
            var combinedRepo = new CombinedClusterRepository(_repo);
            
            // Initialize External Event Handler
            _requestHandler = new CombinedSleeveRequestHandler();
            _externalEvent = Autodesk.Revit.UI.ExternalEvent.Create(_requestHandler);
             
            // Services
            // Passing combinedRepo to DiscoveryService
            _discoveryService = new CombinedClusterDiscoveryService(combinedRepo);

            _formationService = new FormationService(combinedRepo); 
            _paramService = new ParameterAggregatorService();
            _persistService = new CombinedClusterPersistenceService(_repo);

            // Commands
            RunAutoCommand = new RelayCommand(RunAuto);
            SelectByCrossingCommand = new RelayCommand(SelectByCrossing);
            SelectIndividuallyCommand = new RelayCommand(SelectIndividually);
            JoinSelectedCommand = new RelayCommand(JoinSelected);
        }

        // Properties
        public bool IsAutoMode
        {
            get => _isAutoMode;
            set
            {
                _isAutoMode = value;
                OnPropertyChanged(nameof(IsAutoMode));
                OnPropertyChanged(nameof(IsManualMode));
                OnPropertyChanged(nameof(IsAutoModeVisibility));
                OnPropertyChanged(nameof(IsManualModeVisibility));
            }
        }
        public bool IsManualMode { get => !_isAutoMode; set => IsAutoMode = !value; }
        
        public Visibility IsAutoModeVisibility => _isAutoMode ? Visibility.Visible : Visibility.Collapsed;
        public Visibility IsManualModeVisibility => !_isAutoMode ? Visibility.Visible : Visibility.Collapsed;

        public bool IncludeDucts { get; set; } = true;
        public bool IncludePipes { get; set; } = true;
        public bool IncludeCableTrays { get; set; } = true;
        public bool IncludeConduits { get; set; } = true;

        public ObservableCollection<string> SelectedSleevesList { get; set; }
        private List<ElementId> _selectedIds; 

        public bool CanJoin => SelectedSleevesList.Count > 1;

        public string StatusMessage
        {
            get => _statusMessage;
            set { _statusMessage = value; OnPropertyChanged(nameof(StatusMessage)); }
        }

        // Commands
        public ICommand RunAutoCommand { get; }
        public ICommand SelectByCrossingCommand { get; }
        public ICommand SelectIndividuallyCommand { get; }
        public ICommand JoinSelectedCommand { get; }

        // --- Logic Methods ---

        private void RunAuto()
        {
            _executor.ExecuteWithTimeout(() =>
            {
                try
                {
                    StatusMessage = "Detecting candidates...";
                    
                    var categories = new List<string>();
                    if (IncludeDucts) 
                    {
                        categories.Add("Ducts");
                        categories.Add("Duct Accessories");
                    }
                    if (IncludePipes) categories.Add("Pipes");
                    if (IncludeCableTrays) categories.Add("Cable Trays"); 
                    if (IncludeConduits) categories.Add("Conduits");

                    if (categories.Count == 0)
                    {
                        StatusMessage = "Select at least one category.";
                        return Autodesk.Revit.UI.Result.Cancelled;
                    }

                    // Call Discover with filter name (empty or specific) and categories
                    var candidatesList = _discoveryService.Discover("CombinedDiscovery", categories);
                    var candidates = candidatesList.ToList();
                    
                    // Call Formation (Async, so use .Result or wait)
                    var clusters = _formationService.FormCombinedClustersAsync(candidates).Result;
                    
                    if (!clusters.Any())
                    {
                        StatusMessage = "No combined clusters found.";
                        return Autodesk.Revit.UI.Result.Succeeded;
                    }

                    StatusMessage = $"Found {clusters.Count} clusters. Creating...";

                    // Configure Logger
                    DebugLogger.SetCombinedSleeveLogFile();
                    DebugLogger.SetServiceContext("CombinedSleeveAuto");
                    if (OptimizationFlags.UseDiagnosticMode) DebugLogger.IsEnabled = true;

                    // Delegate Transaction to External Event
                    _requestHandler.SetAction((uiapp) =>
                    {
                        Document doc = uiapp.ActiveUIDocument.Document;
                        if (!doc.IsValidObject) return;
                        
                        using (Transaction tx = new Transaction(doc, "Auto-Create Combined Sleeves"))
                        {
                            tx.Start();
                            
                            int count = 0;
                            foreach (var cluster in clusters)
                            {
                                CreateCombinedSleeve(cluster);
                                count++;
                            }
                            
                            tx.Commit();
                            StatusMessage = $"Created {count} combined sleeves.";
                        }
                    });
                    _externalEvent.Raise();
                    return Autodesk.Revit.UI.Result.Succeeded;
                }
                catch (Exception ex)
                {
                    StatusMessage = "Error: " + ex.Message;
                    DebugLogger.Error("Auto-Combine Error: " + ex.ToString());
                    return Autodesk.Revit.UI.Result.Failed;
                }
            }, "Auto-Combine Sleeves");
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

                            double minX = double.MaxValue, maxX = double.MinValue;
                            double minZ = double.MaxValue, maxZ = double.MinValue; // Z is Height
                            // Track Y (Thickness) just in case
                            
                            var cats = new HashSet<string>();
                            var idsToDelete = new List<ElementId>();

                            DebugLogger.Info($"[ManualJoin] Master Sleeve {masterId} Transform Origin: {masterTransform.Origin}");

                            // 3. Process all sleeves to find Union in Local Coordinates
                            foreach (var id in _selectedIds)
                            {
                                var el = doc.GetElement(id);
                                if(el == null) continue;
                                
                                cats.Add(el.Category.Name);

                                // Get Bounds
                                BoundingBoxXYZ bb = el.get_BoundingBox(null);
                                if (bb == null) continue;

                                DebugLogger.Info($"[ManualJoin] Processing {id} ({el.Category.Name}). World Bounds: {bb.Min} to {bb.Max}");

                                var corners = GetBoundingBoxCorners(bb);
                                foreach (var p in corners)
                                {
                                    XYZ localP = inverseTransform.OfPoint(p);
                                    if (localP.X < minX) minX = localP.X;
                                    if (localP.X > maxX) maxX = localP.X;
                                    
                                    // Z is usually Height in Wall families
                                    if (localP.Z < minZ) minZ = localP.Z;
                                    if (localP.Z > maxZ) maxZ = localP.Z;
                                }

                                if (id != masterId) idsToDelete.Add(id);
                            }
                            
                            // Log Calculated Local Bounds
                            DebugLogger.Info($"[ManualJoin] Calculated Local Bounds - X: [{minX:F4}, {maxX:F4}], Z: [{minZ:F4}, {maxZ:F4}]");

                            // 4. Calculate New Size and Offset using DB Data if available
                            double newWidth, newHeight, centerX, centerZ;
                            XYZ worldTranslation;
                            
                            // Define Union variables in outer scope for later use
                            double unionMinX = double.MaxValue, unionMinZ = double.MaxValue, unionMinY = double.MaxValue;
                            double unionMaxX = double.MinValue, unionMaxZ = double.MinValue, unionMaxY = double.MinValue;
                            
                            // Try to get DB records
                            var zones = _repo.GetClashZonesBySleeveIds(_selectedIds.Select(id => id.IntegerValue));
                            
                            // FALLBACK: If DB is empty (e.g. Manual Sleeves not in DB), create Transient Zones from Live Geometry
                            if (zones == null || zones.Count == 0)
                            {
                                DebugLogger.Warning($"[ManualJoin] DB lookup failed. Creating transient zones from Live Geometry (World Axis Aligned).");
                                zones = new List<ClashZone>();
                                foreach(var id in _selectedIds)
                                {
                                    var el = doc.GetElement(id);
                                    if(el == null) continue;
                                    var bb = el.get_BoundingBox(null);
                                    if(bb == null) continue;
                                    
                                    // Construct Transient Zone
                                    var z = new ClashZone();
                                    z.SleeveInstanceId = id.IntegerValue;
                                    z.SleeveWidth = bb.Max.X - bb.Min.X; // Approx
                                    z.SleeveHeight = bb.Max.Z - bb.Min.Z; // Approx
                                    
                                    // Store World Bounds into Cluster Bounds fields (since logic uses them)
                                    z.ClusterSleeveBoundingBoxMinX = bb.Min.X;
                                    z.ClusterSleeveBoundingBoxMinY = bb.Min.Y; 
                                    z.ClusterSleeveBoundingBoxMinZ = bb.Min.Z;
                                    z.ClusterSleeveBoundingBoxMaxX = bb.Max.X;
                                    z.ClusterSleeveBoundingBoxMaxY = bb.Max.Y;
                                    z.ClusterSleeveBoundingBoxMaxZ = bb.Max.Z;
                                    
                                    // Center
                                    z.SleevePlacementPointX = (bb.Min.X + bb.Max.X)/2.0;
                                    z.SleevePlacementPointY = (bb.Min.Y + bb.Max.Y)/2.0;
                                    z.SleevePlacementPointZ = (bb.Min.Z + bb.Max.Z)/2.0;
                                    
                                    zones.Add(z);
                                }
                            }
                            
                            if (zones != null && zones.Count > 0)
                            {
                                DebugLogger.Info($"[ManualJoin] DB/Live Mode: Using {zones.Count} records.");
                                
                                // Map to ClusterSleeveInfo (Simplified)
                                // Variables already defined above
                                
                                foreach (var z in zones)
                                {
                                    // Use Cluster Bounds (which we populated from Live or DB)
                                    // Logic: CombinedClusterFormationService uses these exact fields.
                                    
                                    double zMinX, zMaxX, zMinZ, zMaxZ;
                                    
                                    if (z.ClusterSleeveBoundingBoxMaxX > z.ClusterSleeveBoundingBoxMinX) 
                                    {
                                        zMinX = z.ClusterSleeveBoundingBoxMinX;
                                        zMaxX = z.ClusterSleeveBoundingBoxMaxX;
                                        zMinZ = z.ClusterSleeveBoundingBoxMinZ;
                                        zMaxZ = z.ClusterSleeveBoundingBoxMaxZ;
                                    }
                                    else
                                    {
                                        // Fallback A (DB record exists but no cluster bounds): Use Placement +/- Width
                                        double halfW = (z.SleeveWidth > 0 ? z.SleeveWidth : 1.0) / 2.0;
                                        double halfH = (z.SleeveHeight > 0 ? z.SleeveHeight : 1.0) / 2.0;
                                        double pX = z.SleevePlacementPointX;
                                        double pZ = z.SleevePlacementPointZ;
                                        zMinX = pX - halfW; zMaxX = pX + halfW;
                                        zMinZ = pZ - halfH; zMaxZ = pZ + halfH;
                                    }
                                    
                                    unionMinX = Math.Min(unionMinX, zMinX);
                                    unionMaxX = Math.Max(unionMaxX, zMaxX);
                                    unionMinZ = Math.Min(unionMinZ, zMinZ);
                                    unionMaxZ = Math.Max(unionMaxZ, zMaxZ);
                                }
                                
                                // Center in World (DB Coordinates)
                                double centerWorldX = (unionMinX + unionMaxX) / 2.0;
                                double centerWorldZ = (unionMinZ + unionMaxZ) / 2.0; // Z is Height
                                double centerWorldY = (zones.Min(z => z.SleevePlacementPointY - (z.SleeveWidth/2)) + zones.Max(z => z.SleevePlacementPointY + (z.SleeveWidth/2))) / 2.0; 
                                // Note: Y calc above is rough approx if BBox not stored. 
                                // Better:
                                if(unionMinY == double.MaxValue) 
                                {
                                     // Re-scan for Y if not populated (e.g. only X/Z usage in previous loop?)
                                     // In previous loop we only did X/Z? 
                                     // Wait, I need to check if I populated unionMinY in the previous loop.
                                     // I added definitions for X/Z but not Y in the previous scope fix!
                                     // I need to add Y to definitions and loop.
                                }
                                
                                // RE-DOING LOOP CAREFULLY TO GET Y-BOUNDS TO SUPPORT Y-WALLS
                                unionMinY = double.MaxValue; unionMaxY = double.MinValue;
                                foreach (var z in zones)
                                {
                                     double zMinY, zMaxY;
                                     if(z.ClusterSleeveBoundingBoxMaxY > z.ClusterSleeveBoundingBoxMinY)
                                     {
                                         zMinY = z.ClusterSleeveBoundingBoxMinY;
                                         zMaxY = z.ClusterSleeveBoundingBoxMaxY;
                                     }
                                     else
                                     {
                                         // Fallback Using Placement
                                         // Assume Width applies to X, Length/Depth to Y? 
                                         // OR depends on orientation in DB?
                                         // If DB stored "SleeveWidth", it's the parameter value.
                                         // Use simple point expansion approx?
                                         zMinY = z.SleevePlacementPointY - 0.5; // Arbitrary 1ft spread if unknown
                                         zMaxY = z.SleevePlacementPointY + 0.5;
                                     }
                                     unionMinY = Math.Min(unionMinY, zMinY);
                                     unionMaxY = Math.Max(unionMaxY, zMaxY);
                                }
                                centerWorldY = (unionMinY + unionMaxY) / 2.0;

                                // For translation, we need shift relative to Master.
                                var masterZone = zones.FirstOrDefault(z => z.SleeveInstanceId == masterId.IntegerValue);
                                XYZ masterPoint = null;
                                
                                if (masterZone != null) 
                                    masterPoint = new XYZ(masterZone.SleevePlacementPointX, masterZone.SleevePlacementPointY, masterZone.SleevePlacementPointZ);
                                else
                                    masterPoint = masterTransform.Origin; // Fallback to live

                                // ORIENTATION CHECK & SIZING
                                XYZ basisX = masterTransform.BasisX; // Width Direction
                                XYZ basisY = masterTransform.BasisY; // Thickness Direction
                                XYZ basisZ = masterTransform.BasisZ; // Height Direction (usually 0,0,1)

                                bool useXSpan = true; // Default to X-Span

                                // 1. Determine Wall Orientation
                                // Priority 1: Use Stored DB Info (User Request: "use the dump once")
                                if (masterZone != null && !string.IsNullOrEmpty(masterZone.MepElementOrientationDirection))
                                {
                                    if (masterZone.MepElementOrientationDirection.Equals("X", StringComparison.OrdinalIgnoreCase))
                                    {
                                        useXSpan = true; // DB says X-Oriented (Wall runs X)
                                        DebugLogger.Info($"[ManualJoin] Using Stored DB Orientation: X. Using Delta X.");
                                    }
                                    else if (masterZone.MepElementOrientationDirection.Equals("Y", StringComparison.OrdinalIgnoreCase))
                                    {
                                        useXSpan = false; // DB says Y-Oriented (Wall runs Y)
                                        DebugLogger.Info($"[ManualJoin] Using Stored DB Orientation: Y. Using Delta Y.");
                                    }
                                    else
                                    {
                                        // Specific handling for Floor/Slab if stored? 
                                        // Fallback to Host Logic below.
                                        DebugLogger.Info($"[ManualJoin] Stored Orientation '{masterZone.MepElementOrientationDirection}' not X/Y. Falling back to Host check.");
                                        CombinedSleeveViewModelHelper.DetermineOrientationFromHost(masterId, masterTransform.BasisX, masterTransform.BasisY, doc, ref useXSpan);
                                    }
                                }
                                else
                                {
                                    // Priority 2: Live Host Check (Fallback)
                                    CombinedSleeveViewModelHelper.DetermineOrientationFromHost(masterId, masterTransform.BasisX, masterTransform.BasisY, doc, ref useXSpan);
                                }

                                if (useXSpan)
                                {
                                    newWidth = unionMaxX - unionMinX;
                                }
                                else
                                {
                                    newWidth = unionMaxY - unionMinY;
                                }
                                newHeight = unionMaxZ - unionMinZ;
                                
                                // ROUNDING LOGIC (Nearest Millimeter)
                                // Convert to mm, round, convert back.
                                double widthMM_calc = Math.Round(newWidth * 304.8);
                                double heightMM_calc = Math.Round(newHeight * 304.8);
                                
                                newWidth = widthMM_calc / 304.8;
                                newHeight = heightMM_calc / 304.8;

                                // 2. Calculate Shift (Constrained to Plane)
                                XYZ targetCenter = new XYZ(centerWorldX, centerWorldY, centerWorldZ);
                                XYZ delta = targetCenter - masterPoint;
                                
                                // NEW TRANSLATION LOGIC (Wall Plane Projection)
                                // Instead of projecting onto BasisX/BasisZ (which might be rotated),
                                // We project onto the PLANE defined by the Host Normal (or BasisY).
                                // This allows sliding in any direction along the wall, but prevents moving OUT of the wall.
                                
                                XYZ normalVector = masterTransform.BasisY; // Default fallback (Thickness axis)
                                
                                // Retrieve host element from document
                                Element hostElement = masterSleeve.Host;
                                if (hostElement is Wall w)
                                {
                                    normalVector = w.Orientation;
                                }
                                
                                // Project delta onto the plane perpendicular to normalVector
                                // formula: v_planar = v - (v . n) * n
                                double normalComponent = delta.DotProduct(normalVector);
                                worldTranslation = delta - (normalVector * normalComponent);
                                
                                DebugLogger.Info($"[ManualJoin] Center Shift: Target={targetCenter}, Current={masterPoint}, Delta={delta}.");
                                DebugLogger.Info($"[ManualJoin] Wall Normal={normalVector}. NormalComp={normalComponent}. Final Planar Translation={worldTranslation}.");


                            }
                            else
                            {
                                DebugLogger.Warning("[ManualJoin] DB records not found. Falling back to Geometric Calculation.");
                                newWidth = maxX - minX;
                                newHeight = maxZ - minZ;
                                centerX = (minX + maxX) / 2.0;
                                centerZ = (minZ + maxZ) / 2.0;
                                
                                XYZ localTranslation = new XYZ(centerX, 0, centerZ);
                                worldTranslation = masterTransform.OfVector(localTranslation);
                            }
                            
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
                                // Convert zones to ClusterSleeveInfo
                                var sleeveInfos = new List<ClusterSleeveInfo>();
                                foreach(var z in zones)
                                {
                                    var info = ClusterSleeveInfo.FromClashZone(z);
                                    if(info != null) sleeveInfos.Add(info);
                                }
                                
                                // Create Dummy Candidate
                                var candidate = new CombinedClusterCandidate(0, sleeveInfos);
                                // Set Bounds (Required for service?)
                                // CombinedClusterFormationService usually sets these. 
                                // We populate them so the Aggregator has context if it needs size.
                                candidate.CombinedBoundingBoxMinX = unionMinX;
                                candidate.CombinedBoundingBoxMinY = zones.Min(z => z.SleevePlacementPointY - (z.SleeveWidth/2)); // Rough Logic
                                candidate.CombinedBoundingBoxMinZ = unionMinZ;
                                candidate.CombinedBoundingBoxMaxX = unionMaxX;
                                candidate.CombinedBoundingBoxMaxY = zones.Max(z => z.SleevePlacementPointY + (z.SleeveWidth/2));
                                candidate.CombinedBoundingBoxMaxZ = unionMaxZ;

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
                    _externalEvent.Raise();
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

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

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));

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
            
            // Interfaces for new services
            var combinedRepo = new CombinedClusterRepository(_repo);
             
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
            try
            {
                StatusMessage = "Detecting candidates...";
                
                var categories = new List<string>();
                if (IncludeDucts) categories.Add("Ducts");
                if (IncludePipes) categories.Add("Pipes");
                if (IncludeCableTrays) categories.Add("Cable Trays"); // Adjust string to match Revit category names or usage
                if (IncludeConduits) categories.Add("Conduits");

                if (categories.Count == 0)
                {
                    StatusMessage = "Select at least one category.";
                    return;
                }

                // Call Discover with filter name (empty or specific) and categories
                var candidatesList = _discoveryService.Discover("CombinedDiscovery", categories);
                var candidates = candidatesList.ToList();
                
                // Call Formation (Async, so use .Result or wait)
                var clusters = _formationService.FormCombinedClustersAsync(candidates).Result;
                
                if (!clusters.Any())
                {
                    StatusMessage = "No combined clusters found.";
                    return;
                }

                StatusMessage = $"Found {clusters.Count} clusters. Creating...";

                using (Transaction tx = new Transaction(_document, "Auto-Create Combined Sleeves"))
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
            }
            catch (Exception ex)
            {
                StatusMessage = "Error: " + ex.Message;
                DebugLogger.Error("Auto-Combine Error: " + ex.ToString());
            }
        }

        private void SelectByCrossing()
        {
            try
            {
                StatusMessage = "Select sleeves...";
                var refs = _uiDocument.Selection.PickObjects(SelectionObjectType.Element, new SleeveSelectionFilter(), "Select sleeves to join");
                
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
                StatusMessage = "Pick one by one (Esc to finish)...";
                _selectedIds.Clear();
                SelectedSleevesList.Clear();
                
                while (true)
                {
                     var r = _uiDocument.Selection.PickObject(SelectionObjectType.Element, new SleeveSelectionFilter(), "Pick next sleeve");
                     _selectedIds.Add(r.ElementId);
                     var elem = _document.GetElement(r.ElementId);
                     SelectedSleevesList.Add($"{elem.Category.Name} - {elem.Id}");
                     OnPropertyChanged(nameof(CanJoin));
                }
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                StatusMessage = "Selection finished.";
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
            
            try
            {
                StatusMessage = "Joining...";
                
                using (Transaction tx = new Transaction(_document, "Manual Join Combined Sleeve"))
                {
                    tx.Start();
                    
                    BoundingBoxXYZ combinedBox = null;
                    var cats = new HashSet<string>();
                    
                    foreach (var id in _selectedIds)
                    {
                        var el = _document.GetElement(id);
                        if(el == null) continue;
                        
                        var box = el.get_BoundingBox(null);
                        if (combinedBox == null) combinedBox = box;
                        else
                        {
                            combinedBox.Min = new XYZ(Math.Min(combinedBox.Min.X, box.Min.X), Math.Min(combinedBox.Min.Y, box.Min.Y), Math.Min(combinedBox.Min.Z, box.Min.Z));
                            combinedBox.Max = new XYZ(Math.Max(combinedBox.Max.X, box.Max.X), Math.Max(combinedBox.Max.Y, box.Max.Y), Math.Max(combinedBox.Max.Z, box.Max.Z));
                        }
                        cats.Add(el.Category.Name);
                        
                        _document.Delete(id);
                    }
                    
                    if (combinedBox != null)
                    {
                        string familyName = "RectangularOpeningOnWall";
                        FamilySymbol symbol = new FilteredElementCollector(_document)
                                .OfClass(typeof(FamilySymbol))
                                .Cast<FamilySymbol>()
                                .FirstOrDefault(x => x.Name.Equals(familyName, StringComparison.OrdinalIgnoreCase) || 
                                                     x.Family.Name.Equals(familyName, StringComparison.OrdinalIgnoreCase));
                                                     
                        if (symbol != null)
                        {
                            if(!symbol.IsActive) symbol.Activate();
                            var inst = _document.Create.NewFamilyInstance(combinedBox.Min, symbol, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                            
                            var pWidth = inst.LookupParameter("Width");
                            var pHeight = inst.LookupParameter("Height");
                            if(pWidth!=null) pWidth.Set(combinedBox.Max.X - combinedBox.Min.X);
                            if(pHeight!=null) pHeight.Set(combinedBox.Max.Z - combinedBox.Min.Z);
                            
                            var pComm = inst.LookupParameter("Comments");
                            if(pComm!=null) pComm.Set($"Manual Join: {string.Join(",", cats)}");
                        }
                    }

                    tx.Commit();
                    
                    _selectedIds.Clear();
                    SelectedSleevesList.Clear();
                    OnPropertyChanged(nameof(CanJoin));
                    StatusMessage = "Join complete.";
                }
            }
            catch (Exception ex)
            {
                StatusMessage = "Join failed: " + ex.Message;
            }
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

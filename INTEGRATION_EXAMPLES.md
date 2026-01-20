# Integration Examples for Opening Auto-Update System

## 🎯 **Overview**
This document shows how to integrate the new opening auto-update services with your existing MEP openings code.

## 📋 **Integration Points**

### **1. Service Registration and Initialization**

```csharp
// In your main application startup or service container
public class ApplicationServices
{
    private readonly StatusManager _statusManager;
    private readonly OpeningTrackingService _trackingService;
    private readonly LinkedFileReloadService _linkedFileService;

    public ApplicationServices()
    {
        // Initialize services
        _statusManager = new StatusManager();
        _trackingService = new OpeningTrackingService(_statusManager);
        _linkedFileService = new LinkedFileReloadService(_trackingService, _statusManager);
        
        // Subscribe to events
        _trackingService.OpeningStatusUpdated += OnOpeningStatusUpdated;
        _trackingService.OpeningMoved += OnOpeningMoved;
        _linkedFileService.LinkedFileReloaded += OnLinkedFileReloaded;
    }

    private void OnOpeningStatusUpdated(object? sender, OpeningStatusUpdatedEventArgs e)
    {
        // Handle opening status updates
        DebugLogger.Info($"Opening {e.Status.OpeningId.IntegerValue} status updated: {e.Status.Status}");
    }

    private void OnOpeningMoved(object? sender, OpeningMovedEventArgs e)
    {
        // Handle opening movement
        DebugLogger.Info($"Opening {e.Status.OpeningId.IntegerValue} moved {e.MovementDistance:F2} mm");
    }

    private void OnLinkedFileReloaded(object? sender, LinkedFileReloadedEventArgs e)
    {
        // Handle linked file reload
        DebugLogger.Info($"Linked file reloaded, updating openings in document: {e.Document.Title}");
    }
}
```

### **2. Integration with Existing DuctSleevePlacerService**

```csharp
// Modify your existing DuctSleevePlacerService.cs
public class DuctSleevePlacerService
{
    private readonly OpeningTrackingService _trackingService;
    private readonly StatusManager _statusManager;
    
    // ... existing code ...

    public DuctSleevePlacerService(OpeningTrackingService trackingService, StatusManager statusManager)
    {
        _trackingService = trackingService;
        _statusManager = statusManager;
        // ... existing initialization ...
    }

    public void CreateOpenings(List<OpeningRequirement> requirements, Document doc)
    {
        try
        {
            _statusManager.UpdateStatus("Creating duct openings...", StatusType.Processing);
            
            var createdOpenings = new List<ElementId>();
            
            using (var transaction = new Transaction(doc, "Create Duct Openings"))
            {
                transaction.Start();
                
                foreach (var requirement in requirements)
                {
                    // Create the opening (your existing logic)
                    var openingId = CreateSingleOpening(requirement, doc);
                    if (openingId != null)
                    {
                        createdOpenings.Add(openingId.Value);
                    }
                }
                
                transaction.Commit();
            }
            
            // Register created openings for tracking
            foreach (var requirement in requirements)
            {
                _trackingService.RegisterOpening(requirement);
            }
            
            _statusManager.UpdateStatus($"Created {createdOpenings.Count} duct openings", StatusType.Success);
        }
        catch (Exception ex)
        {
            _statusManager.UpdateStatus($"Error creating duct openings: {ex.Message}", StatusType.Error);
        }
    }
    
    private ElementId? CreateSingleOpening(OpeningRequirement requirement, Document doc)
    {
        // Your existing opening creation logic here
        // Return the ElementId of the created opening
        return null; // Placeholder
    }
}
```

### **3. Integration with Existing PipeSleevePlacerService**

```csharp
// Modify your existing PipeSleevePlacerService.cs
public class PipeSleevePlacerService
{
    private readonly OpeningTrackingService _trackingService;
    private readonly StatusManager _statusManager;
    
    // ... existing code ...

    public void CreateOpenings(List<OpeningRequirement> requirements, Document doc)
    {
        try
        {
            _statusManager.UpdateStatus("Creating pipe openings...", StatusType.Processing);
            
            var createdOpenings = new List<ElementId>();
            
            using (var transaction = new Transaction(doc, "Create Pipe Openings"))
            {
                transaction.Start();
                
                foreach (var requirement in requirements)
                {
                    // Create the opening (your existing logic)
                    var openingId = CreateSingleOpening(requirement, doc);
                    if (openingId != null)
                    {
                        createdOpenings.Add(openingId.Value);
                        
                        // Register for tracking immediately after creation
                        _trackingService.RegisterOpening(openingId.Value, requirement.ElementId.Value, requirement.WallId.Value, doc);
                    }
                }
                
                transaction.Commit();
            }
            
            _statusManager.UpdateStatus($"Created {createdOpenings.Count} pipe openings", StatusType.Success);
        }
        catch (Exception ex)
        {
            _statusManager.UpdateStatus($"Error creating pipe openings: {ex.Message}", StatusType.Error);
        }
    }
}
```

### **4. Integration with Main Dialog**

```csharp
// Modify your existing MainDialogViewModel.cs
public partial class MainDialogViewModel : ObservableObject
{
    private readonly OpeningTrackingService _trackingService;
    private readonly LinkedFileReloadService _linkedFileService;
    private readonly StatusManager _statusManager;
    
    // ... existing code ...

    public MainDialogViewModel(OpeningTrackingService trackingService, 
                              LinkedFileReloadService linkedFileService,
                              StatusManager statusManager)
    {
        _trackingService = trackingService;
        _linkedFileService = linkedFileService;
        _statusManager = statusManager;
        
        // ... existing initialization ...
        
        // Add new commands
        CheckForUpdatesCommand = new RelayCommand(OnCheckForUpdates);
        ShowStatusCommand = new RelayCommand(OnShowStatus);
    }

    [ObservableProperty]
    private bool _autoUpdateEnabled = true;

    public IRelayCommand CheckForUpdatesCommand { get; }
    public IRelayCommand ShowStatusCommand { get; }

    private void OnCheckForUpdates()
    {
        try
        {
            _statusManager.UpdateStatus("Checking for MEP element movements...", StatusType.Processing);
            
            var doc = GetCurrentDocument();
            if (doc != null)
            {
                _trackingService.UpdateOpeningsForMovedMepElements(doc);
                _statusManager.UpdateStatus("Opening position check completed", StatusType.Success);
            }
        }
        catch (Exception ex)
        {
            _statusManager.UpdateStatus($"Error checking updates: {ex.Message}", StatusType.Error);
        }
    }

    private void OnShowStatus()
    {
        var dialog = new OpeningStatusDialog();
        dialog.DataContext = new OpeningStatusDialogViewModel(_trackingService, _statusManager);
        dialog.ShowDialog();
    }

    private Document? GetCurrentDocument()
    {
        // Implement based on your Revit context
        // This might come from ExternalEventManager or similar
        return null; // Placeholder
    }
}
```

### **5. Integration with ExternalEventManager**

```csharp
// Modify your existing ExternalEventManager.cs
public class ExternalEventManager
{
    private readonly OpeningTrackingService _trackingService;
    private readonly LinkedFileReloadService _linkedFileService;
    private readonly StatusManager _statusManager;
    
    // ... existing code ...

    public ExternalEventManager(OpeningTrackingService trackingService,
                                LinkedFileReloadService linkedFileService,
                                StatusManager statusManager)
    {
        _trackingService = trackingService;
        _linkedFileService = linkedFileService;
        _statusManager = statusManager;
        
        // ... existing initialization ...
    }

    // Add new external event for checking updates
    public class CheckForUpdatesExternalEvent : IExternalEventHandler
    {
        private readonly OpeningTrackingService _trackingService;
        private readonly StatusManager _statusManager;

        public CheckForUpdatesExternalEvent(OpeningTrackingService trackingService, StatusManager statusManager)
        {
            _trackingService = trackingService;
            _statusManager = statusManager;
        }

        public void Execute(UIApplication app)
        {
            try
            {
                var doc = app.ActiveUIDocument.Document;
                _trackingService.UpdateOpeningsForMovedMepElements(doc);
            }
            catch (Exception ex)
            {
                _statusManager.UpdateStatus($"Error in external event: {ex.Message}", StatusType.Error);
            }
        }

        public string GetName()
        {
            return "Check for Opening Updates";
        }
    }
}
```

### **6. Integration with Existing Commands**

```csharp
// Modify your existing command classes
[Transaction(TransactionMode.Manual)]
public class DuctSleeveCommand : IExternalCommand
{
    private readonly OpeningTrackingService _trackingService;
    private readonly StatusManager _statusManager;

    public DuctSleeveCommand(OpeningTrackingService trackingService, StatusManager statusManager)
    {
        _trackingService = trackingService;
        _statusManager = statusManager;
    }

    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        try
        {
            var uiDoc = commandData.Application.ActiveUIDocument;
            var doc = uiDoc.Document;

            // Your existing opening creation logic
            var requirements = GetOpeningRequirements(doc);
            var service = new DuctSleevePlacerService(_trackingService, _statusManager);
            service.CreateOpenings(requirements, doc);

            // Check for linked file updates after creation
            var linkedFileService = new LinkedFileReloadService(_trackingService, _statusManager);
            linkedFileService.CheckForLinkedFileUpdates(doc);

            return Result.Succeeded;
        }
        catch (Exception ex)
        {
            message = ex.Message;
            return Result.Failed;
        }
    }

    private List<OpeningRequirement> GetOpeningRequirements(Document doc)
    {
        // Your existing logic to get opening requirements
        return new List<OpeningRequirement>();
    }
}
```

## 🔄 **Usage Patterns**

### **Pattern 1: Automatic Registration**
```csharp
// When creating openings, automatically register them
public void CreateAndTrackOpenings(List<OpeningRequirement> requirements, Document doc)
{
    // Create openings
    var createdOpenings = CreateOpenings(requirements, doc);
    
    // Register for tracking
    foreach (var requirement in requirements)
    {
        _trackingService.RegisterOpening(requirement);
    }
}
```

### **Pattern 2: Manual Updates**
```csharp
// Manually trigger updates when needed
public void ManualUpdateCheck(Document doc)
{
    _trackingService.UpdateOpeningsForMovedMepElements(doc);
}
```

### **Pattern 3: Linked File Monitoring**
```csharp
// Monitor linked files for changes
public void MonitorLinkedFiles(Document doc)
{
    _linkedFileService.CheckForLinkedFileUpdates(doc);
}
```

### **Pattern 4: Status Monitoring**
```csharp
// Get current status of all tracked openings
public void ShowOpeningStatus()
{
    var statuses = _trackingService.GetAllStatuses();
    foreach (var status in statuses)
    {
        DebugLogger.Info($"Opening {status.OpeningId.IntegerValue}: {status.Status} - {status.MovementDistanceText}");
    }
}
```

## 🎯 **Key Integration Points**

1. **Service Initialization** - Set up services in your application startup
2. **Opening Creation** - Register openings immediately after creation
3. **Event Handling** - Subscribe to status and movement events
4. **Manual Triggers** - Provide UI buttons for manual update checks
5. **Linked File Monitoring** - Check for linked file changes periodically
6. **Error Handling** - Handle errors gracefully with status updates

This integration approach ensures that your existing code continues to work while adding the new auto-update functionality seamlessly.



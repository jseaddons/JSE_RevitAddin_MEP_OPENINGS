# Simple Opening Auto-Update Implementation Plan

## 🎯 **Goal**
Create a simple system that:
1. **Tracks opening status relative to MEP elements**
2. **Automatically updates opening positions when MEP elements move in linked files**
3. **Triggers updates when linked files are reloaded**

## 📋 **Current System Analysis**

### **Existing Components:**
- ✅ `OpeningRequirement.cs` - Contains MEP element relationships
- ✅ `StatusManager.cs` - Tracks operation status
- ✅ `ExternalEventManager.cs` - Handles Revit API operations
- ✅ Linked file detection in existing services

### **Key Data Points Available:**
- `ElementId` - MEP element ID
- `WallId` - Host wall ID  
- `Location` - Opening position
- `MepType` - Type of MEP element
- `ElementType` - Element type information

## 🏗️ **Simple Implementation**

### **1. Opening Status Model**

```csharp
// Models/OpeningStatus.cs
public class OpeningStatus
{
    public ElementId OpeningId { get; set; }
    public ElementId MepElementId { get; set; }
    public ElementId WallId { get; set; }
    public XYZ OriginalLocation { get; set; }
    public XYZ CurrentLocation { get; set; }
    public string MepType { get; set; } = string.Empty;
    public DateTime LastUpdated { get; set; }
    public bool IsAutoUpdateEnabled { get; set; } = true;
    public string Status { get; set; } = "Active"; // Active, Moved, Error, Disabled
}
```

### **2. Simple Tracking Service**

```csharp
// Services/OpeningTrackingService.cs
public class OpeningTrackingService
{
    private readonly Dictionary<ElementId, OpeningStatus> _openingStatuses;
    private readonly StatusManager _statusManager;
    
    public OpeningTrackingService(StatusManager statusManager)
    {
        _openingStatuses = new Dictionary<ElementId, OpeningStatus>();
        _statusManager = statusManager;
    }
    
    /// <summary>
    /// Register an opening with its associated MEP element
    /// </summary>
    public void RegisterOpening(OpeningRequirement opening)
    {
        if (opening.ElementId == null || opening.WallId == null) return;
        
        var status = new OpeningStatus
        {
            OpeningId = GetOpeningElementId(opening), // You'll need to implement this
            MepElementId = opening.ElementId.Value,
            WallId = opening.WallId.Value,
            OriginalLocation = opening.Location ?? XYZ.Zero,
            CurrentLocation = opening.Location ?? XYZ.Zero,
            MepType = opening.MepType ?? "Unknown",
            LastUpdated = DateTime.Now,
            Status = "Active"
        };
        
        _openingStatuses[status.OpeningId] = status;
        _statusManager.UpdateStatus($"Registered opening for {opening.MepType}", StatusType.Info);
    }
    
    /// <summary>
    /// Update opening positions when MEP elements move
    /// </summary>
    public void UpdateOpeningsForMovedMepElements(Document doc)
    {
        var updatedCount = 0;
        
        foreach (var status in _openingStatuses.Values.Where(s => s.IsAutoUpdateEnabled))
        {
            try
            {
                var mepElement = doc.GetElement(status.MepElementId);
                if (mepElement == null) continue;
                
                var newLocation = GetMepElementLocation(mepElement);
                if (newLocation == null) continue;
                
                // Check if MEP element moved significantly
                var distance = status.CurrentLocation.DistanceTo(newLocation);
                if (distance > 0.01) // 1cm threshold
                {
                    UpdateOpeningPosition(doc, status, newLocation);
                    status.CurrentLocation = newLocation;
                    status.LastUpdated = DateTime.Now;
                    status.Status = "Moved";
                    updatedCount++;
                }
            }
            catch (Exception ex)
            {
                status.Status = "Error";
                _statusManager.UpdateStatus($"Error updating opening {status.OpeningId}: {ex.Message}", StatusType.Error);
            }
        }
        
        if (updatedCount > 0)
        {
            _statusManager.UpdateStatus($"Updated {updatedCount} opening positions", StatusType.Success);
        }
    }
    
    private void UpdateOpeningPosition(Document doc, OpeningStatus status, XYZ newLocation)
    {
        using (var transaction = new Transaction(doc, "Update Opening Position"))
        {
            transaction.Start();
            
            var opening = doc.GetElement(status.OpeningId);
            if (opening != null)
            {
                // Move the opening to match the new MEP element position
                var transform = Transform.CreateTranslation(newLocation - status.CurrentLocation);
                ElementTransformUtils.MoveElement(doc, status.OpeningId, transform);
            }
            
            transaction.Commit();
        }
    }
    
    private XYZ? GetMepElementLocation(Element mepElement)
    {
        // Get the location based on MEP element type
        if (mepElement is Pipe pipe)
        {
            var curve = pipe.Location as LocationCurve;
            return curve?.Curve?.GetEndPoint(0);
        }
        else if (mepElement is Duct duct)
        {
            var curve = duct.Location as LocationCurve;
            return curve?.Curve?.GetEndPoint(0);
        }
        else if (mepElement is CableTray cableTray)
        {
            var curve = cableTray.Location as LocationCurve;
            return curve?.Curve?.GetEndPoint(0);
        }
        
        return null;
    }
}
```

### **3. Linked File Reload Detection**

```csharp
// Services/LinkedFileReloadService.cs
public class LinkedFileReloadService
{
    private readonly OpeningTrackingService _trackingService;
    private readonly StatusManager _statusManager;
    private Document _lastDocument;
    private DateTime _lastCheckTime;
    
    public LinkedFileReloadService(OpeningTrackingService trackingService, StatusManager statusManager)
    {
        _trackingService = trackingService;
        _statusManager = statusManager;
        _lastCheckTime = DateTime.Now;
    }
    
    /// <summary>
    /// Check for linked file reloads and update openings if needed
    /// </summary>
    public void CheckForLinkedFileUpdates(Document doc)
    {
        if (doc == null) return;
        
        // Check if this is a different document or if enough time has passed
        if (_lastDocument != doc || DateTime.Now - _lastCheckTime > TimeSpan.FromMinutes(1))
        {
            _lastDocument = doc;
            _lastCheckTime = DateTime.Now;
            
            // Check for linked file changes
            var linkedFilesChanged = CheckLinkedFilesChanged(doc);
            
            if (linkedFilesChanged)
            {
                _statusManager.UpdateStatus("Linked files reloaded - updating opening positions", StatusType.Info);
                _trackingService.UpdateOpeningsForMovedMepElements(doc);
            }
        }
    }
    
    private bool CheckLinkedFilesChanged(Document doc)
    {
        try
        {
            // Get all RevitLinkInstances
            var linkInstances = new FilteredElementCollector(doc)
                .OfClass(typeof(RevitLinkInstance))
                .Cast<RevitLinkInstance>()
                .ToList();
            
            // Check if any linked files have been reloaded recently
            foreach (var link in linkInstances)
            {
                var linkDoc = link.GetLinkDocument();
                if (linkDoc != null)
                {
                    // Check if the linked document has been modified recently
                    var lastModified = GetDocumentLastModified(linkDoc);
                    if (lastModified > _lastCheckTime.AddMinutes(-2)) // Within last 2 minutes
                    {
                        return true;
                    }
                }
            }
        }
        catch (Exception ex)
        {
            _statusManager.UpdateStatus($"Error checking linked files: {ex.Message}", StatusType.Warning);
        }
        
        return false;
    }
    
    private DateTime GetDocumentLastModified(Document doc)
    {
        // This is a simplified approach - you might need to implement
        // a more sophisticated way to detect document changes
        return DateTime.Now; // Placeholder
    }
}
```

### **4. Integration with Existing System**

```csharp
// Extend existing services to use the tracking system

// In your existing opening creation services, add:
public class DuctSleevePlacerService
{
    private readonly OpeningTrackingService _trackingService;
    
    // ... existing code ...
    
    public void CreateOpenings(List<OpeningRequirement> requirements)
    {
        // ... existing opening creation logic ...
        
        // Register each created opening for tracking
        foreach (var requirement in requirements)
        {
            _trackingService.RegisterOpening(requirement);
        }
    }
}
```

### **5. Simple UI Design**

#### **A. Add to Existing Main Dialog**

```xml
<!-- Add to your existing MainDialog.xaml -->
<!-- In the bottom toolbar section, add auto-update controls -->
<StackPanel Grid.Column="0" Orientation="Horizontal">
    <!-- Existing buttons -->
    <Button Style="{StaticResource IconButtonStyle}" Content="←" ToolTip="Previous"/>
    <Button Style="{StaticResource IconButtonStyle}" Content="→" ToolTip="Next"/>
    <Button Style="{StaticResource IconButtonStyle}" Content="⚡" ToolTip="Process"/>
    <Button Style="{StaticResource IconButtonStyle}" Content="🔧" ToolTip="Settings"/>
    <Button Style="{StaticResource IconButtonStyle}" Content="🔄" ToolTip="Refresh"/>
    
    <!-- NEW: Auto-Update Controls -->
    <Separator Margin="10,0"/>
    <Button Style="{StaticResource IconButtonStyle}" 
            Content="🔗" 
            ToolTip="Check for MEP Movement"
            Command="{Binding CheckForUpdatesCommand}"
            Margin="5,0"/>
    <Button Style="{StaticResource IconButtonStyle}" 
            Content="📊" 
            ToolTip="Show Opening Status"
            Command="{Binding ShowStatusCommand}"
            Margin="5,0"/>
    <CheckBox Content="Auto-Update" 
              IsChecked="{Binding AutoUpdateEnabled}"
              VerticalAlignment="Center"
              Margin="10,0,0,0"/>
</StackPanel>
```

#### **B. Simple Status Panel**

```xml
<!-- Views/OpeningStatusPanel.xaml -->
<UserControl x:Class="JSE_RevitAddin_MEP_OPENINGS.Views.OpeningStatusPanel">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <!-- Header -->
        <TextBlock Grid.Row="0" Text="Opening Status" 
                   FontSize="16" FontWeight="Bold" Margin="0,0,0,10"/>
        
        <!-- Status List -->
        <ListBox Grid.Row="1" ItemsSource="{Binding OpeningStatuses}">
            <ListBox.ItemTemplate>
                <DataTemplate>
                    <Expander Margin="5" IsExpanded="False">
                        <Expander.Header>
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                
                                <!-- Status Icon -->
                                <TextBlock Grid.Column="0" 
                                           Text="{Binding StatusIcon}" 
                                           FontSize="16" 
                                           Margin="0,0,10,0"
                                           VerticalAlignment="Center"/>
                                
                                <!-- Summary Info -->
                                <StackPanel Grid.Column="1">
                                    <TextBlock Text="{Binding OpeningIdText}" FontWeight="Bold"/>
                                    <TextBlock Text="{Binding MepElementType}" FontSize="10" Foreground="Gray"/>
                                </StackPanel>
                                
                                <!-- Movement Distance -->
                                <TextBlock Grid.Column="2" 
                                           Text="{Binding MovementDistance}" 
                                           FontSize="10" 
                                           Foreground="Orange"
                                           Margin="10,0"
                                           VerticalAlignment="Center"/>
                                
                                <!-- Action Button -->
                                <Button Grid.Column="3" 
                                        Content="Update" 
                                        Width="60" 
                                        Height="25"
                                        Command="{Binding DataContext.UpdateOpeningCommand, RelativeSource={RelativeSource AncestorType=UserControl}}"
                                        CommandParameter="{Binding}"
                                        IsEnabled="{Binding CanUpdate}"/>
                            </Grid>
                        </Expander.Header>
                        
                        <!-- Detailed Information -->
                        <Grid Margin="10,5">
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>
                            
                            <!-- Opening Details -->
                            <GroupBox Grid.Row="0" Header="Opening Details" Margin="0,0,0,5">
                                <StackPanel Margin="10">
                                    <TextBlock Text="{Binding OpeningIdText}" FontWeight="Bold"/>
                                    <TextBlock Text="{Binding OpeningDimensions}" Margin="0,2,0,0"/>
                                    <TextBlock Text="{Binding OpeningLocation}" Margin="0,2,0,0"/>
                                    <TextBlock Text="{Binding OpeningFamilyName}" Margin="0,2,0,0"/>
                                </StackPanel>
                            </GroupBox>
                            
                            <!-- MEP Element Details -->
                            <GroupBox Grid.Row="1" Header="MEP Element Details" Margin="0,0,0,5">
                                <StackPanel Margin="10">
                                    <TextBlock Text="{Binding MepElementIdText}" FontWeight="Bold"/>
                                    <TextBlock Text="{Binding MepElementType}" Margin="0,2,0,0"/>
                                    <TextBlock Text="{Binding MepElementSize}" Margin="0,2,0,0"/>
                                    <TextBlock Text="{Binding MepElementLocation}" Margin="0,2,0,0"/>
                                </StackPanel>
                            </GroupBox>
                            
                            <!-- Wall Details -->
                            <GroupBox Grid.Row="2" Header="Host Wall Details" Margin="0,0,0,5">
                                <StackPanel Margin="10">
                                    <TextBlock Text="{Binding WallIdText}" FontWeight="Bold"/>
                                    <TextBlock Text="{Binding WallType}" Margin="0,2,0,0"/>
                                </StackPanel>
                            </GroupBox>
                            
                            <!-- Status Information -->
                            <GroupBox Grid.Row="3" Header="Status Information">
                                <StackPanel Margin="10">
                                    <TextBlock Text="{Binding Status}" FontWeight="Bold"/>
                                    <TextBlock Text="{Binding LastUpdatedText}" Margin="0,2,0,0"/>
                                    <TextBlock Text="{Binding MovementDistance}" Margin="0,2,0,0"/>
                                </StackPanel>
                            </GroupBox>
                        </Grid>
                    </Expander>
                </DataTemplate>
            </ListBox.ItemTemplate>
        </ListBox>
        
        <!-- Summary -->
        <Border Grid.Row="2" Background="#F8F9FA" Padding="10" Margin="0,10,0,0">
            <StackPanel Orientation="Horizontal">
                <TextBlock Text="Total Openings: " FontWeight="Bold"/>
                <TextBlock Text="{Binding TotalOpenings}" FontWeight="Bold" Margin="0,0,20,0"/>
                <TextBlock Text="Active: " FontWeight="Bold"/>
                <TextBlock Text="{Binding ActiveOpenings}" FontWeight="Bold" Foreground="Green" Margin="0,0,20,0"/>
                <TextBlock Text="Moved: " FontWeight="Bold"/>
                <TextBlock Text="{Binding MovedOpenings}" FontWeight="Bold" Foreground="Orange" Margin="0,0,20,0"/>
                <TextBlock Text="Errors: " FontWeight="Bold"/>
                <TextBlock Text="{Binding ErrorOpenings}" FontWeight="Bold" Foreground="Red"/>
            </StackPanel>
        </Border>
    </Grid>
</UserControl>
```

#### **C. Simple Status Dialog**

```xml
<!-- Views/OpeningStatusDialog.xaml -->
<Window x:Class="JSE_RevitAddin_MEP_OPENINGS.Views.OpeningStatusDialog"
        Title="Opening Status Monitor"
        Width="600" Height="400"
        WindowStartupLocation="CenterScreen">
    
    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <!-- Header with Controls -->
        <Grid Grid.Row="0" Margin="0,0,0,15">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            
            <TextBlock Grid.Column="0" Text="Opening Status Monitor" 
                       FontSize="18" FontWeight="Bold"/>
            
            <StackPanel Grid.Column="1" Orientation="Horizontal">
                <Button Content="Refresh" 
                        Command="{Binding RefreshCommand}"
                        Margin="0,0,10,0"/>
                <Button Content="Check Updates" 
                        Command="{Binding CheckForUpdatesCommand}"
                        Margin="0,0,10,0"/>
                <Button Content="Close" 
                        Click="OnCloseClick"/>
            </StackPanel>
        </Grid>
        
        <!-- Status Panel -->
        <local:OpeningStatusPanel Grid.Row="1" DataContext="{Binding StatusPanelViewModel}"/>
        
        <!-- Bottom Info -->
        <Border Grid.Row="2" Background="#E9ECEF" Padding="10" Margin="0,15,0,0">
            <StackPanel Orientation="Horizontal">
                <TextBlock Text="Auto-Update: " FontWeight="Bold"/>
                <TextBlock Text="{Binding AutoUpdateStatus}" Margin="0,0,20,0"/>
                <TextBlock Text="Last Check: " FontWeight="Bold"/>
                <TextBlock Text="{Binding LastCheckTime}"/>
            </StackPanel>
        </Border>
    </Grid>
</Window>
```

#### **D. ViewModel for Status Panel**

```csharp
// ViewModels/OpeningStatusPanelViewModel.cs
public class OpeningStatusPanelViewModel : ObservableObject
{
    private readonly OpeningTrackingService _trackingService;
    private readonly StatusManager _statusManager;
    
    public ObservableCollection<OpeningStatusItem> OpeningStatuses { get; } = new();
    
    [ObservableProperty]
    private int _totalOpenings;
    
    [ObservableProperty]
    private int _activeOpenings;
    
    [ObservableProperty]
    private int _movedOpenings;
    
    [ObservableProperty]
    private int _errorOpenings;
    
    public IRelayCommand UpdateOpeningCommand { get; }
    public IRelayCommand RefreshCommand { get; }
    
    public OpeningStatusPanelViewModel(OpeningTrackingService trackingService, StatusManager statusManager)
    {
        _trackingService = trackingService;
        _statusManager = statusManager;
        
        UpdateOpeningCommand = new RelayCommand<OpeningStatusItem>(OnUpdateOpening);
        RefreshCommand = new RelayCommand(OnRefresh);
        
        // Subscribe to status updates
        _statusManager.StatusUpdated += OnStatusUpdated;
    }
    
    private void OnRefresh()
    {
        RefreshStatusList();
    }
    
    private void OnUpdateOpening(OpeningStatusItem item)
    {
        if (item != null)
        {
            // Update specific opening
            _trackingService.UpdateSingleOpening(item.OpeningId);
        }
    }
    
    private void RefreshStatusList()
    {
        OpeningStatuses.Clear();
        
        var statuses = _trackingService.GetAllStatuses();
        foreach (var status in statuses)
        {
            OpeningStatuses.Add(new OpeningStatusItem(status));
        }
        
        UpdateCounts();
    }
    
    private void UpdateCounts()
    {
        TotalOpenings = OpeningStatuses.Count;
        ActiveOpenings = OpeningStatuses.Count(s => s.Status == "Active");
        MovedOpenings = OpeningStatuses.Count(s => s.Status == "Moved");
        ErrorOpenings = OpeningStatuses.Count(s => s.Status == "Error");
    }
}

// ViewModels/OpeningStatusItem.cs
public class OpeningStatusItem : ObservableObject
{
    public ElementId OpeningId { get; }
    public ElementId MepElementId { get; }
    public ElementId WallId { get; }
    
    // Opening Details
    public string OpeningIdText { get; }
    public string OpeningDimensions { get; }
    public string OpeningLocation { get; }
    public string OpeningFamilyName { get; }
    
    // MEP Element Details
    public string MepElementIdText { get; }
    public string MepElementType { get; }
    public string MepElementSize { get; }
    public string MepElementLocation { get; }
    
    // Wall Details
    public string WallIdText { get; }
    public string WallType { get; }
    
    // Status Information
    public string Status { get; }
    public string StatusIcon => Status switch
    {
        "Active" => "✓",
        "Moved" => "↗",
        "Error" => "✗",
        "Disabled" => "⏸",
        _ => "?"
    };
    public string LastUpdatedText { get; }
    public bool CanUpdate { get; }
    public string MovementDistance { get; }
    
    public OpeningStatusItem(OpeningStatus status, Document doc)
    {
        OpeningId = status.OpeningId;
        MepElementId = status.MepElementId;
        WallId = status.WallId;
        
        // Opening Details
        OpeningIdText = $"Opening ID: {status.OpeningId.IntegerValue}";
        OpeningDimensions = $"{status.Width:F1} × {status.Height:F1} mm";
        OpeningLocation = $"X:{status.CurrentLocation.X:F2}, Y:{status.CurrentLocation.Y:F2}, Z:{status.CurrentLocation.Z:F2}";
        OpeningFamilyName = GetOpeningFamilyName(status.OpeningId, doc);
        
        // MEP Element Details
        MepElementIdText = $"MEP ID: {status.MepElementId.IntegerValue}";
        MepElementType = status.MepType;
        MepElementSize = GetMepElementSize(status.MepElementId, doc);
        MepElementLocation = GetMepElementLocation(status.MepElementId, doc);
        
        // Wall Details
        WallIdText = $"Wall ID: {status.WallId.IntegerValue}";
        WallType = GetWallType(status.WallId, doc);
        
        // Status Information
        Status = status.Status;
        LastUpdatedText = status.LastUpdated.ToString("HH:mm:ss");
        CanUpdate = status.Status != "Active";
        
        // Calculate movement distance
        var distance = status.OriginalLocation.DistanceTo(status.CurrentLocation);
        MovementDistance = distance > 0.01 ? $"{distance:F2} mm" : "No movement";
    }
    
    private string GetOpeningFamilyName(ElementId openingId, Document doc)
    {
        try
        {
            var opening = doc.GetElement(openingId) as FamilyInstance;
            return opening?.Symbol?.Family?.Name ?? "Unknown";
        }
        catch
        {
            return "Unknown";
        }
    }
    
    private string GetMepElementSize(ElementId mepId, Document doc)
    {
        try
        {
            var element = doc.GetElement(mepId);
            
            if (element is Pipe pipe)
            {
                var diameter = pipe.LookupParameter("Diameter")?.AsDouble() ?? 0;
                return $"Ø{diameter * 304.8:F1} mm"; // Convert feet to mm
            }
            else if (element is Duct duct)
            {
                var width = duct.LookupParameter("Width")?.AsDouble() ?? 0;
                var height = duct.LookupParameter("Height")?.AsDouble() ?? 0;
                return $"{width * 304.8:F1} × {height * 304.8:F1} mm";
            }
            else if (element is CableTray cableTray)
            {
                var width = cableTray.LookupParameter("Width")?.AsDouble() ?? 0;
                var height = cableTray.LookupParameter("Height")?.AsDouble() ?? 0;
                return $"{width * 304.8:F1} × {height * 304.8:F1} mm";
            }
            
            return "Unknown size";
        }
        catch
        {
            return "Unknown size";
        }
    }
    
    private string GetMepElementLocation(ElementId mepId, Document doc)
    {
        try
        {
            var element = doc.GetElement(mepId);
            var location = element?.Location;
            
            if (location is LocationPoint point)
            {
                var xyz = point.Point;
                return $"X:{xyz.X:F2}, Y:{xyz.Y:F2}, Z:{xyz.Z:F2}";
            }
            else if (location is LocationCurve curve)
            {
                var start = curve.Curve.GetEndPoint(0);
                return $"Start: X:{start.X:F2}, Y:{start.Y:F2}, Z:{start.Z:F2}";
            }
            
            return "Unknown location";
        }
        catch
        {
            return "Unknown location";
        }
    }
    
    private string GetWallType(ElementId wallId, Document doc)
    {
        try
        {
            var wall = doc.GetElement(wallId) as Wall;
            return wall?.WallType?.Name ?? "Unknown wall";
        }
        catch
        {
            return "Unknown wall";
        }
    }
}
```

#### **E. Integration with Main Dialog**

```csharp
// Add to your existing MainDialogViewModel.cs
public partial class MainDialogViewModel : ObservableObject
{
    // ... existing code ...
    
    [ObservableProperty]
    private bool _autoUpdateEnabled = true;
    
    public IRelayCommand CheckForUpdatesCommand { get; }
    public IRelayCommand ShowStatusCommand { get; }
    
    public MainDialogViewModel()
    {
        // ... existing code ...
        
        CheckForUpdatesCommand = new RelayCommand(OnCheckForUpdates);
        ShowStatusCommand = new RelayCommand(OnShowStatus);
    }
    
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
}
```

## 🎨 **Simple UI Features**

### **Main Dialog Integration:**
- **🔗 Check for MEP Movement** - Manual trigger to check for updates
- **📊 Show Opening Status** - Opens status monitor dialog
- **Auto-Update checkbox** - Enable/disable automatic updates

### **Status Monitor Dialog:**
- **List of all tracked openings** - Shows status, MEP element info, last updated
- **Summary counts** - Total, Active, Moved, Error openings
- **Individual update buttons** - Update specific openings
- **Refresh button** - Reload status list
- **Auto-update status** - Shows if auto-update is enabled

### **Status Icons:**
- **✓ Active** - Opening is up to date
- **↗ Moved** - Opening was moved due to MEP element movement
- **✗ Error** - Error occurred during update

This UI is much simpler than the complex BCF export interface and focuses on the core functionality you need: monitoring and updating opening positions when MEP elements move.

## 🚀 **Implementation Steps**

### **Step 1: Create Models (1 day)**
- [ ] Create `OpeningStatus.cs` model
- [ ] Add simple status tracking properties

### **Step 2: Create Tracking Service (2 days)**
- [ ] Implement `OpeningTrackingService.cs`
- [ ] Add MEP element location detection
- [ ] Add opening position update logic

### **Step 3: Create Reload Detection (1 day)**
- [ ] Implement `LinkedFileReloadService.cs`
- [ ] Add linked file change detection
- [ ] Integrate with tracking service

### **Step 4: Integrate with Existing System (1 day)**
- [ ] Modify existing opening creation services
- [ ] Add registration calls after opening creation
- [ ] Add simple UI button for manual checking

### **Step 5: Test (1 day)**
- [ ] Test with moved MEP elements
- [ ] Test with linked file reloads
- [ ] Verify opening position updates

## 📋 **Simple Usage**

1. **Create openings** - System automatically registers them for tracking
2. **Move MEP elements** - System detects movement and updates opening positions
3. **Reload linked files** - System checks for changes and updates openings
4. **Manual check** - User can click button to force update check

## 🎯 **Key Benefits**

- **Simple** - Minimal code, easy to understand
- **Automatic** - No user intervention required
- **Reliable** - Uses existing Revit API patterns
- **Integrated** - Works with your existing system
- **Lightweight** - Minimal performance impact

This approach gives you exactly what you need: automatic opening position updates when MEP elements move, with minimal complexity and maximum reliability.

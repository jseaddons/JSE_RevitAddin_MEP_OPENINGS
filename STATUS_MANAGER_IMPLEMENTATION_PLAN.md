# StatusManager Service Implementation Plan

## 🎯 **Overview**
This document provides a detailed implementation plan for creating a comprehensive StatusManager service that provides real-time updates, progress tracking, and enhanced user feedback for the MEP Openings application.

## 📋 **Current State Analysis**

### **Existing Infrastructure:**
- ✅ **DebugLogger.cs** - Robust logging system with file-based output
- ✅ **MVVM Pattern** - Using CommunityToolkit.Mvvm with ObservableObject
- ✅ **Service Architecture** - Well-structured services in Services/ folder
- ✅ **Event-Driven Design** - Some services use Action<string> delegates for logging

### **Current Limitations:**
- ❌ **No Real-time UI Updates** - Logging only goes to files
- ❌ **No Progress Tracking** - Services don't report progress to UI
- ❌ **No Status Management** - No centralized status system
- ❌ **Limited User Feedback** - Users don't see what's happening during operations

## 🏗️ **Implementation Architecture**

### **1. Core StatusManager Service**

```csharp
// Services/StatusManager.cs
public class StatusManager : INotifyPropertyChanged
{
    // Events for real-time updates
    public event EventHandler<StatusUpdateEventArgs> StatusUpdated;
    public event EventHandler<ProgressUpdateEventArgs> ProgressUpdated;
    public event EventHandler<OperationCompletedEventArgs> OperationCompleted;
    
    // Properties for UI binding
    public string CurrentStatus { get; private set; }
    public StatusType StatusType { get; private set; }
    public int ProgressPercentage { get; private set; }
    public string ProgressMessage { get; private set; }
    public bool IsOperationRunning { get; private set; }
    public ObservableCollection<StatusItem> StatusHistory { get; }
    
    // Methods for status updates
    public void UpdateStatus(string message, StatusType type);
    public void UpdateProgress(int current, int total, string operation);
    public void StartOperation(string operationName);
    public void CompleteOperation(string operationName, bool success);
    public void ClearHistory();
}
```

### **2. Supporting Classes**

```csharp
// Models/StatusItem.cs
public class StatusItem
{
    public DateTime Timestamp { get; set; }
    public string Message { get; set; }
    public StatusType Type { get; set; }
    public string Operation { get; set; }
}

// Models/StatusUpdateEventArgs.cs
public class StatusUpdateEventArgs : EventArgs
{
    public string Message { get; set; }
    public StatusType Type { get; set; }
    public DateTime Timestamp { get; set; }
}

// Models/ProgressUpdateEventArgs.cs
public class ProgressUpdateEventArgs : EventArgs
{
    public int Current { get; set; }
    public int Total { get; set; }
    public string Operation { get; set; }
    public double Percentage { get; set; }
}

// Enums/StatusType.cs
public enum StatusType
{
    Info,
    Warning,
    Error,
    Success,
    Processing,
    Debug
}
```

## 🚀 **Implementation Phases**

### **Phase 1: Core StatusManager Foundation (Week 1)**

#### **Step 1.1: Create Core Classes**
- [ ] Create `StatusManager.cs` in Services folder
- [ ] Create `StatusItem.cs` in Models folder
- [ ] Create event argument classes
- [ ] Create `StatusType` enum

#### **Step 1.2: Implement Basic Functionality**
- [ ] Implement INotifyPropertyChanged
- [ ] Add status update methods
- [ ] Add progress tracking methods
- [ ] Add operation lifecycle management

#### **Step 1.3: Integration with Existing Logger**
- [ ] Extend DebugLogger to work with StatusManager
- [ ] Create StatusLogger wrapper class
- [ ] Maintain backward compatibility

### **Phase 2: UI Integration (Week 2)**

#### **Step 2.1: Create StatusViewModel**
- [ ] Create `StatusViewModel.cs` in ViewModels folder
- [ ] Implement data binding for status display
- [ ] Add commands for status management

#### **Step 2.2: Create Status UI Components**
- [ ] Create `StatusPanel.xaml` in Views folder
- [ ] Add progress bar with real-time updates
- [ ] Add status message display
- [ ] Add status history list view

#### **Step 2.3: Integrate with Main Dialog**
- [ ] Update main dialog to include status panel
- [ ] Add tabbed interface for different MEP types
- [ ] Implement status display in each tab

### **Phase 3: Service Integration (Week 3)**

#### **Step 3.1: Update Existing Services**
- [ ] Modify `ProgressiveMepSleeveService` to use StatusManager
- [ ] Update `DuctSleevePlacerService` with progress tracking
- [ ] Update `PipeSleevePlacerService` with status updates
- [ ] Update `CableTraySleevePlacer` with real-time feedback

#### **Step 3.2: Add Progress Tracking**
- [ ] Implement progress reporting in batch operations
- [ ] Add cancellation support for long-running operations
- [ ] Add error handling with user-friendly messages

#### **Step 3.3: Enhanced Error Handling**
- [ ] Create error recovery mechanisms
- [ ] Add user-friendly error messages
- [ ] Implement operation rollback where possible

### **Phase 4: Advanced Features (Week 4)**

#### **Step 4.1: Advanced UI Features**
- [ ] Add status filtering (Info, Warning, Error, etc.)
- [ ] Add status export functionality
- [ ] Add status search and filtering
- [ ] Add status statistics and reporting

#### **Step 4.2: Performance Optimizations**
- [ ] Implement background processing for heavy operations
- [ ] Add memory management for large batch operations
- [ ] Optimize UI updates for better performance

#### **Step 4.3: Settings and Configuration**
- [ ] Add status display preferences
- [ ] Add log level configuration
- [ ] Add status history retention settings

## 📁 **File Structure**

```
Services/
├── StatusManager.cs                 # Core status management service
├── StatusLogger.cs                  # Logger wrapper for StatusManager
├── ProgressTracker.cs               # Progress tracking utilities
└── OperationScheduler.cs            # Operation scheduling and management

Models/
├── StatusItem.cs                    # Status item model
├── StatusUpdateEventArgs.cs         # Status update event arguments
├── ProgressUpdateEventArgs.cs       # Progress update event arguments
└── OperationCompletedEventArgs.cs   # Operation completion event arguments

ViewModels/
├── StatusViewModel.cs               # Status display view model
├── ProgressViewModel.cs             # Progress display view model
└── MainDialogViewModel.cs           # Updated main dialog view model

Views/
├── StatusPanel.xaml                 # Status display panel
├── ProgressPanel.xaml               # Progress display panel
├── StatusHistoryView.xaml           # Status history viewer
└── MainDialog.xaml                  # Updated main dialog

Enums/
└── StatusType.cs                    # Status type enumeration
```

## 🔧 **Technical Implementation Details**

### **1. StatusManager Core Implementation**

```csharp
public class StatusManager : INotifyPropertyChanged
{
    private readonly object _lockObject = new object();
    private readonly ObservableCollection<StatusItem> _statusHistory;
    private string _currentStatus = "Ready";
    private StatusType _statusType = StatusType.Info;
    private int _progressPercentage = 0;
    private string _progressMessage = "";
    private bool _isOperationRunning = false;

    public StatusManager()
    {
        _statusHistory = new ObservableCollection<StatusItem>();
        StatusHistory = new ReadOnlyObservableCollection<StatusItem>(_statusHistory);
    }

    public void UpdateStatus(string message, StatusType type)
    {
        lock (_lockObject)
        {
            CurrentStatus = message;
            StatusType = type;
            
            var statusItem = new StatusItem
            {
                Timestamp = DateTime.Now,
                Message = message,
                Type = type,
                Operation = _currentOperation
            };
            
            _statusHistory.Add(statusItem);
            
            // Limit history size to prevent memory issues
            if (_statusHistory.Count > 1000)
            {
                _statusHistory.RemoveAt(0);
            }
            
            OnPropertyChanged(nameof(CurrentStatus));
            OnPropertyChanged(nameof(StatusType));
            
            StatusUpdated?.Invoke(this, new StatusUpdateEventArgs
            {
                Message = message,
                Type = type,
                Timestamp = DateTime.Now
            });
        }
    }

    public void UpdateProgress(int current, int total, string operation)
    {
        lock (_lockObject)
        {
            ProgressPercentage = total > 0 ? (int)((double)current / total * 100) : 0;
            ProgressMessage = $"{operation}: {current}/{total}";
            
            OnPropertyChanged(nameof(ProgressPercentage));
            OnPropertyChanged(nameof(ProgressMessage));
            
            ProgressUpdated?.Invoke(this, new ProgressUpdateEventArgs
            {
                Current = current,
                Total = total,
                Operation = operation,
                Percentage = ProgressPercentage
            });
        }
    }
}
```

### **2. Service Integration Pattern**

```csharp
public class DuctSleevePlacerService
{
    private readonly StatusManager _statusManager;
    
    public DuctSleevePlacerService(Document doc, StatusManager statusManager, ...)
    {
        _statusManager = statusManager;
        // ... other initialization
    }
    
    public void PlaceAllDuctSleeves()
    {
        _statusManager.StartOperation("Duct Sleeve Placement");
        _statusManager.UpdateStatus("Starting duct sleeve placement...", StatusType.Info);
        
        try
        {
            int total = _ductTuples.Count;
            for (int i = 0; i < total; i++)
            {
                var (duct, transform) = _ductTuples[i];
                
                _statusManager.UpdateProgress(i + 1, total, "Placing duct sleeves");
                _statusManager.UpdateStatus($"Processing duct {duct.Id}...", StatusType.Processing);
                
                // ... placement logic
                
                if (success)
                {
                    PlacedCount++;
                }
                else
                {
                    SkippedCount++;
                    _statusManager.UpdateStatus($"Skipped duct {duct.Id}: {reason}", StatusType.Warning);
                }
            }
            
            _statusManager.CompleteOperation("Duct Sleeve Placement", true);
            _statusManager.UpdateStatus($"Duct sleeve placement completed. Placed: {PlacedCount}, Skipped: {SkippedCount}", StatusType.Success);
        }
        catch (Exception ex)
        {
            _statusManager.CompleteOperation("Duct Sleeve Placement", false);
            _statusManager.UpdateStatus($"Error during duct sleeve placement: {ex.Message}", StatusType.Error);
            throw;
        }
    }
}
```

### **3. UI Integration Pattern**

```xml
<!-- StatusPanel.xaml -->
<UserControl x:Class="JSE_RevitAddin_MEP_OPENINGS.Views.StatusPanel"
             xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
             xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
        </Grid.RowDefinitions>
        
        <!-- Current Status -->
        <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="5">
            <TextBlock Text="Status: " FontWeight="Bold"/>
            <TextBlock Text="{Binding CurrentStatus}" 
                       Foreground="{Binding StatusColor}"/>
        </StackPanel>
        
        <!-- Progress Bar -->
        <StackPanel Grid.Row="1" Margin="5">
            <ProgressBar Value="{Binding ProgressPercentage}" 
                         Maximum="100" 
                         Height="20"/>
            <TextBlock Text="{Binding ProgressMessage}" 
                       HorizontalAlignment="Center" 
                       FontSize="10"/>
        </StackPanel>
        
        <!-- Status History -->
        <ListView Grid.Row="2" 
                  ItemsSource="{Binding StatusHistory}"
                  Margin="5">
            <ListView.ItemTemplate>
                <DataTemplate>
                    <StackPanel Orientation="Horizontal">
                        <TextBlock Text="{Binding Timestamp, StringFormat='{}{0:HH:mm:ss}'}" 
                                   Width="60" 
                                   FontSize="10"/>
                        <TextBlock Text="{Binding Type}" 
                                   Width="80" 
                                   FontSize="10"/>
                        <TextBlock Text="{Binding Message}" 
                                   FontSize="10"/>
                    </StackPanel>
                </DataTemplate>
            </ListView.ItemTemplate>
        </ListView>
    </Grid>
</UserControl>
```

## 🎯 **Integration Points**

### **1. Existing Services Integration**
- **ProgressiveMepSleeveService** - Add progress tracking for batch operations
- **DuctSleevePlacerService** - Add real-time status updates
- **PipeSleevePlacerService** - Add progress reporting
- **CableTraySleevePlacer** - Add status feedback

### **2. UI Integration**
- **Main Dialog** - Add status panel to main interface
- **Tabbed Interface** - Show status for each MEP type
- **Progress Indicators** - Real-time progress bars
- **Status History** - Scrollable status log

### **3. Logging Integration**
- **DebugLogger** - Extend to work with StatusManager
- **File Logging** - Maintain existing file-based logging
- **UI Logging** - Add real-time UI status updates

## 📊 **Success Metrics**

### **Phase 1 Success Criteria:**
- [ ] StatusManager service created and functional
- [ ] Basic status updates working
- [ ] Progress tracking implemented
- [ ] Integration with existing logger complete

### **Phase 2 Success Criteria:**
- [ ] Status UI components created
- [ ] Real-time status updates in UI
- [ ] Progress bars working
- [ ] Status history display functional

### **Phase 3 Success Criteria:**
- [ ] All existing services updated
- [ ] Progress tracking in all operations
- [ ] Error handling improved
- [ ] User feedback enhanced

### **Phase 4 Success Criteria:**
- [ ] Advanced UI features working
- [ ] Performance optimized
- [ ] Settings and configuration complete
- [ ] Full integration successful

## 🚨 **Risk Mitigation**

### **Technical Risks:**
- **UI Thread Issues** - Use Dispatcher for UI updates
- **Memory Leaks** - Implement proper cleanup and disposal
- **Performance Impact** - Use background processing for heavy operations
- **Thread Safety** - Use locks and thread-safe collections

### **Integration Risks:**
- **Breaking Changes** - Maintain backward compatibility
- **Service Dependencies** - Use dependency injection
- **UI Responsiveness** - Use async/await patterns
- **Error Propagation** - Implement proper error handling

## 📅 **Timeline**

- **Week 1:** Core StatusManager implementation
- **Week 2:** UI integration and components
- **Week 3:** Service integration and progress tracking
- **Week 4:** Advanced features and optimization

## 🎉 **Expected Benefits**

1. **Enhanced User Experience** - Real-time feedback on operations
2. **Better Error Handling** - Clear error messages and recovery options
3. **Improved Debugging** - Comprehensive status logging and history
4. **Professional UI** - Modern, responsive interface with progress tracking
5. **Better Performance** - Background processing and progress reporting
6. **Easier Maintenance** - Centralized status management and logging

---

*This implementation plan provides a comprehensive roadmap for creating a professional-grade status management system that will significantly enhance the user experience and maintainability of the MEP Openings application.*

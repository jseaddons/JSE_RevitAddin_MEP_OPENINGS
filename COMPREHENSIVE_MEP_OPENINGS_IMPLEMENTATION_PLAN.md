# Comprehensive MEP Openings Implementation Plan

## 🎯 **Overview**
This document provides a detailed implementation plan for creating a professional MEP Openings application that matches the sophisticated conVoid interface, including profile management, discipline-based workflows, dynamic opening relationships, and real-time status tracking.

## 📺 **Complete System Analysis**

### **Key Features Discovered:**
1. **Profile-Based User Management** - User identification and responsibility tracking
2. **Discipline-Based Workflows** - Role-specific interfaces and permissions
3. **Dynamic Opening Relationships** - Auto-updating openings when MEP elements move
4. **Three-Panel Professional Interface** - Advanced filtering and configuration
5. **Real-Time Status Management** - Progress tracking and user feedback
6. **Configuration Persistence** - Save/load opening configurations
7. **Multi-Project Support** - Work with linked models and multiple projects

## 🏗️ **System Architecture**

### **1. Profile Management System**

#### **A. Profile Setup Dialog**
```xml
<!-- ProfileSetupDialog.xaml -->
<Window x:Class="JSE_RevitAddin_MEP_OPENINGS.Views.ProfileSetupDialog"
        Title="JSE MEP Openings - Profile Setup"
        Width="450" Height="600"
        WindowStartupLocation="CenterScreen"
        ResizeMode="NoResize">
    
    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>  <!-- Header -->
            <RowDefinition Height="Auto"/>  <!-- Language Settings -->
            <RowDefinition Height="Auto"/>  <!-- Profile Name -->
            <RowDefinition Height="Auto"/>  <!-- Discipline Selection -->
            <RowDefinition Height="Auto"/>  <!-- Warning Message -->
            <RowDefinition Height="*"/>     <!-- Spacer -->
            <RowDefinition Height="Auto"/>  <!-- Action Buttons -->
        </Grid.RowDefinitions>
        
        <!-- Header -->
        <TextBlock Grid.Row="0" Text="Profile Setup" FontSize="18" FontWeight="Bold" Margin="0,0,0,20"/>
        
        <!-- Language Settings -->
        <GroupBox Grid.Row="1" Header="Language Settings" Margin="0,0,0,15">
            <Grid Margin="10">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                
                <Grid Grid.Row="0" Margin="0,0,0,10">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    
                    <TextBlock Grid.Column="0" Text="Language User Interface:" VerticalAlignment="Center" Margin="0,0,10,0"/>
                    <ComboBox Grid.Column="1" 
                              ItemsSource="{Binding AvailableLanguages}"
                              SelectedItem="{Binding SelectedLanguage}"/>
                </Grid>
                
                <Grid Grid.Row="1">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    
                    <TextBlock Grid.Column="0" Text="Project Language*:" VerticalAlignment="Center" Margin="0,0,10,0"/>
                    <ComboBox Grid.Column="1" 
                              ItemsSource="{Binding ProjectLanguages}"
                              SelectedItem="{Binding SelectedProjectLanguage}"/>
                </Grid>
            </Grid>
        </GroupBox>
        
        <!-- Profile Name -->
        <GroupBox Grid.Row="2" Header="User Profile" Margin="0,0,0,15">
            <Grid Margin="10">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                
                <TextBlock Grid.Column="0" Text="Profile Name*:" VerticalAlignment="Center" Margin="0,0,10,0"/>
                <TextBox Grid.Column="1" 
                         Text="{Binding ProfileName}" 
                         PlaceholderText="Enter your profile name (e.g., John_Architect)"/>
            </Grid>
        </GroupBox>
        
        <!-- Discipline Selection -->
        <GroupBox Grid.Row="3" Header="Discipline Selection*" Margin="0,0,0,15">
            <Grid Margin="10">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                
                <!-- Primary Disciplines (Single Selection) -->
                <TextBlock Grid.Row="0" Text="Primary Disciplines (Select One):" FontWeight="Bold" Margin="0,0,0,10"/>
                <StackPanel Grid.Row="1" Margin="0,0,0,15">
                    <RadioButton Content="Architectural" 
                                 GroupName="PrimaryDiscipline" 
                                 IsChecked="{Binding IsArchitecturalSelected}"/>
                    <RadioButton Content="Structural" 
                                 GroupName="PrimaryDiscipline" 
                                 IsChecked="{Binding IsStructuralSelected}"/>
                    <RadioButton Content="Technical Planner" 
                                 GroupName="PrimaryDiscipline" 
                                 IsChecked="{Binding IsTechnicalPlannerSelected}"/>
                    <RadioButton Content="Coordination" 
                                 GroupName="PrimaryDiscipline" 
                                 IsChecked="{Binding IsCoordinationSelected}"/>
                </StackPanel>
                
                <!-- MEP Disciplines (Multi Selection) -->
                <TextBlock Grid.Row="2" Text="MEP Disciplines (Can Select Multiple):" FontWeight="Bold" Margin="0,0,0,10"/>
                <StackPanel Grid.Row="3">
                    <CheckBox Content="Mechanical" IsChecked="{Binding IsMechanicalSelected}"/>
                    <CheckBox Content="Electrical" IsChecked="{Binding IsElectricalSelected}"/>
                    <CheckBox Content="Plumbing" IsChecked="{Binding IsPlumbingSelected}"/>
                    <CheckBox Content="Fire Protection" IsChecked="{Binding IsFireProtectionSelected}"/>
                </StackPanel>
            </Grid>
        </GroupBox>
        
        <!-- Warning Message -->
        <Border Grid.Row="4" 
                Background="#FFF3CD" 
                BorderBrush="#FFEAA7" 
                BorderThickness="1" 
                CornerRadius="5" 
                Padding="10" 
                Margin="0,0,0,15">
            <TextBlock Text="⚠️ You should not change the Settings in an ongoing project because your old initials will remain in the void elements." 
                       FontSize="11" 
                       TextWrapping="Wrap" 
                       Foreground="#856404"/>
        </Border>
        
        <!-- Action Buttons -->
        <StackPanel Grid.Row="6" Orientation="Horizontal" HorizontalAlignment="Right">
            <Button Content="OK" 
                    Command="{Binding CreateProfileCommand}" 
                    Style="{StaticResource PrimaryButtonStyle}" 
                    Width="80"/>
        </StackPanel>
    </Grid>
</Window>
```

#### **B. Profile Management Service**
```csharp
// Services/ProfileManagementService.cs
public class ProfileManagementService
{
    private readonly string _profileFilePath;
    private UserProfile _currentProfile;
    
    public event EventHandler<ProfileChangedEventArgs> ProfileChanged;
    
    public ProfileManagementService()
    {
        _profileFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), 
                                       "JSE_MEP_Openings", "profiles.xml");
    }
    
    public UserProfile CreateProfile(string profileName, List<Discipline> disciplines, string language)
    {
        // Validate discipline selection
        if (!ValidateDisciplineSelection(disciplines))
        {
            throw new InvalidOperationException("A combination of these disciplines is not allowed.");
        }
        
        var profile = new UserProfile
        {
            Id = Guid.NewGuid(),
            Name = profileName,
            Disciplines = disciplines,
            Language = language,
            CreatedDate = DateTime.Now,
            IsActive = true
        };
        
        SaveProfile(profile);
        _currentProfile = profile;
        
        ProfileChanged?.Invoke(this, new ProfileChangedEventArgs(profile));
        return profile;
    }
    
    private bool ValidateDisciplineSelection(List<Discipline> disciplines)
    {
        var primaryDisciplines = disciplines.Where(d => d.IsPrimary).ToList();
        return primaryDisciplines.Count <= 1; // Only one primary discipline allowed
    }
}

// Models/UserProfile.cs
public class UserProfile
{
    public Guid Id { get; set; }
    public string Name { get; set; }
    public List<Discipline> Disciplines { get; set; }
    public string Language { get; set; }
    public DateTime CreatedDate { get; set; }
    public bool IsActive { get; set; }
}

// Models/Discipline.cs
public class Discipline
{
    public string Name { get; set; }
    public bool IsPrimary { get; set; }
    public bool IsSelected { get; set; }
    public List<string> AllowedCombinations { get; set; }
}
```

### **2. Main Application Interface**

#### **A. Main Window with Three-Panel Layout**
```xml
<!-- MainDialog.xaml -->
<Window x:Class="JSE_RevitAddin_MEP_OPENINGS.Views.MainDialog"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="JSE MEP Openings - Professional Edition"
        Width="1400" Height="900"
        WindowStartupLocation="CenterScreen"
        ResizeMode="CanResize">
    
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>      <!-- Header -->
            <RowDefinition Height="*"/>         <!-- Main Content -->
            <RowDefinition Height="Auto"/>      <!-- Bottom Toolbar -->
        </Grid.RowDefinitions>
        
        <!-- Header Section -->
        <Border Grid.Row="0" Background="#2E3440" Padding="15">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                
                <StackPanel Grid.Column="0" Orientation="Horizontal">
                    <TextBlock Text="JSE MEP Openings" 
                               FontSize="20" 
                               FontWeight="Bold" 
                               Foreground="White" 
                               VerticalAlignment="Center"/>
                    <TextBlock Text="Professional Edition" 
                               FontSize="12" 
                               Foreground="#D8DEE9" 
                               VerticalAlignment="Center" 
                               Margin="10,0,0,0"/>
                </StackPanel>
                
                <!-- User Profile Info -->
                <StackPanel Grid.Column="1" Orientation="Horizontal">
                    <TextBlock Text="Profile:" 
                               Foreground="#D8DEE9" 
                               VerticalAlignment="Center" 
                               Margin="0,0,5,0"/>
                    <TextBlock Text="{Binding CurrentProfile.Name}" 
                               Foreground="White" 
                               FontWeight="Bold" 
                               VerticalAlignment="Center" 
                               Margin="0,0,10,0"/>
                    <Button Content="Change Profile" 
                            Command="{Binding ChangeProfileCommand}" 
                            Style="{StaticResource SecondaryButtonStyle}"/>
                </StackPanel>
            </Grid>
        </Border>
        
        <!-- Main Content Area - Three Panel Layout -->
        <Grid Grid.Row="1">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="300"/>  <!-- Left Panel -->
                <ColumnDefinition Width="5"/>    <!-- Splitter -->
                <ColumnDefinition Width="400"/>  <!-- Middle Panel -->
                <ColumnDefinition Width="5"/>    <!-- Splitter -->
                <ColumnDefinition Width="*"/>    <!-- Right Panel -->
            </Grid.ColumnDefinitions>
            
            <!-- Left Panel: Project & Element Management -->
            <local:LeftPanel Grid.Column="0" DataContext="{Binding LeftPanelViewModel}"/>
            
            <!-- Splitter 1 -->
            <GridSplitter Grid.Column="1" HorizontalAlignment="Stretch" Background="#4C566A"/>
            
            <!-- Middle Panel: Categories & Opening Types -->
            <local:MiddlePanel Grid.Column="2" DataContext="{Binding MiddlePanelViewModel}"/>
            
            <!-- Splitter 2 -->
            <GridSplitter Grid.Column="3" HorizontalAlignment="Stretch" Background="#4C566A"/>
            
            <!-- Right Panel: Conditions & Configuration -->
            <local:RightPanel Grid.Column="4" DataContext="{Binding RightPanelViewModel}"/>
        </Grid>
        
        <!-- Bottom Toolbar -->
        <Border Grid.Row="2" Background="#F8F9FA" BorderBrush="#DEE2E6" BorderThickness="0,1,0,0">
            <Grid Margin="15,10">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                
                <!-- Left Actions -->
                <StackPanel Grid.Column="0" Orientation="Horizontal">
                    <Button Style="{StaticResource IconButtonStyle}" Content="←" ToolTip="Previous" Command="{Binding PreviousCommand}"/>
                    <Button Style="{StaticResource IconButtonStyle}" Content="→" ToolTip="Next" Command="{Binding NextCommand}"/>
                    <Button Style="{StaticResource IconButtonStyle}" Content="⚡" ToolTip="Process Openings" Command="{Binding ProcessCommand}"/>
                    <Button Style="{StaticResource IconButtonStyle}" Content="🔧" ToolTip="Settings" Command="{Binding SettingsCommand}"/>
                    <Button Style="{StaticResource IconButtonStyle}" Content="🔄" ToolTip="Refresh" Command="{Binding RefreshCommand}"/>
                    <ComboBox ItemsSource="{Binding OpeningModes}" 
                              SelectedItem="{Binding SelectedOpeningMode}" 
                              Width="120" 
                              Margin="10,0,0,0"/>
                </StackPanel>
                
                <!-- Center Progress -->
                <StackPanel Grid.Column="1" Orientation="Horizontal" HorizontalAlignment="Center">
                    <ProgressBar Value="{Binding ProgressPercentage}" 
                                 Maximum="100" 
                                 Width="200" 
                                 Height="20" 
                                 Margin="0,0,10,0"/>
                    <TextBlock Text="{Binding ProgressText}" 
                               VerticalAlignment="Center" 
                               FontWeight="Bold"/>
                </StackPanel>
                
                <!-- Right Actions -->
                <StackPanel Grid.Column="2" Orientation="Horizontal">
                    <Button Content="OK" 
                            Command="{Binding OkCommand}" 
                            Style="{StaticResource PrimaryButtonStyle}" 
                            Margin="0,0,10,0"/>
                    <Button Content="Cancel" 
                            Command="{Binding CancelCommand}" 
                            Style="{StaticResource SecondaryButtonStyle}" 
                            Margin="0,0,10,0"/>
                    <Button Content="Save" 
                            Command="{Binding SaveCommand}" 
                            Style="{StaticResource SecondaryButtonStyle}"/>
                </StackPanel>
            </Grid>
        </Border>
    </Grid>
</Window>
```

#### **B. Left Panel - Project & Element Management**
```xml
<!-- LeftPanel.xaml -->
<UserControl x:Class="JSE_RevitAddin_MEP_OPENINGS.Views.LeftPanel">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>  <!-- Filters -->
            <RowDefinition Height="Auto"/>  <!-- Projects -->
            <RowDefinition Height="Auto"/>  <!-- Reference Elements -->
            <RowDefinition Height="Auto"/>  <!-- Host Elements -->
            <RowDefinition Height="*"/>     <!-- Spacer -->
            <RowDefinition Height="Auto"/>  <!-- Bottom Icons -->
        </Grid.RowDefinitions>
        
        <!-- Filters Section -->
        <GroupBox Grid.Row="0" Header="Filters" Margin="0,0,0,10">
            <Grid Margin="10">
                <TextBox Text="{Binding FilterText}" 
                         PlaceholderText="Enter filter text..."
                         Margin="0,0,0,10"/>
                <ListBox ItemsSource="{Binding FilterResults}" 
                         Height="150"
                         SelectionMode="Multiple"/>
            </Grid>
        </GroupBox>
        
        <!-- Projects and Categories -->
        <GroupBox Grid.Row="1" Header="Projects and Categories" Margin="0,0,0,10">
            <Grid Margin="10">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                
                <!-- Project 1 -->
                <Grid Grid.Row="0" Margin="0,0,0,10">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="60"/>
                        <ColumnDefinition Width="80"/>
                    </Grid.ColumnDefinitions>
                    
                    <ComboBox Grid.Column="0" 
                              ItemsSource="{Binding ProjectSections}"
                              SelectedItem="{Binding SelectedSection1}"
                              Margin="0,0,5,0"/>
                    <TextBox Grid.Column="1" 
                             Text="{Binding Offset1}" 
                             Margin="0,0,5,0"/>
                    <TextBox Grid.Column="2" 
                             Text="{Binding TopValue1}" 
                             Tag="Top"/>
                </Grid>
                
                <!-- Project 2 -->
                <Grid Grid.Row="1">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="60"/>
                        <ColumnDefinition Width="80"/>
                    </Grid.ColumnDefinitions>
                    
                    <ComboBox Grid.Column="0" 
                              ItemsSource="{Binding ProjectSections}"
                              SelectedItem="{Binding SelectedSection2}"
                              Margin="0,0,5,0"/>
                    <TextBox Grid.Column="1" 
                             Text="{Binding Offset2}" 
                             Margin="0,0,5,0"/>
                    <TextBox Grid.Column="2" 
                             Text="{Binding BottomValue2}" 
                             Tag="Bottom"/>
                </Grid>
            </Grid>
        </GroupBox>
        
        <!-- Reference Elements (MEP Elements to Track) -->
        <GroupBox Grid.Row="2" Header="Reference Elements of the Openings" Margin="0,0,0,10">
            <ListBox ItemsSource="{Binding ReferenceElements}" 
                     SelectionMode="Multiple"
                     Height="120">
                <ListBox.ItemTemplate>
                    <DataTemplate>
                        <CheckBox Content="{Binding Name}" 
                                  IsChecked="{Binding IsSelected}"/>
                    </DataTemplate>
                </ListBox.ItemTemplate>
            </ListBox>
        </GroupBox>
        
        <!-- Host Elements (Structural Elements for Openings) -->
        <GroupBox Grid.Row="3" Header="Host Elements of the Openings" Margin="0,0,0,10">
            <ListBox ItemsSource="{Binding HostElements}" 
                     SelectionMode="Multiple"
                     Height="120">
                <ListBox.ItemTemplate>
                    <DataTemplate>
                        <CheckBox Content="{Binding Name}" 
                                  IsChecked="{Binding IsSelected}"/>
                    </DataTemplate>
                </ListBox.ItemTemplate>
            </ListBox>
        </GroupBox>
        
        <!-- Bottom Icons -->
        <StackPanel Grid.Row="5" Orientation="Horizontal" HorizontalAlignment="Center" VerticalAlignment="Bottom">
            <Button Style="{StaticResource IconButtonStyle}" Content="⚙️" ToolTip="Settings" Command="{Binding SettingsCommand}"/>
            <Button Style="{StaticResource IconButtonStyle}" Content="📄" ToolTip="Document" Command="{Binding DocumentCommand}"/>
            <Button Style="{StaticResource IconButtonStyle}" Content="📋" ToolTip="Clipboard" Command="{Binding ClipboardCommand}"/>
            <Button Style="{StaticResource IconButtonStyle}" Content="🤖" ToolTip="AI Assistant" Command="{Binding AIAssistantCommand}"/>
            <Button Style="{StaticResource IconButtonStyle}" Content="💾" ToolTip="Save Configuration" Command="{Binding SaveConfigCommand}"/>
        </StackPanel>
    </Grid>
</UserControl>
```

### **3. Dynamic Opening Relationship System**

#### **A. Opening Relationship Service**
```csharp
// Services/OpeningRelationshipService.cs
public class OpeningRelationshipService
{
    private readonly Document _document;
    private readonly Dictionary<ElementId, List<ElementId>> _openingToReferenceMap;
    private readonly Dictionary<ElementId, ElementId> _referenceToOpeningMap;
    private readonly Dictionary<ElementId, OpeningRelationship> _relationships;
    
    public event EventHandler<ElementMovedEventArgs> ElementMoved;
    public event EventHandler<OpeningUpdatedEventArgs> OpeningUpdated;
    
    public OpeningRelationshipService(Document document)
    {
        _document = document;
        _openingToReferenceMap = new Dictionary<ElementId, List<ElementId>>();
        _referenceToOpeningMap = new Dictionary<ElementId, ElementId>();
        _relationships = new Dictionary<ElementId, OpeningRelationship>();
    }
    
    public void CreateRelationship(ElementId openingId, ElementId referenceElementId, ElementId hostElementId, string profileName)
    {
        var relationship = new OpeningRelationship
        {
            OpeningId = openingId,
            ReferenceElementId = referenceElementId,
            HostElementId = hostElementId,
            CreatedBy = profileName,
            CreatedDate = DateTime.Now,
            LastUpdated = DateTime.Now
        };
        
        _relationships[openingId] = relationship;
        
        // Update maps
        if (!_openingToReferenceMap.ContainsKey(openingId))
            _openingToReferenceMap[openingId] = new List<ElementId>();
        _openingToReferenceMap[openingId].Add(referenceElementId);
        
        _referenceToOpeningMap[referenceElementId] = openingId;
        
        // Tag the opening with profile information
        TagOpeningWithProfile(openingId, profileName);
    }
    
    public void OnElementMoved(ElementId elementId, Transform newTransform)
    {
        if (_referenceToOpeningMap.TryGetValue(elementId, out var openingId))
        {
            UpdateOpeningPosition(openingId, newTransform);
            ElementMoved?.Invoke(this, new ElementMovedEventArgs(elementId, openingId, newTransform));
        }
    }
    
    private void UpdateOpeningPosition(ElementId openingId, Transform newTransform)
    {
        try
        {
            using (var transaction = new Transaction(_document, "Update Opening Position"))
            {
                transaction.Start();
                
                var opening = _document.GetElement(openingId);
                if (opening != null)
                {
                    // Update opening position based on new reference element position
                    // Implementation depends on opening type and family
                    UpdateOpeningGeometry(opening, newTransform);
                    
                    // Update relationship timestamp
                    if (_relationships.TryGetValue(openingId, out var relationship))
                    {
                        relationship.LastUpdated = DateTime.Now;
                    }
                }
                
                transaction.Commit();
                OpeningUpdated?.Invoke(this, new OpeningUpdatedEventArgs(openingId, "Position updated due to reference element movement"));
            }
        }
        catch (Exception ex)
        {
            // Log error and notify user
            DebugLogger.Error($"Failed to update opening {openingId}: {ex.Message}");
        }
    }
    
    private void TagOpeningWithProfile(ElementId openingId, string profileName)
    {
        var opening = _document.GetElement(openingId);
        if (opening != null)
        {
            // Set custom parameters to track profile information
            var createdByParam = opening.LookupParameter("Created By");
            if (createdByParam != null && !createdByParam.IsReadOnly)
            {
                createdByParam.Set(profileName);
            }
            
            var createdDateParam = opening.LookupParameter("Created Date");
            if (createdDateParam != null && !createdDateParam.IsReadOnly)
            {
                createdDateParam.Set(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            }
        }
    }
}

// Models/OpeningRelationship.cs
public class OpeningRelationship
{
    public ElementId OpeningId { get; set; }
    public ElementId ReferenceElementId { get; set; }
    public ElementId HostElementId { get; set; }
    public string CreatedBy { get; set; }
    public DateTime CreatedDate { get; set; }
    public DateTime LastUpdated { get; set; }
    public string RelationshipType { get; set; }
}
```

#### **B. Change Detection Service**
```csharp
// Services/ElementChangeDetectionService.cs
public class ElementChangeDetectionService
{
    private readonly Document _document;
    private readonly Dictionary<ElementId, Transform> _elementPositions;
    private Timer _monitoringTimer;
    
    public event EventHandler<ElementMovedEventArgs> ElementMoved;
    public event EventHandler<ElementDeletedEventArgs> ElementDeleted;
    public event EventHandler<ElementAddedEventArgs> ElementAdded;
    
    public ElementChangeDetectionService(Document document)
    {
        _document = document;
        _elementPositions = new Dictionary<ElementId, Transform>();
        InitializeElementTracking();
    }
    
    public void StartMonitoring()
    {
        _monitoringTimer = new Timer(CheckForChanges, null, TimeSpan.Zero, TimeSpan.FromSeconds(5));
    }
    
    public void StopMonitoring()
    {
        _monitoringTimer?.Dispose();
    }
    
    private void InitializeElementTracking()
    {
        // Track initial positions of all MEP elements
        var mepElements = new FilteredElementCollector(_document)
            .OfClass(typeof(Element))
            .WhereElementIsNotElementType()
            .Where(e => IsMepElement(e));
        
        foreach (Element element in mepElements)
        {
            var transform = GetElementTransform(element);
            if (transform != null)
            {
                _elementPositions[element.Id] = transform;
            }
        }
    }
    
    private void CheckForChanges(object state)
    {
        try
        {
            var currentElements = new FilteredElementCollector(_document)
                .OfClass(typeof(Element))
                .WhereElementIsNotElementType()
                .Where(e => IsMepElement(e));
            
            var currentElementIds = currentElements.Select(e => e.Id).ToHashSet();
            
            // Check for moved elements
            foreach (var element in currentElements)
            {
                var currentTransform = GetElementTransform(element);
                if (currentTransform != null)
                {
                    if (_elementPositions.TryGetValue(element.Id, out var previousTransform))
                    {
                        if (!TransformsEqual(currentTransform, previousTransform))
                        {
                            _elementPositions[element.Id] = currentTransform;
                            ElementMoved?.Invoke(this, new ElementMovedEventArgs(element.Id, element.Id, currentTransform));
                        }
                    }
                    else
                    {
                        _elementPositions[element.Id] = currentTransform;
                        ElementAdded?.Invoke(this, new ElementAddedEventArgs(element.Id));
                    }
                }
            }
            
            // Check for deleted elements
            var deletedElements = _elementPositions.Keys.Except(currentElementIds).ToList();
            foreach (var deletedId in deletedElements)
            {
                _elementPositions.Remove(deletedId);
                ElementDeleted?.Invoke(this, new ElementDeletedEventArgs(deletedId));
            }
        }
        catch (Exception ex)
        {
            DebugLogger.Error($"Error in change detection: {ex.Message}");
        }
    }
    
    private bool IsMepElement(Element element)
    {
        // Check if element is MEP-related (ducts, pipes, cable trays, etc.)
        var category = element.Category;
        return category != null && (
            category.Id.IntegerValue == (int)BuiltInCategory.OST_DuctCurves ||
            category.Id.IntegerValue == (int)BuiltInCategory.OST_PipeCurves ||
            category.Id.IntegerValue == (int)BuiltInCategory.OST_ElectricalFixtures ||
            category.Id.IntegerValue == (int)BuiltInCategory.OST_CableTray ||
            category.Id.IntegerValue == (int)BuiltInCategory.OST_Conduit
        );
    }
    
    private Transform GetElementTransform(Element element)
    {
        // Get element's transform based on its type
        if (element is FamilyInstance familyInstance)
        {
            return familyInstance.GetTransform();
        }
        else if (element is CurveElement curveElement)
        {
            // For curves, get the transform of the first point
            var curve = curveElement.GeometryCurve;
            if (curve != null)
            {
                var startPoint = curve.GetEndPoint(0);
                return Transform.CreateTranslation(startPoint);
            }
        }
        
        return null;
    }
    
    private bool TransformsEqual(Transform t1, Transform t2)
    {
        const double tolerance = 0.001;
        return Math.Abs(t1.Origin.X - t2.Origin.X) < tolerance &&
               Math.Abs(t1.Origin.Y - t2.Origin.Y) < tolerance &&
               Math.Abs(t1.Origin.Z - t2.Origin.Z) < tolerance;
    }
}
```

### **4. Status Management System**

#### **A. Enhanced Status Manager**
```csharp
// Services/StatusManager.cs
public class StatusManager : INotifyPropertyChanged
{
    private readonly object _lockObject = new object();
    private readonly ObservableCollection<StatusItem> _statusHistory;
    private readonly Dictionary<string, OperationProgress> _operationProgress;
    
    private string _currentStatus = "Ready";
    private StatusType _statusType = StatusType.Info;
    private int _overallProgressPercentage = 0;
    private string _progressMessage = "";
    private bool _isOperationRunning = false;
    private string _currentOperation = "";
    
    public event EventHandler<StatusUpdateEventArgs> StatusUpdated;
    public event EventHandler<ProgressUpdateEventArgs> ProgressUpdated;
    public event EventHandler<OperationCompletedEventArgs> OperationCompleted;
    
    public StatusManager()
    {
        _statusHistory = new ObservableCollection<StatusItem>();
        _operationProgress = new Dictionary<string, OperationProgress>();
        StatusHistory = new ReadOnlyObservableCollection<StatusItem>(_statusHistory);
    }
    
    public string CurrentStatus
    {
        get => _currentStatus;
        private set
        {
            if (_currentStatus != value)
            {
                _currentStatus = value;
                OnPropertyChanged();
            }
        }
    }
    
    public StatusType StatusType
    {
        get => _statusType;
        private set
        {
            if (_statusType != value)
            {
                _statusType = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(StatusColor));
                OnPropertyChanged(nameof(StatusBackgroundColor));
            }
        }
    }
    
    public int OverallProgressPercentage
    {
        get => _overallProgressPercentage;
        private set
        {
            if (_overallProgressPercentage != value)
            {
                _overallProgressPercentage = value;
                OnPropertyChanged();
            }
        }
    }
    
    public string ProgressMessage
    {
        get => _progressMessage;
        private set
        {
            if (_progressMessage != value)
            {
                _progressMessage = value;
                OnPropertyChanged();
            }
        }
    }
    
    public bool IsOperationRunning
    {
        get => _isOperationRunning;
        private set
        {
            if (_isOperationRunning != value)
            {
                _isOperationRunning = value;
                OnPropertyChanged();
            }
        }
    }
    
    public ReadOnlyObservableCollection<StatusItem> StatusHistory { get; }
    
    public Brush StatusColor
    {
        get
        {
            return StatusType switch
            {
                StatusType.Success => Brushes.Green,
                StatusType.Warning => Brushes.Orange,
                StatusType.Error => Brushes.Red,
                StatusType.Processing => Brushes.Blue,
                StatusType.Info => Brushes.DarkBlue,
                _ => Brushes.Black
            };
        }
    }
    
    public Brush StatusBackgroundColor
    {
        get
        {
            return StatusType switch
            {
                StatusType.Success => new SolidColorBrush(Color.FromRgb(163, 190, 140)), // Green
                StatusType.Warning => new SolidColorBrush(Color.FromRgb(235, 203, 139)), // Yellow
                StatusType.Error => new SolidColorBrush(Color.FromRgb(191, 97, 106)),   // Red
                StatusType.Processing => new SolidColorBrush(Color.FromRgb(136, 192, 208)), // Blue
                StatusType.Info => new SolidColorBrush(Color.FromRgb(94, 129, 172)),    // Dark Blue
                _ => new SolidColorBrush(Color.FromRgb(216, 222, 233)) // Light Gray
            };
        }
    }
    
    public void StartOperation(string operationName)
    {
        lock (_lockObject)
        {
            _currentOperation = operationName;
            IsOperationRunning = true;
            
            _operationProgress[operationName] = new OperationProgress
            {
                OperationName = operationName,
                StartTime = DateTime.Now,
                CurrentStep = 0,
                TotalSteps = 0
            };
            
            UpdateStatus($"Starting {operationName}...", StatusType.Processing);
        }
    }
    
    public void UpdateProgress(string operationName, int current, int total, string stepDescription = "")
    {
        lock (_lockObject)
        {
            if (_operationProgress.TryGetValue(operationName, out var progress))
            {
                progress.CurrentStep = current;
                progress.TotalSteps = total;
                progress.CurrentStepDescription = stepDescription;
                progress.LastUpdated = DateTime.Now;
                
                // Calculate overall progress
                CalculateOverallProgress();
                
                ProgressMessage = string.IsNullOrEmpty(stepDescription) 
                    ? $"{operationName}: {current}/{total}" 
                    : $"{stepDescription} ({current}/{total})";
                
                OnPropertyChanged(nameof(ProgressMessage));
                OnPropertyChanged(nameof(OverallProgressPercentage));
                
                ProgressUpdated?.Invoke(this, new ProgressUpdateEventArgs
                {
                    OperationName = operationName,
                    Current = current,
                    Total = total,
                    Percentage = OverallProgressPercentage,
                    StepDescription = stepDescription
                });
            }
        }
    }
    
    public void CompleteOperation(string operationName, bool success, string message = "")
    {
        lock (_lockObject)
        {
            if (_operationProgress.TryGetValue(operationName, out var progress))
            {
                progress.IsCompleted = true;
                progress.EndTime = DateTime.Now;
                progress.Success = success;
                
                var statusType = success ? StatusType.Success : StatusType.Error;
                var statusMessage = string.IsNullOrEmpty(message) 
                    ? $"{(success ? "Completed" : "Failed")} {operationName}" 
                    : message;
                
                UpdateStatus(statusMessage, statusType);
                
                OperationCompleted?.Invoke(this, new OperationCompletedEventArgs
                {
                    OperationName = operationName,
                    Success = success,
                    Message = statusMessage,
                    Duration = progress.EndTime - progress.StartTime
                });
            }
            
            // Check if all operations are complete
            if (_operationProgress.Values.All(p => p.IsCompleted))
            {
                IsOperationRunning = false;
                _currentOperation = "";
            }
        }
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
                Timestamp = DateTime.Now,
                Operation = _currentOperation
            });
        }
    }
    
    private void CalculateOverallProgress()
    {
        if (_operationProgress.Count == 0)
        {
            OverallProgressPercentage = 0;
            return;
        }
        
        var totalProgress = _operationProgress.Values
            .Where(p => !p.IsCompleted)
            .Sum(p => p.TotalSteps > 0 ? (double)p.CurrentStep / p.TotalSteps : 0);
        
        OverallProgressPercentage = (int)(totalProgress / _operationProgress.Count * 100);
    }
    
    public void ClearHistory()
    {
        lock (_lockObject)
        {
            _statusHistory.Clear();
            _operationProgress.Clear();
            CurrentStatus = "Ready";
            StatusType = StatusType.Info;
            OverallProgressPercentage = 0;
            ProgressMessage = "";
            IsOperationRunning = false;
            _currentOperation = "";
        }
    }
    
    public event PropertyChangedEventHandler PropertyChanged;
    
    protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}

// Models/OperationProgress.cs
public class OperationProgress
{
    public string OperationName { get; set; }
    public DateTime StartTime { get; set; }
    public DateTime? EndTime { get; set; }
    public DateTime LastUpdated { get; set; }
    public int CurrentStep { get; set; }
    public int TotalSteps { get; set; }
    public string CurrentStepDescription { get; set; }
    public bool IsCompleted { get; set; }
    public bool Success { get; set; }
}
```

## 🚀 **Implementation Phases**

### **Phase 1: Foundation & Profile System (Week 1-2)**

#### **Week 1: Core Infrastructure**
- [ ] Create project structure and dependencies
- [ ] Implement ProfileManagementService
- [ ] Create ProfileSetupDialog with discipline validation
- [ ] Implement basic StatusManager
- [ ] Set up MVVM infrastructure

#### **Week 2: Profile Integration**
- [ ] Integrate profile system with main application
- [ ] Implement discipline-based UI customization
- [ ] Create profile persistence and loading
- [ ] Add profile change functionality
- [ ] Implement basic validation and error handling

### **Phase 2: Main Interface & Three-Panel Layout (Week 3-4)**

#### **Week 3: UI Structure**
- [ ] Create MainDialog with three-panel layout
- [ ] Implement LeftPanel with project and element management
- [ ] Create MiddlePanel with categories and opening types
- [ ] Build RightPanel with conditions and configuration
- [ ] Add modern styling and themes

#### **Week 4: Data Binding & ViewModels**
- [ ] Create ViewModels for all panels
- [ ] Implement data binding for all controls
- [ ] Add parameter filtering functionality
- [ ] Create opening configuration logic
- [ ] Implement save/load configuration

### **Phase 3: Dynamic Relationships & Change Detection (Week 5-6)**

#### **Week 5: Relationship System**
- [ ] Implement OpeningRelationshipService
- [ ] Create relationship tracking and management
- [ ] Add element tagging with profile information
- [ ] Implement relationship persistence
- [ ] Create relationship visualization

#### **Week 6: Change Detection**
- [ ] Implement ElementChangeDetectionService
- [ ] Add real-time monitoring of element changes
- [ ] Create automatic opening updates
- [ ] Implement conflict detection and resolution
- [ ] Add change notification system

### **Phase 4: Advanced Features & Polish (Week 7-8)**

#### **Week 7: Advanced Features**
- [ ] Implement multi-project support
- [ ] Add linked model management
- [ ] Create batch processing capabilities
- [ ] Implement advanced filtering and search
- [ ] Add export/import functionality

#### **Week 8: Testing & Optimization**
- [ ] Comprehensive testing of all features
- [ ] Performance optimization
- [ ] Error handling and recovery
- [ ] User documentation and help system
- [ ] Final polish and deployment preparation

## 📊 **Success Metrics**

### **Phase 1 Success Criteria:**
- [ ] Profile system working with discipline validation
- [ ] Basic status management functional
- [ ] Profile persistence and loading working
- [ ] Error handling for invalid discipline combinations

### **Phase 2 Success Criteria:**
- [ ] Three-panel interface fully functional
- [ ] All data binding working correctly
- [ ] Configuration save/load working
- [ ] Modern UI styling applied

### **Phase 3 Success Criteria:**
- [ ] Opening relationships tracked and managed
- [ ] Change detection working in real-time
- [ ] Automatic opening updates functional
- [ ] Profile-based element tagging working

### **Phase 4 Success Criteria:**
- [ ] Multi-project support working
- [ ] All advanced features functional
- [ ] Performance optimized for large projects
- [ ] Ready for production deployment

## 🎯 **Key Business Logic Features**

### **1. Profile-Based Workflow:**
- **User identification** and responsibility tracking
- **Discipline-specific** interfaces and permissions
- **Element ownership** and accountability
- **Multi-user coordination** support

### **2. Dynamic Opening Management:**
- **Real-time relationship** tracking
- **Automatic updates** when MEP elements move
- **Change detection** and notification
- **Conflict resolution** and validation

### **3. Professional BIM Integration:**
- **Multi-project support** with linked models
- **Discipline coordination** and workflow management
- **Audit trail** and change tracking
- **Industry-standard** practices and compliance

### **4. Advanced Configuration:**
- **Parameter-based filtering** and selection
- **Template management** and reuse
- **Batch processing** capabilities
- **Export/import** functionality

## 🎉 **Expected Benefits**

1. **Professional BIM Tool** - Industry-standard workflow management
2. **Multi-User Coordination** - Team collaboration and responsibility tracking
3. **Dynamic Relationships** - Automatic updates and change management
4. **Scalable Architecture** - Support for large, complex projects
5. **User-Friendly Interface** - Intuitive three-panel design
6. **Comprehensive Status Tracking** - Real-time feedback and progress monitoring
7. **Configuration Management** - Save/load and template support
8. **Future-Proof Design** - Extensible architecture for new features

---

*This comprehensive implementation plan provides a complete roadmap for creating a professional-grade MEP Openings application that matches the sophistication of conVoid while incorporating modern development practices and user experience design.*

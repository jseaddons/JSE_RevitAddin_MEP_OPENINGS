# CONVOID-Style UI Implementation Plan (Revised)

## 🎯 **Overview**
This document provides a detailed plan for implementing a professional MEP Openings interface based on the complete CONVOID architecture from all 12 official lessons, featuring proper separation of Profile, Settings, Filters, and Conditions.

## 🔄 Complete CONVOID Architecture (All 12 Lessons + BCF Analysis)

### **👤 PROFILE** (Session 3 - Create Profile)
- **Profile Name** - Unique user identifier (e.g., "John Smith")
- **Discipline** - User's engineering discipline (HVAC, Plumbing, Electrical)
- **Profile Initials** - Auto-assigned for tracking (e.g., "JS")
- **Purpose**: Track who created/modified/approved openings

### **🔍 FILTER** (Session 3 - Create Openings + BCF Integration)
- **File Selections** - Reference files (Architectural, Structural)
- **Host Files** - Linked Revit models for clash detection
- **BCF Clash Data** - Detected intersections and clashes
- **BCF Status Data** - Approval statuses and comments
- **BCF Coordination** - Multi-disciplinary coordination data
- **Purpose**: Project state management and coordination

### **📋 BCF (BIM Collaboration Format)** (Session 10 - Coordination and Approval)
- **Clash Zones** - Detected intersections between MEP and structural elements
- **Status History** - Approval tracking and comments
- **MEP Modifications** - Change detection for element updates
- **Coordination Data** - Multi-disciplinary approval workflow
- **3D Viewpoints** - Camera positions for clash visualization
- **Purpose**: Industry-standard BIM coordination and collaboration

### **⚙️ OPENING CONDITIONS** (Session 3 - Create Openings)
- **Clearance values** - Space around MEP elements (50mm, 100mm, 150mm)
- **Oversize settings** - Additional clearance beyond standard
- **Level adjustments** - Elevation modifications
- **Purpose**: Define how openings are sized and placed

## 🔄 Alignment with CONVOID_NOTES.md

- **Tech parity**: CONVOID uses WinForms with modeless forms and ExternalEvents; we will keep our WPF MVVM UI but adopt the same Revit integration pattern (modeless window + ExternalEvents).
- **Revit integration**: Add an `IExternalApplication` startup wiring and a ribbon button to open our modeless `MainDialog`. All Revit-modifying actions must run via ExternalEvent handlers.
- **Core manager mapping**: CONVOID's `VoidManagerContent`, `ParameterManager`, `StatusManager`, and `DisciplineManager` map to our Services (`DuctSleevePlacerService`, `PipeSleevePlacerService`, `OpeningDuplicationChecker`, etc.). We will centralize placement orchestration behind a single façade called `OpeningManager` that coordinates these services.
- **Status and discipline**: Add approval status and discipline toggles to the Right panel. Bind to our `StatusManager` equivalent and pass through to placement services.
- **Linked models**: Always resolve `RevitLinkInstance` transforms for hosts (floors, walls, framing) before computing placements. This constraint is mandatory and part of service calls.
- **Performance and transactions**: Pre-compute outside transactions, mutate inside minimal `Transaction`/`TransactionGroup`. Batch work. No multi-threaded API calls.
- **Logging**: Append to `Logs/placement_YYYYMMDD.log` during placement and adjustments (doc title, link context, element ids, timestamps).
- **Licensing**: CONVOID uses Cryptolens; we will not implement licensing here. Out of scope.

The rest of this plan keeps our WPF structure and augments it with the ExternalEvent and ribbon wiring modeled after CONVOID's approach.

## 🏗️ **Revised Architecture Implementation**

### **Data Model Separation (Revised with BCF)**

```csharp
// Profile = User Identity Only
public class UserProfile
{
    public string ProfileName { get; set; }        // "John Smith"
    public string Discipline { get; set; }         // "HVAC", "Plumbing", "Electrical"
    public string ProfileInitials { get; set; }    // "JS", "SJ", "MC"
    public DateTime CreatedDate { get; set; }
    public DateTime LastModified { get; set; }
}

// Filter = File Selection + BCF Data (Project State)
public class ProjectFilter
{
    public string FilterName { get; set; }                   // "Fire Fighting Project"
    public List<string> ReferenceFiles { get; set; }        // Architectural files
    public List<string> HostFiles { get; set; }             // Structural files
    public BcfProjectData BcfData { get; set; }             // BCF coordination data
    public DateTime CreatedDate { get; set; }
    public DateTime LastModified { get; set; }
}

// BCF Data = Industry Standard Coordination
public class BcfProjectData
{
    public List<BcfTopic> ClashTopics { get; set; }         // Detected clashes
    public List<BcfViewpoint> Viewpoints { get; set; }      // 3D visualization
    public List<BcfComment> StatusHistory { get; set; }     // Approval tracking
    public Dictionary<string, string> CoordinationData { get; set; }  // Multi-disciplinary
    public string BcfVersion { get; set; } = "2.1";
    public DateTime LastExport { get; set; }
}

// Opening Conditions = How to Place (Clearance Rules)
public class OpeningConditions
{
    public double DuctClearance { get; set; }                 // 50mm, 100mm, 150mm
    public double PipeClearance { get; set; }
    public double CableTrayClearance { get; set; }
    public double OversizeValue { get; set; }                 // Additional clearance
    public double LevelAdjustment { get; set; }               // Elevation modifications
    public Dictionary<string, double> DisciplineClearances { get; set; }  // Per-discipline
}
```

### **UI Panel Mapping (Revised with BCF)**

```
Left Panel (Project Filter + BCF Data)
├── Project Filter
│   ├── Reference Files (Architectural)
│   ├── Host Files (Structural)
│   └── Filter Name & Settings
└── BCF Coordination Data
    ├── Clash Topics (Detected intersections)
    ├── Status History (Approval tracking)
    ├── Viewpoints (3D visualization)
    └── Multi-disciplinary Labels

Middle Panel (Categories)
├── Reference Categories
└── Host Categories

Right Panel (Opening Conditions)
├── Horizontal Opening Conditions
│   ├── Clearance Values (Duct/Pipe/CableTray)
│   ├── Oversize Settings
│   └── Level Adjustments
└── Vertical Opening Conditions
    ├── Clearance Values
    ├── Oversize Settings
    └── Level Adjustments

Profile Panel (User Identity)
├── Profile Name
├── Discipline
└── Profile Initials (Auto-assigned)
```

### **Save/Load Architecture (Revised with BCF)**

```
Profile Management
├── Save Profile → User identity only (name, discipline, initials)
├── Load Profile → User identity only
└── Profile Initials → Auto-assigned from name

Project Filter Management
├── Save Filter → File selections + BCF data
├── Load Filter → File selections + BCF data
└── Filter Persistence → Project state + coordination data

BCF Data Management
├── Export BCF → Industry-standard coordination format
├── Import BCF → External coordination data
├── BCF Clash Topics → Detected intersections
├── BCF Status History → Approval tracking
└── BCF Viewpoints → 3D visualization data

Opening Conditions Management
├── Save Conditions → Clearance rules (how to place)
├── Load Conditions → Clearance rules (how to place)
└── Conditions Persistence → Duct/Pipe/CableTray clearances
```

## 📺 **CONVOID Interface Analysis**

### **Key UI Features Observed:**
1. **Three-Panel Layout** - Left (Filters/Projects), Middle (Categories), Right (Conditions)
2. **Advanced Filtering** - Text search, project selection, element filtering
3. **Dual Opening Types** - Horizontal and Vertical openings with separate configurations
4. **Parameter Filtering** - Dynamic parameter management with visual indicators
5. **Professional Toolbar** - Icon-based actions with progress tracking
6. **Modern Design** - Clean layout with proper spacing and visual hierarchy

## 🏗️ **UI Architecture Plan**

### Revit Integration Architecture (from ConVoid notes)

- **Startup**: Implement `IExternalApplication` to create a ribbon panel and a button that opens the modeless `MainDialog`.
- **Modeless safety**: Define ExternalEvent handlers for any action that touches the Revit API (placement, parameter sync, filters, dimensions, selection boxes, status sync).
- **Handlers** (initial set):
  - `RunPlacementExternalEvent` → orchestrates opening creation via `OpeningManager`
  - `SyncParametersExternalEvent` → pushes parameter updates
  - `UpdateStatusExternalEvent` → applies approval status
  - `RefreshSelectionExternalEvent` → refreshes cached element sets
- **Manager façade**: `OpeningManager` wraps existing services and ensures: linked-transform resolution, transaction scoping, duplication checks, parameter/status application, and logging.

### **1. Main Window Structure**

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
            <StackPanel Orientation="Horizontal">
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
            
            <!-- Left Panel: Filters and Projects -->
            <local:LeftPanel Grid.Column="0" DataContext="{Binding LeftPanelViewModel}"/>
            
            <!-- Splitter 1 -->
            <GridSplitter Grid.Column="1" HorizontalAlignment="Stretch" Background="#4C566A"/>
            
            <!-- Middle Panel: Categories -->
            <local:MiddlePanel Grid.Column="2" DataContext="{Binding MiddlePanelViewModel}"/>
            
            <!-- Splitter 2 -->
            <GridSplitter Grid.Column="3" HorizontalAlignment="Stretch" Background="#4C566A"/>
            
            <!-- Right Panel: Conditions -->
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
                    <Button Style="{StaticResource IconButtonStyle}" Content="←" ToolTip="Previous"/>
                    <Button Style="{StaticResource IconButtonStyle}" Content="→" ToolTip="Next"/>
                    <Button Style="{StaticResource IconButtonStyle}" Content="⚡" ToolTip="Process"/>
                    <Button Style="{StaticResource IconButtonStyle}" Content="🔧" ToolTip="Settings"/>
                    <Button Style="{StaticResource IconButtonStyle}" Content="🔄" ToolTip="Refresh"/>
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

### **2. Left Panel - Filters and Projects**

```xml
<!-- LeftPanel.xaml -->
<UserControl x:Class="JSE_RevitAddin_MEP_OPENINGS.Views.LeftPanel">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>  <!-- Filters -->
            <RowDefinition Height="Auto"/>  <!-- Projects -->
            <RowDefinition Height="Auto"/>  <!-- Reference Elements -->
            <RowDefinition Height="Auto"/>  <!-- Host Elements -->
            <RowDefinition Height="*"/>     <!-- Bottom Icons -->
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
        
        <!-- Reference Elements -->
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
        
        <!-- Host Elements -->
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
        <StackPanel Grid.Row="4" Orientation="Horizontal" HorizontalAlignment="Center" VerticalAlignment="Bottom">
            <Button Style="{StaticResource IconButtonStyle}" Content="⚙️" ToolTip="Settings"/>
            <Button Style="{StaticResource IconButtonStyle}" Content="📄" ToolTip="Document"/>
            <Button Style="{StaticResource IconButtonStyle}" Content="📋" ToolTip="Clipboard"/>
            <Button Style="{StaticResource IconButtonStyle}" Content="🤖" ToolTip="AI Assistant"/>
            <Button Style="{StaticResource IconButtonStyle}" Content="💾" ToolTip="Save"/>
        </StackPanel>
    </Grid>
</UserControl>
```

### **3. Middle Panel - Categories**

```xml
<!-- MiddlePanel.xaml -->
<UserControl x:Class="JSE_RevitAddin_MEP_OPENINGS.Views.MiddlePanel">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>  <!-- Reference Category -->
            <RowDefinition Height="Auto"/>  <!-- Create Horizontal Openings -->
            <RowDefinition Height="*"/>     <!-- Spacer -->
        </Grid.RowDefinitions>
        
        <!-- Reference Category -->
        <GroupBox Grid.Row="0" Header="Reference Category" Margin="0,0,0,10">
            <ListBox ItemsSource="{Binding ReferenceCategories}" 
                     SelectionMode="Multiple"
                     Height="200">
                <ListBox.ItemTemplate>
                    <DataTemplate>
                        <CheckBox Content="{Binding Name}" 
                                  IsChecked="{Binding IsSelected}"/>
                    </DataTemplate>
                </ListBox.ItemTemplate>
            </ListBox>
        </GroupBox>
        
        <!-- Create Horizontal Openings -->
        <GroupBox Grid.Row="1" Header="Create Horizontal Openings in">
            <ListBox ItemsSource="{Binding HorizontalOpeningCategories}" 
                     SelectionMode="Multiple"
                     Height="300">
                <ListBox.ItemTemplate>
                    <DataTemplate>
                        <CheckBox Content="{Binding Name}" 
                                  IsChecked="{Binding IsSelected}"/>
                    </DataTemplate>
                </ListBox.ItemTemplate>
            </ListBox>
        </GroupBox>
    </Grid>
</UserControl>
```

### **4. Right Panel - Conditions**

```xml
<!-- RightPanel.xaml -->
<UserControl x:Class="JSE_RevitAddin_MEP_OPENINGS.Views.RightPanel">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>  <!-- Horizontal Openings -->
            <RowDefinition Height="Auto"/>  <!-- Vertical Openings -->
        </Grid.RowDefinitions>
        
        <!-- Horizontal Openings -->
        <GroupBox Grid.Row="0" Header="Conditions - Horizontal openings" Margin="0,0,0,10">
            <Grid Margin="10">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>  <!-- Opening Section -->
                    <RowDefinition Height="Auto"/>  <!-- Parameter Filter -->
                </Grid.RowDefinitions>
                
                <!-- Opening Section -->
                <GroupBox Grid.Row="0" Header="Opening" Margin="0,0,0,10">
                    <Grid Margin="10">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>
                        
                        <!-- Level and Mode -->
                        <Grid Grid.Row="0" Margin="0,0,0,10">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="Auto"/>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            
                            <TextBlock Grid.Column="0" Text="Level:" VerticalAlignment="Center" Margin="0,0,10,0"/>
                            <ComboBox Grid.Column="1" 
                                      ItemsSource="{Binding Levels}"
                                      SelectedItem="{Binding SelectedLevel}"
                                      Margin="0,0,10,0"/>
                            <StackPanel Grid.Column="2" Orientation="Horizontal">
                                <Button Style="{StaticResource ModeButtonStyle}" 
                                        Content="↔" 
                                        IsChecked="{Binding Mode1}" 
                                        ToolTip="Mode 1"/>
                                <Button Style="{StaticResource ModeButtonStyle}" 
                                        Content="↕" 
                                        IsChecked="{Binding Mode2}" 
                                        ToolTip="Mode 2"/>
                                <Button Style="{StaticResource ModeButtonStyle}" 
                                        Content="⚡" 
                                        IsChecked="{Binding Mode3}" 
                                        ToolTip="Mode 3"/>
                            </StackPanel>
                        </Grid>
                        
                        <!-- Oversize -->
                        <Grid Grid.Row="1" Margin="0,0,0,10">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="Auto"/>
                                <ColumnDefinition Width="80"/>
                                <ColumnDefinition Width="Auto"/>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                                <ColumnDefinition Width="*"/>
                            </Grid.ColumnDefinitions>
                            
                            <TextBlock Grid.Column="0" Text="Oversize:" VerticalAlignment="Center" Margin="0,0,10,0"/>
                            <TextBox Grid.Column="1" Text="{Binding Oversize}" Margin="0,0,10,0"/>
                            <TextBlock Grid.Column="2" Text="Rectangle:" VerticalAlignment="Center" Margin="0,0,10,0"/>
                            <ComboBox Grid.Column="3" 
                                      ItemsSource="{Binding RectangleOptions}"
                                      SelectedItem="{Binding SelectedRectangle}"
                                      Margin="0,0,10,0"/>
                            <TextBlock Grid.Column="4" Text="Circular:" VerticalAlignment="Center" Margin="0,0,10,0"/>
                            <ComboBox Grid.Column="5" 
                                      ItemsSource="{Binding CircularOptions}"
                                      SelectedItem="{Binding SelectedCircular}"/>
                        </Grid>
                    </Grid>
                </GroupBox>
                
                <!-- Parameter Filter -->
                <GroupBox Grid.Row="1" Header="Parameter Filter">
                    <Grid Margin="10">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>
                        
                        <!-- Filter Row 1 -->
                        <Grid Grid.Row="0" Margin="0,0,0,5">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            
                            <ComboBox Grid.Column="0" 
                                      ItemsSource="{Binding ParameterOptions}"
                                      SelectedItem="{Binding SelectedParameter1}"
                                      Margin="0,0,5,0"/>
                            <Button Grid.Column="1" 
                                    Style="{StaticResource AddRemoveButtonStyle}" 
                                    Content="+" 
                                    Background="Green" 
                                    Foreground="White"/>
                        </Grid>
                        
                        <!-- Filter Row 2 -->
                        <Grid Grid.Row="1" Margin="0,0,0,5">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            
                            <ComboBox Grid.Column="0" 
                                      ItemsSource="{Binding ParameterOptions}"
                                      SelectedItem="{Binding SelectedParameter2}"
                                      Margin="0,0,5,0"/>
                            <Button Grid.Column="1" 
                                    Style="{StaticResource AddRemoveButtonStyle}" 
                                    Content="-" 
                                    Background="Red" 
                                    Foreground="White"/>
                        </Grid>
                        
                        <!-- Filter Row 3 -->
                        <Grid Grid.Row="2">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            
                            <ComboBox Grid.Column="0" 
                                      ItemsSource="{Binding ParameterOptions}"
                                      SelectedItem="{Binding SelectedParameter3}"
                                      Margin="0,0,5,0"/>
                            <Button Grid.Column="1" 
                                    Style="{StaticResource AddRemoveButtonStyle}" 
                                    Content="-" 
                                    Background="Red" 
                                    Foreground="White"/>
                        </Grid>
                    </Grid>
                </GroupBox>
            </Grid>
        </GroupBox>
        
        <!-- Vertical Openings -->
        <GroupBox Grid.Row="1" Header="Conditions - Vertical openings">
            <Grid Margin="10">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>  <!-- Opening Section -->
                    <RowDefinition Height="Auto"/>  <!-- Parameter Filter -->
                </Grid.RowDefinitions>
                
                <!-- Opening Section -->
                <GroupBox Grid.Row="0" Header="Opening" Margin="0,0,0,10">
                    <Grid Margin="10">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>
                        
                        <!-- Level and Mode -->
                        <Grid Grid.Row="0" Margin="0,0,0,10">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="Auto"/>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            
                            <TextBlock Grid.Column="0" Text="Level:" VerticalAlignment="Center" Margin="0,0,10,0"/>
                            <ComboBox Grid.Column="1" 
                                      ItemsSource="{Binding VerticalLevels}"
                                      SelectedItem="{Binding SelectedVerticalLevel}"
                                      Margin="0,0,10,0"/>
                            <StackPanel Grid.Column="2" Orientation="Horizontal">
                                <Button Style="{StaticResource ModeButtonStyle}" 
                                        Content="↔" 
                                        IsChecked="{Binding VerticalMode1}" 
                                        ToolTip="Mode 1"/>
                                <Button Style="{StaticResource ModeButtonStyle}" 
                                        Content="↕" 
                                        IsChecked="{Binding VerticalMode2}" 
                                        ToolTip="Mode 2"/>
                                <Button Style="{StaticResource ModeButtonStyle}" 
                                        Content="⚡" 
                                        IsChecked="{Binding VerticalMode3}" 
                                        ToolTip="Mode 3"/>
                            </StackPanel>
                        </Grid>
                        
                        <!-- Oversize -->
                        <Grid Grid.Row="1">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="Auto"/>
                                <ColumnDefinition Width="80"/>
                                <ColumnDefinition Width="Auto"/>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                                <ColumnDefinition Width="*"/>
                            </Grid.ColumnDefinitions>
                            
                            <TextBlock Grid.Column="0" Text="Oversize:" VerticalAlignment="Center" Margin="0,0,10,0"/>
                            <TextBox Grid.Column="1" Text="{Binding VerticalOversize}" Margin="0,0,10,0"/>
                            <TextBlock Grid.Column="2" Text="Rectangle:" VerticalAlignment="Center" Margin="0,0,10,0"/>
                            <ComboBox Grid.Column="3" 
                                      ItemsSource="{Binding VerticalRectangleOptions}"
                                      SelectedItem="{Binding SelectedVerticalRectangle}"
                                      Margin="0,0,10,0"/>
                            <TextBlock Grid.Column="4" Text="Circular:" VerticalAlignment="Center" Margin="0,0,10,0"/>
                            <ComboBox Grid.Column="5" 
                                      ItemsSource="{Binding VerticalCircularOptions}"
                                      SelectedItem="{Binding SelectedVerticalCircular}"/>
                        </Grid>
                    </Grid>
                </GroupBox>
                
                <!-- Parameter Filter -->
                <GroupBox Grid.Row="1" Header="Parameter Filter">
                    <Grid Margin="10">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>
                        
                        <!-- Filter Row 1 -->
                        <Grid Grid.Row="0" Margin="0,0,0,5">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            
                            <ComboBox Grid.Column="0" 
                                      ItemsSource="{Binding VerticalParameterOptions}"
                                      SelectedItem="{Binding SelectedVerticalParameter1}"
                                      Margin="0,0,5,0"/>
                            <Button Grid.Column="1" 
                                    Style="{StaticResource AddRemoveButtonStyle}" 
                                    Content="+" 
                                    Background="Green" 
                                    Foreground="White"/>
                        </Grid>
                        
                        <!-- Filter Row 2 -->
                        <Grid Grid.Row="1">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            
                            <ComboBox Grid.Column="0" 
                                      ItemsSource="{Binding VerticalParameterOptions}"
                                      SelectedItem="{Binding SelectedVerticalParameter2}"
                                      Margin="0,0,5,0"/>
                            <Button Grid.Column="1" 
                                    Style="{StaticResource AddRemoveButtonStyle}" 
                                    Content="+" 
                                    Background="Green" 
                                    Foreground="White"/>
                        </Grid>
                    </Grid>
                </GroupBox>
            </Grid>
        </GroupBox>
    </Grid>
</UserControl>
```

### UI Additions Based on ConVoid

- **Right Panel → Opening section**: add fields for Approval Status (ComboBox) and Elevation Adjustment (TextBox). Include Discipline toggles (e.g., checkboxes for HVAC, Piping, Electrical) that gate categories and parameter filters.
- **Parameter Filters**: reflect per-discipline parameter sets. Visual indicator when a filter is active.
- **Bottom Toolbar**: add Opening Mode (Horizontal/Vertical), Process, Refresh, and Settings, all invoking ExternalEvents.

### Ribbon and ExternalEvent Wiring

- Add a ribbon panel “JSE Openings” with a “Openings Manager” button. Clicking it opens `MainDialog` modelessly.
- On dialog actions (OK/Process/Save/Status change), raise the corresponding ExternalEvent to execute Revit API work on the Revit thread.
- Create singletons for ExternalEvent handlers at startup to avoid repeated allocations.

## 🎨 **Modern Styling for conVoid Interface**

```xml
<!-- ConVoidStyles.xaml -->
<ResourceDictionary xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
                    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml">
    
    <!-- Icon Button Style -->
    <Style x:Key="IconButtonStyle" TargetType="Button">
        <Setter Property="Width" Value="30"/>
        <Setter Property="Height" Value="30"/>
        <Setter Property="Background" Value="Transparent"/>
        <Setter Property="BorderThickness" Value="1"/>
        <Setter Property="BorderBrush" Value="#DEE2E6"/>
        <Setter Property="FontSize" Value="14"/>
        <Setter Property="Template">
            <Setter.Value>
                <ControlTemplate TargetType="Button">
                    <Border Background="{TemplateBinding Background}" 
                            BorderBrush="{TemplateBinding BorderBrush}" 
                            BorderThickness="{TemplateBinding BorderThickness}"
                            CornerRadius="3">
                        <ContentPresenter HorizontalAlignment="Center" 
                                          VerticalAlignment="Center"/>
                    </Border>
                    <ControlTemplate.Triggers>
                        <Trigger Property="IsMouseOver" Value="True">
                            <Setter Property="Background" Value="#F8F9FA"/>
                        </Trigger>
                    </ControlTemplate.Triggers>
                </ControlTemplate>
            </Setter.Value>
        </Setter>
    </Style>
    
    <!-- Mode Button Style -->
    <Style x:Key="ModeButtonStyle" TargetType="Button">
        <Setter Property="Width" Value="25"/>
        <Setter Property="Height" Value="25"/>
        <Setter Property="Background" Value="Transparent"/>
        <Setter Property="BorderThickness" Value="1"/>
        <Setter Property="BorderBrush" Value="#DEE2E6"/>
        <Setter Property="FontSize" Value="12"/>
        <Setter Property="Template">
            <Setter.Value>
                <ControlTemplate TargetType="Button">
                    <Border Background="{TemplateBinding Background}" 
                            BorderBrush="{TemplateBinding BorderBrush}" 
                            BorderThickness="{TemplateBinding BorderThickness}"
                            CornerRadius="2">
                        <ContentPresenter HorizontalAlignment="Center" 
                                          VerticalAlignment="Center"/>
                    </Border>
                    <ControlTemplate.Triggers>
                        <Trigger Property="IsChecked" Value="True">
                            <Setter Property="Background" Value="#007ACC"/>
                            <Setter Property="Foreground" Value="White"/>
                        </Trigger>
                        <Trigger Property="IsMouseOver" Value="True">
                            <Setter Property="Background" Value="#F8F9FA"/>
                        </Trigger>
                    </ControlTemplate.Triggers>
                </ControlTemplate>
            </Setter.Value>
        </Setter>
    </Style>
    
    <!-- Add/Remove Button Style -->
    <Style x:Key="AddRemoveButtonStyle" TargetType="Button">
        <Setter Property="Width" Value="20"/>
        <Setter Property="Height" Value="20"/>
        <Setter Property="BorderThickness" Value="0"/>
        <Setter Property="FontSize" Value="12"/>
        <Setter Property="FontWeight" Value="Bold"/>
        <Setter Property="Template">
            <Setter.Value>
                <ControlTemplate TargetType="Button">
                    <Border Background="{TemplateBinding Background}" 
                            CornerRadius="10">
                        <ContentPresenter HorizontalAlignment="Center" 
                                          VerticalAlignment="Center"/>
                    </Border>
                </ControlTemplate>
            </Setter.Value>
        </Setter>
    </Style>
    
</ResourceDictionary>
```

## 🚀 **Revised Implementation Phases**

### **Phase 1: Data Model Restructure with BCF (Week 1)**
- [ ] **Separate Profile from Filter**: Profile = user identity only, Filter = project state + BCF data
- [ ] **Create new data models**: UserProfile, ProjectFilter, BcfProjectData, OpeningConditions
- [ ] **Implement BCF integration**: BcfTopic, BcfViewpoint, BcfComment models
- [ ] **Update existing models**: Remove clearance from Profile, add to OpeningConditions
- [ ] **Fix XML serialization**: Ensure all new models are properly serializable
- [ ] **Update save/load logic**: Separate profile, filter (with BCF), and conditions persistence

### **Phase 2: UI Panel Restructure with BCF (Week 2)**
- [ ] **Left Panel**: Project Filter (file selections) + BCF Data (clash topics, status history)
- [ ] **Middle Panel**: Keep categories as-is (reference and host categories)
- [ ] **Right Panel**: Split into Horizontal and Vertical Opening Conditions (clearance values)
- [ ] **Profile Panel**: Simple user identity (name, discipline, initials)
- [ ] **BCF Export/Import**: Add BCF coordination buttons
- [ ] **Add ribbon button** to open modeless `MainDialog`

### **Phase 3: Data Binding & BCF Logic (Week 3)**
- [ ] **Create ViewModels**: ProfileViewModel, ProjectFilterViewModel, BcfDataViewModel, OpeningConditionsViewModel
- [ ] **Implement data binding**: Bind each panel to its respective ViewModel
- [ ] **Update clearance flow**: Clearance values flow from OpeningConditions to placement services
- [ ] **Implement BCF services**: BcfExportService, BcfImportService, BcfTopicGenerator
- [ ] **Implement ExternalEvent handlers**: Profile, Filter (with BCF), and Conditions management
- [ ] **Add opening placement logic**: Use OpeningConditions for clearance values

### **Phase 4: Advanced BCF Features (Week 4)**
- [ ] **Profile management**: Auto-assign initials, discipline tracking
- [ ] **Project Filter management**: File selection, BCF data storage
- [ ] **BCF coordination**: Multi-disciplinary approval workflow
- [ ] **BCF clash detection**: Automated clash topic generation
- [ ] **BCF viewpoints**: 3D visualization for clash zones
- [ ] **BCF export/import**: Industry-standard coordination format
- [ ] **Opening Conditions management**: Clearance values, oversize settings, level adjustments
- [ ] **Progress tracking**: Status updates for each management type
- [ ] **Transaction scoping**: Proper Revit API transaction management
- [ ] **Logging**: Separate logs for profile, filter (BCF), and conditions operations

### **Phase 5: Polish & BCF Testing (Week 5)**
- [ ] **Validation**: Ensure Profile, Project Filter (BCF), and Opening Conditions are properly validated
- [ ] **BCF compatibility**: Test BCF export/import with standard BCF viewers
- [ ] **Error handling**: Proper error messages for each management type
- [ ] **Performance optimization**: Efficient save/load for each data type
- [ ] **BCF testing**: Verify BCF coordination workflow works correctly
- [ ] **Integration testing**: Test Profile + Filter (BCF) + Conditions separation
- [ ] **Documentation**: Update all documentation to reflect BCF-integrated architecture

---

*This implementation plan provides a comprehensive roadmap for creating a professional MEP Openings interface that matches the conVoid application's sophisticated design and functionality.*

using System.Xml.Serialization;

namespace JSE_RevitAddin_MEP_OPENINGS.Models
{
    /// <summary>
    /// Global Application Configuration - Filter-independent, user-wide settings
    /// Location: C:\Users\[Username]\AppData\Roaming\JSE_MEP_Openings\GlobalSettings.xml
    /// </summary>
    [XmlRoot("GlobalApplicationSettings")]
    public class GlobalApplicationSettings
    {
        // Installation and Paths (Filter-Independent)
        public string DefaultFamiliesPath { get; set; } = @"C:\Users\[Username]\AppData\Roaming\JSE_MEP_Openings\Families";
        public string DefaultBCFPath { get; set; } = @"C:\Users\[Username]\AppData\Roaming\JSE_MEP_Openings\BCF";
        public string DefaultFiltersPath { get; set; } = @"C:\Users\[Username]\AppData\Roaming\JSE_MEP_Openings\Projects\[ProjectName]\Filters";
        public string DefaultProjectsPath { get; set; } = @"C:\Users\[Username]\AppData\Roaming\JSE_MEP_Openings\Projects";
        
        // Application Preferences (Filter-Independent)
        public string UILanguage { get; set; } = "English";
        public bool AutoSaveEnabled { get; set; } = true;
        public int AutoSaveIntervalSeconds { get; set; } = 30;
        public bool ShowDebugLogs { get; set; } = false;
        
        // CONFIGURATION Defaults (Advanced Processing Rules)
        public ConfigurationDefaults ConfigurationDefaults { get; set; } = new ConfigurationDefaults();
    }

    /// <summary>
    /// Default CONFIGURATION settings that apply globally unless overridden by Project Filter
    /// </summary>
    public class ConfigurationDefaults
    {
        // LIMITS Section (REQUIRED - Highlighted in CONVOID)
        public double MaxOpeningWidth { get; set; } = 2000.0; // Maximum opening width (mm)
        public double MaxOpeningHeight { get; set; } = 2000.0; // Maximum opening height (mm)
        public double MinOpeningWidth { get; set; } = 50.0; // Minimum opening width (mm)
        public double MinOpeningHeight { get; set; } = 50.0; // Minimum opening height (mm)
        public double MaxOpeningDepth { get; set; } = 500.0; // Maximum opening depth (mm)
        public double MinOpeningDepth { get; set; } = 50.0; // Minimum opening depth (mm)
        
        // ELEMENTS Section (REQUIRED - Highlighted in CONVOID)
        public bool ProcessDucts { get; set; } = true; // Process duct elements
        public bool ProcessPipes { get; set; } = true; // Process pipe elements
        public bool ProcessCableTrays { get; set; } = true; // Process cable tray elements
        public bool ProcessConduits { get; set; } = true; // Process conduit elements
        public bool ProcessDuctInsulation { get; set; } = true; // Process duct insulation
        public bool ProcessPipeInsulation { get; set; } = true; // Process pipe insulation
        
        // MANAGE Section (EXPANDED - Important for workflow)
        public bool ResetApprovalStatusOnChanges { get; set; } = true; // Reset approval status of openings when changes occur
        public double DimensionChangeThreshold { get; set; } = 1.0; // Openings won't be marked as changed if change in dimensions is less than (mm)
        public double LocationChangeThreshold { get; set; } = 1.0; // Openings won't be marked as changed if change in location is less than (mm)
        public bool CutOpeningWithHosts { get; set; } = false; // Cut opening with Hosts
        public bool CreateConstraintBetweenOpeningsAndHosts { get; set; } = true; // Create a constraint between openings and Hosts
        public bool AdoptProvisionForVoidsFromLinkedModel { get; set; } = false; // Adopt Provision for Voids (Openings) from linked Model
        public bool IncludeHostElementsNotVisibleIn3DView { get; set; } = true; // Include Host Elements not visible in the selected 3D view
        public bool IncludeReferenceElementsNotVisibleIn3DView { get; set; } = true; // Include Reference Elements not visible in the selected 3D view
        public bool IncludeHostElementsInDemolishedPhase { get; set; } = false; // Include Host Elements in demolished Phase
        public double IgnoreOpeningsSmallerThan { get; set; } = 1.0; // Ignore openings smaller than (mm)
        public double RoundOpeningsBecomeRectangularIfDiameterGreaterThan { get; set; } = 200.0; // Round openings become rectangular if diameter is greater than (mm)
        public double JoinOpeningsIfDistanceLessThan { get; set; } = 100.0; // Join openings if their distance is less than (mm)
        public double IgnoreOpeningsWithAngleGreaterThan { get; set; } = 45.0; // Ignore openings with an angle greater than (degrees)
    }
}

using System;

namespace JSE_RevitAddin_MEP_OPENINGS.Models
{
    /// <summary>
    /// Model class to hold MEP Openings settings/configuration
    /// </summary>
    public class SettingsModel
    {
        // Manage Section Properties
        public bool ResetApprovalStatus { get; set; } = true;
        public double DimensionChangeThreshold { get; set; } = 1.0;
        public double LocationChangeThreshold { get; set; } = 1.0;

        // Elements Section Properties
        public bool CutOpeningWithHosts { get; set; } = false;
        public bool CreateConstraint { get; set; } = true;
        public bool PipeOpeningTypeRectangular { get; set; } = false; // Default to circular (false = circular, true = rectangular)
        public bool CreateVerticalOpenings { get; set; } = false;
        public bool CreateHorizontalOpenings { get; set; } = false;
        public bool AdoptProvisionForVoids { get; set; } = false;

        // Element Filter Section Properties
        public string ElementFilter { get; set; } = "{3D}";
        public bool IncludeHostElementsNotVisible { get; set; } = true;
        public bool IncludeReferenceElementsNotVisible { get; set; } = true;
        public bool IncludeHostElementsDemolished { get; set; } = false;

        // Limits Section Properties
        public double IgnoreOpeningsSmallerThan { get; set; } = 0.1;
        public double RoundOpeningsRectangular { get; set; } = 200.0;
        public double JoinOpeningsDistance { get; set; } = 200.0;
        public double IgnoreOpeningsAngle { get; set; } = 45.0;
        public bool CreateOpeningsWithSlope { get; set; } = true;
        public double RoundingValue { get; set; } = 5.0; // Default rounding value in mm
        public bool RoundAlwaysUp { get; set; } = false; // If true, always round up; if false, round to nearest
        public double MinWallThickness { get; set; } = 0.0; // Minimum wall thickness in mm - walls below this are ignored
        public bool IgnoreArchitecturalFloors { get; set; } = false; // If true, ignore floors without Structural parameter checked

        public SettingsModel()
        {
            // Default constructor with default values
        }

        /// <summary>
        /// Creates a copy of the current settings
        /// </summary>
        public SettingsModel Clone()
        {
            return new SettingsModel
            {
                ResetApprovalStatus = this.ResetApprovalStatus,
                DimensionChangeThreshold = this.DimensionChangeThreshold,
                LocationChangeThreshold = this.LocationChangeThreshold,
                CutOpeningWithHosts = this.CutOpeningWithHosts,
                CreateConstraint = this.CreateConstraint,
                PipeOpeningTypeRectangular = this.PipeOpeningTypeRectangular,
                CreateVerticalOpenings = this.CreateVerticalOpenings,
                CreateHorizontalOpenings = this.CreateHorizontalOpenings,
                AdoptProvisionForVoids = this.AdoptProvisionForVoids,
                ElementFilter = this.ElementFilter,
                IncludeHostElementsNotVisible = this.IncludeHostElementsNotVisible,
                IncludeReferenceElementsNotVisible = this.IncludeReferenceElementsNotVisible,
                IncludeHostElementsDemolished = this.IncludeHostElementsDemolished,
                IgnoreOpeningsSmallerThan = this.IgnoreOpeningsSmallerThan,
                RoundOpeningsRectangular = this.RoundOpeningsRectangular,
                JoinOpeningsDistance = this.JoinOpeningsDistance,
                IgnoreOpeningsAngle = this.IgnoreOpeningsAngle,
                CreateOpeningsWithSlope = this.CreateOpeningsWithSlope,
                RoundingValue = this.RoundingValue,
                RoundAlwaysUp = this.RoundAlwaysUp,
                MinWallThickness = this.MinWallThickness,
                IgnoreArchitecturalFloors = this.IgnoreArchitecturalFloors
            };
        }
    }
}

using System;

namespace JSE_RevitAddin_MEP_OPENINGS.Models
{
    /// <summary>
    /// Represents the size and shape of an MEP element
    /// Used by universal sleeve placement strategies
    /// </summary>
    public class MepElementSize
    {
        /// <summary>
        /// Shape of the MEP element
        /// </summary>
        public string Shape { get; set; } = "Rectangular"; // "Round", "Rectangular", "Oval"
        
        /// <summary>
        /// Diameter for round elements (in feet - Revit internal units)
        /// </summary>
        public double Diameter { get; set; }
        
        /// <summary>
        /// Width for rectangular elements (in feet - Revit internal units)
        /// </summary>
        public double Width { get; set; }
        
        /// <summary>
        /// Height for rectangular elements (in feet - Revit internal units)
        /// </summary>
        public double Height { get; set; }
        
        /// <summary>
        /// Whether this element has insulation
        /// </summary>
        public bool IsInsulated { get; set; }
        
        /// <summary>
        /// Insulation thickness (in feet - Revit internal units)
        /// </summary>
        public double InsulationThickness { get; set; }
        
        /// <summary>
        /// Damper type for duct accessories (MSFD, MSD, MD, Motorized, etc.)
        /// </summary>
        public string DamperType { get; set; } = "";
        
        /// <summary>
        /// Formatted size string for display (e.g., "Ø300", "400×200")
        /// </summary>
        public string FormattedSize
        {
            get
            {
                if (Shape == "Round" || Shape == "Circular")
                {
                    double diamMm = Autodesk.Revit.DB.UnitUtils.ConvertFromInternalUnits(Diameter, Autodesk.Revit.DB.UnitTypeId.Millimeters);
                    return $"Ø{diamMm:F0}";
                }
                else
                {
                    double widthMm = Autodesk.Revit.DB.UnitUtils.ConvertFromInternalUnits(Width, Autodesk.Revit.DB.UnitTypeId.Millimeters);
                    double heightMm = Autodesk.Revit.DB.UnitUtils.ConvertFromInternalUnits(Height, Autodesk.Revit.DB.UnitTypeId.Millimeters);
                    return $"{widthMm:F0}×{heightMm:F0}";
                }
            }
        }
    }
}


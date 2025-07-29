using Autodesk.Revit.DB;

namespace JSE_RevitAddin_MEP_OPENINGS.Models
{    /// <summary>
     /// Represents the requirements for an opening in a wall
     /// </summary>
    public class OpeningRequirement
    {
        public ElementId ElementId { get; set; }
        public double Width { get; set; }
        public double Height { get; set; }
        public XYZ Location { get; set; }
        public XYZ Direction { get; set; }
        public ElementId WallId { get; set; }
        public ElementId TypeId { get; set; }
        public string ElementType { get; set; }
        public double ElementDiameter { get; set; }
        public double InsulationThickness { get; set; }
        public double IntersectionX { get; set; }
        public double IntersectionY { get; set; }
        public double IntersectionZ { get; set; }
        public double PipeDiameter { get; set; }
        public double DuctWidth { get; set; }
        public double DuctHeight { get; set; }
        public string MepType { get; set; }
        public long WallIdValue { get; set; } // Numeric property for use in CSV export

        public OpeningRequirement()
        {
        }

        public OpeningRequirement(ElementId elementId, double width, double height, XYZ location, XYZ direction, ElementId wallId, ElementId typeId, string elementType)
        {
            ElementId = elementId;
            Width = width;
            Height = height;
            Location = location;
            Direction = direction;
            WallId = wallId;
            TypeId = typeId;
            ElementType = elementType;
        }

        // Constructor for pipe openings
        public OpeningRequirement(ElementId elementId, double diameter, double insulationThickness, XYZ location, XYZ direction, ElementId wallId, ElementId typeId, string elementType, bool isPipe)
        {
            ElementId = elementId;
            ElementDiameter = diameter;
            InsulationThickness = insulationThickness;
            Location = location;
            Direction = direction;
            WallId = wallId;
            TypeId = typeId;
            ElementType = elementType;
        }
    }
}

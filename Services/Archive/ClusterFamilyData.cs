using Autodesk.Revit.DB;

namespace JSE_RevitAddin_MEP_OPENINGS.Commands
{
    public class ClusterFamilyData
    {
        public ElementId Id { get; set; }
    public XYZ? Position { get; set; }
        public double Width { get; set; }
        public double Height { get; set; }
        public bool IsFlipped { get; set; }
        public FamilyInstance Instance { get; set; }

        public ClusterFamilyData(FamilyInstance instance)
        {
            Instance = instance;
            Id = instance.Id;
            var locPoint = instance.Location as LocationPoint;
            Position = locPoint?.Point;
            Width = GetParameter(instance, "Width");
            Height = GetParameter(instance, "Height");
            IsFlipped = instance.FacingFlipped;
        }

        private double GetParameter(FamilyInstance instance, string paramName)
        {
            Parameter param = instance.LookupParameter(paramName);
            return param != null ? param.AsDouble() : 0.0;
        }
    }
}

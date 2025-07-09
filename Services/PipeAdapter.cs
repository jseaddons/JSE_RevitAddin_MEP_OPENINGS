using Autodesk.Revit.DB;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    public class PipeAdapter : IMepElementAdapter
    {
        public Element Element { get; }
        public string ElementTypeName => "Pipe";

        public PipeAdapter(Element element)
        {
            Element = element;
        }

        public Curve GetCurve()
        {
            return (Element.Location as LocationCurve)?.Curve;
        }

        public double GetDiameter()
        {
            var param = Element.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER);
            return param != null ? param.AsDouble() : 0.0;
        }

        public double GetWidth()
        {
            // Not applicable for pipes
            return 0.0;
        }

        public double GetHeight()
        {
            // Not applicable for pipes
            return 0.0;
        }
        public XYZ GetDirection()
        {
            var curve = GetCurve() as Line;
            return curve?.Direction;
        }
    }
}

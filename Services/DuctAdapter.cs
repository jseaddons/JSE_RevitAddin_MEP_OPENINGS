using Autodesk.Revit.DB;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    public class DuctAdapter : IMepElementAdapter
    {
        public Element Element { get; }
        public string ElementTypeName => "Duct";

        public DuctAdapter(Element element)
        {
            Element = element;
        }

        public Curve GetCurve()
        {
            return (Element.Location as LocationCurve)?.Curve;
        }
        public double GetDiameter()
        {
            // Not applicable for ducts
            return 0.0;
        }
        // For ducts, you may want to return the width or height as the "diameter" equivalent.
        // Here, we return the width by default. Adjust as needed for your use case.
        public double GetWidth()
        {
            var widthParam = Element.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM);
            return widthParam != null ? widthParam.AsDouble() : 0.0;
        }

        public double GetHeight()
        {
            var heightParam = Element.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM);
            return heightParam != null ? heightParam.AsDouble() : 0.0;
        }

        public XYZ GetDirection()
        {
            var curve = GetCurve() as Line;
            return curve?.Direction;
        }
    }
}

using Autodesk.Revit.DB;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    public interface IMepElementAdapter
    {
        Element Element { get; }
        Curve GetCurve();
        double GetDiameter();
        double GetWidth();
        double GetHeight();
        XYZ GetDirection();
        string ElementTypeName { get; }
    }
}

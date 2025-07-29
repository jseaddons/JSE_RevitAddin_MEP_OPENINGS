using System;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Services;

public static class OpeningPlacerFactory
{
    public static object GetPlacer(Document doc, string elementTypeName)
    {
        return elementTypeName switch
        {
            "Pipe" => new PipeSleevePlacer(doc), // Commented out: PipeSleevePlacer not available
            "Duct" => new DuctSleevePlacer(doc),
            "CableTray" => new CableTraySleevePlacer(doc),
            _ => throw new NotSupportedException($"Unsupported MEP type: {elementTypeName}")
        };
    }
}

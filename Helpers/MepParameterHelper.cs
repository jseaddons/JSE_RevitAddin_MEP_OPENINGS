using Autodesk.Revit.DB;
using System.Collections.Generic;

namespace JSE_RevitAddin_MEP_OPENINGS.Helpers
{
    /// <summary>
    /// Helper to efficiently extract all relevant MEP parameters from a FamilyInstance in a single API call.
    /// </summary>
    public static class MepParameterHelper
    {
        public class MepParameterSnapshot
        {
            public double? Width { get; set; }
            public double? Height { get; set; }
            public double? Diameter { get; set; }
            public Dictionary<string, object> AllParameters { get; set; } = new();
        }

        // For dampers: prioritize 'Damper Width' and 'Damper Height', fallback to 'Width', 'Height', 'DIMENSION_WIDTH', 'DIMENSION_HEIGHT'
        private static readonly string[] DamperWidthMain = new[] { "Damper Width" };
        private static readonly string[] DamperWidthFallback = new[] { "Width", "width", "DIMENSIONS_WIDTH" };
        private static readonly string[] DamperHeightMain = new[] { "Damper Height" };
        private static readonly string[] DamperHeightFallback = new[] { "Height", "height", "DIMENSIONS_HEIGHT" };

        public static MepParameterSnapshot CaptureParameters(FamilyInstance instance)
        {
            var snapshot = new MepParameterSnapshot();
            if (instance == null) return snapshot;

            Parameter widthParam = null;
            Parameter heightParam = null;
            foreach (Parameter param in instance.Parameters)
            {
                string name = param.Definition?.Name ?? string.Empty;
                object value = null;
                if (param.StorageType == StorageType.Double)
                    value = param.AsDouble();
                else if (param.StorageType == StorageType.Integer)
                    value = param.AsInteger();
                else if (param.StorageType == StorageType.String)
                    value = param.AsString();
                else if (param.StorageType == StorageType.ElementId)
                    value = param.AsElementId();
                snapshot.AllParameters[name] = value;

                // Prioritize main damper width/height, then fallback
                if (widthParam == null && DamperWidthMain.Contains(name))
                    widthParam = param;
                if (heightParam == null && DamperHeightMain.Contains(name))
                    heightParam = param;
                // Fallbacks only if not already set
                if (widthParam == null && DamperWidthFallback.Contains(name))
                    widthParam = param;
                if (heightParam == null && DamperHeightFallback.Contains(name))
                    heightParam = param;

                if (name == "Diameter" || name == "DIMENSION_DIAMETER" || param.Id.IntegerValue == (int)BuiltInParameter.RBS_PIPE_DIAMETER_PARAM)
                    snapshot.Diameter = param.AsDouble();
            }
            if (widthParam != null)
                snapshot.Width = widthParam.AsDouble();
            if (heightParam != null)
                snapshot.Height = heightParam.AsDouble();
            return snapshot;
        }
    }
}

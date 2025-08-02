using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    public static class GlobalFilteredElements
    {
        public static List<(Element element, Transform? transform)> FilteredDucts = new();
        public static List<(Element element, Transform? transform)> FilteredCableTrays = new();
        public static List<(Element element, Transform? transform)> FilteredPipes = new();
    }
}

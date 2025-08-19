using System.Collections.Generic;
using Autodesk.Revit.DB;

namespace JSE_RevitAddin_MEP_OPENINGS.Models
{
    /// <summary>
    /// Compares references to determine if they are equal
    /// Used for comparing references in intersection points
    /// </summary>
    public class ReferenceComparer : IEqualityComparer<Reference>
    {
        public bool Equals(Reference x, Reference y)
        {
            if (x == null || y == null)
                return false;

            return x.ElementId.IntegerValue == y.ElementId.IntegerValue;
        }

        public int GetHashCode(Reference obj)
        {
            if (obj == null)
                return 0;

            return obj.ElementId.IntegerValue.GetHashCode();
        }
    }
}

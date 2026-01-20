using Autodesk.Revit.DB;
using System;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    public interface IRevitUnitConversionService
    {
        double ToInternalMillimeters(double value);
        double FromInternalMillimeters(double value);
        double ToInternalFeet(double value);
        double FromInternalFeet(double value);
    }

    public sealed class RevitUnitConversionService : IRevitUnitConversionService
    {
        public static IRevitUnitConversionService Instance { get; } = new RevitUnitConversionService();

        private const double MillimetersPerFoot = 304.8;
        private const double FeetPerMillimeter = 1.0 / MillimetersPerFoot;

        private readonly bool _useForgeTypeId;

        private RevitUnitConversionService()
        {
            // ✅ CRITICAL: Default to manual conversion to ensure constructor never throws
            _useForgeTypeId = false;
            
            try
            {
                // Use reflection to safely test UnitTypeId API availability without JIT errors
                var unitTypeIdType = typeof(UnitTypeId);
                if (unitTypeIdType != null)
                {
                    var millimetersProperty = unitTypeIdType.GetProperty("Millimeters", 
                        System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static);
                    
                    if (millimetersProperty != null)
                    {
                        // Try to actually get the value via reflection
                        var testValue = millimetersProperty.GetValue(null);
                        if (testValue != null)
                        {
                            _useForgeTypeId = true;
                        }
                    }
                }
            }
            catch (Exception)
            {
                // Silently fall back to manual conversion - NEVER throw from constructor
                _useForgeTypeId = false;
            }
        }

        public double ToInternalMillimeters(double value)
        {
            if (_useForgeTypeId)
            {
                try { return UnitUtils.ConvertToInternalUnits(value, UnitTypeId.Millimeters); }
                catch { /* silent fallback */ }
            }
            return value * FeetPerMillimeter;
        }

        public double FromInternalMillimeters(double value)
        {
            if (_useForgeTypeId)
            {
                try { return UnitUtils.ConvertFromInternalUnits(value, UnitTypeId.Millimeters); }
                catch { /* silent fallback */ }
            }
            return value * MillimetersPerFoot;
        }

        public double ToInternalFeet(double value)
        {
            if (_useForgeTypeId)
            {
                try { return UnitUtils.ConvertToInternalUnits(value, UnitTypeId.Feet); }
                catch { /* silent fallback */ }
            }
            return value;
        }

        public double FromInternalFeet(double value)
        {
            if (_useForgeTypeId)
            {
                try { return UnitUtils.ConvertFromInternalUnits(value, UnitTypeId.Feet); }
                catch { /* silent fallback */ }
            }
            return value;
        }
    }
}
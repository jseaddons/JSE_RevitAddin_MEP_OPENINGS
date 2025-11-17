using Autodesk.Revit.DB;
using System;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Centralised unit conversion service to abstract Revit API differences between versions.
    /// Provides a single OOP entry point that can be expanded or swapped as needed.
    /// </summary>
    public interface IRevitUnitConversionService
    {
        double ToInternalMillimeters(double value);
        double FromInternalMillimeters(double value);
        double ToInternalFeet(double value);
        double FromInternalFeet(double value);
    }

    /// <summary>
    /// Runtime-detected implementation that prefers Forge Type IDs (Revit 2021+) and
    /// gracefully falls back to manual conversions for earlier versions.
    /// </summary>
    public sealed class RevitUnitConversionService : IRevitUnitConversionService
    {
        public static IRevitUnitConversionService Instance { get; } = new RevitUnitConversionService();

        private const double MillimetersPerFoot = 304.8;
        private const double FeetPerMillimeter = 1.0 / MillimetersPerFoot;

        private readonly bool _useForgeTypeId;

        private RevitUnitConversionService()
        {
            try
            {
                // In Revit 2021+ UnitTypeId lives in Autodesk.Revit.DB and exposes Forge-type IDs.
                var unitTypeIdType = typeof(UnitTypeId);
                _useForgeTypeId = unitTypeIdType != null;
            }
            catch
            {
                _useForgeTypeId = false;
            }
        }

        public double ToInternalMillimeters(double value)
        {
            if (_useForgeTypeId)
            {
                try
                {
                    return UnitUtils.ConvertToInternalUnits(value, UnitTypeId.Millimeters);
                }
                catch (Exception ex)
                {
                    DebugLogger.Error($"[UnitConverter] ForgeTypeId conversion failed in ToInternalMillimeters: {ex.Message}");
                }
            }

            // Revit internal units are feet, so convert mm -> feet manually.
            return value * FeetPerMillimeter;
        }

        public double FromInternalMillimeters(double value)
        {
            if (_useForgeTypeId)
            {
                try
                {
                    return UnitUtils.ConvertFromInternalUnits(value, UnitTypeId.Millimeters);
                }
                catch (Exception ex)
                {
                    DebugLogger.Error($"[UnitConverter] ForgeTypeId conversion failed in FromInternalMillimeters: {ex.Message}");
                }
            }

            // Convert feet -> mm manually.
            return value * MillimetersPerFoot;
        }

        public double ToInternalFeet(double value)
        {
            if (_useForgeTypeId)
            {
                try
                {
                    return UnitUtils.ConvertToInternalUnits(value, UnitTypeId.Feet);
                }
                catch (Exception ex)
                {
                    DebugLogger.Error($"[UnitConverter] ForgeTypeId conversion failed in ToInternalFeet: {ex.Message}");
                }
            }

            // Already in feet.
            return value;
        }

        public double FromInternalFeet(double value)
        {
            if (_useForgeTypeId)
            {
                try
                {
                    return UnitUtils.ConvertFromInternalUnits(value, UnitTypeId.Feet);
                }
                catch (Exception ex)
                {
                    DebugLogger.Error($"[UnitConverter] ForgeTypeId conversion failed in FromInternalFeet: {ex.Message}");
                }
            }

            // Already in feet.
            return value;
        }
    }
}


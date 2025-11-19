using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    public enum LinkedFileType
    {
        Electrical,    // Contains "EL"
        Mechanical,    // Contains "ME" 
        Plumbing,      // Contains "PH"
        FireProtection,  // Contains "FP" or "FF"
        Architectural, // Contains "ARC"
        Structural,    // Contains "STR"
        Unknown
    }

    public enum MepCategory
    {
        Pipes,
        Ducts,
        DuctAccessories,
        DuctFittings,
        CableTrays,
        Conduits
    }

    public enum HostCategory
    {
        // Architectural
        Walls,
        Ceilings,
        
        // Structural
        StructuralFraming,
        Floors
    }

    public class LinkedFileDetectionService
    {
        public static LinkedFileType DetectFileType(string fileName)
        {
            if (string.IsNullOrEmpty(fileName))
                return LinkedFileType.Unknown;

            var upperFileName = fileName.ToUpper();
            
            if (upperFileName.Contains("EL") || upperFileName.Contains("EE"))
                return LinkedFileType.Electrical;
            if (upperFileName.Contains("ME"))
                return LinkedFileType.Mechanical;
            if (upperFileName.Contains("PH"))
                return LinkedFileType.Plumbing;
            if (upperFileName.Contains("FP") || upperFileName.Contains("FF"))
                return LinkedFileType.FireProtection;
            if (upperFileName.Contains("ARC") || upperFileName.Contains("AR"))
                return LinkedFileType.Architectural;
            if (upperFileName.Contains("STR") || upperFileName.Contains("-ST-") || upperFileName.Contains("ST"))
                return LinkedFileType.Structural;
                
            return LinkedFileType.Unknown;
        }

        public static List<MepCategory> GetAvailableCategories(LinkedFileType fileType)
        {
            return fileType switch
            {
                LinkedFileType.Electrical => new List<MepCategory> 
                { 
                    MepCategory.CableTrays
                    // Removed MepCategory.Conduits
                },
                LinkedFileType.Mechanical => new List<MepCategory> 
                { 
                    MepCategory.Ducts, 
                    MepCategory.DuctAccessories
                    // Removed MepCategory.DuctFittings
                },
                LinkedFileType.Plumbing => new List<MepCategory> 
                { 
                    MepCategory.Pipes 
                },
                LinkedFileType.FireProtection => new List<MepCategory> 
                { 
                    MepCategory.Pipes 
                },
                LinkedFileType.Architectural or LinkedFileType.Structural => new List<MepCategory>(),
                _ => new List<MepCategory>()
            };
        }

        public static List<HostCategory> GetAvailableHostCategories(LinkedFileType fileType)
        {
            return fileType switch
            {
                LinkedFileType.Architectural => new List<HostCategory>
                {
                    HostCategory.Walls,
                    HostCategory.Ceilings,
                    HostCategory.Floors  // ✅ ADDED: Enable Floors for architectural linked files
                },
                LinkedFileType.Structural => new List<HostCategory>
                {
                    HostCategory.StructuralFraming,
                    HostCategory.Floors
                },
                _ => new List<HostCategory>()
            };
        }

        public static bool IsHostOpeningFile(LinkedFileType fileType)
        {
            return fileType == LinkedFileType.Architectural || fileType == LinkedFileType.Structural;
        }

        public static bool RequiresDamperClearance(MepCategory category)
        {
            return category == MepCategory.DuctAccessories;
        }

        public static string GetCategoryDisplayName(MepCategory category)
        {
            return category switch
            {
                MepCategory.Pipes => "Pipes",
                MepCategory.Ducts => "Ducts",
                MepCategory.DuctAccessories => "Duct Accessories",
                MepCategory.DuctFittings => "Duct Fittings",
                MepCategory.CableTrays => "Cable Trays",
                MepCategory.Conduits => "Conduits",
                _ => category.ToString()
            };
        }

        public static string GetHostCategoryDisplayName(HostCategory category)
        {
            return category switch
            {
                HostCategory.Walls => "Walls",
                HostCategory.Ceilings => "Ceilings",
                HostCategory.StructuralFraming => "Structural Framing",
                HostCategory.Floors => "Floors",
                _ => category.ToString()
            };
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// ✅ CRITICAL: DO NOT MODIFY THIS WORKING LOGIC WITHOUT USER CONSENT
    /// This class handles matching existing clash zones by (MEP Element ID, Structural Element ID) pair.
    /// PROTECTED LOGIC: Tested and working - any changes must be discussed first.
    /// </summary>
    public class ClashZonePairMatcher
    {
        private readonly Document _document;
        
        public ClashZonePairMatcher(Document document)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
        }
        
        /// <summary>
        /// ✅ CRITICAL PROTECTED LOGIC: Finds existing clash zone by (MEP Element ID, Structural Element ID) pair ONLY.
        /// 
        /// PROTECTION RULES:
        /// 1. Only matches by MEP ID + Structural ID (no intersection point check)
        /// 2. If 3-point validation is NOT enabled AND pair exists → skip clash zone creation
        /// 3. DO NOT touch flags or parameters of existing clash zone
        /// 4. This prevents duplicate clash zone creation with new GUIDs
        /// 
        /// ⚠️ DO NOT MODIFY THIS LOGIC WITHOUT:
        /// - User discussion and consent
        /// - Understanding why it was implemented this way
        /// - Testing impact on duplicate prevention
        /// </summary>
        /// <param name="mepElementId">MEP element ID</param>
        /// <param name="structuralElementId">Structural element ID</param>
        /// <param name="existingClashZones">List of existing clash zones to search</param>
        /// <returns>Existing clash zone if found, null otherwise</returns>
        public ClashZone FindExistingClashZoneByPair(
            ElementId mepElementId, 
            ElementId structuralElementId, 
            List<ClashZone> existingClashZones)
        {
            if (mepElementId == null || structuralElementId == null)
                return null;
                
            if (existingClashZones == null || existingClashZones.Count == 0)
                return null;
            
            // ✅ CRITICAL: Only match by (MEP ID, Structural ID) pair - NO intersection point check
            // This prevents creating new clash zones with new GUIDs when pair already exists
            int mepIdValue = mepElementId.IntegerValue;
            int structuralIdValue = structuralElementId.IntegerValue;
            
            if (mepIdValue <= 0 || structuralIdValue <= 0)
                return null;
            
            // ✅ PROTECTED LOGIC: Find by IntegerValue to handle XML deserialization cases
            // Matches by exact pair only - no tolerance, no point matching
            var existing = existingClashZones.FirstOrDefault(cz =>
            {
                if (cz == null) return false;
                
                int czMepId = cz.MepElementId?.IntegerValue ?? cz.MepElementIdValue;
                int czStructuralId = cz.StructuralElementId?.IntegerValue ?? cz.StructuralElementIdValue;
                
                return czMepId == mepIdValue && czStructuralId == structuralIdValue;
            });
            
            return existing;
        }
        
        /// <summary>
        /// ✅ CRITICAL PROTECTED LOGIC: Checks if a (MEP ID, Structural ID) pair already exists in existing clash zones.
        /// 
        /// PROTECTION RULES:
        /// 1. Returns true if pair exists (regardless of flags or parameters)
        /// 2. Used BEFORE clash zone creation to prevent duplicates
        /// 3. When pair exists AND 3-point validation is NOT enabled → skip creation
        /// 
        /// ⚠️ DO NOT MODIFY THIS LOGIC WITHOUT USER CONSENT
        /// </summary>
        /// <param name="mepElementId">MEP element ID</param>
        /// <param name="structuralElementId">Structural element ID</param>
        /// <param name="existingClashZones">List of existing clash zones to search</param>
        /// <returns>True if pair exists, false otherwise</returns>
        public bool PairExists(ElementId mepElementId, ElementId structuralElementId, List<ClashZone> existingClashZones)
        {
            return FindExistingClashZoneByPair(mepElementId, structuralElementId, existingClashZones) != null;
        }
    }
}


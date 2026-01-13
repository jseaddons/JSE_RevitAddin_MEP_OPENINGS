namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Shared grouping key for clustering sleeves (host/system/orientation).
    /// Extracted from UniversalClusterService for Phase 8 algorithm service reuse.
    /// </summary>
    public struct SleeveGroupKey
    {
        public string hostType;
        public string systemType;
        public string orientation;
        public int HostElementId; // ✅ NEW: Track host element ID to prevent cross-wall clustering

        public SleeveGroupKey(string hostType, string systemType, string orientation, int hostElementId = -1)
        {
            this.hostType = hostType;
            this.systemType = systemType;
            this.orientation = orientation;
            this.HostElementId = hostElementId;
        }

        public override bool Equals(object obj)
        {
            if (!(obj is SleeveGroupKey)) return false;
            var other = (SleeveGroupKey)obj;
            // ✅ Include HostId in equality check to separate sleeves in different walls
            return hostType == other.hostType && 
                   systemType == other.systemType && 
                   orientation == other.orientation &&
                   HostElementId == other.HostElementId;
        }

        public override int GetHashCode()
        {
            return (hostType, systemType, orientation, HostElementId).GetHashCode();
        }
    }
}

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

        public SleeveGroupKey(string hostType, string systemType, string orientation)
        {
            this.hostType = hostType;
            this.systemType = systemType;
            this.orientation = orientation;
        }

        public override bool Equals(object obj)
        {
            if (!(obj is SleeveGroupKey)) return false;
            var other = (SleeveGroupKey)obj;
            return hostType == other.hostType && systemType == other.systemType && orientation == other.orientation;
        }

        public override int GetHashCode()
        {
            return (hostType, systemType, orientation).GetHashCode();
        }
    }
}

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
        public int SpatialX;
        public int SpatialY;
        public int SpatialZ;

        public SleeveGroupKey(string hostType, string systemType, string orientation, int spatialX, int spatialY, int spatialZ)
        {
            this.hostType = hostType;
            this.systemType = systemType;
            this.orientation = orientation;
            this.SpatialX = spatialX;
            this.SpatialY = spatialY;
            this.SpatialZ = spatialZ;
        }

        public override bool Equals(object obj)
        {
            if (!(obj is SleeveGroupKey)) return false;
            var other = (SleeveGroupKey)obj;
            // ✅ SPATIAL GROUPING: Group by category/orientation/spatial-bucket instead of HostID
            // This ensures sleeves on the same wall/floor AND nearby are grouped together, preventing duplicates
            return hostType == other.hostType && 
                   systemType == other.systemType && 
                   orientation == other.orientation &&
                   SpatialX == other.SpatialX &&
                   SpatialY == other.SpatialY &&
                   SpatialZ == other.SpatialZ;
        }

        public override int GetHashCode()
        {
            return (hostType, systemType, orientation, SpatialX, SpatialY, SpatialZ).GetHashCode();
        }
    }
}

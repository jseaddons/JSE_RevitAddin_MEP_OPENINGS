namespace JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor
{
    /// <summary>
    /// Team B: Configuration abstraction for optimization features.
    /// Groups related flags into immutable records for better organization and testability.
    /// Provides coexistence with existing OptimizationFlags static class.
    /// </summary>
    public interface IOptimizationConfig
    {
        /// <summary>
        /// Placement-related optimization options
        /// </summary>
        PlacementOptions Placement { get; }
        
        /// <summary>
        /// Detection-related optimization options (spatial indexing, filtering)
        /// </summary>
        DetectionOptions Detection { get; }
        
        /// <summary>
        /// Performance monitoring and logging options
        /// </summary>
        PerformanceOptions Performance { get; }
        
        /// <summary>
        /// Configuration version for future migrations
        /// </summary>
        int Version { get; }
    }
    
    /// <summary>
    /// Immutable class for placement optimization options.
    /// Uses constructor initialization for immutability (read-only properties).
    /// </summary>
    public class PlacementOptions
    {
        public bool UseParallelClearance { get; }
        public bool UseParameterBatching { get; }
        public bool UseSmartReplay { get; }
        public bool UseNewSleevePlacerService { get; }
        public bool UseNewSleeveRepository { get; }
        
        public PlacementOptions(
            bool useParallelClearance,
            bool useParameterBatching,
            bool useSmartReplay,
            bool useNewSleevePlacerService,
            bool useNewSleeveRepository)
        {
            UseParallelClearance = useParallelClearance;
            UseParameterBatching = useParameterBatching;
            UseSmartReplay = useSmartReplay;
            UseNewSleevePlacerService = useNewSleevePlacerService;
            UseNewSleeveRepository = useNewSleeveRepository;
        }
    }
    
    /// <summary>
    /// Immutable class for detection optimization options.
    /// Uses constructor initialization for immutability (read-only properties).
    /// </summary>
    public class DetectionOptions
    {
        public bool UseRTreeIndex { get; }
        public bool UseSpatialGrid { get; }
        public bool UseRTreeDatabaseIndex { get; }
        public bool UseBoundingBoxSectionBoxFilter { get; }
        public bool UseCurveInBoundingBoxFilter { get; }
        public bool UseViewIndependentFilter { get; }
        
        public DetectionOptions(
            bool useRTreeIndex,
            bool useSpatialGrid,
            bool useRTreeDatabaseIndex,
            bool useBoundingBoxSectionBoxFilter,
            bool useCurveInBoundingBoxFilter,
            bool useViewIndependentFilter)
        {
            UseRTreeIndex = useRTreeIndex;
            UseSpatialGrid = useSpatialGrid;
            UseRTreeDatabaseIndex = useRTreeDatabaseIndex;
            UseBoundingBoxSectionBoxFilter = useBoundingBoxSectionBoxFilter;
            UseCurveInBoundingBoxFilter = useCurveInBoundingBoxFilter;
            UseViewIndependentFilter = useViewIndependentFilter;
        }
    }
    
    /// <summary>
    /// Immutable class for performance monitoring options.
    /// Uses constructor initialization for immutability (read-only properties).
    /// </summary>
    public class PerformanceOptions
    {
        public bool EnablePerformanceLogging { get; }
        public bool EnableDetailedTiming { get; }
        
        public PerformanceOptions(
            bool enablePerformanceLogging,
            bool enableDetailedTiming)
        {
            EnablePerformanceLogging = enablePerformanceLogging;
            EnableDetailedTiming = enableDetailedTiming;
        }
    }
}


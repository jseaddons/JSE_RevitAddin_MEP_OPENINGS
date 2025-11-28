using System.Collections.Generic;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Placement
{
    /// <summary>
    /// Result of a single placement stage execution.
    /// Encapsulates success status, updated context, and errors for fail-safe semantics.
    /// </summary>
    public sealed class StageResult
    {
        public bool Success { get; }
        public PlacementContext Context { get; }
        public IReadOnlyList<string> Errors { get; }

        public StageResult(bool success, PlacementContext context, IReadOnlyList<string> errors)
        {
            Success = success;
            Context = context;
            Errors = errors ?? System.Array.Empty<string>();
        }

        /// <summary>
        /// Create a successful stage result with updated context.
        /// </summary>
        public static StageResult Succeeded(PlacementContext context) =>
            new StageResult(true, context, System.Array.Empty<string>());

        /// <summary>
        /// Create a failed stage result with error messages.
        /// </summary>
        public static StageResult Failed(PlacementContext context, params string[] errors) =>
            new StageResult(false, context, errors);

        /// <summary>
        /// Create a failed stage result with error list.
        /// </summary>
        public static StageResult Failed(PlacementContext context, IReadOnlyList<string> errors) =>
            new StageResult(false, context, errors);
    }
}

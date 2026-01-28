using JSE_RevitAddin_MEP_OPENINGS.Services.Caching;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces
{
    public interface ILevelCacheConsumer
    {
        void SetGlobalLevelCache(GlobalLevelCache cache);
    }
}

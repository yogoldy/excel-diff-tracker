using ExcelDiffTracker.Storage;
using ExcelDiffTracker.Tracking;

namespace ExcelDiffTracker.App.Services;

public sealed class AppServices : IAsyncDisposable
{
    public AppServices()
    {
        Store = new HistoryStore(AppPaths.DatabasePath);
        Coordinator = new TrackingCoordinator(Store);
        Startup = new StartupManager();
        Themes = new ThemeManager();
    }

    public HistoryStore Store { get; }
    public TrackingCoordinator Coordinator { get; }
    public StartupManager Startup { get; }
    public ThemeManager Themes { get; }

    public async ValueTask DisposeAsync()
    {
        await Coordinator.DisposeAsync().ConfigureAwait(false);
        Themes.Dispose();
    }
}

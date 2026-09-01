using System.Threading;
using System.Windows;
using ExcelDiffTracker.App.Services;

namespace ExcelDiffTracker.App;

public partial class App : Application
{
    private Mutex? _singleInstance;
    private AppServices? _services;
    private TrayIconService? _tray;
    private EventWaitHandle? _activationEvent;
    private ManualResetEvent? _activationStop;
    private Task? _activationTask;
    private bool _exiting;

    public bool IsExiting => _exiting;

    protected override async void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);
        _singleInstance = new Mutex(initiallyOwned: true, "Local\\ExcelDiffTracker.Singleton", out var isFirstInstance);
        if (!isFirstInstance)
        {
            try
            {
                using var activation = EventWaitHandle.OpenExisting("Local\\ExcelDiffTracker.Activate");
                activation.Set();
            }
            catch (WaitHandleCannotBeOpenedException)
            {
            }
            Shutdown();
            return;
        }

        _activationEvent = new EventWaitHandle(false, EventResetMode.AutoReset, "Local\\ExcelDiffTracker.Activate");
        _activationStop = new ManualResetEvent(false);

        try
        {
            _services = new AppServices();
            await _services.Coordinator.InitializeAsync();

            var storedTheme = await _services.Store.GetSettingAsync("theme");
            var selectedTheme = Enum.TryParse<AppTheme>(storedTheme, out var parsedTheme) ? parsedTheme : AppTheme.System;
            _services.Themes.Apply(selectedTheme);

            var onboardingComplete = string.Equals(await _services.Store.GetSettingAsync("onboarding_complete"), "true", StringComparison.OrdinalIgnoreCase);
            var startupSetting = await _services.Store.GetSettingAsync("start_with_windows");
            if (onboardingComplete && startupSetting is null)
            {
                try
                {
                    _services.Startup.SetEnabled(true);
                    await _services.Store.SetSettingAsync("start_with_windows", "true");
                }
                catch
                {
                    await _services.Store.SetSettingAsync("start_with_windows", "false");
                }
            }

            var mainWindow = new MainWindow(_services);
            MainWindow = mainWindow;
            _tray = new TrayIconService(mainWindow, _services.Store, _services.Coordinator, ExitApplicationAsync);
            _activationTask = Task.Run(() => ActivationLoop(mainWindow));

            if (!onboardingComplete)
            {
                var onboarding = new OnboardingWindow(_services);
                _ = onboarding.ShowDialog();
            }

            await mainWindow.ViewModel.RefreshAsync();
            if (!e.Args.Contains("--background", StringComparer.OrdinalIgnoreCase))
                mainWindow.Show();
            await _services.Coordinator.StartAsync();
        }
        catch (Exception exception)
        {
            MessageBox.Show(
                $"Excel Diff Tracker could not start.\n\n{exception.Message}",
                "Startup problem",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
            await ExitApplicationAsync();
        }
    }

    public async Task ExitApplicationAsync()
    {
        if (_exiting)
            return;
        _exiting = true;
        _activationStop?.Set();
        _activationEvent?.Set();
        _tray?.Dispose();
        if (_services is not null)
            await _services.DisposeAsync();
        _singleInstance?.ReleaseMutex();
        _singleInstance?.Dispose();
        if (_activationTask is not null)
            await _activationTask;
        _activationEvent?.Dispose();
        _activationStop?.Dispose();
        Shutdown();
    }

    private void ActivationLoop(MainWindow window)
    {
        if (_activationEvent is null || _activationStop is null)
            return;
        var handles = new WaitHandle[] { _activationEvent, _activationStop };
        while (WaitHandle.WaitAny(handles) == 0)
        {
            if (_exiting)
                return;
            _ = Dispatcher.BeginInvoke(() =>
            {
                window.Show();
                if (window.WindowState == WindowState.Minimized)
                    window.WindowState = WindowState.Normal;
                window.Activate();
            });
        }
    }
}

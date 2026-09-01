using System.Windows;
using ExcelDiffTracker.Core;
using ExcelDiffTracker.Storage;
using ExcelDiffTracker.Tracking;

namespace ExcelDiffTracker.App.Services;

public sealed class TrayIconService : IDisposable
{
    private readonly MainWindow _window;
    private readonly HistoryStore _store;
    private readonly TrackingCoordinator _coordinator;
    private readonly Func<Task> _exit;
    private readonly System.Windows.Forms.NotifyIcon _icon;
    private System.Drawing.Icon? _stateIcon;

    public TrayIconService(MainWindow window, HistoryStore store, TrackingCoordinator coordinator, Func<Task> exit)
    {
        _window = window;
        _store = store;
        _coordinator = coordinator;
        _exit = exit;
        _stateIcon = AppIconFactory.Create(TrayState.Idle);
        _icon = new System.Windows.Forms.NotifyIcon
        {
            Icon = _stateIcon,
            Text = "Excel Diff Tracker — tracking is active",
            Visible = true,
            ContextMenuStrip = BuildMenu()
        };
        _icon.DoubleClick += (_, _) => ShowWindow();
        _coordinator.CaptureOccurred += OnCaptureOccurred;
        _ = UpdateStateAsync();
    }

    public void Dispose()
    {
        _coordinator.CaptureOccurred -= OnCaptureOccurred;
        _icon.Visible = false;
        _icon.Dispose();
        _stateIcon?.Dispose();
    }

    private System.Windows.Forms.ContextMenuStrip BuildMenu()
    {
        var menu = new System.Windows.Forms.ContextMenuStrip();
        menu.Items.Add("Open Excel Diff Tracker", null, (_, _) => ShowWindow());
        menu.Items.Add("Add workbook…", null, (_, _) =>
        {
            ShowWindow();
            _window.ViewModel.Navigate("Workbooks");
            _window.ViewModel.AddWorkbookCommand.Execute(null);
        });
        menu.Items.Add(new System.Windows.Forms.ToolStripSeparator());
        menu.Items.Add("Exit", null, async (_, _) => await _exit());
        return menu;
    }

    private void OnCaptureOccurred(object? sender, CaptureEvent captureEvent)
    {
        _ = Application.Current.Dispatcher.BeginInvoke(() => _ = HandleCaptureEventAsync(captureEvent));
    }

    private async Task HandleCaptureEventAsync(CaptureEvent captureEvent)
    {
        try
        {
            await _window.ViewModel.RefreshAsync();
            await UpdateStateAsync();
            if (captureEvent.Kind != CaptureEventKind.DuplicateIgnored &&
                !(captureEvent.Kind == CaptureEventKind.StatusChanged && captureEvent.Message?.StartsWith("Processing", StringComparison.Ordinal) == true))
                _window.ViewModel.ShowToast(captureEvent.Message ?? "Workbook history updated.", captureEvent.Kind == CaptureEventKind.Failed);
            if (captureEvent.Kind is CaptureEventKind.Captured or CaptureEventKind.NoTrackedChanges or CaptureEventKind.Failed)
            {
                _icon.BalloonTipTitle = captureEvent.Kind == CaptureEventKind.Failed ? "Capture needs attention" : "Workbook save captured";
                _icon.BalloonTipText = captureEvent.Message ?? "Excel Diff Tracker updated the local history.";
                _icon.BalloonTipIcon = captureEvent.Kind == CaptureEventKind.Failed ? System.Windows.Forms.ToolTipIcon.Warning : System.Windows.Forms.ToolTipIcon.Info;
                _icon.ShowBalloonTip(4_000);
            }
        }
        catch (Exception exception)
        {
            _window.ViewModel.ShowToast(exception.Message, isError: true);
        }
    }

    private async Task UpdateStateAsync()
    {
        try
        {
            var workbooks = await _store.GetTrackedWorkbooksAsync();
            var state = workbooks.Any(item => item.Status == TrackingStatus.Processing)
                ? TrayState.Processing
                : workbooks.Any(item => item.Status is TrackingStatus.Warning or TrackingStatus.Missing)
                    ? TrayState.Warning
                    : workbooks.Count > 0 && workbooks.All(item => !item.IsEnabled)
                        ? TrayState.Paused
                        : TrayState.Idle;
            SetState(state, workbooks.Count == 0);
        }
        catch
        {
            SetState(TrayState.Warning);
        }
    }

    private void SetState(TrayState state, bool noWorkbooks = false)
    {
        var replacement = AppIconFactory.Create(state);
        _icon.Icon = replacement;
        _stateIcon?.Dispose();
        _stateIcon = replacement;
        _icon.Text = state switch
        {
            TrayState.Processing => "Excel Diff Tracker — processing",
            TrayState.Paused => "Excel Diff Tracker — paused",
            TrayState.Warning => "Excel Diff Tracker — attention needed",
            _ when noWorkbooks => "Excel Diff Tracker — no workbooks yet",
            _ => "Excel Diff Tracker — tracking is active"
        };
    }

    private void ShowWindow()
    {
        _window.Show();
        if (_window.WindowState == WindowState.Minimized)
            _window.WindowState = WindowState.Normal;
        _window.Activate();
        _window.Topmost = true;
        _window.Topmost = false;
    }
}

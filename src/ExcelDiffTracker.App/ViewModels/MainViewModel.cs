using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Windows;
using System.Windows.Input;
using System.Windows.Data;
using ExcelDiffTracker.App.Services;
using ExcelDiffTracker.Core;
using Microsoft.Win32;

namespace ExcelDiffTracker.App.ViewModels;

public sealed class MainViewModel : ObservableObject
{
    private readonly AppServices _services;
    private string _currentPage = "Dashboard";
    private string _selectedTheme = AppTheme.System.ToString();
    private bool _startWithWindows;
    private string _defaultReportDirectory = AppPaths.SuggestedReportsDirectory;
    private bool _toastVisible;
    private bool _toastIsError;
    private string _toastMessage = string.Empty;
    private CancellationTokenSource? _toastCancellation;
    private string _selectedHistoryWorkbook = "All workbooks";
    private bool _showValueHistory = true;
    private bool _showFormulaHistory = true;
    private bool _showSheetHistory = true;
    private bool _showErrors = true;

    public MainViewModel(AppServices services)
    {
        _services = services;
        NavigateCommand = new RelayCommand(parameter => Navigate(parameter?.ToString() ?? "Dashboard"));
        AddWorkbookCommand = new AsyncRelayCommand(_ => AddWorkbookAsync());
        RefreshCommand = new AsyncRelayCommand(_ => RefreshAsync());
        ToggleTrackingCommand = new AsyncRelayCommand(ToggleTrackingAsync, parameter => parameter is TrackedWorkbook);
        PurgeCommand = new AsyncRelayCommand(PurgeAsync, parameter => parameter is TrackedWorkbook);
        OpenReportCommand = new RelayCommand(OpenReport, parameter => parameter is VersionRecord { ReportPath: not null });
        ExportFullCommand = new AsyncRelayCommand(ExportFullAsync, parameter => parameter is VersionRecord);
        ChooseDefaultReportFolderCommand = new AsyncRelayCommand(_ => ChooseDefaultReportFolderAsync());
        ApplyThemeCommand = new AsyncRelayCommand(_ => ApplyThemeAsync());
        SaveStartupCommand = new AsyncRelayCommand(_ => SaveStartupAsync());
        ChangeWorkbookReportFolderCommand = new AsyncRelayCommand(ChangeWorkbookReportFolderAsync, parameter => parameter is TrackedWorkbook);
        VersionsView = CollectionViewSource.GetDefaultView(Versions);
        VersionsView.Filter = FilterVersion;
        ErrorsView = CollectionViewSource.GetDefaultView(Errors);
        ErrorsView.Filter = FilterError;
    }

    public ObservableCollection<TrackedWorkbook> Workbooks { get; } = [];
    public ObservableCollection<VersionRecord> Versions { get; } = [];
    public ObservableCollection<CaptureErrorRecord> Errors { get; } = [];
    public ObservableCollection<string> HistoryWorkbookOptions { get; } = ["All workbooks"];
    public IReadOnlyList<string> Themes { get; } = Enum.GetNames<AppTheme>();
    public ICollectionView VersionsView { get; }
    public ICollectionView ErrorsView { get; }

    public ICommand NavigateCommand { get; }
    public ICommand AddWorkbookCommand { get; }
    public ICommand RefreshCommand { get; }
    public ICommand ToggleTrackingCommand { get; }
    public ICommand PurgeCommand { get; }
    public ICommand OpenReportCommand { get; }
    public ICommand ExportFullCommand { get; }
    public ICommand ChooseDefaultReportFolderCommand { get; }
    public ICommand ApplyThemeCommand { get; }
    public ICommand SaveStartupCommand { get; }
    public ICommand ChangeWorkbookReportFolderCommand { get; }

    public string CurrentPage
    {
        get => _currentPage;
        private set => SetProperty(ref _currentPage, value);
    }

    public string SelectedTheme
    {
        get => _selectedTheme;
        set => SetProperty(ref _selectedTheme, value);
    }

    public bool StartWithWindows
    {
        get => _startWithWindows;
        set => SetProperty(ref _startWithWindows, value);
    }

    public string DefaultReportDirectory
    {
        get => _defaultReportDirectory;
        set => SetProperty(ref _defaultReportDirectory, value);
    }

    public bool ToastVisible
    {
        get => _toastVisible;
        private set => SetProperty(ref _toastVisible, value);
    }

    public bool ToastIsError
    {
        get => _toastIsError;
        private set => SetProperty(ref _toastIsError, value);
    }

    public string ToastMessage
    {
        get => _toastMessage;
        private set => SetProperty(ref _toastMessage, value);
    }

    public string SelectedHistoryWorkbook
    {
        get => _selectedHistoryWorkbook;
        set { if (SetProperty(ref _selectedHistoryWorkbook, value)) RefreshHistoryFilters(); }
    }

    public bool ShowValueHistory
    {
        get => _showValueHistory;
        set { if (SetProperty(ref _showValueHistory, value)) VersionsView.Refresh(); }
    }

    public bool ShowFormulaHistory
    {
        get => _showFormulaHistory;
        set { if (SetProperty(ref _showFormulaHistory, value)) VersionsView.Refresh(); }
    }

    public bool ShowSheetHistory
    {
        get => _showSheetHistory;
        set { if (SetProperty(ref _showSheetHistory, value)) VersionsView.Refresh(); }
    }

    public bool ShowErrors
    {
        get => _showErrors;
        set { if (SetProperty(ref _showErrors, value)) ErrorsView.Refresh(); }
    }

    public string ActiveSummary => Workbooks.Count == 0
        ? "No workbooks are being tracked yet."
        : $"{Workbooks.Count(item => item.IsEnabled):N0} active of {Workbooks.Count:N0} workbooks";

    public void Navigate(string page) => CurrentPage = page;

    public async Task RefreshAsync()
    {
        var workbooks = await _services.Store.GetTrackedWorkbooksAsync();
        var versions = await _services.Store.GetVersionsAsync(limit: 10_000);
        var errors = await _services.Store.GetErrorsAsync(limit: 10_000);
        Replace(Workbooks, workbooks);
        Replace(Versions, versions);
        Replace(Errors, errors);
        var selected = SelectedHistoryWorkbook;
        HistoryWorkbookOptions.Clear();
        HistoryWorkbookOptions.Add("All workbooks");
        foreach (var path in workbooks.Select(item => item.Path).Distinct(StringComparer.OrdinalIgnoreCase))
            HistoryWorkbookOptions.Add(path);
        SelectedHistoryWorkbook = HistoryWorkbookOptions.Contains(selected) ? selected : "All workbooks";
        RefreshHistoryFilters();
        StartWithWindows = _services.Startup.IsEnabled;
        DefaultReportDirectory = await _services.Store.GetSettingAsync("default_report_directory") ?? AppPaths.SuggestedReportsDirectory;
        SelectedTheme = await _services.Store.GetSettingAsync("theme") ?? AppTheme.System.ToString();
        OnPropertyChanged(nameof(ActiveSummary));
    }

    public void ShowToast(string message, bool isError = false)
    {
        _toastCancellation?.Cancel();
        _toastCancellation?.Dispose();
        _toastCancellation = new CancellationTokenSource();
        ToastMessage = message;
        ToastIsError = isError;
        ToastVisible = true;
        _ = HideToastAsync(_toastCancellation.Token);
    }

    private async Task AddWorkbookAsync()
    {
        try
        {
            var picker = new OpenFileDialog
            {
                Title = "Choose a workbook to track",
                Filter = "Excel workbooks (*.xlsx;*.xlsm)|*.xlsx;*.xlsm",
                Multiselect = false,
                CheckFileExists = true
            };
            if (picker.ShowDialog() != true)
                return;

            var reportDirectory = await _services.Store.GetSettingAsync("default_report_directory") ?? DefaultReportDirectory;
            Directory.CreateDirectory(reportDirectory);
            await _services.Coordinator.AddWorkbookAsync(picker.FileName, reportDirectory);
            await RefreshAsync();
            Navigate("Dashboard");
        }
        catch (Exception exception)
        {
            ShowToast(exception.Message, isError: true);
        }
    }

    private async Task ToggleTrackingAsync(object? parameter)
    {
        if (parameter is not TrackedWorkbook workbook)
            return;
        try
        {
            await _services.Coordinator.SetEnabledAsync(workbook.Id, !workbook.IsEnabled);
            await RefreshAsync();
        }
        catch (Exception exception)
        {
            ShowToast(exception.Message, isError: true);
        }
    }

    private async Task PurgeAsync(object? parameter)
    {
        if (parameter is not TrackedWorkbook workbook)
            return;
        var answer = MessageBox.Show(
            $"Permanently delete all local history and generated reports for:\n\n{workbook.Path}\n\nThe workbook itself will not be changed.",
            "Permanently delete history?",
            MessageBoxButton.YesNo,
            MessageBoxImage.Warning,
            MessageBoxResult.No);
        if (answer != MessageBoxResult.Yes)
            return;
        try
        {
            await _services.Coordinator.PurgeAsync(workbook.Id, confirmed: true);
            await RefreshAsync();
            ShowToast("Workbook history was permanently deleted.");
        }
        catch (Exception exception)
        {
            ShowToast(exception.Message, isError: true);
        }
    }

    private static void OpenReport(object? parameter)
    {
        if (parameter is not VersionRecord { ReportPath: { } path })
            return;
        if (!File.Exists(path))
        {
            MessageBox.Show("That automatic report is missing. Use Full export to regenerate it from the local history database.", "Report missing", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }
        _ = Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
    }

    private async Task ExportFullAsync(object? parameter)
    {
        if (parameter is not VersionRecord version)
            return;
        var picker = new SaveFileDialog
        {
            Title = "Export complete Markdown report",
            Filter = "Markdown (*.md)|*.md",
            FileName = $"{Path.GetFileNameWithoutExtension(version.WorkbookPath)}-version-{version.Sequence:D6}-full.md",
            AddExtension = true
        };
        if (picker.ShowDialog() != true)
            return;
        try
        {
            await _services.Coordinator.ExportFullReportAsync(version, picker.FileName);
            ShowToast("Full Markdown report exported.");
        }
        catch (Exception exception)
        {
            ShowToast(exception.Message, isError: true);
        }
    }

    private async Task ChooseDefaultReportFolderAsync()
    {
        using var picker = new System.Windows.Forms.FolderBrowserDialog
        {
            Description = "Choose where Excel Diff Tracker should save Markdown reports",
            InitialDirectory = DefaultReportDirectory,
            ShowNewFolderButton = true,
            UseDescriptionForTitle = true
        };
        if (picker.ShowDialog() != System.Windows.Forms.DialogResult.OK)
            return;
        DefaultReportDirectory = picker.SelectedPath;
        await _services.Store.SetSettingAsync("default_report_directory", picker.SelectedPath);
        ShowToast("Default report folder updated.");
    }

    private async Task ApplyThemeAsync()
    {
        if (!Enum.TryParse<AppTheme>(SelectedTheme, out var theme))
            theme = AppTheme.System;
        _services.Themes.Apply(theme);
        await _services.Store.SetSettingAsync("theme", theme.ToString());
    }

    private async Task SaveStartupAsync()
    {
        try
        {
            _services.Startup.SetEnabled(StartWithWindows);
            await _services.Store.SetSettingAsync("start_with_windows", StartWithWindows ? "true" : "false");
            ShowToast(StartWithWindows ? "Excel Diff Tracker will start with Windows." : "Windows startup disabled.");
        }
        catch (Exception exception)
        {
            ShowToast(exception.Message, isError: true);
        }
    }

    private async Task ChangeWorkbookReportFolderAsync(object? parameter)
    {
        if (parameter is not TrackedWorkbook workbook)
            return;
        using var picker = new System.Windows.Forms.FolderBrowserDialog
        {
            Description = $"Choose the Markdown report folder for {Path.GetFileName(workbook.Path)}",
            InitialDirectory = workbook.ReportDirectory,
            ShowNewFolderButton = true,
            UseDescriptionForTitle = true
        };
        if (picker.ShowDialog() != System.Windows.Forms.DialogResult.OK)
            return;
        await _services.Store.UpdateReportDirectoryAsync(workbook.Id, picker.SelectedPath);
        await RefreshAsync();
        ShowToast("Workbook report folder updated.");
    }

    private bool FilterVersion(object item)
    {
        if (item is not VersionRecord version || !MatchesWorkbook(version.WorkbookPath))
            return false;
        if (ShowValueHistory && ShowFormulaHistory && ShowSheetHistory)
            return true;
        return ShowValueHistory && (version.LiteralChangeCount > 0 || version.CellTypeChangeCount > 0)
            || ShowFormulaHistory && (version.FormulaChangeCount > 0 || version.FormulaResultChangeCount > 0)
            || ShowSheetHistory && version.SheetChangeCount > 0;
    }

    private bool FilterError(object item) => ShowErrors && item is CaptureErrorRecord error && MatchesWorkbook(error.WorkbookPath);
    private bool MatchesWorkbook(string path) => SelectedHistoryWorkbook == "All workbooks" || string.Equals(SelectedHistoryWorkbook, path, StringComparison.OrdinalIgnoreCase);
    private void RefreshHistoryFilters()
    {
        VersionsView.Refresh();
        ErrorsView.Refresh();
    }

    private async Task HideToastAsync(CancellationToken cancellationToken)
    {
        try
        {
            await Task.Delay(TimeSpan.FromSeconds(5), cancellationToken);
            ToastVisible = false;
        }
        catch (OperationCanceledException)
        {
        }
    }

    private static void Replace<T>(ObservableCollection<T> target, IEnumerable<T> values)
    {
        target.Clear();
        foreach (var value in values)
            target.Add(value);
    }
}

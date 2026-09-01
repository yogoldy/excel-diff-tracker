using System.ComponentModel;
using System.Windows;
using ExcelDiffTracker.App.Services;
using Microsoft.Win32;

namespace ExcelDiffTracker.App;

public partial class OnboardingWindow : Window, INotifyPropertyChanged
{
    private readonly AppServices _services;
    private readonly FrameworkElement[] _steps;
    private int _step = 1;
    private string _reportFolder = AppPaths.SuggestedReportsDirectory;
    private string _workbookPath = string.Empty;
    private bool _completed;

    public OnboardingWindow(AppServices services)
    {
        _services = services;
        InitializeComponent();
        _steps = [StepOne, StepTwo, StepThree, StepFour, StepFive];
        DataContext = this;
        UpdateStep();
    }

    public event PropertyChangedEventHandler? PropertyChanged;
    public int StepNumber => _step;
    public string StepLabel => $"Step {_step} of 5";
    public string ReportFolder
    {
        get => _reportFolder;
        private set { _reportFolder = value; PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(ReportFolder))); }
    }
    public string WorkbookPath
    {
        get => _workbookPath;
        private set
        {
            _workbookPath = value;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(WorkbookPath)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(WorkbookName)));
        }
    }
    public string WorkbookName => string.IsNullOrWhiteSpace(WorkbookPath) ? "No workbook selected" : Path.GetFileName(WorkbookPath);

    private void ChooseReportFolder(object sender, RoutedEventArgs args)
    {
        using var picker = new System.Windows.Forms.FolderBrowserDialog
        {
            Description = "Choose where Markdown reports should be saved",
            InitialDirectory = ReportFolder,
            ShowNewFolderButton = true,
            UseDescriptionForTitle = true
        };
        if (picker.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            ReportFolder = picker.SelectedPath;
    }

    private void ChooseWorkbook(object sender, RoutedEventArgs args)
    {
        var picker = new OpenFileDialog
        {
            Title = "Choose a workbook to track",
            Filter = "Excel workbooks (*.xlsx;*.xlsm)|*.xlsx;*.xlsm",
            CheckFileExists = true
        };
        if (picker.ShowDialog() == true)
            WorkbookPath = picker.FileName;
    }

    private void Back(object sender, RoutedEventArgs args)
    {
        if (_step > 1 && _step < 5)
        {
            _step--;
            UpdateStep();
        }
    }

    private async void Next(object sender, RoutedEventArgs args)
    {
        if (_step == 3 && string.IsNullOrWhiteSpace(WorkbookPath))
        {
            StepThreeStatus.Text = "Choose an .xlsx or .xlsm workbook before continuing.";
            return;
        }

        if (_step == 4)
        {
            NextButton.IsEnabled = false;
            BackButton.IsEnabled = false;
            SetupStatus.Text = "Creating the local baseline…";
            try
            {
                Directory.CreateDirectory(ReportFolder);
                await _services.Store.SetSettingAsync("default_report_directory", ReportFolder);
                var startup = StartupCheck.IsChecked == true;
                _services.Startup.SetEnabled(startup);
                await _services.Store.SetSettingAsync("start_with_windows", startup ? "true" : "false");
                await _services.Coordinator.AddWorkbookAsync(WorkbookPath, ReportFolder);
                await _services.Store.SetSettingAsync("onboarding_complete", "true");
                _completed = true;
            }
            catch (Exception exception)
            {
                SetupStatus.Text = exception.Message;
                NextButton.IsEnabled = true;
                BackButton.IsEnabled = true;
                return;
            }
        }

        if (_step == 5)
        {
            await _services.Store.SetSettingAsync("onboarding_complete", "true");
            _completed = true;
            DialogResult = true;
            Close();
            return;
        }

        _step++;
        UpdateStep();
    }

    private void UpdateStep()
    {
        for (var index = 0; index < _steps.Length; index++)
            _steps[index].Visibility = index == _step - 1 ? Visibility.Visible : Visibility.Collapsed;
        BackButton.Visibility = _step is > 1 and < 5 ? Visibility.Visible : Visibility.Hidden;
        NextButton.Content = _step switch { 4 => "Activate tracking", 5 => "Open dashboard", _ => "Continue" };
        NextButton.IsEnabled = true;
        BackButton.IsEnabled = true;
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(StepNumber)));
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(StepLabel)));
    }

    protected override void OnClosing(CancelEventArgs args)
    {
        if (!_completed)
        {
            var answer = MessageBox.Show("Finish setup later? Excel Diff Tracker will remain available, but no workbook will be tracked until onboarding is complete.", "Leave setup?", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No);
            if (answer != MessageBoxResult.Yes)
            {
                args.Cancel = true;
                return;
            }
        }
        base.OnClosing(args);
    }
}

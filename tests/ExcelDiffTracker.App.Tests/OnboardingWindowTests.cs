using System.Runtime.ExceptionServices;
using System.Threading;
using System.IO;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Automation.Peers;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Media;
using ExcelDiffTracker.App.Services;
using Xunit;

namespace ExcelDiffTracker.App.Tests;

public sealed class OnboardingWindowTests
{
    [Fact]
    public void LegacyCleanupDeletesOnlyDatabaseFilesAndLeavesMarkdownEvidenceUntouched()
    {
        var directory = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"scenario-cleanup-{Guid.NewGuid():N}");
        Directory.CreateDirectory(directory);
        try
        {
            foreach (var name in new[] { "history.db", "history.db-shm", "history.db-wal" })
                File.WriteAllText(System.IO.Path.Combine(directory, name), "old local database");
            var markdown = System.IO.Path.Combine(directory, "saved-evidence.md");
            File.WriteAllText(markdown, "retain");

            LegacyDataCleanup.DeleteDatabaseFiles(directory);

            Assert.Empty(LegacyDataCleanup.FindDatabaseFiles(directory));
            Assert.True(File.Exists(markdown));
            Assert.Equal("retain", File.ReadAllText(markdown));
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void LegacyCleanupPreflightsEveryDatabaseFileBeforeDeletingAny()
    {
        var directory = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"scenario-cleanup-lock-{Guid.NewGuid():N}");
        Directory.CreateDirectory(directory);
        var paths = new[] { "history.db", "history.db-shm", "history.db-wal" }
            .Select(name => System.IO.Path.Combine(directory, name))
            .ToArray();
        try
        {
            foreach (var path in paths)
                File.WriteAllText(path, "old local database");
            using (File.Open(paths[1], FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                Assert.Throws<IOException>(() => LegacyDataCleanup.DeleteDatabaseFiles(directory));
            Assert.All(paths, path => Assert.True(File.Exists(path)));
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void ApplicationResourcesAndPrimaryWindowsPreserveThemeLayoutAndAutomationContracts()
    {
        Exception? failure = null;
        var thread = new Thread(() =>
        {
            try
            {
                var application = new App();
                application.InitializeComponent();
                using var services = new TestServices();
                var window = new OnboardingWindow(services.Value);
                var progress = Assert.IsType<ProgressBar>(window.FindName("StepProgress"));
                var binding = BindingOperations.GetBinding(progress, ProgressBar.ValueProperty);

                Assert.NotNull(binding);
                Assert.Equal(BindingMode.OneWay, binding.Mode);

                var mutedStyle = Assert.IsType<Style>(application.Resources["MutedText"]);
                var pageTitleStyle = Assert.IsType<Style>(application.Resources["PageTitle"]);
                var sectionTitleStyle = Assert.IsType<Style>(application.Resources["SectionTitle"]);
                Assert.NotNull(mutedStyle.BasedOn);
                Assert.NotNull(pageTitleStyle.BasedOn);
                Assert.NotNull(sectionTitleStyle.BasedOn);
                Assert.Contains(pageTitleStyle.Setters.OfType<Setter>(), setter => setter.Property == TextBlock.FontSizeProperty && Equals(setter.Value, 32d));

                var contentScroller = Assert.IsType<ScrollViewer>(window.FindName("StepContentScrollViewer"));
                Assert.Equal(ScrollBarVisibility.Auto, contentScroller.VerticalScrollBarVisibility);
                Assert.Equal(ScrollBarVisibility.Disabled, contentScroller.HorizontalScrollBarVisibility);
                Assert.True(double.IsNaN(Assert.IsType<Border>(window.FindName("ReportFolderCard")).Width));
                Assert.True(double.IsNaN(Assert.IsType<Border>(window.FindName("WorkbookCard")).Width));
                Assert.True(double.IsNaN(Assert.IsType<Border>(window.FindName("OptionsCard")).Width));
                Assert.Equal("OnboardingWindow", AutomationProperties.GetAutomationId(window));
                Assert.Equal(window.Title, new WindowAutomationPeer(window).GetName());
                Assert.Equal("OnboardingNextButton", AutomationProperties.GetAutomationId(Assert.IsType<Button>(window.FindName("NextButton"))));

                services.Value.Themes.Apply(AppTheme.Dark);
                if (!SystemParameters.HighContrast)
                {
                    Assert.Equal((Color)ColorConverter.ConvertFromString("#EDF2EE"), GetResourceColor(application, "TextBrush"));
                    Assert.Equal((Color)ColorConverter.ConvertFromString("#2D332F"), GetResourceColor(application, "CardBrush"));
                    Assert.Equal((Color)ColorConverter.ConvertFromString("#10221D"), GetResourceColor(application, "PrimaryForegroundBrush"));
                }

                var mainWindow = new MainWindow(services.Value);
                Assert.Equal("MainWindow", AutomationProperties.GetAutomationId(mainWindow));
                Assert.Equal(mainWindow.Title, new WindowAutomationPeer(mainWindow).GetName());
                AssertTemplateAutomation(mainWindow, "DashboardWorkbooks", 1);
                AssertTemplateAutomation(mainWindow, "WorkbookRows", 5);
                AssertTemplateAutomation(mainWindow, "HistoryVersions", 5);
                Assert.Equal("SettingsNavigationButton", AutomationProperties.GetAutomationId(Assert.IsType<Button>(mainWindow.FindName("SettingsNavigationButton"))));
                var themeComboBox = Assert.IsType<ComboBox>(mainWindow.FindName("ThemeComboBox"));
                Assert.Equal("ThemeComboBox", AutomationProperties.GetAutomationId(themeComboBox));
                Assert.NotNull(themeComboBox.Template);
                Assert.Equal(ScrollBarVisibility.Auto, Assert.IsType<ScrollViewer>(mainWindow.FindName("SettingsScrollViewer")).VerticalScrollBarVisibility);
                Assert.Equal(ScrollBarVisibility.Auto, Assert.IsType<ScrollViewer>(mainWindow.FindName("AboutScrollViewer")).VerticalScrollBarVisibility);
                Assert.Equal("Excel Scenario Analysis Tool", mainWindow.Title);
                Assert.Equal("Welcome to Excel Scenario Analysis Tool", window.Title);
                if (!SystemParameters.HighContrast)
                {
                    Assert.Equal(GetResourceColor(application, "TextBrush"), Assert.IsType<SolidColorBrush>(themeComboBox.Foreground).Color);
                    Assert.Equal(GetResourceColor(application, "CardBrush"), Assert.IsType<SolidColorBrush>(themeComboBox.Background).Color);
                }
            }
            catch (Exception exception)
            {
                failure = exception;
            }
        });

        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        thread.Join();

        if (failure is not null)
            ExceptionDispatchInfo.Capture(failure).Throw();
    }

    [Theory]
    [InlineData(AppTheme.Light)]
    [InlineData(AppTheme.Dark)]
    public void ThemePalettesMeetNormalTextContrastRequirements(AppTheme theme)
    {
        var palette = ThemeManager.GetPalette(theme);

        AssertContrast(palette, "TextBrush", "AppBackgroundBrush");
        AssertContrast(palette, "MutedTextBrush", "AppBackgroundBrush");
        AssertContrast(palette, "TextBrush", "CardBrush");
        AssertContrast(palette, "MutedTextBrush", "CardBrush");
        AssertContrast(palette, "PrimaryForegroundBrush", "AccentBrush");
        AssertContrast(palette, "AccentBrush", "AccentSoftBrush");
        AssertContrast(palette, "WarningBrush", "CardBrush");
        AssertContrast(palette, "WarningBrush", "AppBackgroundBrush");
        AssertContrast(palette, "ErrorBrush", "CardBrush");
    }

    [Theory]
    [InlineData(AppTheme.Light)]
    [InlineData(AppTheme.Dark)]
    public void ThemePalettesMeetEssentialUiContrastRequirements(AppTheme theme)
    {
        var palette = ThemeManager.GetPalette(theme);

        AssertContrast(palette, "BorderBrush", "CardBrush", 3);
        AssertContrast(palette, "BorderBrush", "AppBackgroundBrush", 3);
        AssertContrast(palette, "AccentBrush", "CardBrush", 3);
        AssertContrast(palette, "AccentBrush", "SidebarBrush", 3);
        AssertContrast(palette, "AccentBrush", "AppBackgroundBrush", 3);
        AssertContrast(palette, "PrimaryForegroundBrush", "AccentBrush", 3);
    }

    private static void AssertTemplateAutomation(MainWindow window, string controlName, int minimumActionCount)
    {
        var items = Assert.IsAssignableFrom<ItemsControl>(window.FindName(controlName));
        var actionIds = new HashSet<string>();
        for (var record = 1; record <= 2; record++)
        {
            var path = @"C:\Acceptance\deliberately long workbook paths\" + new string('x', 90) + record + ".xlsx";
            var root = Assert.IsAssignableFrom<FrameworkElement>(items.ItemTemplate.LoadContent());
            root.DataContext = new
            {
                Id = record,
                Path = path,
                FileName = System.IO.Path.GetFileName(path),
                Directory = System.IO.Path.GetDirectoryName(path),
                WorkbookPath = path,
                ReportDirectory = @"C:\Acceptance\reports",
                IsEnabled = true,
                Status = ExcelDiffTracker.Core.TrackingStatus.Active,
                CurrentSequence = 1,
                Sequence = 1,
                LastSuccessfulCaptureUtc = DateTime.UtcNow,
                CapturedUtc = DateTime.UtcNow,
                LastSummary = "One cell changed",
                Summary = "One cell changed",
                LastError = (string?)null,
                ComparisonSummary = "One changed cell",
                ComparisonBaselineLabel = "Previous save"
            };
            root.Measure(new Size(760, 520));
            root.Arrange(new Rect(0, 0, 760, 520));
            root.UpdateLayout();
            var elements = Descendants(root).ToArray();
            var buttons = elements.OfType<Button>().ToArray();
            Assert.True(buttons.Length >= minimumActionCount, $"Expected at least {minimumActionCount} actions, found {buttons.Length}.");
            foreach (var button in buttons)
            {
                var peer = new ButtonAutomationPeer(button);
                Assert.False(string.IsNullOrWhiteSpace(peer.GetAutomationId()));
                Assert.True(actionIds.Add(peer.GetAutomationId()), "Repeated records must expose distinct action IDs.");
                Assert.False(string.IsNullOrWhiteSpace(peer.GetName()));
                Assert.True(button.Focusable);
            }
            Assert.Contains(elements.OfType<TextBlock>(), text =>
                new TextBlockAutomationPeer(text).GetHelpText() == path && Equals(text.ToolTip, path));
        }
    }

    private static IEnumerable<DependencyObject> Descendants(DependencyObject root)
    {
        foreach (var child in LogicalTreeHelper.GetChildren(root).OfType<DependencyObject>())
        {
            yield return child;
            foreach (var descendant in Descendants(child))
                yield return descendant;
        }
    }

    private static Color GetResourceColor(Application application, string key) =>
        Assert.IsType<SolidColorBrush>(application.Resources[key]).Color;

    private static void AssertContrast(IReadOnlyDictionary<string, string> palette, string foreground, string background, double minimum = 4.5)
    {
        var foregroundColor = (Color)ColorConverter.ConvertFromString(palette[foreground]);
        var backgroundColor = (Color)ColorConverter.ConvertFromString(palette[background]);
        var lighter = Math.Max(RelativeLuminance(foregroundColor), RelativeLuminance(backgroundColor));
        var darker = Math.Min(RelativeLuminance(foregroundColor), RelativeLuminance(backgroundColor));
        Assert.True((lighter + 0.05) / (darker + 0.05) >= minimum, $"{foreground} on {background} did not meet {minimum:N1}:1 contrast.");
    }

    private static double RelativeLuminance(Color color) =>
        0.2126 * Linearize(color.R / 255d) + 0.7152 * Linearize(color.G / 255d) + 0.0722 * Linearize(color.B / 255d);

    private static double Linearize(double channel) =>
        channel <= 0.04045 ? channel / 12.92 : Math.Pow((channel + 0.055) / 1.055, 2.4);

    private sealed class TestServices : IDisposable
    {
        private readonly string? _previousDirectory;
        private readonly string _testDirectory;

        public TestServices()
        {
            _testDirectory = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"scenario-app-tests-{Guid.NewGuid():N}");
            _previousDirectory = Environment.GetEnvironmentVariable("EXCEL_DIFF_TRACKER_TEST_DATA_DIRECTORY");
            Environment.SetEnvironmentVariable("EXCEL_DIFF_TRACKER_TEST_DATA_DIRECTORY", _testDirectory);
            Value = new AppServices();
        }

        public AppServices Value { get; }

        public void Dispose()
        {
            Value.DisposeAsync().AsTask().GetAwaiter().GetResult();
            Environment.SetEnvironmentVariable("EXCEL_DIFF_TRACKER_TEST_DATA_DIRECTORY", _previousDirectory);
            if (Directory.Exists(_testDirectory))
                Directory.Delete(_testDirectory, recursive: true);
        }
    }
}

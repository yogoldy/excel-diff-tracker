using System.Runtime.ExceptionServices;
using System.Threading;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Media;
using ExcelDiffTracker.App.Services;
using Xunit;

namespace ExcelDiffTracker.App.Tests;

public sealed class OnboardingWindowTests
{
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

                var contentScroller = Assert.IsType<ScrollViewer>(window.FindName("StepContentScrollViewer"));
                Assert.Equal(ScrollBarVisibility.Auto, contentScroller.VerticalScrollBarVisibility);
                Assert.Equal(ScrollBarVisibility.Disabled, contentScroller.HorizontalScrollBarVisibility);
                Assert.True(double.IsNaN(Assert.IsType<Border>(window.FindName("ReportFolderCard")).Width));
                Assert.True(double.IsNaN(Assert.IsType<Border>(window.FindName("WorkbookCard")).Width));
                Assert.True(double.IsNaN(Assert.IsType<Border>(window.FindName("OptionsCard")).Width));
                Assert.Equal("OnboardingWindow", AutomationProperties.GetAutomationId(window));
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
                Assert.Equal("SettingsNavigationButton", AutomationProperties.GetAutomationId(Assert.IsType<Button>(mainWindow.FindName("SettingsNavigationButton"))));
                var themeComboBox = Assert.IsType<ComboBox>(mainWindow.FindName("ThemeComboBox"));
                Assert.Equal("ThemeComboBox", AutomationProperties.GetAutomationId(themeComboBox));
                Assert.NotNull(themeComboBox.Template);
                Assert.Equal(ScrollBarVisibility.Auto, Assert.IsType<ScrollViewer>(mainWindow.FindName("SettingsScrollViewer")).VerticalScrollBarVisibility);
                Assert.Equal(ScrollBarVisibility.Auto, Assert.IsType<ScrollViewer>(mainWindow.FindName("AboutScrollViewer")).VerticalScrollBarVisibility);
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
        AssertContrast(palette, "ErrorBrush", "CardBrush");
    }

    private static Color GetResourceColor(Application application, string key) =>
        Assert.IsType<SolidColorBrush>(application.Resources[key]).Color;

    private static void AssertContrast(IReadOnlyDictionary<string, string> palette, string foreground, string background)
    {
        var foregroundColor = (Color)ColorConverter.ConvertFromString(palette[foreground]);
        var backgroundColor = (Color)ColorConverter.ConvertFromString(palette[background]);
        var lighter = Math.Max(RelativeLuminance(foregroundColor), RelativeLuminance(backgroundColor));
        var darker = Math.Min(RelativeLuminance(foregroundColor), RelativeLuminance(backgroundColor));
        Assert.True((lighter + 0.05) / (darker + 0.05) >= 4.5, $"{foreground} on {background} did not meet 4.5:1 contrast.");
    }

    private static double RelativeLuminance(Color color) =>
        0.2126 * Linearize(color.R / 255d) + 0.7152 * Linearize(color.G / 255d) + 0.0722 * Linearize(color.B / 255d);

    private static double Linearize(double channel) =>
        channel <= 0.04045 ? channel / 12.92 : Math.Pow((channel + 0.055) / 1.055, 2.4);

    private sealed class TestServices : IDisposable
    {
        public TestServices() => Value = new AppServices();

        public AppServices Value { get; }

        public void Dispose() => Value.DisposeAsync().AsTask().GetAwaiter().GetResult();
    }
}

using System.Windows;
using System.Windows.Media;
using Microsoft.Win32;

namespace ExcelDiffTracker.App.Services;

public enum AppTheme
{
    System,
    Light,
    Dark
}

public sealed class ThemeManager : IDisposable
{
    private static readonly IReadOnlyDictionary<string, string> DarkColors = new Dictionary<string, string>
    {
        ["AppBackgroundBrush"] = "#242825", ["SidebarBrush"] = "#1D211F", ["CardBrush"] = "#2D332F",
        ["CardHoverBrush"] = "#35423C", ["TextBrush"] = "#EDF2EE", ["MutedTextBrush"] = "#A9B4AF",
        ["BorderBrush"] = "#465049", ["AccentBrush"] = "#65CBB0", ["AccentSoftBrush"] = "#314D45",
        ["WarningBrush"] = "#E5A64B", ["ErrorBrush"] = "#F07D7D", ["PrimaryForegroundBrush"] = "#10221D"
    };

    private static readonly IReadOnlyDictionary<string, string> LightColors = new Dictionary<string, string>
    {
        ["AppBackgroundBrush"] = "#F6F4EF", ["SidebarBrush"] = "#ECE9E1", ["CardBrush"] = "#FFFEFB",
        ["CardHoverBrush"] = "#F3F7F4", ["TextBrush"] = "#24312D", ["MutedTextBrush"] = "#5E6965",
        ["BorderBrush"] = "#D9DED9", ["AccentBrush"] = "#236B5C", ["AccentSoftBrush"] = "#DDEDE7",
        ["WarningBrush"] = "#A86413", ["ErrorBrush"] = "#B33A3A", ["PrimaryForegroundBrush"] = "#FFFFFF"
    };

    private AppTheme _selected = AppTheme.System;

    public ThemeManager() => SystemEvents.UserPreferenceChanged += OnUserPreferenceChanged;

    public void Apply(AppTheme theme)
    {
        _selected = theme;
        var dark = theme == AppTheme.Dark || theme == AppTheme.System && IsSystemDark();
        var colors = GetPalette(dark ? AppTheme.Dark : AppTheme.Light);

        if (SystemParameters.HighContrast)
        {
            Application.Current.Resources["AppBackgroundBrush"] = SystemColors.WindowBrush;
            Application.Current.Resources["SidebarBrush"] = SystemColors.ControlBrush;
            Application.Current.Resources["CardBrush"] = SystemColors.WindowBrush;
            Application.Current.Resources["CardHoverBrush"] = SystemColors.ControlLightBrush;
            Application.Current.Resources["TextBrush"] = SystemColors.WindowTextBrush;
            Application.Current.Resources["MutedTextBrush"] = SystemColors.GrayTextBrush;
            Application.Current.Resources["BorderBrush"] = SystemColors.ActiveBorderBrush;
            Application.Current.Resources["AccentBrush"] = SystemColors.HighlightBrush;
            Application.Current.Resources["AccentSoftBrush"] = SystemColors.ControlLightBrush;
            Application.Current.Resources["WarningBrush"] = SystemColors.WindowTextBrush;
            Application.Current.Resources["ErrorBrush"] = SystemColors.WindowTextBrush;
            Application.Current.Resources["PrimaryForegroundBrush"] = SystemColors.HighlightTextBrush;
            return;
        }

        foreach (var (key, value) in colors)
            Application.Current.Resources[key] = new SolidColorBrush((Color)ColorConverter.ConvertFromString(value));
    }

    public void Dispose() => SystemEvents.UserPreferenceChanged -= OnUserPreferenceChanged;

    internal static IReadOnlyDictionary<string, string> GetPalette(AppTheme theme) =>
        theme == AppTheme.Dark ? DarkColors : LightColors;

    private void OnUserPreferenceChanged(object sender, UserPreferenceChangedEventArgs args)
    {
        if (Application.Current is null)
            return;
        _ = Application.Current.Dispatcher.BeginInvoke(() => Apply(_selected));
    }

    private static bool IsSystemDark()
    {
        using var key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize");
        return key?.GetValue("AppsUseLightTheme") is int value && value == 0;
    }
}

using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Media;

namespace ExcelDiffTracker.App.Services;

public sealed class WindowThemeService : IDisposable
{
    private const int UseImmersiveDarkMode = 20;
    private const int CaptionColor = 35;
    private const int TextColor = 36;
    private const uint DefaultColor = 0xFFFFFFFF;
    private readonly Window _window;
    private readonly ThemeManager _themes;

    private WindowThemeService(Window window, ThemeManager themes)
    {
        _window = window;
        _themes = themes;
        _window.SourceInitialized += OnSourceInitialized;
        _themes.ThemeChanged += OnThemeChanged;
    }

    public static WindowThemeService Attach(Window window, ThemeManager themes) => new(window, themes);

    public void Dispose()
    {
        _window.SourceInitialized -= OnSourceInitialized;
        _themes.ThemeChanged -= OnThemeChanged;
    }

    private void OnSourceInitialized(object? sender, EventArgs args) => Apply();
    private void OnThemeChanged(object? sender, EventArgs args) => Apply();

    private void Apply()
    {
        var handle = new WindowInteropHelper(_window).Handle;
        if (handle == nint.Zero)
            return;
        var dark = _themes.IsDark && !SystemParameters.HighContrast ? 1 : 0;
        _ = DwmSetWindowAttribute(handle, UseImmersiveDarkMode, ref dark, sizeof(int));
        if (SystemParameters.HighContrast)
        {
            var reset = DefaultColor;
            _ = DwmSetWindowAttribute(handle, CaptionColor, ref reset, sizeof(uint));
            _ = DwmSetWindowAttribute(handle, TextColor, ref reset, sizeof(uint));
            return;
        }

        var caption = ToColorRef(((SolidColorBrush)Application.Current.Resources["SidebarBrush"]).Color);
        var text = ToColorRef(((SolidColorBrush)Application.Current.Resources["TextBrush"]).Color);
        _ = DwmSetWindowAttribute(handle, CaptionColor, ref caption, sizeof(uint));
        _ = DwmSetWindowAttribute(handle, TextColor, ref text, sizeof(uint));
    }

    private static uint ToColorRef(Color color) => (uint)(color.R | color.G << 8 | color.B << 16);

    [DllImport("dwmapi.dll")]
    private static extern int DwmSetWindowAttribute(nint window, int attribute, ref int value, int size);

    [DllImport("dwmapi.dll")]
    private static extern int DwmSetWindowAttribute(nint window, int attribute, ref uint value, int size);
}

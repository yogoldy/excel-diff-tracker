using System.Reflection;

namespace ExcelDiffTracker.App.Services;

public static class AppBrand
{
    public const string ProductName = "Excel Scenario Analysis Tool";
    public const string ShortName = "Scenario Analysis";

    public static string Version
    {
        get
        {
            var version = Assembly.GetEntryAssembly()?.GetName().Version;
            return version is null ? "Unknown" : $"{version.Major}.{version.Minor}.{Math.Max(0, version.Build)}";
        }
    }

    public static string Platform => Environment.Is64BitProcess ? "Windows 11 ARM64" : "Windows";
}

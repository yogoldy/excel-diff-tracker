namespace ExcelDiffTracker.App.Services;

public static class AppPaths
{
    public static string LocalDataDirectory { get; } = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "Excel Diff Tracker");

    public static string DatabasePath => Path.Combine(LocalDataDirectory, "history.db");

    public static string SuggestedReportsDirectory => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
        "Excel Diff Tracker Reports");
}

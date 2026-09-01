namespace ExcelDiffTracker.App.Services;

public static class AppPaths
{
    public static string LocalDataDirectory { get; } = ResolveLocalDataDirectory();

    public static string DatabasePath => Path.Combine(LocalDataDirectory, "history.db");

    public static string SuggestedReportsDirectory => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
        "Excel Diff Tracker Reports");

    private static string ResolveLocalDataDirectory()
    {
        var testDirectory = Environment.GetEnvironmentVariable("EXCEL_DIFF_TRACKER_TEST_DATA_DIRECTORY");
        return string.IsNullOrWhiteSpace(testDirectory)
            ? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Excel Diff Tracker")
            : Path.GetFullPath(testDirectory);
    }
}

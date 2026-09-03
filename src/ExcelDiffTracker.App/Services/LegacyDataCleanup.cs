using System.Windows;

namespace ExcelDiffTracker.App.Services;

public static class LegacyDataCleanup
{
    private static readonly string[] DatabaseFiles = ["history.db-wal", "history.db-shm", "history.db"];

    public static void OfferOnFirstLaunch()
    {
        if (!string.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable("EXCEL_DIFF_TRACKER_TEST_DATA_DIRECTORY")))
            return;
        var legacyDirectory = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "Excel Diff Tracker");
        var existing = FindDatabaseFiles(legacyDirectory);
        if (existing.Length == 0)
            return;

        var answer = MessageBox.Show(
            "Excel Scenario Analysis Tool starts with fresh local history. Delete the old Excel Diff Tracker database files now?\n\nExisting Markdown reports and workbooks will not be moved, changed, or deleted.",
            "Remove old local app data?",
            MessageBoxButton.YesNo,
            MessageBoxImage.Question,
            MessageBoxResult.No);
        if (answer != MessageBoxResult.Yes)
            return;

        try
        {
            DeleteDatabaseFiles(legacyDirectory);
        }
        catch (Exception exception)
        {
            MessageBox.Show(
                $"Cleanup could not finish. No Markdown reports or workbooks were touched, but some old SQLite files may remain. Close the older app and try again.\n\n{exception.Message}",
                "Old database cleanup incomplete",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
        }
    }

    internal static string[] FindDatabaseFiles(string legacyDirectory) => DatabaseFiles
        .Select(name => Path.Combine(legacyDirectory, name))
        .Where(File.Exists)
        .ToArray();

    internal static void DeleteDatabaseFiles(string legacyDirectory)
    {
        var existing = FindDatabaseFiles(legacyDirectory);
        var locks = new List<FileStream>();
        try
        {
            foreach (var path in existing)
            {
                if ((File.GetAttributes(path) & FileAttributes.ReadOnly) != 0)
                    throw new IOException($"The old SQLite file is read-only: {Path.GetFileName(path)}");
                locks.Add(File.Open(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None));
            }
        }
        finally
        {
            foreach (var stream in locks)
                stream.Dispose();
        }

        foreach (var path in existing)
            File.Delete(path);
        if (Directory.Exists(legacyDirectory) && !Directory.EnumerateFileSystemEntries(legacyDirectory).Any())
            Directory.Delete(legacyDirectory);
    }
}

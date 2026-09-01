namespace ExcelDiffTracker.Tests;

public sealed class TestDirectory : IDisposable
{
    public TestDirectory()
    {
        Path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"ExcelDiffTrackerTests-{Guid.NewGuid():N}");
        Directory.CreateDirectory(Path);
    }
    public string Path { get; }
    public void Dispose()
    {
        try
        {
            if (Directory.Exists(Path))
                Directory.Delete(Path, recursive: true);
        }
        catch (IOException)
        {
        }
        catch (UnauthorizedAccessException)
        {
        }
    }
}

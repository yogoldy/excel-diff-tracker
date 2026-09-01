using System.Security.Cryptography;

namespace ExcelDiffTracker.Tracking;

public sealed record StableCopyResult : IDisposable
{
    public required string TemporaryPath { get; init; }
    public required string Sha256 { get; init; }
    public required DateTime SourceLastWriteUtc { get; init; }
    public required long SourceLength { get; init; }

    public void Dispose()
    {
        try
        {
            if (File.Exists(TemporaryPath))
                File.Delete(TemporaryPath);
        }
        catch (IOException)
        {
            // The OS temp directory will eventually clean up a copy that is still in use.
        }
        catch (UnauthorizedAccessException)
        {
            // Cleanup failure must not hide the capture result.
        }
    }
}

public sealed class StableWorkbookCopy
{
    private readonly TimeSpan _stableInterval;
    private readonly TimeSpan _retryInterval;
    private readonly TimeSpan _timeout;

    public StableWorkbookCopy(TimeSpan? stableInterval = null, TimeSpan? retryInterval = null, TimeSpan? timeout = null)
    {
        _stableInterval = stableInterval ?? TimeSpan.FromMilliseconds(300);
        _retryInterval = retryInterval ?? TimeSpan.FromSeconds(1);
        _timeout = timeout ?? TimeSpan.FromSeconds(60);
    }

    public async Task<StableCopyResult> CreateAsync(string sourcePath, CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sourcePath);
        var fullPath = Path.GetFullPath(sourcePath);
        var started = DateTime.UtcNow;
        Exception? lastError = null;

        while (DateTime.UtcNow - started < _timeout)
        {
            cancellationToken.ThrowIfCancellationRequested();
            try
            {
                var first = ReadMetadata(fullPath);
                await Task.Delay(_stableInterval, cancellationToken).ConfigureAwait(false);
                var second = ReadMetadata(fullPath);
                if (first.Length != second.Length || first.LastWriteUtc != second.LastWriteUtc)
                    continue;

                var temporaryPath = Path.Combine(Path.GetTempPath(), $"ExcelDiffTracker-{Guid.NewGuid():N}{Path.GetExtension(fullPath)}");
                try
                {
                    await using (var source = new FileStream(
                        fullPath,
                        FileMode.Open,
                        FileAccess.Read,
                        FileShare.Read | FileShare.Delete,
                        1024 * 1024,
                        FileOptions.Asynchronous | FileOptions.SequentialScan))
                    await using (var target = new FileStream(
                        temporaryPath,
                        FileMode.CreateNew,
                        FileAccess.Write,
                        FileShare.None,
                        1024 * 1024,
                        FileOptions.Asynchronous | FileOptions.SequentialScan))
                    {
                        await source.CopyToAsync(target, 1024 * 1024, cancellationToken).ConfigureAwait(false);
                        await target.FlushAsync(cancellationToken).ConfigureAwait(false);
                    }

                    var after = ReadMetadata(fullPath);
                    var copyLength = new FileInfo(temporaryPath).Length;
                    if (second.Length != after.Length || second.LastWriteUtc != after.LastWriteUtc || copyLength != second.Length)
                    {
                        File.Delete(temporaryPath);
                        continue;
                    }

                    var sha256 = await ComputeHashAsync(temporaryPath, cancellationToken).ConfigureAwait(false);
                    return new StableCopyResult
                    {
                        TemporaryPath = temporaryPath,
                        Sha256 = sha256,
                        SourceLastWriteUtc = after.LastWriteUtc,
                        SourceLength = after.Length
                    };
                }
                catch
                {
                    if (File.Exists(temporaryPath))
                        File.Delete(temporaryPath);
                    throw;
                }
            }
            catch (Exception exception) when (exception is IOException or UnauthorizedAccessException or FileNotFoundException or DirectoryNotFoundException)
            {
                lastError = exception;
                await Task.Delay(_retryInterval, cancellationToken).ConfigureAwait(false);
            }
        }

        throw new IOException($"The workbook did not become readable and stable within {_timeout.TotalSeconds:N0} seconds.", lastError);
    }

    private static (long Length, DateTime LastWriteUtc) ReadMetadata(string path)
    {
        var file = new FileInfo(path);
        if (!file.Exists)
            throw new FileNotFoundException("Workbook not found.", path);
        return (file.Length, file.LastWriteTimeUtc);
    }

    private static async Task<string> ComputeHashAsync(string path, CancellationToken cancellationToken)
    {
        await using var stream = new FileStream(
            path,
            FileMode.Open,
            FileAccess.Read,
            FileShare.Read,
            1024 * 1024,
            FileOptions.Asynchronous | FileOptions.SequentialScan);
        var hash = await SHA256.HashDataAsync(stream, cancellationToken).ConfigureAwait(false);
        return Convert.ToHexStringLower(hash);
    }
}

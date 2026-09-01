using System.IO.Compression;

namespace ExcelDiffTracker.Core;

// Safety limits and package-first validation are adapted from xlsx-review
// (MIT, Copyright 2026 CinciNeuro / Henry Bloomingdale), commit c39bcf7.
internal static class WorkbookPackageGuard
{
    private const int MaxEntryCount = 20_000;
    private const long MaxTotalUncompressedBytes = 512L * 1024 * 1024;
    private const long MaxEntryUncompressedBytes = 256L * 1024 * 1024;
    private const long CompressionRatioThresholdBytes = 64L * 1024;
    private const double MaxCompressionRatio = 150d;

    public static void Validate(string path)
    {
        var extension = Path.GetExtension(path);
        if (!extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase) &&
            !extension.Equals(".xlsm", StringComparison.OrdinalIgnoreCase))
        {
            throw new UnsupportedWorkbookException("Only .xlsx and .xlsm workbooks are supported.");
        }

        try
        {
            using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
            using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false);

            if (archive.Entries.Count == 0)
                throw new UnsafeWorkbookException("The workbook package is empty or encrypted.");
            if (archive.Entries.Count > MaxEntryCount)
                throw new UnsafeWorkbookException($"The workbook contains too many package parts ({archive.Entries.Count:N0}).");

            long total = 0;
            long totalCompressed = 0;
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var entry in archive.Entries)
            {
                var normalizedName = entry.FullName.Replace('\\', '/').TrimStart('/');
                if (normalizedName.Contains("../", StringComparison.Ordinal) || !names.Add(normalizedName))
                    throw new UnsafeWorkbookException($"The workbook contains an unsafe or duplicate package path: {normalizedName}");

                total = checked(total + entry.Length);
                totalCompressed = checked(totalCompressed + entry.CompressedLength);
                if (total > MaxTotalUncompressedBytes)
                    throw new UnsafeWorkbookException("The workbook exceeds the 512 MB uncompressed safety limit.");
                if (entry.Length > MaxEntryUncompressedBytes)
                    throw new UnsafeWorkbookException($"Workbook part '{normalizedName}' is too large.");
                if (entry.CompressedLength > 0 && entry.Length >= CompressionRatioThresholdBytes &&
                    (double)entry.Length / entry.CompressedLength > MaxCompressionRatio)
                {
                    throw new UnsafeWorkbookException($"Workbook part '{normalizedName}' has an unsafe compression ratio.");
                }
            }

            if (total >= CompressionRatioThresholdBytes && totalCompressed > 0 &&
                (double)total / totalCompressed > MaxCompressionRatio)
            {
                throw new UnsafeWorkbookException("The workbook package has an unsafe overall compression ratio.");
            }

            if (!names.Contains("[Content_Types].xml") || !names.Contains("_rels/.rels"))
                throw new UnsafeWorkbookException("The file is not a readable Office Open XML workbook.");
        }
        catch (UnsafeWorkbookException)
        {
            throw;
        }
        catch (InvalidDataException exception)
        {
            throw new UnsafeWorkbookException($"The workbook is corrupt or encrypted: {exception.Message}");
        }
    }
}

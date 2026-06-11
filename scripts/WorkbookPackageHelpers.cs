using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;

static class WorkbookPackageHelpers
{
    private const string ConfigFileName = "vba-dev-pack.json";

    public static VbaDevPackConfig ReadConfig(string baseDir)
    {
        var configPath = Path.Combine(baseDir, ConfigFileName);
        if (!File.Exists(configPath))
        {
            return new VbaDevPackConfig();
        }

        try
        {
            var json = File.ReadAllText(configPath, new UTF8Encoding(false));
            return JsonSerializer.Deserialize<VbaDevPackConfig>(json) ?? new VbaDevPackConfig();
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Warning: Could not read {ConfigFileName}: {ex.Message}. Using default codepage.");
            return new VbaDevPackConfig();
        }
    }

    private static readonly string[] CustomUiCandidates =
    {
        "customUI/customUI.xml",
        "customUI/customUI14.xml"
    };

    public static bool TryExportCustomUi(string workbookPath, string outputPath, out string message)
    {
        if (!TryReadCustomUi(workbookPath, out var xmlText, out var partName))
        {
            message = "customUI part was not found in workbook.";
            return false;
        }

        var outputDir = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrEmpty(outputDir))
        {
            Directory.CreateDirectory(outputDir);
        }

        File.WriteAllText(outputPath, xmlText!, new UTF8Encoding(false));
        message = $"Exported customUI from '{partName}' to: {outputPath}";
        return true;
    }

    public static void ImportCustomUi(string workbookPath, string sourcePath)
    {
        if (!File.Exists(sourcePath))
        {
            throw new FileNotFoundException("customUI source file not found.", sourcePath);
        }

        var xmlText = File.ReadAllText(sourcePath, new UTF8Encoding(false, true));

        using var archive = ZipFile.Open(workbookPath, ZipArchiveMode.Update);
        var existingPart = FindCustomUiPart(archive);
        var targetPartName = existingPart?.FullName ?? "customUI/customUI.xml";

        existingPart?.Delete();

        var entry = archive.CreateEntry(targetPartName, CompressionLevel.Optimal);
        using var stream = entry.Open();
        using var writer = new StreamWriter(stream, new UTF8Encoding(false));
        writer.Write(xmlText);
    }

    public static void EnsureWorkbookClosed(string workbookPath)
    {
        try
        {
            using var _ = new FileStream(workbookPath, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        }
        catch (IOException ex)
        {
            throw new InvalidOperationException(
                "Workbook is currently open or locked. Close the workbook in Excel and run the script again.",
                ex);
        }
    }

    public static string CreateTempBackupDirectory(string scenario)
    {
        var safeScenario = string.IsNullOrWhiteSpace(scenario) ? "backup" : scenario.Trim().ToLowerInvariant();
        var dir = Path.Combine(
            Path.GetTempPath(),
            "vba-dev-pack-backups",
            safeScenario,
            DateTime.Now.ToString("yyyyMMdd-HHmmss") + "-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        return dir;
    }

    public static string BackupFile(string sourcePath, string backupRoot, string? relativeTarget = null)
    {
        var targetPath = relativeTarget is null
            ? Path.Combine(backupRoot, Path.GetFileName(sourcePath))
            : Path.Combine(backupRoot, relativeTarget);

        var targetDir = Path.GetDirectoryName(targetPath);
        if (!string.IsNullOrEmpty(targetDir))
        {
            Directory.CreateDirectory(targetDir);
        }

        File.Copy(sourcePath, targetPath, true);
        return targetPath;
    }

    public static string BackupDirectory(string sourceDir, string backupRoot, string relativeTarget)
    {
        var targetDir = Path.Combine(backupRoot, relativeTarget);
        CopyDirectoryRecursive(sourceDir, targetDir);
        return targetDir;
    }

    public static void EnsureVbProjectAccessible(object vbProject)
    {
        object? vbComponents = null;
        try
        {
            vbComponents = ((dynamic)vbProject).VBComponents;
            _ = (int)((dynamic)vbComponents).Count;
        }
        catch (COMException ex)
        {
            var hexHresult = $"0x{(uint)ex.ErrorCode:X8}";
            throw new InvalidOperationException(
                "Cannot access VBProject via COM. Enable: Excel -> File -> Options -> Trust Center -> Trust Center Settings -> Macro Settings -> Trust access to the VBA project object model. " +
                $"HRESULT: {hexHresult}",
                ex);
        }
        finally
        {
            if (vbComponents is not null && Marshal.IsComObject(vbComponents))
            {
                Marshal.FinalReleaseComObject(vbComponents);
            }
        }
    }

    private static void CopyDirectoryRecursive(string sourceDir, string targetDir)
    {
        Directory.CreateDirectory(targetDir);

        foreach (var file in Directory.EnumerateFiles(sourceDir, "*", SearchOption.TopDirectoryOnly))
        {
            var targetFile = Path.Combine(targetDir, Path.GetFileName(file));
            File.Copy(file, targetFile, true);
        }

        foreach (var subDir in Directory.EnumerateDirectories(sourceDir, "*", SearchOption.TopDirectoryOnly))
        {
            var childTarget = Path.Combine(targetDir, Path.GetFileName(subDir));
            CopyDirectoryRecursive(subDir, childTarget);
        }
    }

    private static bool TryReadCustomUi(string workbookPath, out string? xmlText, out string? partName)
    {
        using var archive = ZipFile.OpenRead(workbookPath);
        var entry = FindCustomUiPart(archive);

        if (entry is null)
        {
            xmlText = null;
            partName = null;
            return false;
        }

        using var stream = entry.Open();
        using var reader = new StreamReader(stream, new UTF8Encoding(false, true));
        xmlText = reader.ReadToEnd();
        partName = entry.FullName;
        return true;
    }

    private static ZipArchiveEntry? FindCustomUiPart(ZipArchive archive)
    {
        foreach (var candidate in CustomUiCandidates)
        {
            var entry = archive.GetEntry(candidate);
            if (entry is not null)
            {
                return entry;
            }
        }

        return null;
    }
}

sealed class VbaDevPackConfig
{
    [JsonPropertyName("codepage")]
    public int Codepage { get; set; } = 1251;

    public Encoding GetCodepageEncoding()
    {
        try
        {
            return Encoding.GetEncoding(Codepage);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException(
                $"Codepage {Codepage} specified in vba-dev-pack.json is not supported. " +
                "Use a valid Windows codepage number (e.g. 1251 for Cyrillic, 1252 for Western European, 1250 for Central European).",
                ex);
        }
    }
}

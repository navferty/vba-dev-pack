using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Xml.Linq;

static class WorkbookPackageHelpers
{
    private const string ConfigFileName = "vba-dev-pack.json";
    private const string PackageRelationshipsPath = "_rels/.rels";
    private const string UiExtensibilityRelationshipType = "http://schemas.microsoft.com/office/2006/relationships/ui/extensibility";
    private static readonly XNamespace RelationshipNs = "http://schemas.openxmlformats.org/package/2006/relationships";
    private static readonly UTF8Encoding Utf8NoBom = new(false);
    private static readonly UTF8Encoding Utf8NoBomStrict = new(false, true);
    private static readonly HashSet<string> ConvertibleExtensions = new(StringComparer.OrdinalIgnoreCase)
    {
        ".txt", ".csv", ".md", ".bas", ".cls", ".frm", ".doccls", ".vbs", ".ps1"
    };


    public static EncodingScriptOptions ParseEncodingScriptOptions(string[] args)
    {
        var options = ParseNamedArguments(args);
        var showHelp = options.ContainsKey("help") || options.ContainsKey("h");
        if (showHelp)
        {
            return EncodingScriptOptions.HelpRequested();
        }

        if (options.TryGetValue("error", out var parseError))
        {
            return EncodingScriptOptions.FromError(parseError ?? "Invalid arguments.");
        }

        string? sourceValue;
        if (!TryGetOptionValue(options, "source", "s", out sourceValue))
        {
            sourceValue = Path.Combine(Environment.CurrentDirectory, "source");
        }

        string? filesValue;
        if (!TryGetOptionValue(options, "files", "f", out filesValue))
        {
            filesValue = null;
        }

        var files = new List<string>();
        if (!string.IsNullOrWhiteSpace(filesValue))
        {
            files.AddRange(
                filesValue
                    .Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries));
        }

        return EncodingScriptOptions.FromValues(Path.GetFullPath(sourceValue!), files);
    }

    public static Dictionary<string, string?> ParseNamedArguments(string[] args, HashSet<string>? valuelessOptions = null)
    {
        valuelessOptions ??= new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "help", "h"
        };

        var result = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);
        for (var i = 0; i < args.Length; i++)
        {
            var arg = args[i];
            if (!arg.StartsWith('-'))
            {
                result["error"] = $"Unexpected positional argument: {arg}";
                return result;
            }

            var key = arg.TrimStart('-');
            if (key.Length == 0)
            {
                result["error"] = "Invalid argument syntax.";
                return result;
            }

            if (valuelessOptions.Contains(key))
            {
                result[key] = null;
                continue;
            }

            if (i + 1 >= args.Length || args[i + 1].StartsWith('-'))
            {
                result["error"] = $"Option '--{key}' requires a value.";
                return result;
            }

            result[key] = args[i + 1];
            i++;
        }

        return result;
    }

    public static string BuildEncodingHelp(string scriptName, string direction)
    {
        return $"""
{scriptName}
Converts text files {direction}.

Usage:
    dotnet .\\scripts\\{scriptName}.cs [--source <path>] [--files <file1,file2,...>]

Options:
  --source, -s   Source directory (default: .\\source)
  --files, -f    Comma or semicolon-separated file names
  --help, -h     Show this help

Examples:
    dotnet .\\scripts\\{scriptName}.cs
    dotnet .\\scripts\\{scriptName}.cs -- --files SampleModule.bas,ThisWorkbook.cls
""";
    }

    public static Encoding InitializeNativeEncoding(string baseDir)
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        var config = ReadConfig(baseDir);
        return config.GetCodepageEncoding();
    }

    public static VbaDevPackConfig ReadConfig(string baseDir)
    {
        var configPath = Path.Combine(baseDir, ConfigFileName);
        if (!File.Exists(configPath))
        {
            return new VbaDevPackConfig();
        }

        try
        {
            var json = File.ReadAllText(configPath, Utf8NoBom);
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

        File.WriteAllText(outputPath, xmlText!, Utf8NoBom);
        message = $"Exported customUI from '{partName}' to: {outputPath}";
        return true;
    }

    public static void ImportCustomUi(string workbookPath, string sourcePath)
    {
        if (!File.Exists(sourcePath))
        {
            throw new FileNotFoundException("customUI source file not found.", sourcePath);
        }

        var xmlText = File.ReadAllText(sourcePath, Utf8NoBomStrict);

        using var archive = ZipFile.Open(workbookPath, ZipArchiveMode.Update);
        var existingPart = FindCustomUiPart(archive);
        var targetPartName = existingPart?.FullName ?? "customUI/customUI.xml";

        existingPart?.Delete();

        WriteArchiveEntryText(archive, targetPartName, xmlText);
        var relsChange = EnsureCustomUiPackageReferences(archive, targetPartName);

        if (relsChange.Added)
        {
            Console.WriteLine($"Added package rels entry for customUI: _rels/.rels ({relsChange.RelationshipId} -> {relsChange.Target})");
        }
        else if (relsChange.Updated)
        {
            Console.WriteLine($"Updated package rels entry for customUI: _rels/.rels ({relsChange.RelationshipId} -> {relsChange.Target})");
        }
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

    public static string CreateTempBackupDirectory(string scenario, VbaDevPackConfig? config = null, string? baseDir = null)
    {
        var safeScenario = string.IsNullOrWhiteSpace(scenario) ? "backup" : scenario.Trim().ToLowerInvariant();
        var backupBaseDir = ResolveBackupBaseDirectory(config, baseDir ?? Environment.CurrentDirectory);
        var dir = Path.Combine(
            backupBaseDir,
            safeScenario,
            DateTime.Now.ToString("yyyyMMdd-HHmmss") + "-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        return dir;
    }

    public static string ResolveBackupBaseDirectory(VbaDevPackConfig? config, string baseDir)
    {
        var configured = config?.BackupRoot;
        if (!string.IsNullOrWhiteSpace(configured))
        {
            var expanded = Environment.ExpandEnvironmentVariables(configured);
            return Path.IsPathRooted(expanded)
                ? expanded
                : Path.GetFullPath(Path.Combine(baseDir, expanded));
        }

        return Path.Combine(Path.GetTempPath(), "vba-dev-pack-backups");
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

    public static void SafeFinalReleaseComObject(object? comObject)
    {
        if (comObject is not null && Marshal.IsComObject(comObject))
        {
            Marshal.FinalReleaseComObject(comObject);
        }
    }

    public static void ForceComCleanup()
    {
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
        GC.WaitForPendingFinalizers();
    }

    public static Exception BuildFriendlyComException(COMException ex, string action)
    {
        var hexHresult = $"0x{(uint)ex.ErrorCode:X8}";
        var msg = ex.Message ?? string.Empty;
        if (ex.ErrorCode == unchecked((int)0x800A03EC))
        {
            return new InvalidOperationException(
                "Excel denied access to VBProject. Enable: Excel -> File -> Options -> Trust Center -> Trust Center Settings -> Macro Settings -> Trust access to the VBA project object model. " +
                $"HRESULT: {hexHresult}",
                ex);
        }

        return new InvalidOperationException($"COM error while {action}. HRESULT: {hexHresult}. Details: {msg}", ex);
    }

    public static bool IsSupportedConvertibleExtension(string extension)
    {
        return ConvertibleExtensions.Contains(extension);
    }

    public static List<string> ResolveEncodingFiles(EncodingScriptOptions options)
    {
        var resolvedFiles = new List<string>();

        if (options.Files.Count == 0)
        {
            resolvedFiles.AddRange(
                Directory
                    .EnumerateFiles(options.SourceDir, "*.*", SearchOption.TopDirectoryOnly)
                    .Where(path => IsSupportedConvertibleExtension(Path.GetExtension(path)))
                    .OrderBy(path => path, StringComparer.OrdinalIgnoreCase));

            return resolvedFiles;
        }

        foreach (var rawPath in options.Files)
        {
            var fullPath = Path.IsPathRooted(rawPath)
                ? Path.GetFullPath(rawPath)
                : Path.GetFullPath(Path.Combine(options.SourceDir, rawPath));

            if (!File.Exists(fullPath))
            {
                Console.WriteLine($"Skipped, file not found: {fullPath}");
                continue;
            }

            if (!IsSupportedConvertibleExtension(Path.GetExtension(fullPath)))
            {
                Console.WriteLine($"Skipped, unsupported extension: {fullPath}");
                continue;
            }

            resolvedFiles.Add(fullPath);
        }

        return resolvedFiles;
    }

    public static bool IsValidUtf8(byte[] bytes)
    {
        try
        {
            var text = Utf8NoBomStrict.GetString(bytes);
            var roundTrip = Utf8NoBom.GetBytes(text);
            return bytes.SequenceEqual(roundTrip);
        }
        catch
        {
            return false;
        }
    }

    public static bool EnsureApprovedOrExit(bool force)
    {
        if (force)
        {
            return true;
        }

        if (Console.IsInputRedirected)
        {
            Console.Error.WriteLine("Interactive confirmation is required. Re-run with --force in non-interactive mode.");
            Environment.Exit(5);
        }

        Console.Write("Approve diff? [Y]es / [n]o (default: yes): ");
        var input = Console.ReadLine();
        if (string.IsNullOrWhiteSpace(input))
        {
            return true;
        }

        var normalized = input.Trim().ToLowerInvariant();
        return normalized is "y" or "yes";
    }

    private static bool TryGetOptionValue(Dictionary<string, string?> options, string longName, string shortName, out string? value)
    {
        if (options.TryGetValue(longName, out value))
        {
            return true;
        }

        if (options.TryGetValue(shortName, out value))
        {
            return true;
        }

        value = null;
        return false;
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
        using var reader = new StreamReader(stream, Utf8NoBomStrict);
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

    private static PackageRelsChange EnsureCustomUiPackageReferences(ZipArchive archive, string partName)
    {
        var normalizedPartName = partName.Replace('\\', '/').TrimStart('/');
        return EnsurePackageRelationship(archive, normalizedPartName);
    }

    private static PackageRelsChange EnsurePackageRelationship(ZipArchive archive, string targetPartName)
    {
        var relDoc = ReadXmlEntry(archive, PackageRelationshipsPath)
            ?? throw new InvalidOperationException($"Workbook package is missing required part: {PackageRelationshipsPath}");

        var root = relDoc.Root
            ?? throw new InvalidOperationException($"Workbook package part is invalid: {PackageRelationshipsPath}");

        var relationships = root.Elements(RelationshipNs + "Relationship").ToList();
        var uiRelationships = relationships
            .Where(x => string.Equals((string?)x.Attribute("Type"), UiExtensibilityRelationshipType, StringComparison.Ordinal))
            .ToList();

        if (uiRelationships.Count == 0)
        {
            var newId = BuildNextRelationshipId(relationships);
            root.Add(new XElement(
                RelationshipNs + "Relationship",
                new XAttribute("Id", newId),
                new XAttribute("Type", UiExtensibilityRelationshipType),
                new XAttribute("Target", targetPartName)));

            WriteXmlEntry(archive, PackageRelationshipsPath, relDoc);
            return new PackageRelsChange(true, false, newId, targetPartName);
        }

        var primary = uiRelationships[0];
        var changed = false;
        var relationshipId = (string?)primary.Attribute("Id");
        if (string.IsNullOrWhiteSpace(relationshipId))
        {
            relationshipId = BuildNextRelationshipId(relationships);
            primary.SetAttributeValue("Id", relationshipId);
            changed = true;
        }

        if (!string.Equals((string?)primary.Attribute("Type"), UiExtensibilityRelationshipType, StringComparison.Ordinal))
        {
            primary.SetAttributeValue("Type", UiExtensibilityRelationshipType);
            changed = true;
        }

        if (!string.Equals((string?)primary.Attribute("Target"), targetPartName, StringComparison.Ordinal))
        {
            primary.SetAttributeValue("Target", targetPartName);
            changed = true;
        }

        if (uiRelationships.Count > 1)
        {
            for (var i = 1; i < uiRelationships.Count; i++)
            {
                uiRelationships[i].Remove();
            }

            changed = true;
        }

        if (changed)
        {
            WriteXmlEntry(archive, PackageRelationshipsPath, relDoc);
        }

        return new PackageRelsChange(false, changed, relationshipId!, targetPartName);
    }

    private static string BuildNextRelationshipId(List<XElement> relationships)
    {
        var used = new HashSet<int>();
        foreach (var rel in relationships)
        {
            var id = (string?)rel.Attribute("Id");
            if (id is null || !id.StartsWith("rId", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (int.TryParse(id[3..], out var number) && number > 0)
            {
                used.Add(number);
            }
        }

        var next = 1;
        while (used.Contains(next))
        {
            next++;
        }

        return $"rId{next}";
    }

    private static XDocument? ReadXmlEntry(ZipArchive archive, string fullName)
    {
        var entry = archive.GetEntry(fullName);
        if (entry is null)
        {
            return null;
        }

        using var stream = entry.Open();
        using var reader = new StreamReader(stream, Utf8NoBomStrict);
        var text = reader.ReadToEnd();
        return XDocument.Parse(text, LoadOptions.PreserveWhitespace);
    }

    private static void WriteXmlEntry(ZipArchive archive, string fullName, XDocument document)
    {
        WriteArchiveEntryText(archive, fullName, document.ToString(SaveOptions.DisableFormatting));
    }

    private static void WriteArchiveEntryText(ZipArchive archive, string fullName, string text)
    {
        archive.GetEntry(fullName)?.Delete();

        var entry = archive.CreateEntry(fullName, CompressionLevel.Optimal);
        using var stream = entry.Open();
        using var writer = new StreamWriter(stream, Utf8NoBom);
        writer.Write(text);
    }

    private readonly record struct PackageRelsChange(bool Added, bool Updated, string RelationshipId, string Target);
}

sealed class EncodingScriptOptions
{
    public bool ShowHelp { get; init; }
    public bool IsValid { get; init; }
    public string? Error { get; init; }
    public string SourceDir { get; init; } = string.Empty;
    public List<string> Files { get; init; } = new();

    public static EncodingScriptOptions HelpRequested() => new() { ShowHelp = true, IsValid = true };

    public static EncodingScriptOptions FromError(string error) => new() { IsValid = false, Error = error };

    public static EncodingScriptOptions FromValues(string sourceDir, List<string> files) =>
        new()
        {
            IsValid = true,
            SourceDir = sourceDir,
            Files = files
        };
}

sealed class VbaDevPackConfig
{
    [JsonPropertyName("workbook")]
    public string Workbook { get; set; } = string.Empty;

    [JsonPropertyName("codepage")]
    public int Codepage { get; set; } = 1251;

    [JsonPropertyName("backupRoot")]
    public string? BackupRoot { get; set; }

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

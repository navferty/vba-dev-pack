#!/usr/bin/env -S dotnet --
#:sdk Microsoft.NET.Sdk
#:property TargetFramework=net10.0-windows
#:property PublishAot=false
#:property PublishTrimmed=false
#:property IsAotCompatible=false
#:property EnableTrimAnalyzer=false
#:property EnableAotAnalyzer=false
#:property BuiltInComInteropSupport=true
#:include WorkbookPackageHelpers.cs
#:include WorkbookSyncHelpers.cs
#:include VbaDiffEngine.cs

using System.Runtime.InteropServices;
using System.Text;

var config = WorkbookPackageHelpers.ReadConfig(Environment.CurrentDirectory);
var options = WorkbookSyncHelpers.ParseApplyScriptOptions(args, config, Environment.CurrentDirectory);
if (options.ShowHelp)
{
    Console.WriteLine(WorkbookSyncHelpers.BuildImportHelp(config, Environment.CurrentDirectory));
    return;
}

if (!options.IsValid)
{
    Console.Error.WriteLine(options.Error);
    Console.Error.WriteLine("Use --help to see available options.");
    Environment.Exit(1);
}

var nativeEncoding = WorkbookPackageHelpers.InitializeNativeEncoding(Environment.CurrentDirectory);
var workbookPath = options.Paths.WorkbookPath;
var sourceDir = options.Paths.SourceDir;
var customUiPath = options.Paths.CustomUiPath;

if (!File.Exists(workbookPath))
{
    Console.Error.WriteLine($"Workbook not found: {workbookPath}");
    Environment.Exit(1);
}

if (!Directory.Exists(sourceDir))
{
    Console.Error.WriteLine($"Source directory not found: {sourceDir}");
    Environment.Exit(1);
}

try
{
    WorkbookPackageHelpers.EnsureWorkbookClosed(workbookPath);
}
catch (Exception ex)
{
    Console.Error.WriteLine(ex.Message);
    Environment.Exit(1);
}

var preflightExportDir = Path.Combine(Path.GetTempPath(), "vba-import-preflight-" + Guid.NewGuid().ToString("N"));
var preflightCustomUiPath = Path.Combine(preflightExportDir, "customUI.xml");
var reportPath = Path.Combine(Path.GetTempPath(), $"vba-import-preflight-{DateTime.Now:yyyyMMdd-HHmmss}.html");

if (!RunPreflight(workbookPath, sourceDir, customUiPath, nativeEncoding, preflightExportDir, preflightCustomUiPath, reportPath, options.NoOpenReport, options.Force))
{
    Console.WriteLine("Import cancelled.");
    Environment.Exit(4);
}

var backupRoot = WorkbookPackageHelpers.CreateTempBackupDirectory("import", config, Environment.CurrentDirectory);
var workbookBackupPath = WorkbookPackageHelpers.BackupFile(
    workbookPath,
    backupRoot,
    Path.Combine("workbook", Path.GetFileName(workbookPath)));
Console.WriteLine($"Backup (workbook): {workbookBackupPath}");
Console.WriteLine($"Backup root: {backupRoot}");

if (File.Exists(customUiPath))
{
    WorkbookPackageHelpers.ImportCustomUi(workbookPath, customUiPath);
    Console.WriteLine($"Imported customUI from: {customUiPath}");
}
else
{
    Console.WriteLine($"Skipped customUI import, file not found: {customUiPath}");
}

var moduleFiles = Directory
    .EnumerateFiles(sourceDir, "*.*", SearchOption.TopDirectoryOnly)
    .Where(path => IsSupportedImportExtension(Path.GetExtension(path)))
    .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
    .ToList();

if (moduleFiles.Count == 0)
{
    Console.WriteLine($"No importable files found in: {sourceDir}");
    Environment.Exit(0);
}

object? excel = null;
object? workbooks = null;
object? workbook = null;
object? vbProject = null;
object? components = null;
var tempDir = Path.Combine(Path.GetTempPath(), "vba-import-" + Guid.NewGuid().ToString("N"));

try
{
    Directory.CreateDirectory(tempDir);

    var excelType = Type.GetTypeFromProgID("Excel.Application");
    if (excelType is null)
    {
        throw new InvalidOperationException("Excel COM type is not available. Is Microsoft Excel installed?");
    }

    excel = Activator.CreateInstance(excelType);
    ((dynamic)excel!).Visible = false;
    ((dynamic)excel!).DisplayAlerts = false;

    workbooks = ((dynamic)excel!).Workbooks;
    workbook = ((dynamic)workbooks!).Open(workbookPath, false, false);
    vbProject = ((dynamic)workbook!).VBProject;
    WorkbookPackageHelpers.EnsureVbProjectAccessible(vbProject);
    components = ((dynamic)vbProject!).VBComponents;

    foreach (var sourcePath in moduleFiles)
    {
        var ext = Path.GetExtension(sourcePath).ToLowerInvariant();
        var baseName = Path.GetFileNameWithoutExtension(sourcePath);

        if (ext == ".cls" && TryGetComponentByName((dynamic)components!, baseName, out dynamic existingByName))
        {
            try
            {
                if ((int)existingByName.Type == 100)
                {
                    UpdateDocumentModule(existingByName, sourcePath);
                    Console.WriteLine($"Updated document module: {baseName}");
                    continue;
                }
            }
            finally
            {
                Marshal.FinalReleaseComObject(existingByName);
            }
        }

        ImportAsRegularComponent((dynamic)components!, sourcePath, tempDir, nativeEncoding);
    }

    ((dynamic)workbook!).Save();
    Console.WriteLine($"Done. Imported modules from: {sourceDir}");
}
catch (COMException ex)
{
    Console.Error.WriteLine(WorkbookPackageHelpers.BuildFriendlyComException(ex, "importing VBA modules").Message);
    Environment.Exit(2);
}
catch (Exception ex)
{
    Console.Error.WriteLine(ex.Message);
    Environment.Exit(3);
}
finally
{
    WorkbookPackageHelpers.SafeFinalReleaseComObject(components);
    WorkbookPackageHelpers.SafeFinalReleaseComObject(vbProject);

    if (workbook is not null)
    {
        try
        {
            ((dynamic)workbook).Close(true);
        }
        catch
        {
            // No-op.
        }

        WorkbookPackageHelpers.SafeFinalReleaseComObject(workbook);
    }

    WorkbookPackageHelpers.SafeFinalReleaseComObject(workbooks);

    if (excel is not null)
    {
        try
        {
            ((dynamic)excel).Quit();
        }
        catch
        {
            // No-op.
        }

        WorkbookPackageHelpers.SafeFinalReleaseComObject(excel);
    }

    WorkbookPackageHelpers.ForceComCleanup();

    try
    {
        if (Directory.Exists(tempDir))
        {
            Directory.Delete(tempDir, true);
        }
    }
    catch
    {
        // No-op.
    }
}

static bool IsSupportedImportExtension(string ext)
{
    return ext.Equals(".bas", StringComparison.OrdinalIgnoreCase)
        || ext.Equals(".cls", StringComparison.OrdinalIgnoreCase)
        || ext.Equals(".frm", StringComparison.OrdinalIgnoreCase);
}

static void ImportAsRegularComponent(dynamic components, string sourcePath, string tempDir, Encoding nativeEncoding)
{
    var ext = Path.GetExtension(sourcePath).ToLowerInvariant();
    var moduleName = Path.GetFileNameWithoutExtension(sourcePath);
    var expectedType = ext switch
    {
        ".bas" => 1,
        ".cls" => 2,
        ".frm" => 3,
        _ => throw new InvalidOperationException($"Unsupported import extension: {ext}")
    };

    if (TryGetComponentByName(components, moduleName, out dynamic existing))
    {
        try
        {
            var existingType = (int)existing.Type;
            if (existingType == 100)
            {
                throw new InvalidOperationException(
                    $"Component '{moduleName}' is a document module in workbook and cannot be replaced by file '{Path.GetFileName(sourcePath)}'.");
            }

            if (existingType != expectedType)
            {
                throw new InvalidOperationException(
                    $"Component type mismatch for '{moduleName}'. Existing type={existingType}, file extension={ext}.");
            }

            components.Remove(existing);
            Console.WriteLine($"Removed existing: {moduleName}");
        }
        finally
        {
            Marshal.FinalReleaseComObject(existing);
        }
    }

    var nativePath = CreateNativeEncodingImportCopy(sourcePath, tempDir, nativeEncoding);
    dynamic imported = components.Import(nativePath);
    try
    {
        Console.WriteLine($"Imported: {Path.GetFileName(sourcePath)} as {imported.Name}");
    }
    finally
    {
        Marshal.FinalReleaseComObject(imported);
    }
}

static string CreateNativeEncodingImportCopy(string sourcePath, string tempDir, Encoding nativeEncoding)
{
    var ext = Path.GetExtension(sourcePath).ToLowerInvariant();
    var fileName = Path.GetFileName(sourcePath);
    var nativePath = Path.Combine(tempDir, fileName);

    var utf8NoBom = new UTF8Encoding(false, true);

    var text = File.ReadAllText(sourcePath, utf8NoBom);
    File.WriteAllText(nativePath, text, nativeEncoding);

    if (ext == ".frm")
    {
        var sourceFrx = Path.Combine(Path.GetDirectoryName(sourcePath)!, Path.GetFileNameWithoutExtension(sourcePath) + ".frx");
        if (File.Exists(sourceFrx))
        {
            var targetFrx = Path.Combine(tempDir, Path.GetFileName(sourceFrx));
            File.Copy(sourceFrx, targetFrx, true);
        }
    }

    return nativePath;
}

static void UpdateDocumentModule(dynamic component, string sourcePath)
{
    var utf8NoBom = new UTF8Encoding(false, true);
    var fullText = File.ReadAllText(sourcePath, utf8NoBom);
    var codeOnly = ExtractDocumentCode(fullText);

    dynamic codeModule = component.CodeModule;
    try
    {
        var lines = (int)codeModule.CountOfLines;
        if (lines > 0)
        {
            codeModule.DeleteLines(1, lines);
        }

        if (!string.IsNullOrWhiteSpace(codeOnly))
        {
            codeModule.AddFromString(codeOnly);
        }
    }
    finally
    {
        Marshal.FinalReleaseComObject(codeModule);
    }
}

static string ExtractDocumentCode(string text)
{
    var normalized = text.Replace("\r\n", "\n").Replace('\r', '\n');
    var lines = normalized.Split('\n');

    var lastAttributeIndex = -1;
    for (var i = 0; i < lines.Length; i++)
    {
        var trimmed = lines[i].TrimStart();
        if (trimmed.StartsWith("Attribute ", StringComparison.OrdinalIgnoreCase))
        {
            lastAttributeIndex = i;
        }
    }

    if (lastAttributeIndex >= 0 && lastAttributeIndex + 1 < lines.Length)
    {
        return string.Join(Environment.NewLine, lines[(lastAttributeIndex + 1)..]).TrimStart('\r', '\n');
    }

    return normalized;
}

static bool TryGetComponentByName(dynamic components, string name, out dynamic component)
{
    component = null!;
    try
    {
        component = components.Item(name);
        return true;
    }
    catch
    {
        return false;
    }
}

static bool RunPreflight(
    string workbookPath,
    string sourceDir,
    string customUiPath,
    Encoding nativeEncoding,
    string tempExportDir,
    string tempCustomUiPath,
    string reportPath,
    bool noOpenReport,
    bool force)
{
    object? excel = null;
    object? workbooks = null;
    object? workbook = null;

    try
    {
        Directory.CreateDirectory(tempExportDir);

        _ = WorkbookPackageHelpers.TryExportCustomUi(workbookPath, tempCustomUiPath, out _);

        var excelType = Type.GetTypeFromProgID("Excel.Application");
        if (excelType is null)
        {
            throw new InvalidOperationException("Excel COM type is not available. Is Microsoft Excel installed?");
        }

        excel = Activator.CreateInstance(excelType);
        ((dynamic)excel!).Visible = false;
        ((dynamic)excel!).DisplayAlerts = false;

        workbooks = ((dynamic)excel!).Workbooks;
        workbook = ((dynamic)workbooks!).Open(workbookPath, false, true);
        WorkbookPackageHelpers.EnsureVbProjectAccessible(((dynamic)workbook!).VBProject);

        VbaDiffEngine.ExportWorkbookModulesToTemp((dynamic)workbook!, tempExportDir, nativeEncoding);

        var compare = VbaDiffEngine.BuildComparison(sourceDir, tempExportDir, customUiPath, tempCustomUiPath, DiffDirection.Import);
        var html = VbaDiffEngine.BuildHtmlReport(compare, workbookPath, sourceDir, tempExportDir, DiffDirection.Import);
        File.WriteAllText(reportPath, html, new UTF8Encoding(false));

        if (!noOpenReport)
        {
            _ = VbaDiffEngine.TryOpenReport(reportPath);
        }

        Console.WriteLine($"Preflight report path: {reportPath}");
        Console.WriteLine(VbaDiffEngine.BuildConsoleSummary(compare, DiffDirection.Import));

        return WorkbookPackageHelpers.EnsureApprovedOrExit(force);
    }
    catch (COMException ex)
    {
        Console.Error.WriteLine(WorkbookPackageHelpers.BuildFriendlyComException(ex, "preflight import diff").Message);
        Environment.Exit(2);
        return false;
    }
    catch (Exception ex)
    {
        Console.Error.WriteLine(ex.Message);
        Environment.Exit(3);
        return false;
    }
    finally
    {
        if (workbook is not null)
        {
            try
            {
                ((dynamic)workbook).Close(false);
            }
            catch
            {
                // No-op.
            }

            WorkbookPackageHelpers.SafeFinalReleaseComObject(workbook);
        }

        WorkbookPackageHelpers.SafeFinalReleaseComObject(workbooks);

        if (excel is not null)
        {
            try
            {
                ((dynamic)excel).Quit();
            }
            catch
            {
                // No-op.
            }

            WorkbookPackageHelpers.SafeFinalReleaseComObject(excel);
        }

        WorkbookPackageHelpers.ForceComCleanup();

        try
        {
            if (Directory.Exists(tempExportDir))
            {
                Directory.Delete(tempExportDir, true);
            }
        }
        catch
        {
            // No-op.
        }
    }
}

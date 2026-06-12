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

using System.Runtime.InteropServices;
using System.Text;

var config = WorkbookPackageHelpers.ReadConfig(Environment.CurrentDirectory);
var parsedPaths = WorkbookPackageHelpers.ParseScriptPaths(args, config, Environment.CurrentDirectory);
if (parsedPaths.ShowHelp)
{
    Console.WriteLine(WorkbookPackageHelpers.BuildCommonPathsHelp(
        "ExportVbaModules",
        "Exports workbook VBA modules and custom UI to repository files.",
        config,
        Environment.CurrentDirectory));
    return;
}

if (!parsedPaths.IsValid)
{
    Console.Error.WriteLine(parsedPaths.Error);
    Console.Error.WriteLine("Use --help to see available options.");
    Environment.Exit(1);
}

var nativeEncoding = WorkbookPackageHelpers.InitializeNativeEncoding(Environment.CurrentDirectory);
var workbookPath = parsedPaths.WorkbookPath;
var sourceDir = parsedPaths.SourceDir;
var customUiPath = parsedPaths.CustomUiPath;

if (!File.Exists(workbookPath))
{
    Console.Error.WriteLine($"Workbook not found: {workbookPath}");
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

Directory.CreateDirectory(sourceDir);

var backupRoot = WorkbookPackageHelpers.CreateTempBackupDirectory("export", config, Environment.CurrentDirectory);
var workbookBackupPath = WorkbookPackageHelpers.BackupFile(
    workbookPath,
    backupRoot,
    Path.Combine("workbook", Path.GetFileName(workbookPath)));
Console.WriteLine($"Backup (workbook): {workbookBackupPath}");

var sourceBackupPath = WorkbookPackageHelpers.BackupDirectory(sourceDir, backupRoot, "source");
Console.WriteLine($"Backup (source): {sourceBackupPath}");

if (File.Exists(customUiPath))
{
    var customUiBackupPath = WorkbookPackageHelpers.BackupFile(
        customUiPath,
        backupRoot,
        Path.Combine("customUI", Path.GetFileName(customUiPath)));
    Console.WriteLine($"Backup (customUI): {customUiBackupPath}");
}

Console.WriteLine($"Backup root: {backupRoot}");

if (WorkbookPackageHelpers.TryExportCustomUi(workbookPath, customUiPath, out var customUiMessage))
{
    Console.WriteLine(customUiMessage);
}
else
{
    Console.WriteLine("Skipped customUI export: customUI part was not found in workbook.");
}

object? excel = null;
object? workbooks = null;
object? workbook = null;
object? vbProject = null;
object? components = null;

try
{
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
    vbProject = ((dynamic)workbook!).VBProject;
    WorkbookPackageHelpers.EnsureVbProjectAccessible(vbProject);
    components = ((dynamic)vbProject!).VBComponents;

    var extsToClean = new[] { "*.bas", "*.cls", "*.frm", "*.frx" };
    foreach (var pattern in extsToClean)
    {
        foreach (var file in Directory.EnumerateFiles(sourceDir, pattern, SearchOption.TopDirectoryOnly))
        {
            File.Delete(file);
        }
    }

    var exportedTextFiles = new List<string>();

    for (var i = 1; i <= (int)((dynamic)components!).Count; i++)
    {
        dynamic component = ((dynamic)components!).Item(i);
        try
        {
            var compName = (string)component.Name;
            var extension = GetExtension((int)component.Type);
            if (extension is null)
            {
                continue;
            }

            var exportPath = Path.Combine(sourceDir, compName + extension);
            if (File.Exists(exportPath))
            {
                File.Delete(exportPath);
            }

            component.Export(exportPath);

            if (extension is ".bas" or ".cls" or ".frm")
            {
                exportedTextFiles.Add(exportPath);
            }

            Console.WriteLine($"Exported: {Path.GetFileName(exportPath)}");
        }
        finally
        {
            Marshal.FinalReleaseComObject(component);
        }
    }

    var utf8NoBom = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);

    foreach (var filePath in exportedTextFiles)
    {
        var bytes = File.ReadAllBytes(filePath);
        var text = nativeEncoding.GetString(bytes);
        File.WriteAllText(filePath, text, utf8NoBom);
        Console.WriteLine($"Converted to UTF-8: {Path.GetFileName(filePath)}");
    }

    Console.WriteLine($"Done. Exported modules to: {sourceDir}");
}
catch (COMException ex)
{
    Console.Error.WriteLine(WorkbookPackageHelpers.BuildFriendlyComException(ex, "exporting VBA modules").Message);
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
}

static string? GetExtension(int componentType)
{
    return componentType switch
    {
        1 => ".bas", // vbext_ct_StdModule
        2 => ".cls", // vbext_ct_ClassModule
        3 => ".frm", // vbext_ct_MSForm
        100 => ".cls", // vbext_ct_Document
        _ => null
    };
}

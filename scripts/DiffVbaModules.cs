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
#:include VbaDiffEngine.cs

using System.Runtime.InteropServices;
using System.Text;

var config = WorkbookPackageHelpers.ReadConfig(Environment.CurrentDirectory);
var options = WorkbookPackageHelpers.ParseDiffScriptOptions(args, config, Environment.CurrentDirectory);
if (options.ShowHelp)
{
    Console.WriteLine(WorkbookPackageHelpers.BuildDiffHelp(config, Environment.CurrentDirectory));
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

var tempExportDir = Path.Combine(Path.GetTempPath(), "vba-diff-export-" + Guid.NewGuid().ToString("N"));
var reportPath = Path.Combine(Path.GetTempPath(), $"vba-diff-report-{DateTime.Now:yyyyMMdd-HHmmss}.html");
var tempCustomUiPath = Path.Combine(tempExportDir, "customUI.xml");

object? excel = null;
object? workbooks = null;
object? workbook = null;

try
{
    Directory.CreateDirectory(tempExportDir);

    if (WorkbookPackageHelpers.TryExportCustomUi(workbookPath, tempCustomUiPath, out var customUiMessage))
    {
        Console.WriteLine(customUiMessage);
    }
    else
    {
        Console.WriteLine("Skipped customUI export for diff: customUI part was not found in workbook.");
    }

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

    var compare = VbaDiffEngine.BuildComparison(sourceDir, tempExportDir, customUiPath, tempCustomUiPath, options.Direction);
    var html = VbaDiffEngine.BuildHtmlReport(compare, workbookPath, sourceDir, tempExportDir, options.Direction);
    File.WriteAllText(reportPath, html, new UTF8Encoding(false));

    if (!options.NoOpenReport)
    {
        _ = VbaDiffEngine.TryOpenReport(reportPath);
    }

    Console.WriteLine($"Report path: {reportPath}");
    Console.WriteLine(VbaDiffEngine.BuildConsoleSummary(compare, options.Direction));
}
catch (COMException ex)
{
    Console.Error.WriteLine(WorkbookPackageHelpers.BuildFriendlyComException(ex, "building diff report").Message);
    Environment.Exit(2);
}
catch (Exception ex)
{
    Console.Error.WriteLine(ex.Message);
    Environment.Exit(3);
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

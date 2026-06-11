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

Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

var workbookPath = args.Length > 0
    ? Path.GetFullPath(args[0])
    : Path.GetFullPath(Path.Combine(Environment.CurrentDirectory, "Sample.xlsm"));

var sourceDir = args.Length > 1
    ? Path.GetFullPath(args[1])
    : Path.GetFullPath(Path.Combine(Environment.CurrentDirectory, "source"));

var customUiPath = args.Length > 2
    ? Path.GetFullPath(args[2])
    : Path.GetFullPath(Path.Combine(Environment.CurrentDirectory, "customUI", "customUI.xml"));

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

var backupRoot = WorkbookPackageHelpers.CreateTempBackupDirectory("export");
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

    var cp1251 = Encoding.GetEncoding(1251);
    var utf8NoBom = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);

    foreach (var filePath in exportedTextFiles)
    {
        var bytes = File.ReadAllBytes(filePath);
        var text = cp1251.GetString(bytes);
        File.WriteAllText(filePath, text, utf8NoBom);
        Console.WriteLine($"Converted to UTF-8: {Path.GetFileName(filePath)}");
    }

    Console.WriteLine($"Done. Exported modules to: {sourceDir}");
}
catch (COMException ex)
{
    Console.Error.WriteLine(BuildFriendlyComException(ex, "exporting VBA modules").Message);
    Environment.Exit(2);
}
catch (Exception ex)
{
    Console.Error.WriteLine(ex.Message);
    Environment.Exit(3);
}
finally
{
    SafeFinalReleaseComObject(components);
    SafeFinalReleaseComObject(vbProject);

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

        SafeFinalReleaseComObject(workbook);
    }

    SafeFinalReleaseComObject(workbooks);

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

        SafeFinalReleaseComObject(excel);
    }

    ForceComCleanup();
}

static void SafeFinalReleaseComObject(object? comObject)
{
    if (comObject is not null && Marshal.IsComObject(comObject))
    {
        Marshal.FinalReleaseComObject(comObject);
    }
}

static void ForceComCleanup()
{
    GC.Collect();
    GC.WaitForPendingFinalizers();
    GC.Collect();
    GC.WaitForPendingFinalizers();
}

static Exception BuildFriendlyComException(COMException ex, string action)
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

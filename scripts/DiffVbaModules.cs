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

using System.Diagnostics;
using System.Net;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;

var config = WorkbookPackageHelpers.ReadConfig(Environment.CurrentDirectory);
var parsedPaths = WorkbookPackageHelpers.ParseScriptPaths(args, config, Environment.CurrentDirectory);
if (parsedPaths.ShowHelp)
{
    Console.WriteLine(WorkbookPackageHelpers.BuildCommonPathsHelp(
        "DiffVbaModules",
        "Builds HTML diff report comparing workbook export and repository source.",
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

    ExportVbaComponentsToTemp((dynamic)workbook!, tempExportDir, nativeEncoding);

    var compare = BuildComparison(sourceDir, tempExportDir, customUiPath, tempCustomUiPath);
    var html = BuildHtmlReport(compare, workbookPath, sourceDir, tempExportDir);
    File.WriteAllText(reportPath, html, new UTF8Encoding(false));

    Process.Start(new ProcessStartInfo
    {
        FileName = reportPath,
        UseShellExecute = true
    });

    Console.WriteLine($"Report created: {reportPath}");
    Console.WriteLine($"Summary: changed={compare.Changed.Count}, missingInWorkbook={compare.OnlyInSource.Count}, missingInSource={compare.OnlyInExport.Count}, equal={compare.Equal.Count}");
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
}

static void ExportVbaComponentsToTemp(dynamic workbook, string exportDir, Encoding nativeEncoding)
{
    dynamic vbProject = workbook.VBProject;
    dynamic components = vbProject.VBComponents;
    try
    {
        var exportedText = new List<string>();

        for (var i = 1; i <= (int)components.Count; i++)
        {
            dynamic component = components.Item(i);
            try
            {
                var type = (int)component.Type;
                var ext = type switch
                {
                    1 => ".bas",
                    2 => ".cls",
                    3 => ".frm",
                    100 => ".cls",
                    _ => null
                };

                if (ext is null)
                {
                    continue;
                }

                var outputPath = Path.Combine(exportDir, (string)component.Name + ext);
                if (File.Exists(outputPath))
                {
                    File.Delete(outputPath);
                }

                component.Export(outputPath);
                Console.WriteLine($"Exported for diff: {Path.GetFileName(outputPath)}");

                if (ext is ".bas" or ".cls" or ".frm")
                {
                    exportedText.Add(outputPath);
                }
            }
            finally
            {
                Marshal.FinalReleaseComObject(component);
            }
        }

        var utf8NoBom = new UTF8Encoding(false);

        foreach (var path in exportedText)
        {
            var bytes = File.ReadAllBytes(path);
            var text = nativeEncoding.GetString(bytes);
            File.WriteAllText(path, text, utf8NoBom);
        }
    }
    finally
    {
        WorkbookPackageHelpers.SafeFinalReleaseComObject(components);
        WorkbookPackageHelpers.SafeFinalReleaseComObject(vbProject);
    }
}

static CompareResult BuildComparison(string sourceDir, string exportDir, string sourceCustomUiPath, string exportCustomUiPath)
{
    var allowed = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
    {
        ".bas", ".cls", ".frm", ".frx"
    };

    var sourceFiles = Directory
        .EnumerateFiles(sourceDir, "*.*", SearchOption.TopDirectoryOnly)
        .Where(path => allowed.Contains(Path.GetExtension(path)))
        .ToDictionary(Path.GetFileName, StringComparer.OrdinalIgnoreCase);

    var exportFiles = Directory
        .EnumerateFiles(exportDir, "*.*", SearchOption.TopDirectoryOnly)
        .Where(path => allowed.Contains(Path.GetExtension(path)))
        .ToDictionary(Path.GetFileName, StringComparer.OrdinalIgnoreCase);

    var allNames = sourceFiles.Keys
        .Union(exportFiles.Keys, StringComparer.OrdinalIgnoreCase)
        .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
        .ToList();

    var result = new CompareResult();

    foreach (var name in allNames)
    {
        var inSource = sourceFiles.TryGetValue(name, out var sourcePath);
        var inExport = exportFiles.TryGetValue(name, out var exportPath);

        if (inSource && !inExport)
        {
            result.OnlyInSource.Add(name);
            continue;
        }

        if (!inSource && inExport)
        {
            result.OnlyInExport.Add(name);
            continue;
        }

        if (!inSource || !inExport)
        {
            continue;
        }

        var ext = Path.GetExtension(name).ToLowerInvariant();
        if (ext == ".frx")
        {
            var sourceHash = Sha256(sourcePath!);
            var exportHash = Sha256(exportPath!);
            if (!sourceHash.Equals(exportHash, StringComparison.OrdinalIgnoreCase))
            {
                result.Changed.Add(new DiffItem
                {
                    FileName = name,
                    IsBinary = true,
                    BinaryHashSource = sourceHash,
                    BinaryHashExport = exportHash
                });
            }
            else
            {
                result.Equal.Add(name);
            }

            continue;
        }

        var sourceText = File.ReadAllText(sourcePath!, new UTF8Encoding(false, true));
        var exportText = File.ReadAllText(exportPath!, new UTF8Encoding(false, true));

        if (NormalizeLineEndings(sourceText) == NormalizeLineEndings(exportText))
        {
            result.Equal.Add(name);
            continue;
        }

        result.Changed.Add(new DiffItem
        {
            FileName = name,
            IsBinary = false,
            UnifiedDiff = BuildUnifiedDiff(name, sourceText, exportText)
        });
    }

    var hasSourceCustomUi = File.Exists(sourceCustomUiPath);
    var hasExportCustomUi = File.Exists(exportCustomUiPath);
    const string customUiLogicalName = "customUI/customUI.xml";

    if (hasSourceCustomUi && hasExportCustomUi)
    {
        var sourceText = File.ReadAllText(sourceCustomUiPath, new UTF8Encoding(false, true));
        var exportText = File.ReadAllText(exportCustomUiPath, new UTF8Encoding(false, true));

        if (NormalizeLineEndings(sourceText) == NormalizeLineEndings(exportText))
        {
            result.Equal.Add(customUiLogicalName);
        }
        else
        {
            result.Changed.Add(new DiffItem
            {
                FileName = customUiLogicalName,
                IsBinary = false,
                UnifiedDiff = BuildUnifiedDiff(customUiLogicalName, sourceText, exportText)
            });
        }
    }
    else if (hasSourceCustomUi)
    {
        result.OnlyInSource.Add(customUiLogicalName);
    }
    else if (hasExportCustomUi)
    {
        result.OnlyInExport.Add(customUiLogicalName);
    }

    return result;
}

static string BuildHtmlReport(CompareResult compare, string workbookPath, string sourceDir, string exportDir)
{
    var sb = new StringBuilder();

    sb.AppendLine("<!doctype html>");
    sb.AppendLine("<html><head><meta charset=\"utf-8\" />");
    sb.AppendLine("<title>VBA Diff Report</title>");
    sb.AppendLine("<link rel=\"stylesheet\" href=\"https://cdn.jsdelivr.net/npm/diff2html/bundles/css/diff2html.min.css\" />");
    sb.AppendLine("<style>");
    sb.AppendLine("body{font-family:Segoe UI,Arial,sans-serif;margin:24px;background:#f7f9fc;color:#1e293b}");
    sb.AppendLine("h1,h2{margin:0 0 12px 0}");
    sb.AppendLine(".meta{margin-bottom:18px;padding:12px;background:#fff;border:1px solid #dbe4f0;border-radius:8px}");
    sb.AppendLine(".grid{display:grid;grid-template-columns:repeat(4,minmax(120px,1fr));gap:10px;margin:12px 0 20px 0}");
    sb.AppendLine(".card{background:#fff;border:1px solid #dbe4f0;border-radius:8px;padding:10px}");
    sb.AppendLine(".ok{color:#0f766e}.warn{color:#b45309}.bad{color:#b91c1c}");
    sb.AppendLine("details{margin:10px 0;background:#fff;border:1px solid #dbe4f0;border-radius:8px}");
    sb.AppendLine("summary{cursor:pointer;padding:10px 12px;font-weight:600;list-style:none;display:flex;align-items:center;gap:8px}");
    sb.AppendLine("summary::-webkit-details-marker{display:none}");
    sb.AppendLine("summary::before{content:'▶';font-size:.75em;transition:transform .15s}");
    sb.AppendLine("details[open]>summary::before{transform:rotate(90deg)}");
    sb.AppendLine(".binary-diff{margin:0;padding:12px;font-family:monospace;font-size:13px;background:#fff8ee;border-top:1px solid #dbe4f0}");
    sb.AppendLine("ul{margin:8px 0 0 20px}");
    sb.AppendLine("li{margin:4px 0}");
    sb.AppendLine(".d2h-wrapper{border-top:1px solid #dbe4f0}");
    sb.AppendLine("</style></head><body>");

    sb.AppendLine("<h1>VBA Diff Report</h1>");
    sb.AppendLine("<div class=\"meta\">");
    sb.AppendLine($"<div><b>Workbook:</b> {Html(workbookPath)}</div>");
    sb.AppendLine($"<div><b>Source:</b> {Html(sourceDir)}</div>");
    sb.AppendLine($"<div><b>Export Temp:</b> {Html(exportDir)}</div>");
    sb.AppendLine($"<div><b>Generated:</b> {Html(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))}</div>");
    sb.AppendLine("</div>");

    sb.AppendLine("<div class=\"grid\">");
    sb.AppendLine($"<div class=\"card\"><div>Changed</div><div class=\"bad\"><b>{compare.Changed.Count}</b></div></div>");
    sb.AppendLine($"<div class=\"card\"><div>Only In Source</div><div class=\"warn\"><b>{compare.OnlyInSource.Count}</b></div></div>");
    sb.AppendLine($"<div class=\"card\"><div>Only In Export</div><div class=\"warn\"><b>{compare.OnlyInExport.Count}</b></div></div>");
    sb.AppendLine($"<div class=\"card\"><div>Equal</div><div class=\"ok\"><b>{compare.Equal.Count}</b></div></div>");
    sb.AppendLine("</div>");

    AppendSimpleList(sb, "Only In Source", compare.OnlyInSource);
    AppendSimpleList(sb, "Only In Export", compare.OnlyInExport);
    AppendSimpleList(sb, "Equal", compare.Equal);

    sb.AppendLine("<h2>Changed Files</h2>");
    if (compare.Changed.Count == 0)
    {
        sb.AppendLine("<p class=\"ok\">No changed files.</p>");
    }
    else
    {
        sb.AppendLine("<div style=\"margin-bottom:12px\">");
        sb.AppendLine("  <button id=\"btn-side\" onclick=\"renderAll('side-by-side')\" style=\"margin-right:6px\">Side-by-side</button>");
        sb.AppendLine("  <button id=\"btn-inline\" onclick=\"renderAll('line-by-line')\">Inline</button>");
        sb.AppendLine("</div>");

        var diffData = new StringBuilder();
        diffData.AppendLine("const diffs = [");
        var idx = 0;
        foreach (var item in compare.Changed)
        {
            sb.AppendLine("<details open>");
            sb.AppendLine($"<summary>{Html(item.FileName)}</summary>");
            if (item.IsBinary)
            {
                sb.AppendLine($"<pre class=\"binary-diff\">Binary file differs (SHA256).\nsource:  {Html(item.BinaryHashSource ?? string.Empty)}\nexport:  {Html(item.BinaryHashExport ?? string.Empty)}</pre>");
                diffData.AppendLine("  null,");
            }
            else
            {
                var divId = $"diff{idx}";
                sb.AppendLine($"<div class=\"d2h-wrapper\" id=\"{divId}\"></div>");
                var jsonDiff = System.Text.Json.JsonSerializer.Serialize(item.UnifiedDiff ?? string.Empty);
                diffData.AppendLine($"  {{id:'{divId}',text:{jsonDiff}}},");
                idx++;
            }

            sb.AppendLine("</details>");
        }

        diffData.AppendLine("];");

        sb.AppendLine("<script src=\"https://cdn.jsdelivr.net/npm/diff2html/bundles/js/diff2html-ui.min.js\"></script>");
        sb.AppendLine("<script>");
        sb.Append(diffData);
        sb.AppendLine("function renderAll(fmt) {");
        sb.AppendLine("  document.getElementById('btn-side').style.fontWeight = fmt==='side-by-side'?'bold':'normal';");
        sb.AppendLine("  document.getElementById('btn-inline').style.fontWeight = fmt==='line-by-line'?'bold':'normal';");
        sb.AppendLine("  diffs.forEach(function(d){if(!d)return;new Diff2HtmlUI(document.getElementById(d.id),d.text,{outputFormat:fmt,matching:'lines',drawFileList:false}).draw();});");
        sb.AppendLine("}");
        sb.AppendLine("renderAll('side-by-side');");
        sb.AppendLine("</script>");
    }

    sb.AppendLine("</body></html>");
    return sb.ToString();
}

static void AppendSimpleList(StringBuilder sb, string title, List<string> items)
{
    sb.AppendLine("<details open>");
    sb.AppendLine($"<summary>{Html(title)} ({items.Count})</summary>");
    if (items.Count == 0)
    {
        sb.AppendLine("<pre>(none)</pre>");
    }
    else
    {
        sb.AppendLine("<ul>");
        foreach (var item in items)
        {
            sb.AppendLine($"<li>{Html(item)}</li>");
        }

        sb.AppendLine("</ul>");
    }

    sb.AppendLine("</details>");
}

static string BuildUnifiedDiff(string fileName, string sourceText, string exportText)
{
    const int context = 3;
    var a = SplitLines(NormalizeLineEndings(sourceText));
    var b = SplitLines(NormalizeLineEndings(exportText));
    var n = a.Length;
    var m = b.Length;

    // LCS table
    var lcs = new int[n + 1, m + 1];
    for (var i = n - 1; i >= 0; i--)
        for (var j = m - 1; j >= 0; j--)
            lcs[i, j] = a[i] == b[j] ? lcs[i + 1, j + 1] + 1 : Math.Max(lcs[i + 1, j], lcs[i, j + 1]);

    // Build flat edit operations
    var ops = new List<(char op, string line, int oldLine, int newLine)>();
    var x = 0;
    var y = 0;
    var oldNum = 1;
    var newNum = 1;
    while (x < n || y < m)
    {
        if (x < n && y < m && a[x] == b[y])
        {
            ops.Add((' ', a[x], oldNum++, newNum++));
            x++; y++;
        }
        else if (y < m && (x >= n || lcs[x, y + 1] >= lcs[x + 1, y]))
        {
            ops.Add(('+', b[y], 0, newNum++));
            y++;
        }
        else
        {
            ops.Add(('-', a[x], oldNum++, 0));
            x++;
        }
    }

    // Find changed positions
    var changedAt = new HashSet<int>();
    for (var i = 0; i < ops.Count; i++)
        if (ops[i].op != ' ') changedAt.Add(i);

    // Group into hunks with context
    var hunks = new List<(int start, int end)>();
    var hunkStart = -1;
    var hunkEnd = -1;
    foreach (var ci in changedAt.Order())
    {
        var s = Math.Max(0, ci - context);
        var e = Math.Min(ops.Count - 1, ci + context);
        if (hunkStart < 0)
        {
            hunkStart = s;
            hunkEnd = e;
        }
        else if (s <= hunkEnd + 1)
        {
            hunkEnd = Math.Max(hunkEnd, e);
        }
        else
        {
            hunks.Add((hunkStart, hunkEnd));
            hunkStart = s;
            hunkEnd = e;
        }
    }
    if (hunkStart >= 0) hunks.Add((hunkStart, hunkEnd));

    var sb = new StringBuilder();
    sb.AppendLine($"--- a/{fileName}");
    sb.AppendLine($"+++ b/{fileName}");

    foreach (var (hs, he) in hunks)
    {
        var hunkOps = ops.Skip(hs).Take(he - hs + 1).ToList();
        var oldStart = hunkOps.FirstOrDefault(o => o.oldLine > 0).oldLine;
        var newStart = hunkOps.FirstOrDefault(o => o.newLine > 0).newLine;
        if (oldStart == 0) oldStart = 1;
        if (newStart == 0) newStart = 1;
        var oldCount = hunkOps.Count(o => o.op != '+');
        var newCount = hunkOps.Count(o => o.op != '-');
        sb.AppendLine($"@@ -{oldStart},{oldCount} +{newStart},{newCount} @@");
        foreach (var (op, line, _, _) in hunkOps)
            sb.AppendLine(op + line);
    }

    return sb.ToString();
}

static string[] SplitLines(string text)
{
    return text.Split('\n');
}

static string NormalizeLineEndings(string text)
{
    return text.Replace("\r\n", "\n").Replace('\r', '\n');
}

static string Sha256(string path)
{
    using var sha = SHA256.Create();
    using var stream = File.OpenRead(path);
    var hash = sha.ComputeHash(stream);
    return Convert.ToHexString(hash);
}

static string Html(string value)
{
    return WebUtility.HtmlEncode(value);
}

sealed class CompareResult
{
    public List<string> OnlyInSource { get; } = new();
    public List<string> OnlyInExport { get; } = new();
    public List<string> Equal { get; } = new();
    public List<DiffItem> Changed { get; } = new();
}

sealed class DiffItem
{
    public string FileName { get; set; } = string.Empty;
    public bool IsBinary { get; set; }
    public string? UnifiedDiff { get; set; }
    public string? BinaryHashSource { get; set; }
    public string? BinaryHashExport { get; set; }
}

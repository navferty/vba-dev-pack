#!/usr/bin/env -S dotnet --
#:sdk Microsoft.NET.Sdk
#:property TargetFramework=net10.0
#:include WorkbookPackageHelpers.cs

using System.Text;

var options = WorkbookPackageHelpers.ParseEncodingScriptOptions(args);
if (options.ShowHelp)
{
    Console.WriteLine(WorkbookPackageHelpers.BuildEncodingHelp(
        "FromUtf",
        "from UTF-8 (no BOM) to the configured native codepage"));
    return;
}

if (!options.IsValid)
{
    Console.Error.WriteLine(options.Error);
    Console.Error.WriteLine("Use --help to see available options.");
    Environment.Exit(1);
}

if (!Directory.Exists(options.SourceDir))
{
    Console.Error.WriteLine($"Source directory not found: {options.SourceDir}");
    Environment.Exit(1);
}

var nativeEncoding = WorkbookPackageHelpers.InitializeNativeEncoding(Environment.CurrentDirectory);
var utf8NoBomStrict = new UTF8Encoding(false, true);
var filesToProcess = ResolveFiles(options);

foreach (var filePath in filesToProcess)
{
    var bytes = File.ReadAllBytes(filePath);
    if (!WorkbookPackageHelpers.IsValidUtf8(bytes))
    {
        Console.WriteLine($"Skipped, already native encoding or not valid UTF-8: {filePath}");
        continue;
    }

    var text = File.ReadAllText(filePath, utf8NoBomStrict);
    File.WriteAllText(filePath, text, nativeEncoding);
    Console.WriteLine($"Converted from UTF-8: {filePath}");
}

static List<string> ResolveFiles(EncodingScriptOptions options)
{
    var resolvedFiles = new List<string>();

    if (options.Files.Count == 0)
    {
        resolvedFiles.AddRange(
            Directory
                .EnumerateFiles(options.SourceDir, "*.*", SearchOption.TopDirectoryOnly)
                .Where(path => WorkbookPackageHelpers.IsSupportedConvertibleExtension(Path.GetExtension(path)))
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

        if (!WorkbookPackageHelpers.IsSupportedConvertibleExtension(Path.GetExtension(fullPath)))
        {
            Console.WriteLine($"Skipped, unsupported extension: {fullPath}");
            continue;
        }

        resolvedFiles.Add(fullPath);
    }

    return resolvedFiles;
}

#!/usr/bin/env -S dotnet --
#:sdk Microsoft.NET.Sdk
#:property TargetFramework=net10.0
#:include WorkbookPackageHelpers.cs

using System.Text;

var options = WorkbookPackageHelpers.ParseEncodingScriptOptions(args);
if (options.ShowHelp)
{
    Console.WriteLine(WorkbookPackageHelpers.BuildEncodingHelp(
        "ToUtf",
        "from the configured native codepage to UTF-8 (no BOM)"));
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
var utf8NoBom = new UTF8Encoding(false);
var filesToProcess = WorkbookPackageHelpers.ResolveEncodingFiles(options);

foreach (var filePath in filesToProcess)
{
    var bytes = File.ReadAllBytes(filePath);
    if (WorkbookPackageHelpers.IsValidUtf8(bytes))
    {
        Console.WriteLine($"Skipped, already UTF-8: {filePath}");
        continue;
    }

    var text = nativeEncoding.GetString(bytes);
    File.WriteAllText(filePath, text, utf8NoBom);
    Console.WriteLine($"Converted to UTF-8: {filePath}");
}

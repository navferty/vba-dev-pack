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
var filesToProcess = WorkbookPackageHelpers.ResolveEncodingFiles(options);

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

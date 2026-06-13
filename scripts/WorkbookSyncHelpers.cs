static class WorkbookSyncHelpers
{
    public static ScriptPaths ParseScriptPaths(string[] args, VbaDevPackConfig config, string baseDir)
    {
        var options = WorkbookPackageHelpers.ParseNamedArguments(args);
        return ParseScriptPaths(options, config, baseDir);
    }

    public static DiffScriptOptions ParseDiffScriptOptions(string[] args, VbaDevPackConfig config, string baseDir)
    {
        var options = WorkbookPackageHelpers.ParseNamedArguments(args, new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "help", "h", "no-open-report"
        });

        if (options.TryGetValue("error", out var parseError))
        {
            return DiffScriptOptions.FromError(parseError ?? "Invalid arguments.");
        }

        var showHelp = options.ContainsKey("help") || options.ContainsKey("h");
        if (showHelp)
        {
            return DiffScriptOptions.HelpRequested();
        }

        var paths = ParseScriptPaths(options, config, baseDir);
        if (!paths.IsValid)
        {
            return DiffScriptOptions.FromError(paths.Error ?? "Invalid path arguments.");
        }

        var direction = DiffDirection.Neutral;
        if (TryGetOptionValue(options, "direction", "d", out var directionValue) && !string.IsNullOrWhiteSpace(directionValue))
        {
            if (directionValue.Equals("export", StringComparison.OrdinalIgnoreCase))
            {
                direction = DiffDirection.Export;
            }
            else if (directionValue.Equals("import", StringComparison.OrdinalIgnoreCase))
            {
                direction = DiffDirection.Import;
            }
            else
            {
                return DiffScriptOptions.FromError("Option '--direction' must be either 'export' or 'import'.");
            }
        }

        var noOpenReport = options.ContainsKey("no-open-report");
        return DiffScriptOptions.FromValues(paths, direction, noOpenReport);
    }

    public static ApplyScriptOptions ParseApplyScriptOptions(string[] args, VbaDevPackConfig config, string baseDir)
    {
        var options = WorkbookPackageHelpers.ParseNamedArguments(args, new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "help", "h", "force", "no-open-report"
        });

        if (options.TryGetValue("error", out var parseError))
        {
            return ApplyScriptOptions.FromError(parseError ?? "Invalid arguments.");
        }

        var showHelp = options.ContainsKey("help") || options.ContainsKey("h");
        if (showHelp)
        {
            return ApplyScriptOptions.HelpRequested();
        }

        var paths = ParseScriptPaths(options, config, baseDir);
        if (!paths.IsValid)
        {
            return ApplyScriptOptions.FromError(paths.Error ?? "Invalid path arguments.");
        }

        var force = options.ContainsKey("force");
        var noOpenReport = options.ContainsKey("no-open-report");
        return ApplyScriptOptions.FromValues(paths, force, noOpenReport);
    }

    public static string BuildCommonPathsHelp(string scriptName, string description, VbaDevPackConfig config, string baseDir)
    {
        var workbookDefault = ResolveDefaultWorkbookPath(config, baseDir) ?? "(set in vba-dev-pack.json: workbook)";
        return $"""
{scriptName}
{description}

Usage:
    dotnet .\\scripts\\{scriptName}.cs -- --workbook <path> --source <path> --custom-ui <path>

Options:
    --workbook, -w   Path to workbook (.xlsm) (default: {workbookDefault})
    --source, -s     Path to source directory (default: .\\source)
    --custom-ui, -c  Path to customUI XML (default: .\\customUI\\customUI.xml)
    --help, -h       Show this help

Example:
    dotnet .\\scripts\\{scriptName}.cs -- --workbook .\\Sample.xlsm --source .\\source --custom-ui .\\customUI\\customUI.xml
""";
    }

    public static string BuildDiffHelp(VbaDevPackConfig config, string baseDir)
    {
        return BuildCommonPathsHelp("DiffVbaModules", "Builds HTML diff report comparing workbook export and repository source.", config, baseDir)
            + "\nAdditional options:\n"
            + "    --direction, -d      Report mode: export|import (default: neutral)\n"
            + "    --no-open-report     Do not open HTML report automatically\n";
    }

    public static string BuildExportHelp(VbaDevPackConfig config, string baseDir)
    {
        return BuildCommonPathsHelp("ExportVbaModules", "Exports workbook VBA modules and custom UI to repository files.", config, baseDir)
            + "\nAdditional options:\n"
            + "    --force              Skip interactive diff confirmation\n"
            + "    --no-open-report     Do not open HTML report automatically\n";
    }

    public static string BuildImportHelp(VbaDevPackConfig config, string baseDir)
    {
        return BuildCommonPathsHelp("ImportVbaModules", "Imports repository VBA modules and custom UI into workbook.", config, baseDir)
            + "\nAdditional options:\n"
            + "    --force              Skip interactive diff confirmation\n"
            + "    --no-open-report     Do not open HTML report automatically\n";
    }

    private static ScriptPaths ParseScriptPaths(Dictionary<string, string?> options, VbaDevPackConfig config, string baseDir)
    {
        var showHelp = options.ContainsKey("help") || options.ContainsKey("h");
        if (showHelp)
        {
            return ScriptPaths.HelpRequested();
        }

        if (options.TryGetValue("error", out var parseError))
        {
            return ScriptPaths.FromError(parseError ?? "Invalid arguments.");
        }

        if (!TryGetOptionValue(options, "workbook", "w", out var workbookValue))
        {
            var workbookFromConfig = ResolveDefaultWorkbookPath(config, baseDir);
            if (workbookFromConfig is null)
            {
                return ScriptPaths.FromError("Missing workbook path. Set 'workbook' in vba-dev-pack.json or pass --workbook.");
            }

            workbookValue = workbookFromConfig;
        }

        if (!TryGetOptionValue(options, "source", "s", out var sourceValue))
        {
            sourceValue = Path.Combine(baseDir, "source");
        }

        if (!TryGetOptionValue(options, "custom-ui", "c", out var customUiValue))
        {
            customUiValue = Path.Combine(baseDir, "customUI", "customUI.xml");
        }

        return ScriptPaths.FromValues(
            Path.GetFullPath(workbookValue!),
            Path.GetFullPath(sourceValue!),
            Path.GetFullPath(customUiValue!));
    }

    private static string? ResolveDefaultWorkbookPath(VbaDevPackConfig config, string baseDir)
    {
        if (string.IsNullOrWhiteSpace(config.Workbook))
        {
            return null;
        }

        var expanded = Environment.ExpandEnvironmentVariables(config.Workbook);
        return Path.IsPathRooted(expanded)
            ? Path.GetFullPath(expanded)
            : Path.GetFullPath(Path.Combine(baseDir, expanded));
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
}

sealed class DiffScriptOptions
{
    public bool ShowHelp { get; init; }
    public bool IsValid { get; init; }
    public string? Error { get; init; }
    public ScriptPaths Paths { get; init; } = new();
    public DiffDirection Direction { get; init; } = DiffDirection.Neutral;
    public bool NoOpenReport { get; init; }

    public static DiffScriptOptions HelpRequested() => new() { ShowHelp = true, IsValid = true };

    public static DiffScriptOptions FromError(string error) => new() { IsValid = false, Error = error };

    public static DiffScriptOptions FromValues(ScriptPaths paths, DiffDirection direction, bool noOpenReport) =>
        new()
        {
            IsValid = true,
            Paths = paths,
            Direction = direction,
            NoOpenReport = noOpenReport
        };
}

sealed class ApplyScriptOptions
{
    public bool ShowHelp { get; init; }
    public bool IsValid { get; init; }
    public string? Error { get; init; }
    public ScriptPaths Paths { get; init; } = new();
    public bool Force { get; init; }
    public bool NoOpenReport { get; init; }

    public static ApplyScriptOptions HelpRequested() => new() { ShowHelp = true, IsValid = true };

    public static ApplyScriptOptions FromError(string error) => new() { IsValid = false, Error = error };

    public static ApplyScriptOptions FromValues(ScriptPaths paths, bool force, bool noOpenReport) =>
        new()
        {
            IsValid = true,
            Paths = paths,
            Force = force,
            NoOpenReport = noOpenReport
        };
}

sealed class ScriptPaths
{
    public bool ShowHelp { get; init; }
    public bool IsValid { get; init; }
    public string? Error { get; init; }
    public string WorkbookPath { get; init; } = string.Empty;
    public string SourceDir { get; init; } = string.Empty;
    public string CustomUiPath { get; init; } = string.Empty;

    public static ScriptPaths HelpRequested() => new() { ShowHelp = true, IsValid = true };

    public static ScriptPaths FromError(string error) => new() { IsValid = false, Error = error };

    public static ScriptPaths FromValues(string workbookPath, string sourceDir, string customUiPath) =>
        new()
        {
            IsValid = true,
            WorkbookPath = workbookPath,
            SourceDir = sourceDir,
            CustomUiPath = customUiPath
        };
}

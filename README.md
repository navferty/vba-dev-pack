# vba-dev-pack

Template repository for Excel XLSM + VBA development with source control.

## Quickstart

```powershell
# 1) Compare workbook vs repository first (safety check)
dotnet .\scripts\DiffVbaModules.cs

# 2) If workbook is the source of truth, export workbook -> repo
dotnet .\scripts\ExportVbaModules.cs

# 3) If repository is the source of truth, import repo -> workbook
dotnet .\scripts\ImportVbaModules.cs
```

Always run diff before import/export to avoid accidentally overwriting newer changes.

## Why use this template

- Keep VBA modules and Ribbon XML as normal text files in Git.
- Round-trip safely between workbook and repository with backup creation.
- Get visual diff report between workbook state and repository state.
- Automate patch version bump + email delivery of versioned XLSM via GitHub Actions.

## What is included

- `Sample.xlsm` - example macro-enabled workbook.
- `source/` - exported VBA modules (`.bas`, `.cls`, `.frm`, `.frx`).
- `customUI/customUI.xml` - Ribbon XML kept in repository.
- `scripts/ExportVbaModules.cs` - export workbook VBA/customUI to repository.
- `scripts/ImportVbaModules.cs` - import repository VBA/customUI into workbook.
- `scripts/DiffVbaModules.cs` - generate and open HTML diff report.
- `scripts/ToUtf.cs` / `scripts/FromUtf.cs` - encoding helpers.
- `.github/workflows/notify-email.yml` - version bump + Gmail notification flow.
- `VERSION` - semantic version used by workflow (patch auto-increment).

## Requirements

- Windows.
- Microsoft Excel desktop installed.
- .NET SDK with support for `net10.0-windows`.
- Workbook closed when running export/import/diff.
- Excel setting enabled:
	- File -> Options -> Trust Center -> Trust Center Settings -> Macro Settings -> Trust access to the VBA project object model.

## Script arguments

`ExportVbaModules.cs`, `ImportVbaModules.cs`, and `DiffVbaModules.cs` use named arguments:

- `--workbook`, `-w` - path to workbook (`.xlsm`), default from `vba-dev-pack.json` field `workbook`
- `--source`, `-s` - path to source directory, default `./source`
- `--custom-ui`, `-c` - path to customUI XML, default `./customUI/customUI.xml`
- `--help`, `-h` - show script help

Example:

```powershell
dotnet .\scripts\ExportVbaModules.cs

# or override workbook explicitly
dotnet .\scripts\ExportVbaModules.cs -- --workbook .\Sample.xlsm
```

Help examples:

```powershell
dotnet .\scripts\ExportVbaModules.cs -- --help
dotnet .\scripts\ImportVbaModules.cs -- --help
dotnet .\scripts\DiffVbaModules.cs -- --help
```

## Daily development flow

1. Pull latest changes.
2. Run diff and decide authoritative side (workbook or repo) before any import/export.
3. Import repo source into workbook only if repo should overwrite workbook.
4. Edit VBA in Excel.
5. Export workbook back to repo only if workbook should overwrite repo.
6. Run diff report again before commit.
7. Commit `source/`, `customUI/`, and workbook changes if needed.

Commands:

```powershell
# Import repo -> workbook
dotnet .\scripts\ImportVbaModules.cs

# Export workbook -> repo
dotnet .\scripts\ExportVbaModules.cs

# Open HTML comparison report
dotnet .\scripts\DiffVbaModules.cs
```

## How export/import behave

- Export removes old module files in `source/` and writes fresh files.
- Export and diff convert VBA text files from the configured codepage to UTF-8 (no BOM) for Git.
- Import converts UTF-8 repo files to the configured codepage for Excel import.
- Import updates document modules (`ThisWorkbook`, sheet classes) in place.
- `customUI/customUI.xml` is exported/imported from workbook package part (`customUI.xml` or `customUI14.xml` when present).
- Export/import create backups under `%TEMP%\vba-dev-pack-backups\...` by default.
- You can override backup root in `vba-dev-pack.json` using `backupRoot`.
- Backup root path is printed in script output (`Backup root: ...`) for recovery.

## Encoding helper scripts

Convert all supported files in `source/`:

```powershell
# native codepage -> UTF-8
dotnet .\scripts\ToUtf.cs

# UTF-8 -> native codepage
dotnet .\scripts\FromUtf.cs
```

Convert selected files:

```powershell
dotnet .\scripts\ToUtf.cs -- --files SampleModule.bas,ThisWorkbook.cls
```

Encoding helper options:

- `--source`, `-s` - source directory (default: `./source`)
- `--files`, `-f` - comma or semicolon-separated file names
- `--help`, `-h` - show script help

## Configure for your project

1. Set your workbook path in `vba-dev-pack.json` field `workbook` (or pass explicit `--workbook` in commands).
2. Keep your VBA modules in `source/` and Ribbon XML in `customUI/customUI.xml`.
3. Set the codepage for your project in `vba-dev-pack.json` (default: `1251`):

```json
{
	"workbook": "Sample.xlsm",
	"codepage": 1251,
	"backupRoot": "C:\\vba-dev-pack-backups"
}
```

The codepage is used when converting VBA text files between the native Excel encoding and UTF-8 during export, import, diff, and encoding helper operations. Set it to the Windows codepage that matches your locale (e.g. `1252` for Western European, `1251` for Cyrillic, `1250` for Central European).

`workbook` is required unless you always pass `--workbook` explicitly.

`backupRoot` is optional. If omitted, backups are created under `%TEMP%\vba-dev-pack-backups`.

4. Set GitHub workflow variables/secrets used by `.github/workflows/notify-email.yml`:

- Repository variable: `GMAIL_USER`
- Repository variable: `NOTIFY_EMAIL`
- Repository secret: `GMAIL_APP_PASSWORD`

Example with GitHub CLI:

```powershell
gh variable set GMAIL_USER --body "your.account@gmail.com"
gh variable set NOTIFY_EMAIL --body "team@company.com"
gh secret set GMAIL_APP_PASSWORD --body "<gmail-app-password>"
```

5. Update CI workbook naming in `.github/workflows/notify-email.yml` (copy/attachment step currently uses `Sample.xlsm`).
6. Ensure your default branch matches workflow trigger (`master`) or update trigger branch.

## GitHub workflow behavior

On each push to `master`, workflow does:

1. Read `VERSION` and bump patch (`x.y.z -> x.y.(z+1)`).
2. Commit and push updated `VERSION`.
3. Create `Sample_v<version>.xlsm` attachment.
4. Send email through Gmail SMTP with workbook attachment and commit metadata.

If your workbook is not `Sample.xlsm`, update the attachment step in workflow.

## Troubleshooting

- `Workbook is currently open or locked`: close workbook in Excel.
- `Excel denied access to VBProject`: enable Trust Center VBA project access setting.
- `Workbook not found` or `Source directory not found`: verify paths or pass explicit arguments.
- 
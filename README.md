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

It is recommended to always run diff before import/export.

## Why use this template

- Keep VBA modules and Ribbon XML as plain text in Git.
- Safe workbook <-> repository round-trip with backups.
- Visual HTML diff report before apply.
- Optional version bump + email delivery via GitHub Actions.

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

- `--workbook`, `-w`: workbook path (`.xlsm`), default from `vba-dev-pack.json` (`workbook`)
- `--source`, `-s`: source directory, default `./source`
- `--custom-ui`, `-c`: customUI XML path, default `./customUI/customUI.xml`
- `--help`, `-h`: show help

Additional options:

- `DiffVbaModules.cs`: `--direction`, `-d` (`export` | `import`), `--no-open-report`
- `ExportVbaModules.cs`: `--force`, `--no-open-report`
- `ImportVbaModules.cs`: `--force`, `--no-open-report`

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

Direction examples:

```powershell
# Neutral comparison (default)
dotnet .\scripts\DiffVbaModules.cs

# Export-oriented preview (workbook -> repo)
dotnet .\scripts\DiffVbaModules.cs -- --direction export

# Import-oriented preview (repo -> workbook)
dotnet .\scripts\DiffVbaModules.cs -- --direction import
```

## Daily development flow

1. Pull latest changes.
2. Run directional diff and choose source of truth.
3. Run apply command (import or export).
4. Review preflight report.
5. Approve apply (`Enter` = yes by default).
6. Run diff again before commit.

Commands:

```powershell
# Import repo -> workbook
dotnet .\scripts\ImportVbaModules.cs

# Export workbook -> repo
dotnet .\scripts\ExportVbaModules.cs

# Open HTML comparison report
dotnet .\scripts\DiffVbaModules.cs
```

## Final sync UX

1. Preview with direction:
	- `dotnet .\scripts\DiffVbaModules.cs -- -d export`
	- `dotnet .\scripts\DiffVbaModules.cs -- -d import`
2. Run apply command:
	- `dotnet .\scripts\ExportVbaModules.cs` or `dotnet .\scripts\ImportVbaModules.cs`
3. Check preflight HTML report.
4. Confirm prompt: `Approve diff? [Y]es / [n]o (default: yes)`.

Apply flags:

- `--force`: skip interactive confirmation
- `--no-open-report`: do not auto-open report in browser

## How export/import behave

- Export rewrites VBA files in `source/`.
- Export and diff convert VBA text to UTF-8 (no BOM).
- Import converts UTF-8 back to configured native codepage.
- Import updates document modules (`ThisWorkbook`, sheet classes) in place.
- `customUI/customUI.xml` is synced with workbook package part.
- Apply commands create backups (default: `%TEMP%\vba-dev-pack-backups`).
- Backup root is printed in output.

## Exit codes

Common:

- `0`: success
- `1`: invalid arguments/paths or pre-check failure
- `2`: Excel COM/VBProject error
- `3`: other runtime error

Apply only (`ExportVbaModules.cs`, `ImportVbaModules.cs`):

- `4`: cancelled by user
- `5`: non-interactive mode without `--force`

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

`codepage` controls VBA text conversion between native encoding and UTF-8.

`workbook` is required unless you always pass `--workbook`.

`backupRoot` is optional. Default is `%TEMP%\vba-dev-pack-backups`.

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

## Security and limitations

- Windows + Excel Desktop only.
- Requires VBA project access in Trust Center.
- Uses Excel COM automation.
- No atomic rollback.
- Recovery is manual from backups.

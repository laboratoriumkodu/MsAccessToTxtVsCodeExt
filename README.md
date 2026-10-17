# Access Source Exporter

Use **Access: Export Source** from the Command Palette to select an Access
database and an output folder. The extension invokes the existing PowerShell
exporter; the standalone script remains at
`../script/Export-AccessSource.ps1`.

## Current status

The command flow is implemented. The build copies the exporter into the
extension so the installed extension does not depend on the repository being
present.

## Requirements

The completed extension is intended for Windows with PowerShell and Microsoft
Access installed. The exporter uses Access COM automation.

## Development

Open this folder in VS Code and run the `Run Extension` launch configuration to
start an Extension Development Host. Run `npm run compile` to build and check
the extension.

# Access Source Exporter

Use **Access: Export Source** from the Command Palette to select an Access
database and an output folder. The extension invokes the existing PowerShell
exporter; the standalone script remains at
`../script/Export-AccessSource.ps1`.

## Current status

The command flow is implemented. The build copies the exporter into the
extension so the installed extension does not depend on the repository being
present. Export progress can be cancelled, and PowerShell output and errors
appear in the **Access Source Export** Output channel.

## Requirements

The extension requires Windows, Windows PowerShell 5.1 (`powershell.exe`), and
Microsoft Access installed and activated. The exporter uses Access COM
automation and its process architecture must be compatible with the Access
installation.

## Development

Open this folder in VS Code and run the `Run Extension` launch configuration to
start an Extension Development Host. Run `npm run compile` to build and check
the extension.
